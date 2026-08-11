import os
import urllib.parse
from datetime import datetime
import pandas as pd
import requests
from bs4 import BeautifulSoup

# 검색 범위를 월/분기별로 세분화하여 1000개 수집 보장
DATE_RANGES = [
    ('2022-01-01', '2022-06-30'),
    ('2022-07-01', '2022-12-31'),
    ('2023-01-01', '2023-06-30'),
    ('2023-07-01', '2023-12-31'),
    ('2024-01-01', '2024-06-30'),
    ('2024-07-01', '2024-12-31'),
    ('2025-01-01', '2025-06-30'),
    ('2025-07-01', '2025-12-31'),
    ('2026-01-01', '2026-08-11'),
]

KEYWORDS_KR = [
    '삼성전자 공정기술',
    'SK하이닉스 양산기술',
    '현대차 개발품질',
    '현대모비스 품질',
    'SK하이닉스 M15',
    'EUV 노광',
    'Dry Etch',
    'ALD 증착',
    'CMP 슬러리',
    'Defect 제어',
    '수율 램프업',
    'IATF16949',
    'AEC-Q100',
    'ISO26262',
    'FMEA',
    '전장 신뢰성',
    'HBM 수율',
    'TSP 공정',
    '반도체 수율',
    '자동차 품질',
]

FILE_PATH = 'accumulated_news.xlsx'


def calculate_importance_score(title):
    score = 10
    tier1 = [
        '수율',
        'Yield',
        'Defect',
        'FMEA',
        'IATF',
        'AEC-Q',
        'EUV',
        'ALD',
        'CMP',
        'ISO26262',
        'HBM',
        '신뢰성',
    ]
    for w in tier1:
        if w.lower() in title.lower():
            score += 15

    tier2 = [
        '양산',
        '품질',
        '식각',
        '증착',
        '파티클',
        '계측',
        'Chamber',
        'Overlay',
        '램프업',
        'M15',
        'APQP',
        'PPAP',
    ]
    for w in tier2:
        if w.lower() in title.lower():
            score += 10

    companies = ['삼성전자', 'SK하이닉스', '현대차', '현대모비스']
    for c in companies:
        if c in title:
            score += 5

    return min(score, 100)


def classify_category(title):
    if any(
        k in title
        for k in ['노광', 'EUV', '식각', 'Etch', '증착', 'ALD', 'CMP', '슬러리']
    ):
        return '반도체 8대공정'
    elif any(
        k in title
        for k in [
            '수율',
            'Yield',
            'Defect',
            '파티클',
            '계측',
            'Chamber',
            '램프업',
            'Overlay',
            'HBM',
            '양산',
        ]
    ):
        return '수율/양산/PEDTC'
    elif any(
        k in title
        for k in [
            'IATF',
            '16949',
            'FMEA',
            'AEC',
            'ISO26262',
            '신뢰성',
            'SQ',
            '전장',
            'ECU',
            '품질',
        ]
    ):
        return '자동차/전장 품질'
    else:
        return '기업일반/동향'


def main():
    all_news = []
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }

    print('2022년~현재 1,000개 수집 작업 시작...')

    for kw in KEYWORDS_KR:
        for start_date, end_date in DATE_RANGES:
            query = f'{kw} after:{start_date} before:{end_date}'
            encoded_query = urllib.parse.quote(query)
            url = f'https://news.google.com/rss/search?q={encoded_query}&hl=ko&gl=KR&ceid=KR:ko'

            try:
                res = requests.get(url, headers=headers, timeout=10)
                soup = BeautifulSoup(res.text, 'xml')

                for item in soup.find_all('item')[:15]:
                    title = item.title.text if item.title else ''
                    link = item.link.text if item.link else ''
                    pub_date = item.pubDate.text if item.pubDate else ''
                    source = (
                        item.source.text if item.source else 'Google News'
                    )

                    spam_keywords = [
                        '특징주',
                        '상한가',
                        '급등주',
                        '리딩방',
                        '목표가',
                        '주가',
                        '증시',
                        '매수',
                    ]
                    if any(spam in title for spam in spam_keywords):
                        continue

                    score = calculate_importance_score(title)
                    all_news.append({
                        '중요도점수': score,
                        '면접활용도': '★ 핵심' if score >= 40 else '일반',
                        '분류태그': classify_category(title),
                        '검색키워드': kw,
                        '언론사': source,
                        '뉴스제목': title,
                        '링크': link,
                        '발행일시': pub_date,
                        '수집일시': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    })
            except Exception as e:
                print(f'[{kw}] 수집 예외 발생: {e}')

    df = pd.DataFrame(all_news)
    df.drop_duplicates(subset=['뉴스제목'], keep='first', inplace=True)
    df.sort_values(by='중요도점수', ascending=False, inplace=True)

    # 상위 1,000개 추출
    top1000_df = df.head(1000)

    # 기존 실시간 시트가 있다면 유지
    realtime_df = pd.DataFrame()
    if os.path.exists(FILE_PATH):
        try:
            realtime_df = pd.read_excel(FILE_PATH, sheet_name='실시간_누적뉴스')
        except Exception:
            pass

    # Sheet1: 역대_핵심뉴스_1000 / Sheet2: 실시간_누적뉴스
    with pd.ExcelWriter(FILE_PATH, engine='openpyxl') as writer:
        top1000_df.to_excel(writer, sheet_name='역대_핵심뉴스_1000', index=False)
        realtime_df.to_excel(writer, sheet_name='실시간_누적뉴스', index=False)

    print(f'완료! Sheet1에 상위 {len(top1000_df)}개 기사가 저장되었습니다.')


if __name__ == '__main__':
    main()
