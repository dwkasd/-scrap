import os
import urllib.parse
from datetime import datetime
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ==========================================
# 1. 2022년~현재 수집용 기간 설정 및 키워드
# ==========================================
DATE_RANGES = [
    ('2022-01-01', '2022-12-31'),
    ('2023-01-01', '2023-12-31'),
    ('2024-01-01', '2024-12-31'),
    ('2025-01-01', '2025-12-31'),
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
]

FILE_PATH = 'accumulated_news.xlsx'


def calculate_importance_score(title):
    """제목 내 주요 기술 및 품질 단어 포함 여부에 따라 중요도 점수(0~100점)를 계산합니다."""
    score = 10  # 기본 점수

    # 1. 초고가치 직무 기술 단어 (+15점씩)
    tier1_words = [
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
    for w in tier1_words:
        if w.lower() in title.lower():
            score += 15

    # 2. 주요 공정/검사/양산 단어 (+10점씩)
    tier2_words = [
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
    for w in tier2_words:
        if w.lower() in title.lower():
            score += 10

    # 3. 타겟 기업명 포함 (+5점씩)
    companies = ['삼성전자', 'SK하이닉스', '현대차', '현대모비스']
    for c in companies:
        if c in title:
            score += 5

    return min(score, 100)


def classify_category(title):
    """뉴스 분류 태그 생성"""
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


def fetch_historical_news():
    all_news = []
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }

    print('2022년~현재 데이터 수집을 시작합니다. (시간이 다소 소요됩니다)...')

    for kw in KEYWORDS_KR:
        for start_date, end_date in DATE_RANGES:
            query = f'{kw} after:{start_date} before:{end_date}'
            encoded_query = urllib.parse.quote(query)
            url = f'https://news.google.com/rss/search?q={encoded_query}&hl=ko&gl=KR&ceid=KR:ko'

            try:
                res = requests.get(url, headers=headers, timeout=10)
                soup = BeautifulSoup(res.text, 'xml')

                for item in soup.find_all('item')[:10]:
                    title = item.title.text if item.title else ''
                    link = item.link.text if item.link else ''
                    pub_date = item.pubDate.text if item.pubDate else ''
                    source = (
                        item.source.text if item.source else 'Google News'
                    )

                    # 주식/스팸 필터링
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
                    category = classify_category(title)

                    all_news.append({
                        '중요도점수': score,
                        '면접활용도': '★ 핵심' if score >= 40 else '일반',
                        '분류태그': category,
                        '검색키워드': kw,
                        '언론사': source,
                        '뉴스제목': title,
                        '링크': link,
                        '발행일시': pub_date,
                        '수집일시': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    })
            except Exception as e:
                print(f'[{kw}] ({start_date}) 수집 중 에러: {e}')

    new_df = pd.DataFrame(all_news)

    # 기존 파일 존재 시 병합
    if os.path.exists(FILE_PATH):
        try:
            existing_df = pd.read_excel(FILE_PATH)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        except Exception:
            combined_df = new_df
    else:
        combined_df = new_df

    # 1. 중복 제거 (제목 기준)
    combined_df.drop_duplicates(subset=['뉴스제목'], keep='first', inplace=True)

    # 2. 중요도 점수 높은 순으로 내림차순 정렬
    combined_df.sort_values(by='중요도점수', ascending=False, inplace=True)

    # 3. 중요도 상위 1,000개만 슬라이싱하여 최종 선별
    final_df = combined_df.head(1000)

    # 4. 저장
    final_df.to_excel(FILE_PATH, index=False)
    print(
        f'완료! 중요도 점수가 가장 높은 상위 {len(final_df)}개 기사가 {FILE_PATH}에 누적 저장되었습니다.'
    )


if __name__ == '__main__':
    fetch_historical_news()
