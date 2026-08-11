import os
import urllib.parse
from datetime import datetime
import pandas as pd
import requests
from bs4 import BeautifulSoup

KEYWORDS_KR = [
    '삼성전자 공정기술',
    'SK하이닉스 양산기술',
    '현대차 개발품질',
    '현대모비스 품질',
    'SK하이닉스 청주',
    'SK하이닉스 M15',
    'HBM 수율',
    'TSP 공정',
    'EUV',
    'High-NA',
    'Dry Etch',
    'ALD 증착',
    'CMP 슬러리',
    'Defect 제어',
    '챔버 파티클',
    '플라즈마 제어',
    'Overlay 제어',
    '수율 램프업',
    'IATF16949',
    'AEC-Q100',
    'ISO26262',
    'FMEA',
    'APQP',
    'PPAP',
    'SQ인증',
    '전장 신뢰성',
]

FILE_PATH = 'accumulated_news.xlsx'


def classify_category(title):
    if any(
        k in title
        for k in ['노광', 'EUV', '식각', 'Etch', '증착', 'ALD', 'CVD', 'CMP', '슬러리']
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
            'MI',
            '챔버',
            '플라즈마',
            '램프업',
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
            'SQ인증',
            '전장',
            'ECU',
            '품질',
        ]
    ):
        return '자동차/전장 품질'
    else:
        return '기업일반/동향'


def fetch_google_news(keyword):
    encoded_keyword = urllib.parse.quote(keyword)
    url = f'https://news.google.com/rss/search?q={encoded_keyword}&hl=ko&gl=KR&ceid=KR:ko'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }

    news_items = []
    try:
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'xml')

        for item in soup.find_all('item')[:5]:
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

            news_items.append({
                '수집일시': datetime.now().strftime('%Y-%m-%d %H:%M'),
                '면접활용도': (
                    '★ 핵심'
                    if any(
                        k in title
                        for k in ['수율', 'Defect', 'FMEA', 'IATF', 'EUV', 'HBM']
                    )
                    else '일반'
                ),
                '분류태그': classify_category(title),
                '검색키워드': keyword,
                '언론사': source,
                '뉴스제목': title,
                '링크': link,
                '발행일시': pub_date,
            })
    except Exception as e:
        print(f'[{keyword}] 수집 중 에러: {e}')

    return news_items


def main():
    print(
        f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 2시간 자동 누적 수집 시작"
    )

    new_items = []
    for kw in KEYWORDS_KR:
        new_items.extend(fetch_google_news(kw))

    new_df = pd.DataFrame(new_items)
    if new_df.empty:
        print('신규 수집 뉴스가 없습니다.')
        return

    historical_df = pd.DataFrame()
    realtime_df = pd.DataFrame()

    # 기존 시트데이터 읽기 (Sheet1 무조건 보존)
    if os.path.exists(FILE_PATH):
        try:
            historical_df = pd.read_excel(
                FILE_PATH, sheet_name='역대_핵심뉴스_1000'
            )
        except Exception:
            pass

        try:
            realtime_df = pd.read_excel(FILE_PATH, sheet_name='실시간_누적뉴스')
        except Exception:
            pass

    # Sheet2 데이터 병합 및 중복 제거
    combined_realtime_df = pd.concat([realtime_df, new_df], ignore_index=True)
    combined_realtime_df.drop_duplicates(
        subset=['뉴스제목'], keep='first', inplace=True
    )

    # 두 시트 모두 유지하면서 파일 저장
    with pd.ExcelWriter(FILE_PATH, engine='openpyxl') as writer:
        historical_df.to_excel(
            writer, sheet_name='역대_핵심뉴스_1000', index=False
        )
        combined_realtime_df.to_excel(
            writer, sheet_name='실시간_누적뉴스', index=False
        )

    print(
        f'업데이트 완료 | Sheet1 보존 | Sheet2(실시간 누적): {len(combined_realtime_df)}건'
    )


if __name__ == '__main__':
    main()
