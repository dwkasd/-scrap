import os
import urllib.parse
from datetime import datetime
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ==========================================
# 수집 타겟 키워드 및 환경 설정
# ==========================================
KEYWORDS_KR = ['삼성전자', 'SK하이닉스', 'DB하이텍', '현대자동차', '현대모비스', '반도체']
KEYWORDS_EN = ['Samsung Electronics', 'SK Hynix', 'TSMC semiconductor', 'Hyundai EV', 'HBM memory']

# 클라우드 저장소 내부 단일 누적 파일 경로
FILE_PATH = 'accumulated_news.xlsx'


def fetch_google_news(keyword, lang='kr'):
    """구글 뉴스 RSS를 통해 해당 키워드의 최신 뉴스를 수집합니다."""
    encoded_keyword = urllib.parse.quote(keyword)

    if lang == 'kr':
        url = f'https://news.google.com/rss/search?q={encoded_keyword}&hl=ko&gl=KR&ceid=KR:ko'
    else:
        url = f'https://news.google.com/rss/search?q={encoded_keyword}&hl=en-US&gl=US&ceid=US:en'

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }

    news_items = []
    try:
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'xml')
        items = soup.find_all('item')

        for item in items[:5]:  # 키워드당 상위 5개 최신 기사 수집
            title = item.title.text if item.title else ''
            link = item.link.text if item.link else ''
            pub_date = item.pubDate.text if item.pubDate else ''
            source = item.source.text if item.source else 'Google News'

            # 노이즈/스팸 기사 1차 정리
            if any(spam in title for spam in ['특징주', '상한가', '급등주', '리딩방']):
                continue

            news_items.append({
                '수집일시': datetime.now().strftime('%Y-%m-%d %H:%M'),
                '구분': '국내' if lang == 'kr' else '해외',
                '키워드': keyword,
                '언론사': source,
                '뉴스제목': title,
                '링크': link,
                '발행일시': pub_date
            })
    except Exception as e:
        print(f"[{keyword}] 수집 중 에러 발생: {e}")

    return news_items


def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 뉴스 수집 및 업데이트 시작")

    all_scraped_news = []

    # 1. 국내외 뉴스 수집
    for kw in KEYWORDS_KR:
        all_scraped_news.extend(fetch_google_news(kw, lang='kr'))

    for kw in KEYWORDS_EN:
        all_scraped_news.extend(fetch_google_news(kw, lang='en'))

    new_df = pd.DataFrame(all_scraped_news)

    if new_df.empty:
        print("신규 수집된 뉴스가 없습니다.")
        return

    # 2. 기존 단일 엑셀 파일 로드 및 누적(Append) 처리
    if os.path.exists(FILE_PATH):
        try:
            existing_df = pd.read_excel(FILE_PATH)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        except Exception as e:
            print(f"기존 파일 읽기 실패 (신규 파일로 새로 시작): {e}")
            combined_df = new_df
    else:
        combined_df = new_df

    # 3. 뉴스제목 기준 중복 제거 (기존 수집 데이터 우선 보존)
    initial_count = len(combined_df)
    combined_df.drop_duplicates(subset=['뉴스제목'], keep='first', inplace=True)
    final_count = len(combined_df)

    # 4. 동일한 파일명으로 갱신 저장
    combined_df.to_excel(FILE_PATH, index=False)

    print(f"업데이트 완료 | 신규 획득: {len(new_df)}건 | 중복 제외: {initial_count - final_count}건 | 총 누적 DB: {final_count}건")


if __name__ == '__main__':
    main()
