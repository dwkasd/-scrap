import os
import urllib.parse
from datetime import datetime
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ==========================================
# 1. 최소 단위(원자) 고유 키워드 설정
# (복합어를 지우고, 독자 검색이 가능한 전문 단어로만 구성)
# ==========================================

KEYWORDS_KR = [
    # [기업 & 사업장 / 제품]
    '삼성전자 공정기술',
    'SK하이닉스 양산기술',
    '현대차 개발품질',
    '현대모비스 품질',
    'SK하이닉스 청주',
    'SK하이닉스 M15',
    'HBM 수율',
    'TSP 공정',
    # [반도체 8대공정 / PEDTC 전문 단어]
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
    # [자동차 / 전장 품질 표준 전문 단어]
    'IATF16949',
    'AEC-Q100',
    'ISO26262',
    'FMEA',
    'APQP',
    'PPAP',
    'SQ인증',
    '전장 신뢰성',
    'DV PV 검증',
    '필드클레임'
]

KEYWORDS_EN = [
    'EUV Lithography',
    'Dry Etching',
    'ALD Deposition',
    'CMP Slurry',
    'Semiconductor Defect',
    'Yield Ramp Up',
    'IATF 16949',
    'AEC-Q100',
    'ISO 26262',
    'FMEA Quality'
]

FILE_PATH = 'accumulated_news.xlsx'


def classify_category(title):
    """최소 단위 단어들의 부분 일치(Substring Match)를 통해 직무별 자동 분류"""
    # 반도체 8대 공정 (단일 키워드 검사)
    if any(k in title for k in ['노광', 'EUV', '식각', 'Etch', '증착', 'ALD', 'CVD', 'CMP', '슬러리', '박막', 'PR', 'High-k']):
        return '반도체 8대공정'
    
    # PEDTC / 수율 / MI / 양산
    elif any(k in title for k in ['수율', 'Yield', 'Defect', '디펙트', '파티클', 'Particle', '계측', 'MI', '챔버', 'Chamber', '플라즈마', 'Plasma', '램프업', 'Overlay', '오버레이', 'TSP', 'HBM', '양산']):
        return '수율/양산/PEDTC'
    
    # 자동차 및 전장 품질
    elif any(k in title for k in ['IATF', '16949', 'FMEA', 'APQP', 'PPAP', 'AEC', 'Q100', 'ISO26262', '신뢰성', 'SQ인증', 'SQ', '전장', 'ECU', 'DV', 'PV', '클레임']):
        return '자동차/전장 품질'
    
    # 라인 / 투자 / Fab
    elif any(k in title for k in ['증설', '투자', '라인', 'Fab', '팹', 'M15', '용인', '평택', '착공']):
        return '라인/투자/Fab'
    
    else:
        return '기업일반/동향'


def is_high_priority(title):
    """핵심 기술 단어가 1개라도 포함되면 ★ 핵심 지정"""
    atomic_high_value = [
        '수율', 'Defect', 'FMEA', 'IATF', 'AEC', 'EUV', '식각', 'ALD', 'CMP', 
        '신뢰성', '양산', '품질', 'ISO26262', '파티클', '계측', 'Chamber', 'HBM'
    ]
    return '★ 핵심' if any(k in title for k in atomic_high_value) else '일반'


def fetch_google_news(keyword, lang='kr'):
    """구글 뉴스 RSS를 통해 원자 키워드 기반 뉴스를 수집합니다."""
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

        for item in items[:5]:  # 원자 키워드가 대폭 늘어났으므로 키워드당 5개씩만 수집 (중복 방지 및 빠른 실행)
            title = item.title.text if item.title else ''
            link = item.link.text if item.link else ''
            pub_date = item.pubDate.text if item.pubDate else ''
            source = item.source.text if item.source else 'Google News'

            # 주식/금융 관련 스팸 필터링
            spam_keywords = ['특징주', '상한가', '급등주', '리딩방', '목표가', '주가', '증시', '매수', '종목', '코스피', '코스닥']
            if any(spam in title for spam in spam_keywords):
                continue

            category = classify_category(title)
            priority = is_high_priority(title)

            news_items.append({
                '수집일시': datetime.now().strftime('%Y-%m-%d %H:%M'),
                '면접활용도': priority,
                '분류태그': category,
                '구분': '국내' if lang == 'kr' else '해외',
                '검색키워드': keyword,
                '언론사': source,
                '뉴스제목': title,
                '링크': link,
                '발행일시': pub_date
            })
    except Exception as e:
        print(f"[{keyword}] 수집 중 에러 발생: {e}")

    return news_items


def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 원자 단위 키워드 기반 뉴스 수집 시작")

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

    # 2. 기존 엑셀 파일 누적 처리
    if os.path.exists(FILE_PATH):
        try:
            existing_df = pd.read_excel(FILE_PATH)
            
            for col in ['면접활용도', '분류태그']:
                if col not in existing_df.columns:
                    existing_df[col] = '일반'

            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        except Exception as e:
            print(f"기존 파일 읽기 실패 (신규 파일로 새로 시작): {e}")
            combined_df = new_df
    else:
        combined_df = new_df

    # 3. 뉴스제목 기준 중복 제거
    initial_count = len(combined_df)
    combined_df.drop_duplicates(subset=['뉴스제목'], keep='first', inplace=True)
    final_count = len(combined_df)

    # 4. 동일 파일 저장
    combined_df.to_excel(FILE_PATH, index=False)

    print(f"업데이트 완료 | 신규 수집: {len(new_df)}건 | 중복 제외: {initial_count - final_count}건 | 총 누적 DB: {final_count}건")


if __name__ == '__main__':
    main()
