import re
import json
import os
import pandas as pd
import scrapetube
from youtube_transcript_api import YouTubeTranscriptApi
import google.generativeai as genai

# 1. 환경 변수에서 Gemini API Key 로드
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    print("[-] ERROR: GEMINI_API_KEY가 설정되지 않았습니다.")
    exit(1)

genai.configure(api_key=GEMINI_API_KEY)

# 설정값
PROCESSED_FILE = "processed_ids.json"
EXCEL_FILE = "resume_youtube_knowledge.xlsx"
JSONL_FILE = "resume_youtube_knowledge.jsonl"
LIMIT_PER_KEYWORD = 10  # 키워드당 수집할 상위 영상 개수 (조회수 순)

def load_processed_ids():
    if os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()

def save_processed_ids(processed_set):
    with open(PROCESSED_FILE, "w", encoding="utf-8") as f:
        json.dump(list(processed_set), f, ensure_ascii=False, indent=2)

def search_top_videos_by_keyword(keyword, processed_ids, limit=10):
    """키워드로 유튜브 검색 후 조회수 높은 순서대로 영상 추출"""
    print(f"\n[+] 키워드 검색 시작: '{keyword}' (조회수 높은 순)")
    target_videos = []
    
    try:
        # sort_by="views" 옵션으로 조회수 높은 순서 정렬
        results = scrapetube.get_search(query=keyword, sort_by="views", limit=limit * 2)
        
        for video in results:
            video_id = video['videoId']
            title = video.get('title', {}).get('runs', [{}])[0].get('text', '')
            
            # 이미 처리된 영상 스킵
            if video_id in processed_ids:
                continue
                
            target_videos.append({
                "id": video_id,
                "title": title,
                "url": f"https://www.youtube.com/watch?v={video_id}",
                "keyword": keyword
            })
            print(f"  ├─ [선별] {title} ({video_id})")
            
            # 목표 개수 달성 시 중단
            if len(target_videos) >= limit:
                break
                
    except Exception as e:
        print(f"[-] 검색 중 오류 발생 ('{keyword}'): {e}")
        
    return target_videos

def get_youtube_transcript(video_id):
    try:
        transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=['ko'])
        return " ".join([item['text'] for item in transcript])
    except Exception as e:
        print(f"  └─ [-] 자막 수집 실패: {e}")
        return None

def analyze_with_gemini(transcript_text, video_url, video_title):
    model = genai.GenerativeModel('gemini-1.5-flash')
    prompt = f"""
    아래는 이력서/자소서 작성 관련 유튜브 영상의 자막입니다.
    이 내용을 정밀하게 분석하여 지식 데이터베이스용 JSON 규격으로 정제해 주세요.

    [영상 제목]: {video_title}
    [자막 내용]
    {transcript_text[:15000]}

    [응답 JSON 규격]
    {{
        "channel_name": "채널명 (추정 가능시 작성, 모르면 '미상')",
        "video_title": "영상 주제 요약 제목",
        "section": "학력 / 경력기술서 / 지원동기 / 프로젝트 / 자기소개서 중 해당 영역",
        "target": "신입 / 경력 / 이직 / 공통 중 선택",
        "keywords": "핵심 키워드 3~5개 (쉼표 구분)",
        "summary": "영상 핵심 내용 1~2문장 요약",
        "do_rules": "이력서 작성 시 반드시 해야 할 권장 사항",
        "dont_rules": "이력서 작성 시 절대로 하지 말아야 할 주의/금지 사항",
        "example_sentence": "영상에서 추천하는 실제 문장 예시"
    }}
    """
    try:
        response = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"}
        )
        result = json.loads(response.text)
        result["video_url"] = video_url
        return result
    except Exception as e:
        print(f"  └─ [-] Gemini API 가공 실패: {e}")
        return None

def append_to_database(new_data_list):
    if not new_data_list:
        return

    # JSONL 저장
    with open(JSONL_FILE, "a", encoding="utf-8") as f:
        for item in new_data_list:
            f.write(json.dumps(item, ensure_ascii=False) + "\n")

    # 엑셀 저장
    new_df = pd.DataFrame(new_data_list)
    column_mapping = {
        "channel_name": "채널명", "video_title": "영상 제목", "section": "이력서 영역",
        "target": "지원 대상", "keywords": "핵심 키워드", "summary": "한 줄 요약",
        "do_rules": "Do (권장 사항)", "dont_rules": "Don't (금지 사항)",
        "example_sentence": "적용 예시 문장", "video_url": "영상 링크"
    }
    new_df = new_df.reindex(columns=list(column_mapping.keys())).rename(columns=column_mapping)

    if os.path.exists(EXCEL_FILE):
        existing_df = pd.read_excel(EXCEL_FILE)
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = new_df

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        combined_df.to_excel(writer, index=False, sheet_name='유튜브 지식베이스')
        worksheet = writer.sheets['유튜브 지식베이스']
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            worksheet.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 60)

    print(f"\n[+] 데이터베이스 저장 완료 (총 {len(combined_df)}건 누적됨)")

if __name__ == "__main__":
    if not os.path.exists("keywords.txt"):
        print("[-] keywords.txt 파일이 없습니다.")
        exit(1)

    processed_ids = load_processed_ids()
    
    with open("keywords.txt", "r", encoding="utf-8") as f:
        keywords = [line.strip() for line in f if line.strip()]

    all_target_videos = []
    for kw in keywords:
        videos = search_top_videos_by_keyword(kw, processed_ids, limit=LIMIT_PER_KEYWORD)
        all_target_videos.extend(videos)

    print(f"\n[=] 총 {len(all_target_videos)}개의 인기 영상을 수집 대상으로 지정했습니다.")

    new_results = []
    for idx, item in enumerate(all_target_videos, 1):
        print(f"\n[{idx}/{len(all_target_videos)}] 처리 중: {item['title']}")
        
        transcript = get_youtube_transcript(item['id'])
        if not transcript:
            processed_ids.add(item['id'])
            continue
            
        data = analyze_with_gemini(transcript, item['url'], item['title'])
        if data:
            new_results.append(data)
            processed_ids.add(item['id'])

    append_to_database(new_results)
    save_processed_ids(processed_ids)
