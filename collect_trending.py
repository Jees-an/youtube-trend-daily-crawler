import os
import re
import json
import base64
import pandas as pd
from datetime import datetime, timezone
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

REGION = 'KR'
MAX_RESULTS = 50
VIDEO_PARTS = 'snippet,statistics,contentDetails,topicDetails,status'
CHANNEL_PARTS = 'snippet,statistics,brandingSettings,topicDetails'
DESCRIPTION_MAX_LENGTH = 500

CATEGORY_KO = {
    "1": "영화/애니메이션", "2": "자동차/차량", "10": "음악",
    "15": "애완동물/동물", "17": "스포츠", "18": "단편 영화",
    "19": "여행/이벤트", "20": "게임", "21": "비디오 블로깅",
    "22": "인물/블로그", "23": "코미디", "24": "엔터테인먼트",
    "25": "뉴스/정치", "26": "사용법/스타일", "27": "교육",
    "28": "과학/기술", "29": "비영리/사회운동"
}


def clean_text(text):
    text = re.sub(r'[\n\r]', ' ', text)
    text = re.sub(r'<br>', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def parse_duration(iso_duration):
    """ISO 8601 duration (PT1H2M3S) -> total seconds."""
    match = re.match(r'PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?', iso_duration or '')
    if not match:
        return 0
    h, m, s = (int(x) if x else 0 for x in match.groups())
    return h * 3600 + m * 60 + s


def fetch_trending_videos(youtube):
    request = youtube.videos().list(
        part=VIDEO_PARTS,
        chart='mostPopular',
        regionCode=REGION,
        maxResults=MAX_RESULTS
    )
    return request.execute()


def extract_video_data(items, collect_time):
    rows = []
    channel_ids = set()

    for rank, item in enumerate(items, start=1):
        snippet = item.get('snippet', {})
        stats = item.get('statistics', {})
        content = item.get('contentDetails', {})
        topic = item.get('topicDetails', {})
        status = item.get('status', {})

        channel_id = snippet.get('channelId', '')
        if channel_id:
            channel_ids.add(channel_id)

        tags = snippet.get('tags', [])
        tags_str = '|'.join(tags) if isinstance(tags, list) else ''
        topic_cats = '|'.join(topic.get('topicCategories', []))
        desc = clean_text(snippet.get('description', ''))[:DESCRIPTION_MAX_LENGTH]
        category_id = snippet.get('categoryId', '')
        duration_raw = content.get('duration', '')

        rows.append({
            'collected_at': collect_time,
            'trending_rank': rank,
            'video_id': item.get('id', ''),
            'title': snippet.get('title', ''),
            'channel_id': channel_id,
            'channel_title': snippet.get('channelTitle', ''),
            'published_at': snippet.get('publishedAt', ''),
            'description': desc,
            'tags': tags_str,
            'category_id': category_id,
            'category_name_ko': CATEGORY_KO.get(category_id, '기타'),
            'default_language': snippet.get('defaultLanguage', ''),
            'default_audio_language': snippet.get('defaultAudioLanguage', ''),
            'live_broadcast_content': snippet.get('liveBroadcastContent', ''),
            'view_count': stats.get('viewCount', ''),
            'like_count': stats.get('likeCount', ''),
            'comment_count': stats.get('commentCount', ''),
            'favorite_count': stats.get('favoriteCount', ''),
            'duration': duration_raw,
            'duration_seconds': parse_duration(duration_raw),
            'dimension': content.get('dimension', ''),
            'definition': content.get('definition', ''),
            'caption': content.get('caption', ''),
            'licensed_content': content.get('licensedContent', ''),
            'projection': content.get('projection', ''),
            'topic_categories': topic_cats,
            'privacy_status': status.get('privacyStatus', ''),
            'license': status.get('license', ''),
            'made_for_kids': status.get('madeForKids', ''),
        })

    return rows, channel_ids


def fetch_channel_data(youtube, channel_ids):
    request = youtube.channels().list(
        part=CHANNEL_PARTS,
        id=','.join(channel_ids)
    )
    return request.execute()


def extract_channel_data(items, collect_time):
    rows = []

    for item in items:
        snippet = item.get('snippet', {})
        stats = item.get('statistics', {})
        branding = item.get('brandingSettings', {}).get('channel', {})
        topic = item.get('topicDetails', {})

        desc = clean_text(snippet.get('description', ''))[:DESCRIPTION_MAX_LENGTH]
        topic_cats = '|'.join(topic.get('topicCategories', []))

        rows.append({
            'channel_id': item.get('id', ''),
            'title': snippet.get('title', ''),
            'description': desc,
            'custom_url': snippet.get('customUrl', ''),
            'published_at': snippet.get('publishedAt', ''),
            'country': snippet.get('country', ''),
            'subscriber_count': stats.get('subscriberCount', ''),
            'total_view_count': stats.get('viewCount', ''),
            'video_count': stats.get('videoCount', ''),
            'hidden_subscriber_count': stats.get('hiddenSubscriberCount', ''),
            'keywords': branding.get('keywords', ''),
            'topic_categories': topic_cats,
            'collected_at': collect_time,
        })

    return rows


def append_to_google_sheets(video_rows, channel_rows, spreadsheet_id, credentials_b64, collect_time):
    """Google Sheets에 데이터 추가. 실패해도 CSV 저장에 영향 없음."""
    import gspread
    from google.oauth2.service_account import Credentials

    creds_dict = json.loads(base64.b64decode(credentials_b64))
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(spreadsheet_id)

    # --- trending_log 월별 탭: append ---
    month_str = collect_time[:7].replace('-', '')  # "20260227T..." -> "202602"
    tab_name = f'trending_log_{month_str}'

    try:
        ws_trending = sh.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        ws_trending = sh.add_worksheet(title=tab_name, rows=1000, cols=30)

    if video_rows:
        headers = list(video_rows[0].keys())
        if not ws_trending.row_values(1):
            ws_trending.append_row(headers, value_input_option='RAW')
        rows = [[str(row.get(h, '')) for h in headers] for row in video_rows]
        ws_trending.append_rows(rows, value_input_option='RAW')
        print(f"  Sheets {tab_name}: {len(rows)}행 추가")

    # --- channels 탭: upsert ---
    try:
        ws_channels = sh.worksheet('channels')
    except gspread.exceptions.WorksheetNotFound:
        ws_channels = sh.add_worksheet(title='channels', rows=1000, cols=15)

    if channel_rows:
        headers = list(channel_rows[0].keys())
        existing = ws_channels.get_all_records()
        existing_ids = {row['channel_id']: idx + 2 for idx, row in enumerate(existing)}

        if not ws_channels.row_values(1):
            ws_channels.append_row(headers, value_input_option='RAW')

        new_rows = []
        batch_updates = []

        for ch in channel_rows:
            cid = ch['channel_id']
            if cid in existing_ids:
                row_num = existing_ids[cid]
                for col_idx, h in enumerate(headers):
                    val = str(ch.get(h, ''))
                    cell = gspread.utils.rowcol_to_a1(row_num, col_idx + 1)
                    batch_updates.append({'range': cell, 'values': [[val]]})
            else:
                new_rows.append([str(ch.get(h, '')) for h in headers])

        if batch_updates:
            ws_channels.batch_update(batch_updates, value_input_option='RAW')
        if new_rows:
            ws_channels.append_rows(new_rows, value_input_option='RAW')

        updated = len(channel_rows) - len(new_rows)
        print(f"  Sheets channels: {len(new_rows)}개 신규, {updated}개 업데이트")


def main():
    API_KEY = os.environ.get('YOUTUBE_API_KEY')
    if not API_KEY:
        raise ValueError("환경변수 'YOUTUBE_API_KEY'가 설정되어 있지 않습니다.")

    now = datetime.now(timezone.utc)
    date_str = now.strftime('%Y%m%d')
    hour_str = now.strftime('%H%M%S')
    collect_time = now.isoformat()

    # 디렉토리 구성
    WORK_DIR = os.path.join('./output', date_str, hour_str)
    CSV_DIR = os.path.join(WORK_DIR, 'trending')
    JSON_DIR = os.path.join(WORK_DIR, 'json')
    LOG_DIR = os.path.join(WORK_DIR, 'log')
    for d in [CSV_DIR, JSON_DIR, LOG_DIR]:
        os.makedirs(d, exist_ok=True)

    video_csv_path = os.path.join(CSV_DIR, f'videos_{date_str}_{hour_str}.csv')
    channel_csv_path = os.path.join(CSV_DIR, f'channels_{date_str}_{hour_str}.csv')
    video_json_path = os.path.join(JSON_DIR, 'videos_response.json')
    channel_json_path = os.path.join(JSON_DIR, 'channels_response.json')
    log_path = os.path.join(LOG_DIR, 'run.log')

    log_lines = []

    try:
        youtube = build('youtube', 'v3', developerKey=API_KEY)

        # 1) 인기급상승 영상 수집
        video_response = fetch_trending_videos(youtube)
        with open(video_json_path, 'w', encoding='utf-8') as f:
            json.dump(video_response, f, ensure_ascii=False, indent=2)

        items = video_response.get('items', [])
        video_rows, channel_ids = extract_video_data(items, collect_time)

        if not video_rows:
            raise ValueError("유효한 영상 데이터가 없음")

        df_videos = pd.DataFrame(video_rows)
        df_videos.to_csv(video_csv_path, index=False, encoding='utf-8-sig')
        print(f"영상 수집 완료: {len(df_videos)}개 → {video_csv_path}")
        log_lines.append(f"[{collect_time}] 영상 수집: {len(df_videos)}개")

        # 2) 채널 메타데이터 수집
        channel_rows = []
        if channel_ids:
            channel_response = fetch_channel_data(youtube, channel_ids)
            with open(channel_json_path, 'w', encoding='utf-8') as f:
                json.dump(channel_response, f, ensure_ascii=False, indent=2)

            ch_items = channel_response.get('items', [])
            channel_rows = extract_channel_data(ch_items, collect_time)

            df_channels = pd.DataFrame(channel_rows)
            df_channels.to_csv(channel_csv_path, index=False, encoding='utf-8-sig')
            print(f"채널 수집 완료: {len(df_channels)}개 → {channel_csv_path}")
            log_lines.append(f"[{collect_time}] 채널 수집: {len(df_channels)}개")

        # 3) Google Sheets 누적 (실패해도 계속 진행)
        sheets_creds = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
        sheet_id = os.environ.get('GOOGLE_SHEET_ID')
        if sheets_creds and sheet_id:
            try:
                append_to_google_sheets(video_rows, channel_rows, sheet_id, sheets_creds, collect_time)
                log_lines.append(f"[{collect_time}] Sheets 업데이트 성공")
            except Exception as e:
                print(f"Sheets 업데이트 실패 (비치명적): {e}")
                log_lines.append(f"[{collect_time}] Sheets 업데이트 실패: {e}")
        else:
            print("Sheets 환경변수 미설정, 건너뜀")
            log_lines.append(f"[{collect_time}] Sheets 건너뜀 (환경변수 없음)")

    except HttpError as e:
        print(f"API 요청 실패: {e}")
        log_lines.append(f"[{collect_time}] API 요청 실패: {e}")

    except Exception as e:
        print(f"예외 발생: {e}")
        log_lines.append(f"[{collect_time}] 예외 발생: {e}")

    finally:
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(log_lines) + '\n')


if __name__ == "__main__":
    main()
