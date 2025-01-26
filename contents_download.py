import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import yt_dlp
import os
import zipfile
import time
import pandas as pd

# [중요] 구글 서비스 계정 키 파일 경로 또는 같은 디렉토리에 있는 경우 파일명
SERVICE_ACCOUNT_FILE = "isbn-search-443005-d650802511ad.json"

# 구글 시트 URL (혹은 ID를 통해 열어도 됩니다)
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1wtZedkiq_jVIRccEdh_3ssb_j6FFpU2K0coldo05EEc/edit?usp=sharing"

def authenticate_gsheet():
    """
    서비스 계정 키 파일을 사용해 구글 스프레드시트에 인증한 client를 반환
    """
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, 
        scopes=scopes
    )
    gc = gspread.authorize(credentials)
    return gc

def download_youtube_videos(video_links, download_dir):
    """
    yt-dlp를 사용하여 유튜브 링크들을 MP4(H.264 코덱)로 지정된 디렉토리에 다운로드
    """
    ydl_opts = {
        'format': 'bestvideo[ext=mp4][vcodec^=avc]+bestaudio[ext=m4a]/best[ext=mp4]',  # H.264 코덱 우선
        'outtmpl': os.path.join(download_dir, '%(title)s.%(ext)s'),
        'merge_output_format': 'mp4',
        'postprocessors': [{
            'key': 'FFmpegVideoConvertor',
            'preferedformat': 'mp4'
        }]
    }

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        ydl.download(video_links)

def create_zip_file(folder_path, zip_path):
    """
    folder_path 내의 모든 파일을 zip_path라는 이름으로 압축(zip) 파일 생성
    """
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                # 압축 시에는 파일명만 보존
                zipf.write(file_path, arcname=file)

def main():
    st.title("YouTube MP4 다운로드 자동화")

    # 사용자로부터 탭(시트) 이름 입력
    sheet_tab = st.text_input("다운로드할 탭(시트) 이름을 입력하세요:", value="예시탭")

    if st.button("다운로드 실행"):
        if not sheet_tab:
            st.warning("탭(시트) 이름을 입력하세요.")
            return

        # Step 1. 구글 스프레드시트 인증
        try:
            gc = authenticate_gsheet()
            spreadsheet = gc.open_by_url(SPREADSHEET_URL)
        except Exception as e:
            st.error(f"구글 시트 연결 오류: {e}")
            return

        # Step 2. 해당 탭(시트) 열기
        try:
            worksheet = spreadsheet.worksheet(sheet_tab)
        except Exception as e:
            st.error(f"'{sheet_tab}' 시트를 찾을 수 없습니다. 오류: {e}")
            return

        # 시트 데이터 가져오기 (첫 번째 행을 헤더로 가정)
        data = worksheet.get_all_records()

        if not data:
            st.warning(f"'{sheet_tab}' 시트에 데이터가 없습니다.")
            return

        df = pd.DataFrame(data)

        # '링크'와 '진행상태' 컬럼이 존재하는지 확인
        if '링크' not in df.columns or '진행상태' not in df.columns:
            st.error("시트에 '링크' 또는 '진행상태' 컬럼이 존재하지 않습니다.")
            return

        # '진행상태'가 '미진행'인 행만 필터링
        df_mijinhaeng = df[df['진행상태'] == '미진행'].copy()

        if df_mijinhaeng.empty:
            st.info("더 이상 '미진행' 상태의 링크가 없습니다.")
            return

        # Step 3. 유튜브 링크들을 다운로드할 임시 디렉토리 생성
        timestamp = int(time.time())
        download_dir = f"downloaded_videos_{timestamp}"
        os.makedirs(download_dir, exist_ok=True)

        # 미진행 상태 링크 수집
        video_links = df_mijinhaeng['링크'].tolist()

        st.write(f"총 {len(video_links)}개의 링크를 다운로드합니다...")
        try:
            download_youtube_videos(video_links, download_dir)
            st.success("모든 동영상 다운로드가 완료되었습니다.")
        except Exception as e:
            st.error(f"동영상 다운로드 중 오류가 발생했습니다: {e}")
            # 다운로드 중 에러가 나면 종료
            return

        # Step 4. 다운로드된 파일들을 ZIP 파일로 묶기
        zip_file_name = f"videos_{timestamp}.zip"
        create_zip_file(download_dir, zip_file_name)

        # Step 5. Streamlit에서 ZIP 파일 다운로드 제공
        with open(zip_file_name, "rb") as f:
            st.download_button(
                label="ZIP 파일 다운로드",
                data=f,
                file_name=zip_file_name,
                mime="application/zip"
            )

        # Step 6. 다운로드 완료된 링크들의 '진행상태'를 '진행완료'로 업데이트
        #   - Google Sheet에서 컬럼 위치 찾기
        header_row = worksheet.row_values(1)  # 첫 번째 행(헤더)
        link_col_idx = header_row.index('링크') + 1
        status_col_idx = header_row.index('진행상태') + 1

        #   - df에서 각각의 행 인덱스에 맞춰 업데이트
        #     (df의 인덱스와 시트의 행 번호를 맞춰주기 위해 +2를 해줌)
        for df_index in df_mijinhaeng.index:
            # 현재 시트에서 (df_index + 2) 행, status_col_idx 열을 '진행완료'로
            worksheet.update_cell(df_index + 2, status_col_idx, '진행완료')

        st.success("구글 시트의 '진행상태'가 '진행완료'로 업데이트되었습니다.")
        
        # 임시 디렉토리와 ZIP 파일 정리(선택사항)
        try:
            os.remove(zip_file_name)
            for f in os.listdir(download_dir):
                os.remove(os.path.join(download_dir, f))
            os.rmdir(download_dir)
        except:
            pass

if __name__ == "__main__":
    main()
