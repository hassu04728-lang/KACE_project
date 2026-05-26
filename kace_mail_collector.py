import imaplib
import email
from email.header import decode_header
import pandas as pd
import hashlib
import re
from datetime import datetime
import os

# ==========================================
# 1. 환경 설정 및 경로
# ==========================================
IMAP_SERVER = "imap.worksmobile.com"
EMAIL_USER = "hassu04@korea.ac.kr" 
EMAIL_PASS = "rDWhYxiw612C" # 수현님의 16자리 앱 비밀번호 입력
EXCEL_FILE = "KACE_통합관리.xlsm"

def extract_date_from_title(title):
    """제목에서 날짜 패턴(MM/DD 등)을 찾아 추출함"""
    # 패턴: 3/31, 03.31, 4월 5일 등
    pattern = r"(\d{1,2}[./월]\s?\d{1,2}(?:일)?)"
    match = re.search(pattern, title)
    if match:
        return match.group(1)
    return "본문 확인 필요"

def get_selected_keywords_with_weight():
    """엑셀 Settings 시트에서 키워드, 사용여부(Y/N), 가중치(1~5)를 로드"""
    try:
        # 엑셀이 열려 있어도 읽기 위해 engine='openpyxl' 사용
        with pd.ExcelFile(EXCEL_FILE, engine='openpyxl') as xls:
            df = pd.read_excel(xls, sheet_name="Settings")
        
        df.columns = [c.strip() for c in df.columns]
        
        # '사용여부'가 'Y'인 행만 추출
        active_df = df[df['사용여부'].str.upper() == 'Y'].copy()
        # 가중치가 비어있거나 숫자가 아니면 1로 기본값 설정
        active_df['가중치'] = pd.to_numeric(active_df['가중치'], errors='coerce').fillna(1).astype(int)
        
        # {키워드: 가중치} 딕셔너리 생성
        return dict(zip(active_df['키워드'], active_df['가중치']))
    except Exception as e:
        print(f"(!) Settings 시트 로드 실패: {e}")
        # 실패 시 시스템 작동을 위한 기본값 반환
        return {"장학": 5, "인턴": 3, "공지": 1}

def fetch_kace_mails():
    try:
        # 1. 가중치 설정 로드
        keyword_map = get_selected_keywords_with_weight()
        keywords = list(keyword_map.keys())
        print(f"[*] 활성화된 키워드 및 가중치: {keyword_map}")

        # 2. 메일 서버 접속
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("INBOX")

        results = []
        status, messages = mail.search(None, 'ALL')
        # 최근 50건의 메일만 조사 (속도 최적화)
        target_ids = messages[0].split()[-50:]

        for m_id in target_ids:
            _, msg_data = mail.fetch(m_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    
                    # 제목 디코딩
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8", errors="ignore")
                    
                    if not subject: continue
                    
                    # 제목에 설정한 키워드가 포함되어 있는지 확인
                    matched = [kw for kw in keywords if str(kw) in subject]
                    
                    if matched:
                        kw = matched[0]
                        weight = keyword_map.get(kw, 1)
                        
                        # 본문 추출 (요약용 100자만)
                        body = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    body = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='ignore')
                                    break
                        else:
                            body = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8', errors='ignore')

                        # 제목에서 날짜 추출 (본문 노이즈 방지)
                        deadline_info = extract_date_from_title(subject)
                        
                        # 중복 제거를 위한 고유 Key 생성 (제목 + 발신일 기준)
                        mail_date = msg.get("Date")
                        unique_key = hashlib.md5(f"{subject}{mail_date}".encode()).hexdigest()

                        results.append({
                            "Module": "C" if any(ck in str(kw) for ck in ["채용", "인턴", "반도체", "공정"]) else "B",
                            "Source": "Email",
                            "Category": kw,
                            "Title": subject.strip(),
                            "Summary": re.sub(r'\s+', ' ', body).strip()[:100],
                            "PostedDate": mail_date,
                            "Deadline": deadline_info,
                            "Status": "미확인",
                            "Key": unique_key,
                            "Weight": weight, # 수현님이 설정한 1~5 가중치
                            "LastSeen": datetime.now().strftime('%Y-%m-%d')
                        })
                        print(f" -> 수집 성공: {subject[:15]}... (가중치: {weight})")

        mail.logout()
        
        # 3. 데이터 저장 및 중복 제거
        if results:
            df_final = pd.DataFrame(results)
            # 중복된 Key가 있다면 하나만 남김
            df_final = df_final.drop_duplicates(subset=['Key'], keep='first')
            
            # 엑셀(VBA)에서 읽기 편하도록 인코딩 설정하여 CSV 저장
            df_final.to_csv("KACE_Mail_DB.csv", index=False, encoding="utf-8-sig")
            print(f"\n[완료] 총 {len(df_final)}건의 중복 없는 데이터를 저장했습니다.")
        else:
            print("\n[알림] 일치하는 키워드의 새로운 메일이 없습니다.")

    except Exception as e:
        print(f"[전체 실행 오류] {e}")

if __name__ == "__main__":
    fetch_kace_mails()