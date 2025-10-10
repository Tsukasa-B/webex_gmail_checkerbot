import os
import os.path
import base64
import re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from webexteamssdk import WebexTeamsAPI

# --- 設定項目 ---
# Gmail APIのスコープ
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
# 検索するメールの条件
GMAIL_QUERY = 'subject:"[実験実習購入]" newer_than:1d is:unread'

# ★★★ WebexのアクセストークンとルームIDを環境変数から取得 ★★★
WEBEX_BOT_TOKEN = os.environ.get('WEBEX_BOT_TOKEN')
WEBEX_ROOM_ID = os.environ.get('WEBEX_ROOM_ID')
# --- 設定項目ここまで ---

def get_gmail_service():
    """Gmail APIサービスへの認証とサービスオブジェクトの取得"""
    creds_info = {
        "token": os.environ.get('GMAIL_TOKEN'),
        "refresh_token": os.environ.get('GMAIL_REFRESH_TOKEN'),
        "token_uri": "https://oauth2.googleapis.com/token",
        "client_id": os.environ.get('GMAIL_CLIENT_ID'),
        "client_secret": os.environ.get('GMAIL_CLIENT_SECRET'),
        "scopes": SCOPES
    }

    try:
        creds = Credentials.from_authorized_user_info(creds_info, SCOPES)

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        
        service = build('gmail', 'v1', credentials=creds)
        print("Gmail APIサービスへの接続に成功しました。")
        return service
    except Exception as e:
        print(f'Gmail APIサービスへの接続中にエラーが発生しました: {e}')
        return None
    
def fetch_emails(service, query):
    """指定されたクエリでメールを検索して取得する"""
    if not service:
        return []
    try:
        result = service.users().messages().list(userId='me', q=query).execute()
        messages_info = result.get('messages', [])

        if not messages_info:
            print(f"条件「{query}」に一致する新しいメールは見つかりませんでした。")
            return []

        emails_data = []
        print(f"--- {len(messages_info)}件のメールが見つかりました ---")
        for msg_info in messages_info:
            msg = service.users().messages().get(userId='me', id=msg_info['id'], format='full').execute()
            payload = msg.get('payload', {})
            headers = payload.get('headers', [])
            
            subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), '件名なし')
            
            body_data = ""
            if 'parts' in payload: # マルチパートメールの場合
                for part in payload['parts']:
                    if part.get('mimeType') == 'text/plain':
                        body_data_encoded = part.get('body', {}).get('data')
                        if body_data_encoded:
                            body_data = base64.urlsafe_b64decode(body_data_encoded.encode('ASCII')).decode('utf-8')
                            break
            elif 'body' in payload and payload['body'].get('data'): # シングルパートメールの場合
                body_data_encoded = payload['body']['data']
                body_data = base64.urlsafe_b64decode(body_data_encoded.encode('ASCII')).decode('utf-8')
            
            emails_data.append({
                'id': msg_info['id'],
                'subject': subject,
                'body': body_data.strip()
            })
            print(f"\n取得したメール件名: {subject}")
            print(f"本文 (最初の100文字): {body_data.strip()[:100]}...")
        return emails_data
    except HttpError as error:
        print(f'メールの取得中にエラーが発生しました: {error}')
        return []
    
def extract_info_from_email(subject, body):
    """メールの件名と本文から情報を抽出する"""
    info = {}
    print(f"\n--- 情報抽出開始 ---")
    print(f"件名: {subject}")

    match_subject = re.search(r'(?:\[daitai:\d+\]\s*)?\[実験実習購入\]\s*(\d+)\s*/\s*([^（]+)(?:\s*（(.*?)）)?', subject)
    if match_subject:
        info['申請番号'] = match_subject.group(1).strip()
        info['品名'] = match_subject.group(2).strip()
        if match_subject.group(3):
            info['備考'] = match_subject.group(3).strip()
        print(f"件名から抽出: 申請番号='{info.get('申請番号')}', 品名='{info.get('品名')}', 備考='{info.get('備考')}'")
    else: 
        match_subject_alt = re.search(r'\[実験実習購入\]\s*(\d+)\s*/\s*([^（]+)(?:\s*（(.*?)）)?', subject)
        if match_subject_alt:
            info['申請番号'] = match_subject_alt.group(1).strip()
            info['品名'] = match_subject_alt.group(2).strip()
            if match_subject_alt.group(3):
                info['備考'] = match_subject_alt.group(3).strip()
            print(f"件名から抽出 (別パターン): 申請番号='{info.get('申請番号')}', 品名='{info.get('品名')}', 備考='{info.get('備考')}'")

    if not info.get('品名'):
        match_body_item = re.search(r'「(.*?)」です。', body)
        if match_body_item:
            info['品名'] = match_body_item.group(1).strip()
            print(f"本文から抽出 (品名): '{info['品名']}'")
    
    if not info.get('申請番号'):
        match_body_id = re.search(r'申請番号：\s*(\d+)', body)
        if match_body_id:
            info['申請番号'] = match_body_id.group(1).strip()
            print(f"本文から抽出 (申請番号): '{info['申請番号']}'")

    if "大至急" in subject or "大至急" in body:
        info['緊急度'] = "大至急"
        print("緊急度: 大至急")

    match_submission = re.search(r'(精密事務室\(\d+号室\))へ提出してください', body)
    if match_submission:
        info['書類提出先'] = match_submission.group(1)
        print(f"書類提出先: {info['書類提出先']}")
    
    if not info.get('品名') and not info.get('申請番号'):
        print("必要な情報が抽出できませんでした。")
        return {}
        
    print(f"最終抽出情報: {info}")
    return info

def send_message_to_webex(info_dict):
    """抽出された情報を使ってWebexにメッセージを送信する"""
    if not info_dict:
        print("Webexに送信する情報がありません (info_dictが空です)。")
        return False

    if not WEBEX_BOT_TOKEN:
        print("エラー: Webex Botのアクセストークンが設定されていません。GitHub Secretsを確認してください。")
        return False
    if not WEBEX_ROOM_ID:
        print("エラー: WebexのルームIDが設定されていません。GitHub Secretsを確認してください。")
        return False

    try:
        api = WebexTeamsAPI(access_token=WEBEX_BOT_TOKEN)
    except Exception as e:
        print(f"WebexTeamsAPIの初期化中にエラーが発生しました: {e}")
        return False

    message_title = "【📢 納品連絡】"
    if info_dict.get('緊急度') == "大至急":
        message_title = "【🚨 大至急！納品連絡】"

    message_parts = [message_title]

    item_name = info_dict.get('品名')
    if item_name:
        message_parts.append(f"- **品名**: {item_name}")
    else:
        message_parts.append(f"- **件名(品名不明)**: {info_dict.get('subject', '件名情報なし')}")

    if info_dict.get('申請番号'):
        message_parts.append(f"- **申請番号**: {info_dict['申請番号']}")
    if info_dict.get('備考'):
        message_parts.append(f"- **その他**: {info_dict['備考']}")

    message_parts.append("\n**メールからの指示・情報：**")
    if info_dict.get('書類提出先'):
        message_parts.append(f"- 書類は **{info_dict['書類提出先']}** へ提出してください。")
    else:
        message_parts.append("- 受け取り後、納品書等の書類は速やかに精密事務室へ提出してください。")
    message_parts.append("- 担当者は内容を確認し、対応をお願いします。")
    message_parts.append("- 取りに行ったら、他の人にも共有しましょう！")

    original_body_text = info_dict.get('original_body')
    if original_body_text:
        message_parts.append("\n--- ▼元のメール本文▼ ---")
        message_parts.append(f"```{original_body_text}```")
        message_parts.append("--- ▲元のメール本文▲ ---")

    message_markdown = "\n".join(message_parts)

    try:
        api.messages.create(roomId=WEBEX_ROOM_ID, markdown=message_markdown)
        print(f"Webexスペース「{WEBEX_ROOM_ID}」にメッセージを送信しました。")
        print(f"送信内容:\n{message_markdown}")
        return True
    except Exception as e:
        print(f"Webexへのメッセージ送信中にエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
        return False

# --- メインの処理 ---
if __name__ == '__main__':
    print("--- 納品連絡自動通知プログラム開始 ---")
    gmail_service = get_gmail_service()
    
    if gmail_service:
        print(f"\nメールを検索します... (検索条件: '{GMAIL_QUERY}')")
        new_emails = fetch_emails(gmail_service, GMAIL_QUERY)
        
        if new_emails:
            print(f"\n--- {len(new_emails)}件のメールから情報を抽出し、Webexに通知します ---")
            processed_count = 0
            for email_data in new_emails:
                print(f"\n処理中のメール件名: {email_data['subject']}")
                extracted_info = extract_info_from_email(email_data['subject'], email_data['body'])
                
                if extracted_info:
                    if 'subject' not in extracted_info:
                        extracted_info['subject'] = email_data['subject']

                    extracted_info['original_body'] = email_data['body']

                    print(">>> 抽出成功:", extracted_info)
                    
                    if send_message_to_webex(extracted_info):
                        processed_count += 1
                        print(f"メール (ID: {email_data['id']}) の通知をWebexに送信しました。")
                        try:
                            gmail_service.users().messages().modify(userId='me', id=email_data['id'], body={'removeLabelIds': ['UNREAD']}).execute()
                            print(f"メール {email_data['id']} を既読にしました。")
                        except HttpError as error:
                            print(f"メール {email_data['id']} の既読化に失敗: {error}")
                    else:
                        print(f"メール (件名: {email_data['subject']}) のWebex通知に失敗しました。")
                else:
                    print(">>> フィルタリングされたか、情報抽出に失敗しました。Webex通知は行いません。")
            
            if processed_count > 0:
                print(f"\n--- {processed_count}件の納品連絡をWebexに通知しました ---")
            else:
                print("\n--- 今回、Webexに通知された納品連絡はありませんでした (フィルタリング結果またはエラーによる) ---")
        else:
            print("\n--- 処理対象のメールはありませんでした ---")
            
    print("\n--- 納品連絡自動通知プログラム終了 ---")
