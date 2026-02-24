import csv
import json
import os
import time
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials
from flask import Flask
import threading

app = Flask(__name__)

LOGIN_URL = "https://hikkoshi-kanri.zba.jp/"
CSV_URL = "https://hikkoshi-kanri.zba.jp/checkbox/company/users/searched/50/1"
CHECK_INTERVAL = 30  # 30秒ごとにチェック

# グローバル変数
previous_data = None
monitoring_active = False


def send_gmail_notification(subject, body):
    """Gmail通知を送信"""
    try:
        gmail_address = os.environ.get("GMAIL_ADDRESS")
        gmail_password = os.environ.get("GMAIL_APP_PASSWORD")
        
        if not gmail_address or not gmail_password:
            print("⚠️ Gmail設定がありません")
            return False

        msg = MIMEMultipart()
        msg['From'] = gmail_address
        msg['To'] = gmail_address
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(gmail_address, gmail_password)
            server.send_message(msg)
        
        print("✅ Gmail通知を送信しました")
        return True
        
    except Exception as e:
        print(f"❌ Gmail通知エラー: {e}")
        return False


def get_current_data():
    """現在のデータを取得"""
    try:
        accounts = [
            (os.environ.get("WC_ID_1"), os.environ.get("WC_PASS_1")),
            (os.environ.get("WC_ID_2"), os.environ.get("WC_PASS_2")),
        ]
        
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            all_data = []
            
            for account_id, account_pass in accounts:
                if not account_id or not account_pass:
                    print(f"⚠️ アカウント情報が設定されていません")
                    continue
                
                page = browser.new_page()
                
                try:
                    # ログイン
                    print(f"🔐 ログイン中: {account_id}")
                    page.goto(LOGIN_URL, wait_until="networkidle", timeout=30000)
                    page.wait_for_selector("input[autocomplete='username']", timeout=15000)
                    page.fill("input[autocomplete='username']", account_id)
                    page.fill("input[autocomplete='current-password']", account_pass)
                    
                    try:
                        page.click("button[type='submit']", timeout=3000)
                    except:
                        page.press("input[autocomplete='current-password']", "Enter")
                    
                    page.wait_for_load_state("networkidle", timeout=30000)
                    
                    # CSV取得
                    print(f"📥 CSVダウンロード中: {account_id}")
                    page.goto(CSV_URL, wait_until="networkidle", timeout=30000)
                    
                    filename = f"/tmp/temp_{account_id}.csv"
                    selectors = [
                        "button:text('CSV')",
                        "button:text('出力')",
                        "button:text('ダウンロード')",
                        "a[href*='export']",
                        "a[href*='csv']",
                        "a:text('CSV')",
                    ]
                    
                    downloaded = False
                    for selector in selectors:
                        try:
                            with page.expect_download(timeout=15000) as dl_info:
                                page.click(selector, timeout=5000)
                            dl = dl_info.value
                            dl.save_as(filename)
                            print(f"✅ CSV保存: {filename}")
                            downloaded = True
                            break
                        except Exception as e:
                            continue
                    
                    if not downloaded:
                        print(f"❌ CSVダウンロード失敗: {account_id}")
                        page.close()
                        continue
                    
                    # CSV読み込み
                    for encoding in ["shift_jis", "cp932", "utf-8-sig", "utf-8"]:
                        try:
                            with open(filename, "r", encoding=encoding) as f:
                                rows = list(csv.reader(f))
                            print(f"✅ CSV読み込み成功: {len(rows)}行")
                            if all_data and len(rows) > 0:
                                all_data.extend(rows[1:])  # ヘッダースキップ
                            elif len(rows) > 0:
                                all_data = rows
                            break
                        except Exception as e:
                            continue
                    
                except Exception as e:
                    print(f"❌ アカウント処理エラー ({account_id}): {e}")
                finally:
                    page.close()
            
            browser.close()
            print(f"📊 合計データ: {len(all_data)}行")
            return all_data if len(all_data) > 1 else None
            
    except Exception as e:
        print(f"❌ データ取得エラー: {e}")
        import traceback
        traceback.print_exc()
        return None


def update_spreadsheet(data):
    """スプレッドシートを更新"""
    try:
        sa_json = os.environ.get("GCP_SERVICE_ACCOUNT")
        if not sa_json:
            print("⚠️ GCP_SERVICE_ACCOUNT が設定されていません")
            return False
        
        sa_info = json.loads(sa_json)
        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
        gc = gspread.authorize(creds)
        
        SHEET_ID = "1zfnTMt8RKAojSBZ51M3M2s73vTneFP8eyyVEYRtxlwM"
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet("row")
        ws.clear()
        ws.append_rows(data)
        
        print(f"✅ スプレッドシート更新: {len(data)}行")
        return True
        
    except Exception as e:
        print(f"❌ スプレッドシート更新エラー: {e}")
        return False


def check_new_cases():
    """新規案件をチェック"""
    global previous_data
    
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"\n{'='*60}")
    print(f"[{timestamp}] チェック開始")
    print(f"{'='*60}")
    
    current_data = get_current_data()
    
    if not current_data:
        print("⚠️ データ取得失敗")
        return
    
    if previous_data and len(previous_data) > 1:
        # 新規案件を検出
        prev_set = set(tuple(row) for row in previous_data[1:])
        new_cases = [row for row in current_data[1:] if tuple(row) not in prev_set]
        
        if new_cases:
            print(f"🆕 新規案件検出: {len(new_cases)}件")
            
            # 通知メッセージを作成
            header = current_data[0]
            subject = f"🚨 緊急！新規案件 {len(new_cases)}件"
            
            body = f"【{timestamp}】\n\n"
            body += f"新しい引越し案件が {len(new_cases)}件 追加されました！\n"
            body += "すぐに対応してください。\n\n"
            body += "="*60 + "\n\n"
            
            for i, case in enumerate(new_cases[:3], 1):
                body += f"【案件 {i}】\n"
                for j, (col_name, value) in enumerate(zip(header, case)):
                    if value and j < 10:
                        body += f"  {col_name}: {value}\n"
                body += "\n" + "-"*40 + "\n\n"
            
            if len(new_cases) > 3:
                body += f"\n※ 他 {len(new_cases) - 3}件の新規案件があります\n\n"
            
            body += "="*60 + "\n"
            body += "スプレッドシートで確認:\n"
            body += f"https://docs.google.com/spreadsheets/d/1zfnTMt8RKAojSBZ51M3M2s73vTneFP8eyyVEYRtxlwM\n"
            
            # Gmail送信
            send_gmail_notification(subject, body)
            
            # スプレッドシート更新
            update_spreadsheet(current_data)
        else:
            print("✓ 新規案件なし")
    else:
        print("ℹ️ 初回チェック（通知スキップ）")
        # 初回でもスプレッドシートは更新
        update_spreadsheet(current_data)
    
    previous_data = current_data


def monitoring_loop():
    """監視ループ（バックグラウンドで実行）"""
    global monitoring_active
    monitoring_active = True
    
    print("🔍 監視開始")
    print(f"⏱️  チェック間隔: {CHECK_INTERVAL}秒")
    
    while monitoring_active:
        try:
            check_new_cases()
            print(f"⏳ {CHECK_INTERVAL}秒待機...")
            time.sleep(CHECK_INTERVAL)
        except Exception as e:
            print(f"❌ 監視ループエラー: {e}")
            import traceback
            traceback.print_exc()
            time.sleep(CHECK_INTERVAL)


@app.route('/')
def index():
    """ヘルスチェック"""
    status = "🟢 稼働中" if monitoring_active else "🔴 停止中"
    return f'''
    <html>
    <head><title>案件監視システム</title></head>
    <body style="font-family: sans-serif; padding: 40px;">
        <h1>🔍 案件監視システム</h1>
        <p>ステータス: <strong>{status}</strong></p>
        <p>チェック間隔: {CHECK_INTERVAL}秒</p>
        <p>最終チェック: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </body>
    </html>
    '''


@app.route('/health')
def health():
    """ヘルスチェック（シンプル）"""
    return 'OK', 200


if __name__ == '__main__':
    # 監視をバックグラウンドスレッドで開始
    monitor_thread = threading.Thread(target=monitoring_loop, daemon=True)
    monitor_thread.start()
    
    # Flaskサーバーを起動（Renderが必要とする）
    port = int(os.environ.get('PORT', 8080))
    print(f"🚀 サーバー起動: ポート {port}")
    app.run(host='0.0.0.0', port=port)
