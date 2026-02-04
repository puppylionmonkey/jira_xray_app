import flet as ft
import requests
import csv
import asyncio
import tkinter as tk
import configparser
import os
from pathlib import Path  # 新增：用於偵測系統路徑
from tkinter import filedialog
from requests.auth import HTTPBasicAuth

# --- 設定檔自動化讀取邏輯 ---
config = configparser.ConfigParser()
CONFIG_FILE = "config.ini"


def load_config():
    if not os.path.exists(CONFIG_FILE):
        config['XRAY'] = {'CLIENT_ID': 'YOUR_XRAY_ID', 'CLIENT_SECRET': 'YOUR_XRAY_SECRET'}
        config['JIRA'] = {'DOMAIN': 'yourname.atlassian.net', 'EMAIL': 'your_email@example.com', 'API_TOKEN': 'your_jira_api_token'}
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            config.write(f)
        return False
    config.read(CONFIG_FILE, encoding='utf-8')
    return True


load_config()

CLIENT_ID = config.get('XRAY', 'CLIENT_ID', fallback="")
CLIENT_SECRET = config.get('XRAY', 'CLIENT_SECRET', fallback="")
JIRA_DOMAIN = config.get('JIRA', 'DOMAIN', fallback="")
JIRA_EMAIL = config.get('JIRA', 'EMAIL', fallback="")
JIRA_API_TOKEN = config.get('JIRA', 'API_TOKEN', fallback="")
BASE_URL = "https://xray.cloud.getxray.app/api/v2"


# --- 自動取得下載資料夾路徑 ---
def get_download_path():
    """取得當前系統的預設下載資料夾"""
    return str(Path.home() / "Downloads")


# --- API 邏輯 (保持原樣) ---
def get_xray_token():
    try:
        res = requests.post(f"{BASE_URL}/authenticate", json={"client_id": CLIENT_ID, "client_secret": CLIENT_SECRET}, timeout=10)
        return res.text.strip().replace('"', '') if res.status_code == 200 else None
    except:
        return None


def fetch_jira_links(key):
    auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)
    try:
        res = requests.get(f"https://{JIRA_DOMAIN}/rest/api/2/issue/{key}", auth=auth, timeout=5)
        if res.status_code == 200:
            links = res.json().get('fields', {}).get('issuelinks', [])
            keys = [(l.get('outwardIssue') or l.get('inwardIssue'))['key'] for l in links if (l.get('outwardIssue') or l.get('inwardIssue'))]
            return ";".join(keys)
    except:
        pass
    return ""


def fetch_xray_data(token, keys):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    all_results = []
    limit = 100
    start = 0

    # 先組裝好 JQL
    formatted_keys = ", ".join([f'"{k}"' for k in keys])
    jql_query = f"key IN ({formatted_keys})"

    while True:
        query = {
            # 這裡的 Int 改成 Int!，表示強制要求數值
            "query": """
                query($jql:String, $start:Int!, $limit:Int!){
                    getTests(jql:$jql, start:$start, limit:$limit){
                        results{
                            jira(fields:["key","summary","priority"])
                            testType{name}
                            folder{path}
                            steps{action data result}
                        }
                    }
                }
            """,
            "variables": {
                "jql": jql_query,
                "start": start,
                "limit": limit
            }
        }
        try:
            res = requests.post(f"{BASE_URL}/graphql", headers=headers, json=query, timeout=20)
            res_json = res.json()

            if "errors" in res_json:
                print(f"GraphQL 錯誤內容: {res_json['errors']}")
                break

            data = res_json.get('data', {}).get('getTests', {}).get('results', [])
            if not data:
                break

            all_results.extend(data)

            # 如果抓回來的數量小於要求數量，代表沒有下一頁了
            if len(data) < limit:
                break

            start += limit
        except Exception as e:
            print(f"連線異常: {e}")
            break
    return all_results


# --- Flet UI 主程式 ---
async def main(page: ft.Page):
    page.title = "Xray CSV Exporter Pro"
    page.theme_mode = ft.ThemeMode.DARK
    page.window.width = 650
    page.window.height = 680

    if "YOUR_" in CLIENT_ID or not JIRA_API_TOKEN:
        page.add(ft.Text("⚠️ 請先在 config.ini 中填寫正確的 API 資訊後重啟程式。", color="orange", size=20))
        page.update()
        return

    # 1. 變數初始化
    import_keys = []

    # 2. 定義 UI 元件（先定義基礎元件，不帶 on_click 的）
    single_key_input = ft.TextField(label="輸入單個 PBPM 編號", hint_text="例如: PBPM-25818", expand=True)
    selected_files_text = ft.Text("尚未選取檔案", color=ft.Colors.GREY_500)
    merge_checkbox = ft.Checkbox(label="合併為單一 CSV 檔案", value=True)
    log_text = ft.Text(size=13)
    loading_ring = ft.ProgressRing(width=30, height=30, stroke_width=3, visible=False, color=ft.Colors.BLUE_400)

    # 3. 定義狀態控制函式
    def set_ui_state(disabled: bool):
        """統一控制 UI 啟用狀態與轉圈圈顯示"""
        export_btn_single.disabled = disabled
        export_btn_batch.disabled = disabled
        pick_file_btn.disabled = disabled
        single_key_input.disabled = disabled
        merge_checkbox.disabled = disabled
        loading_ring.visible = disabled
        page.update()

    # 4. 定義 CSV 寫入函式
    def write_to_csv(tests, filename, start_id_at):
        download_folder = get_download_path()
        full_path = os.path.join(download_folder, filename)
        with open(full_path, mode='w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file)
            writer.writerow(['Test Repo', 'Issue Id', 'Issue key', 'Test type', 'Test Summary', 'Test Priority', 'Action', 'Data', 'Result', 'Links', 'Description', 'Unstructured definition'])
            for current_id, test in enumerate(tests, start=start_id_at):
                key = test['jira']['key']
                summary = test['jira']['summary']
                t_type = test['testType']['name']
                priority = test['jira']['priority']['name'] if test['jira'].get('priority') else ""
                repo_path = test['folder']['path'].lstrip('/') if test.get('folder') else ""
                links = fetch_jira_links(key)
                steps = test.get('steps', [])
                if not steps:
                    writer.writerow([repo_path, current_id, key, t_type, summary, priority, "", "", "", links, "", ""])
                else:
                    for s_idx, step in enumerate(steps):
                        if s_idx == 0:
                            writer.writerow([repo_path, current_id, key, t_type, summary, priority, step.get('action', ''), step.get('data', ''), step.get('result', ''), links, "", ""])
                        else:
                            writer.writerow(["", current_id, "", t_type, "", "", step.get('action', ''), step.get('data', ''), step.get('result', ''), "", "", ""])
        return full_path

    # 5. 定義匯出任務執行函式
    def start_export(keys, is_merge):
        def run_task():
            try:
                clean_keys = [k for k in keys if k and k.strip()]
                if not clean_keys:
                    log_text.value = "請輸入編號或匯入清單"
                    log_text.color = ft.Colors.RED_400
                else:
                    token = get_xray_token()
                    if not token:
                        log_text.value = "認證失敗：請檢查 config.ini"
                        log_text.color = ft.Colors.RED_400
                    else:
                        results = fetch_xray_data(token, clean_keys)
                        if not results:
                            log_text.value = "找不到資料：請檢查 Key 是否正確"
                            log_text.color = ft.Colors.ORANGE_400
                        else:
                            if is_merge:
                                filename = f"Merged_{clean_keys[0]}.csv"
                                write_to_csv(results, filename, 1)
                                log_text.value = f'✅ 已存至"下載"資料夾: {filename}'
                            else:
                                for i, test in enumerate(results, 1):
                                    write_to_csv([test], f"{test['jira']['key']}.csv", i)
                                log_text.value = f'✅ 已產出 {len(results)} 個檔案至"下載"資料夾'
                            log_text.color = ft.Colors.GREEN_400
            except Exception as e:
                log_text.value = f"錯誤: {str(e)}"
                log_text.color = ft.Colors.RED_400

            set_ui_state(False)

        log_text.value = "匯出中..."
        log_text.color = ft.Colors.BLUE_400
        set_ui_state(True)
        import threading
        threading.Thread(target=run_task, daemon=True).start()

    # 6. 定義檔案選取函式 (必須在被引用前定義)
    async def pick_file_click(e):
        def pick_sync():
            root = tk.Tk();
            root.withdraw();
            root.attributes('-topmost', True)
            res = filedialog.askopenfilename(title="選擇 CSV", filetypes=[("CSV Files", "*.csv")])
            root.destroy();
            return res

        loop = asyncio.get_event_loop()
        file_path = await loop.run_in_executor(None, pick_sync)
        if file_path:
            import_keys.clear()
            try:
                with open(file_path, newline='', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if row: import_keys.append(row[0].strip())
                selected_files_text.value = f"已讀取: {len(import_keys)} 筆"
                selected_files_text.color = ft.Colors.GREEN_400
            except:
                selected_files_text.value = "讀取失敗"
                selected_files_text.color = ft.Colors.RED_400
            page.update()

    # 7. 定義按鈕元件 (此時 pick_file_click 和 start_export 均已存在)
    export_btn_single = ft.Button("匯出", on_click=lambda _: start_export([single_key_input.value], False))
    export_btn_batch = ft.Button("匯出", on_click=lambda _: start_export(import_keys, merge_checkbox.value))
    pick_file_btn = ft.Button("選取 CSV", icon=ft.Icons.UPLOAD_FILE, on_click=pick_file_click)

    # 8. 組合畫面佈局
    page.add(
        ft.Row([ft.Icon(ft.Icons.CLOUD_SYNC, color=ft.Colors.BLUE_400), ft.Text("Xray CSV Exporter", size=24, weight="bold")]),
        ft.Divider(),
        ft.Text("匯出單個Test", weight="bold"),
        ft.Row([single_key_input, export_btn_single]),
        ft.Container(height=20),
        ft.Text("匯出多個Test", weight="bold"),
        ft.Row([pick_file_btn, selected_files_text]),
        ft.Row([merge_checkbox, export_btn_batch]),
        ft.Divider(),
        ft.Row([loading_ring], alignment=ft.MainAxisAlignment.CENTER),
        log_text
    )



if __name__ == "__main__":
    ft.run(main)

if __name__ == "__main__":
    ft.run(main)
