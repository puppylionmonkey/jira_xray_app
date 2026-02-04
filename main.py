import flet as ft
import requests
import csv
import asyncio
import tkinter as tk
import configparser
import os
from tkinter import filedialog
from requests.auth import HTTPBasicAuth

# --- 設定檔自動化讀取邏輯 ---
config = configparser.ConfigParser()
CONFIG_FILE = "config.ini"


def load_config():
    # 如果檔案不存在，生成一個範本檔案
    if not os.path.exists(CONFIG_FILE):
        config['XRAY'] = {
            'CLIENT_ID': 'YOUR_XRAY_ID',
            'CLIENT_SECRET': 'YOUR_XRAY_SECRET'
        }
        config['JIRA'] = {
            'DOMAIN': 'yourname.atlassian.net',
            'EMAIL': 'your_email@example.com',
            'API_TOKEN': 'your_jira_api_token'
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            config.write(f)
        return False

    config.read(CONFIG_FILE, encoding='utf-8')
    return True


# 執行讀取
load_config()

# 將變數指向設定檔內容
CLIENT_ID = config.get('XRAY', 'CLIENT_ID', fallback="")
CLIENT_SECRET = config.get('XRAY', 'CLIENT_SECRET', fallback="")
JIRA_DOMAIN = config.get('JIRA', 'DOMAIN', fallback="")
JIRA_EMAIL = config.get('JIRA', 'EMAIL', fallback="")
JIRA_API_TOKEN = config.get('JIRA', 'API_TOKEN', fallback="")
BASE_URL = "https://xray.cloud.getxray.app/api/v2"


# --- API 邏輯 ---
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
    formatted_keys = ", ".join([f'"{k}"' for k in keys])
    query = {
        "query": "query($jql:String){getTests(jql:$jql,limit:100){results{jira(fields:[\"key\",\"summary\",\"priority\"]) testType{name} folder{path} steps{action data result}}}}",
        "variables": {"jql": f"key IN ({formatted_keys})"}
    }
    res = requests.post(f"{BASE_URL}/graphql", headers=headers, json=query)
    return res.json().get('data', {}).get('getTests', {}).get('results', [])


# --- Flet UI 主程式 ---
async def main(page: ft.Page):
    page.title = "Xray CSV Exporter Pro"
    page.theme_mode = ft.ThemeMode.DARK
    page.window.width = 650
    page.window.height = 680

    # 檢查設定檔是否已填寫 (簡單檢查)
    if "YOUR_" in CLIENT_ID or not JIRA_API_TOKEN:
        page.add(ft.Text("⚠️ 請先在 config.ini 中填寫正確的 API 資訊後重啟程式。", color="orange", size=20))
        page.update()
        return

    import_keys = []
    single_key_input = ft.TextField(label="輸入單個 PBPM 編號", hint_text="例如: PBPM-25818", expand=True)
    selected_files_text = ft.Text("尚未選取檔案", color=ft.Colors.GREY_500)
    merge_checkbox = ft.Checkbox(label="合併為單一 CSV 檔案", value=True)
    log_text = ft.Text(size=13)
    pb = ft.ProgressBar(width=600, visible=False, color=ft.Colors.BLUE_400)

    def write_to_csv(tests, filename, start_id_at):
        with open(filename, mode='w', newline='', encoding='utf-8-sig') as file:
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

    def start_export(keys, is_merge):
        clean_keys = [k for k in keys if k and k.strip()]
        if not clean_keys:
            log_text.value = "請輸入編號或匯入清單";
            log_text.color = ft.Colors.RED_400;
            page.update();
            return

        pb.visible = True
        log_text.value = "連線中...";
        page.update()

        token = get_xray_token()
        if not token:
            log_text.value = "認證失敗：請檢查 config.ini 中的 Xray ID/Secret";
            log_text.color = ft.Colors.RED_400;
            pb.visible = False;
            page.update();
            return

        results = fetch_xray_data(token, clean_keys)
        if not results:
            log_text.value = "找不到資料：請檢查 Key 是否正確";
            log_text.color = ft.Colors.ORANGE_400;
            pb.visible = False;
            page.update();
            return

        if is_merge:
            filename = f"Merged_{clean_keys[0]}.csv"
            write_to_csv(results, filename, 1)
            log_text.value = f"✅ 已合併: {filename}"
        else:
            for i, test in enumerate(results, 1):
                write_to_csv([test], f"{test['jira']['key']}.csv", i)
            log_text.value = f"✅ 已產出 {len(results)} 個檔案"

        log_text.color = ft.Colors.GREEN_400;
        pb.visible = False;
        page.update()

    def pick_file_sync():
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        file_path = filedialog.askopenfilename(title="選擇包含 Issue ID 的 CSV 檔案", filetypes=[("CSV Files", "*.csv")])
        root.destroy()
        return file_path

    async def pick_file_click(e):
        nonlocal import_keys
        loop = asyncio.get_event_loop()
        file_path = await loop.run_in_executor(None, pick_file_sync)
        if file_path:
            import_keys = []
            try:
                with open(file_path, newline='', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if row: import_keys.append(row[0].strip())
                selected_files_text.value = f"已讀取: {len(import_keys)} 筆";
                selected_files_text.color = ft.Colors.GREEN_400
            except:
                selected_files_text.value = "讀取失敗";
                selected_files_text.color = ft.Colors.RED_400
            page.update()

    page.add(
        ft.Row([ft.Icon(ft.Icons.CLOUD_SYNC, color=ft.Colors.BLUE_400), ft.Text("Xray CSV Exporter", size=24, weight="bold")]),
        ft.Divider(),
        ft.Text("方案 A: 單個編號", weight="bold"),
        ft.Row([single_key_input, ft.Button("下載", on_click=lambda _: start_export([single_key_input.value], False))]),
        ft.Container(height=20),
        ft.Text("方案 B: 批次匯入", weight="bold"),
        ft.Row([
            ft.Button("選取 CSV", icon=ft.Icons.UPLOAD_FILE, on_click=pick_file_click),
            selected_files_text
        ]),
        ft.Row([merge_checkbox, ft.Button("執行批次", on_click=lambda _: start_export(import_keys, merge_checkbox.value))]),
        ft.Divider(),
        pb,
        log_text
    )


if __name__ == "__main__":
    ft.run(main)