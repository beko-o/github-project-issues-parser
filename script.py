import os
import sys
import re
import logging
import requests
import pandas as pd
import argparse
from dotenv import load_dotenv
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)

# Парсинг аргументов (гибридный режим)
parser = argparse.ArgumentParser(
    description="Генерация отчёта по GitHub Projects для заданного assignee и периода"
)
parser.add_argument("--assignee", "-a", help="GitHub-логин assignee")
parser.add_argument("--start", "-s", help="Дата начала анализа (YYYY-MM-DD)")
parser.add_argument("--end", "-e", help="Дата окончания анализа (YYYY-MM-DD)")
parser.add_argument("--folder-id", "-f", help="ID папки Google Drive для сохранения отчета")
args = parser.parse_args()

# Ввод параметров
assignee_login = args.assignee or input("Введите GitHub-логин assignee: ").strip()

def parse_date(arg, prompt):
    if arg:
        try:
            return datetime.fromisoformat(arg)
        except ValueError:
            parser.error(f"{prompt} должен быть в формате YYYY-MM-DD")
    while True:
        val = input(f"Введите {prompt} (YYYY-MM-DD): ").strip()
        try:
            return datetime.fromisoformat(val)
        except ValueError:
            logger.warning("Некорректный формат, попробуйте ещё раз.")

start_date = parse_date(args.start, "дату начала анализа")
end_date   = parse_date(args.end,   "дату окончания анализа")

# Загрузка GitHub токена
load_dotenv()
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
if not GITHUB_TOKEN:
    logger.error("Не найден GITHUB_TOKEN в окружении.")
    sys.exit(1)
headers = {"Authorization": f"Bearer {GITHUB_TOKEN}"}

# Константы и настройки
ORG = "Maxinum"
ACTUAL_NAMES   = {"actual", "actual hours", "acutal hours"}
ESTIMATE_NAMES = {"estimate", "planned hours", "hours", "estimate hours", "estimates"}
PAGE_SIZE = 100

# GraphQL-запросы
proj_list_query = '''
query($org:String!,$first:Int!,$after:String){
  organization(login:$org){
    projectsV2(first:$first,after:$after){
      pageInfo{hasNextPage,endCursor}
      nodes{number,title,updatedAt}
    }
  }
}'''

items_query = '''
query($org:String!,$projNum:Int!,$first:Int!,$after:String){
  organization(login:$org){
    projectV2(number:$projNum){
      items(first:$first,after:$after){
        pageInfo{hasNextPage,endCursor}
        nodes{
          content{...on Issue{number,title,repository{name},assignees(first:10){nodes{login}},url,createdAt}}
          fieldValues(first:20){nodes{...on ProjectV2ItemFieldNumberValue{field{...on ProjectV2FieldCommon{name}},number}}}
        }
      }
    }
  }
}'''

# Функции для выборки данных

def fetch_all_projects():
    projects, cursor = [], None
    while True:
        resp = requests.post(
            "https://api.github.com/graphql",
            json={"query": proj_list_query, "variables": {"org":ORG, "first":PAGE_SIZE, "after":cursor}},
            headers=headers, timeout=10
        )
        resp.raise_for_status()
        data = resp.json().get("data", {}).get("organization", {}).get("projectsV2", {})
        projects.extend(data.get("nodes", []))
        page = data.get("pageInfo", {})
        if not page.get("hasNextPage"): break
        cursor = page.get("endCursor")
    return projects


def fetch_proj(proj):
    num, title = proj["number"], proj.get("title", "")
    rows, cursor = [], None
    while True:
        resp = requests.post(
            "https://api.github.com/graphql",
            json={"query": items_query, "variables": {"org":ORG, "projNum":num, "first":PAGE_SIZE, "after":cursor}},
            headers=headers, timeout=10
        )
        resp.raise_for_status()
        items = resp.json().get("data", {}).get("organization", {}).get("projectV2", {}).get("items", {})
        for it in items.get("nodes", []):
            issue = it.get("content")
            if not issue or "assignees" not in issue: continue
            if assignee_login not in [a.get("login") for a in issue.get("assignees", {}).get("nodes", [])]: continue
            try:
                created = datetime.fromisoformat(issue.get("createdAt", "").rstrip("Z"))
            except ValueError:
                continue
            if not (start_date <= created <= end_date): continue
            actual = estimate = 0
            for fv in it.get("fieldValues", {}).get("nodes", []):
                fld, val = fv.get("field"), fv.get("number")
                if fld and val is not None:
                    nm = fld.get("name", "").strip().lower()
                    if nm in ACTUAL_NAMES: actual = val
                    if nm in ESTIMATE_NAMES: estimate = val
            rows.append({
                "number":    issue.get("number"),
                "title":     issue.get("title"),
                "repo":      issue.get("repository", {}).get("name"),
                "url":       issue.get("url"),
                "createdAt": created,
                "actual":    actual,
                "estimate":  estimate
            })
        page = items.get("pageInfo", {})
        if not page.get("hasNextPage"): break
        cursor = page.get("endCursor")
    if not rows: return None
    df = pd.DataFrame(rows).fillna(0)
    safe = re.sub(r"[\\\/\?\*\[\]:]", "_", title)[:25]
    return num, safe, df

# Основная логика
all_projects = fetch_all_projects()
projects = [p for p in all_projects if datetime.fromisoformat(p.get("updatedAt","").rstrip("Z")) > start_date]
if not projects:
    logger.info(f"Нет проектов после {start_date.date()}.")
    sys.exit(0)
logger.info(f"Найдено {len(projects)} проектов после {start_date.date()}.")

# Параллельный сбор задач
results = []
with ThreadPoolExecutor(max_workers=5) as executor:
    futures = {executor.submit(fetch_proj, p): p for p in projects}
    for fut in as_completed(futures):
        proj = futures[fut]
        try:
            res = fut.result()
            if res:
                n,s,df = res
                results.append((n,s,df))
                logger.info(f"Проект #{n}: {len(df)} задач")
            else:
                logger.info(f"Проект #{proj['number']}: нет задач")
        except Exception as ex:
            logger.error(f"Ошибка проекта #{proj['number']}: {ex}")

# Интеграция с Google Sheets и создание графиков
try:
    from google.oauth2.service_account import Credentials
    import gspread
    from gspread_dataframe import set_with_dataframe
    from googleapiclient.discovery import build

    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path or not os.path.exists(creds_path):
        logger.error("credentials.json не найден.")
        sys.exit(1)
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    drive = build('drive','v3',credentials=creds)
    sheets = build('sheets','v4',credentials=creds)

    folder_id = args.folder_id or "1J_10hg5Lu1kOk69Oj9Yo2E_xSZefaSSu"
    report_title = f"GitHub Report {assignee_login} {start_date.date()}_{end_date.date()}"
    sh = gc.create(report_title)
    sheet_id = sh.id
    drive.files().update(fileId=sheet_id, addParents=folder_id, removeParents='root', fields='id, parents').execute()
    logger.info(f"Создан Spreadsheet '{report_title}' (ID={sheet_id}) в папке {folder_id}")

    # Удаляем пустой лист
    try: sh.del_worksheet(sh.sheet1)
    except: pass

    sheet_ids, data_map = {}, {}
    for n,s,df in results:
        ws = sh.add_worksheet(title=f"{n}_{s}", rows=df.shape[0]+5, cols=df.shape[1]+2)
        set_with_dataframe(ws, df)
        sid = ws._properties['sheetId']
        sheet_ids[f"{n}_{s}"] = sid
        data_map[f"{n}_{s}"] = df

    if results:
        combined = pd.concat([df.assign(project=f"{n}_{s}") for n,s,df in results], ignore_index=True)
        ws_sum = sh.add_worksheet(title='Summary', rows=combined.shape[0]+5, cols=combined.shape[1]+2)
        set_with_dataframe(ws_sum, combined)
        sheet_ids['Summary'] = ws_sum._properties['sheetId']
        data_map['Summary'] = combined

    # Добавляем диаграммы
    requests = []
    for name, df in data_map.items():
        sid = sheet_ids[name]
        # пишем summary строки
        row = df.shape[0] + 2
        ws = sh.worksheet(name)
        ws.update(range_name=f"A{row}:B{row}", values=[["estimate","actual"]])
        ws.update(range_name=f"A{row+1}:B{row+1}", values=[[df['estimate'].sum(), df['actual'].sum()]])
        requests.append({
            'addChart':{
                'chart':{
                    'spec':{
                        'title':'Сумма часов',
                        'basicChart':{
                            'chartType':'COLUMN','legendPosition':'BOTTOM_LEGEND',
                            'domains':[{'domain':{'sourceRange':{'sources':[{'sheetId':sid,'startRowIndex':row-1,'endRowIndex':row,'startColumnIndex':0,'endColumnIndex':2}]}}}],
                            'series':[{'series':{'sourceRange':{'sources':[{'sheetId':sid,'startRowIndex':row,'endRowIndex':row+1,'startColumnIndex':0,'endColumnIndex':2}]}},'targetAxis':'LEFT_AXIS'}]
                        }
                    },
                    'position':{'overlayPosition':{'anchorCell':{'sheetId':sid,'rowIndex':row,'columnIndex':0}}}
                }
            }
        })
    if requests:
        sheets.spreadsheets().batchUpdate(spreadsheetId=sheet_id,body={'requests':requests}).execute()
        logger.info("Диаграммы добавлены в Google Sheets")

except ImportError as ie:
    logger.error(f"Не установлены библиотеки для Google: {ie}")
    sys.exit(1)
except Exception as e:
    logger.error(f"Ошибка Google интеграции: {e}")
    sys.exit(1)
