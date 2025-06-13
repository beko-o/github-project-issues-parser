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
    description="Генерация отчётов по GitHub Projects для одного или всех assignee за период"
)
parser.add_argument("--assignee", "-a", help="GitHub-логин assignee")
parser.add_argument("--start", "-s", required=True, help="Дата начала анализа (YYYY-MM-DD)")
parser.add_argument("--end", "-e",   required=True, help="Дата окончания анализа (YYYY-MM-DD)")
parser.add_argument("--folder-id", "-f", help="ID папки Google Drive для сохранения отчета")
args = parser.parse_args()

# Валидация дат
try:
    start_date = datetime.fromisoformat(args.start)
    end_date   = datetime.fromisoformat(args.end)
except ValueError:
    parser.error("Даты должны быть в формате YYYY-MM-DD")

# Загрузка GitHub токена
load_dotenv()
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
if not GITHUB_TOKEN:
    logger.error("Не найден GITHUB_TOKEN в окружении.")
    sys.exit(1)
headers = {"Authorization": f"Bearer {GITHUB_TOKEN}"}

# Константы
ORG = "Maxinum"
PAGE_SIZE = 100
ACTUAL_NAMES   = {"actual", "actual hours", "acutal hours"}
ESTIMATE_NAMES = {"estimate", "planned hours", "hours", "estimate hours", "estimates"}

# GraphQL-запросы
proj_query = '''
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

# Пагинация GraphQL

def graphql_paginate(query, variables, path):
    items, cursor = [], None
    while True:
        vars = variables.copy()
        vars.update({"first": PAGE_SIZE, "after": cursor})
        r = requests.post(
            "https://api.github.com/graphql",
            json={"query": query, "variables": vars},
            headers=headers, timeout=10
        )
        r.raise_for_status()
        data = r.json().get("data", {})
        for key in path.split('.'):
            data = data.get(key, {})
        nodes = data.get("nodes", [])
        items.extend(nodes)
        page = data.get("pageInfo", {})
        if not page.get("hasNextPage"): break
        cursor = page.get("endCursor")
    return items

# Сбор проектов, изменённых после start_date
all_projects = graphql_paginate(proj_query, {"org": ORG}, "organization.projectsV2")
projects = []
for p in all_projects:
    try:
        dt = datetime.fromisoformat(p.get("updatedAt", "").rstrip("Z"))
        if dt >= start_date:
            projects.append(p)
    except ValueError:
        logger.warning(f"Невалидный updatedAt в проекте #{p.get('number')}")
if not projects:
    logger.info(f"Нет проектов после {start_date.date()}.")
    sys.exit(0)
logger.info(f"Найдено {len(projects)} проектов после {start_date.date()}.")

# Сбор всех задач за период
def fetch_proj_all(proj):
    num = proj["number"]
    title = re.sub(r"[\\\/\?\*\[\]:]", "_", proj.get("title",""))[:25]
    rows, cursor = [], None
    while True:
        r = requests.post(
            "https://api.github.com/graphql",
            json={"query": items_query, "variables": {"org": ORG, "projNum": num, "first": PAGE_SIZE, "after": cursor}},
            headers=headers, timeout=10
        )
        r.raise_for_status()
        data = r.json().get("data", {})
        items = data.get("organization", {}).get("projectV2", {}).get("items", {})
        for it in items.get("nodes", []):
            issue = it.get("content")
            if not issue: continue
            try:
                created = datetime.fromisoformat(issue.get("createdAt",""
                                                        ).rstrip("Z"))
            except ValueError:
                continue
            if not (start_date <= created <= end_date): continue
            assignees = [a.get("login") for a in issue.get("assignees", {}).get("nodes", [])]
            actual = estimate = 0
            for fv in it.get("fieldValues", {}).get("nodes", []):
                fld, val = fv.get("field"), fv.get("number")
                if fld and val is not None:
                    nm = fld.get("name","" ).strip().lower()
                    if nm in ACTUAL_NAMES: actual = val
                    if nm in ESTIMATE_NAMES: estimate = val
            rows.append({
                "project": f"{num}_{title}",
                "number":  issue.get("number"),
                "title":   issue.get("title"),
                "repo":    issue.get("repository", {}).get("name"),
                "url":     issue.get("url"),
                "createdAt": created,
                "assignees": assignees,
                "actual":    actual,
                "estimate":  estimate
            })
        page = items.get("pageInfo", {})
        if not page.get("hasNextPage"): break
        cursor = page.get("endCursor")
    return pd.DataFrame(rows)

# Собираем все задачи по всем проектам
with ThreadPoolExecutor(max_workers=3) as executor:
    dfs = list(executor.map(fetch_proj_all, projects))
all_tasks = pd.concat(dfs, ignore_index=True)

# Список assignee для отчётов
if args.assignee:
    users = [args.assignee]
else:
    users = sorted({u for lst in all_tasks['assignees'] for u in lst})

# Интеграция с Google Sheets и Drive
from google.oauth2.service_account import Credentials
import gspread
from gspread_dataframe import set_with_dataframe
from googleapiclient.discovery import build

creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
if not creds_path or not os.path.exists(creds_path):
    logger.error("credentials.json не найден.")
    sys.exit(1)
creds = Credentials.from_service_account_file(creds_path, scopes=[
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
])
gc = gspread.authorize(creds)
drive = build('drive','v3',credentials=creds)
sheets = build('sheets','v4',credentials=creds)

# Создание общей папки отчётов при multiple users
parent_folder = args.folder_id or "1J_10hg5Lu1kOk69Oj9Yo2E_xSZefaSSu"
if not args.assignee:
    subfolder = f"Reports_{start_date.date()}_{end_date.date()}"
    meta = {'name': subfolder, 'mimeType': 'application/vnd.google-apps.folder', 'parents':[parent_folder]}
    sf = drive.files().create(body=meta, fields='id').execute()
    parent_folder = sf.get('id')
    logger.info(f"Создана папка для отчётов: {subfolder} (ID={parent_folder})")

# Генерация одного Spreadsheet на пользователя
for user in users:
    df_user = all_tasks[all_tasks['assignees'].apply(lambda lst: user in lst)]
    if df_user.empty:
        logger.info(f"Нет задач для {user}, пропускаем")
        continue
    title = f"GitHub Report {user} {start_date.date()}_{end_date.date()}"
    sh = gc.create(title)
    sid = sh.id
    drive.files().update(fileId=sid, addParents=parent_folder, removeParents='root', fields='id,parents').execute()
    logger.info(f"Создан Spreadsheet '{title}' (ID={sid})")
    try: sh.del_worksheet(sh.sheet1)
    except: pass
    # Добавляем листы проектов
    sheet_ids, data_map = {}, {}
    for proj_name, grp in df_user.groupby('project'):
        ws = sh.add_worksheet(title=proj_name, rows=grp.shape[0]+5, cols=grp.shape[1]-1)
        set_with_dataframe(ws, grp.drop(columns=['project'], errors='ignore'))
        sheet_ids[proj_name] = ws._properties['sheetId']
        data_map[proj_name] = grp
    # Summary лист
    ws_sum = sh.add_worksheet(title='Summary', rows=df_user.shape[0]+5, cols=df_user.shape[1])
    set_with_dataframe(ws_sum, df_user)
    sheet_ids['Summary'] = ws_sum._properties['sheetId']
    data_map['Summary'] = df_user
    # Добавляем диаграммы
    chart_reqs = []
    for name, df in data_map.items():
        sid_sheet = sheet_ids[name]
        row = df.shape[0] + 2
        ws = sh.worksheet(name)
        ws.update(range_name=f"A{row}:B{row}", values=[['estimate','actual']])
        ws.update(range_name=f"A{row+1}:B{row+1}", values=[[df['estimate'].sum(), df['actual'].sum()]])
        chart_reqs.append({
            'addChart':{
                'chart':{
                    'spec':{
                        'title':'Сумма часов',
                        'basicChart':{
                            'chartType':'COLUMN','legendPosition':'BOTTOM_LEGEND',
                            'domains':[{'domain':{'sourceRange':{'sources':[{'sheetId':sid_sheet,'startRowIndex':row-1,'endRowIndex':row,'startColumnIndex':0,'endColumnIndex':2}]}}}],
                            'series':[{'series':{'sourceRange':{'sources':[{'sheetId':sid_sheet,'startRowIndex':row,'endRowIndex':row+1,'startColumnIndex':0,'endColumnIndex':2}]}},'targetAxis':'LEFT_AXIS'}]
                        }
                    },
                    'position':{'overlayPosition':{'anchorCell':{'sheetId':sid_sheet,'rowIndex':row,'columnIndex':0}}}
                }
            }
        })
    if chart_reqs:
        sheets.spreadsheets().batchUpdate(spreadsheetId=sid,body={'requests':chart_reqs}).execute()
        logger.info(f"Диаграммы добавлены для {user}")
