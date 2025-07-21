#!/usr/bin/env python3
import os
import sys
import re
import logging
import requests
import pandas as pd
import argparse
import tempfile
import zipfile
from dotenv import load_dotenv
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

# --- Логирование ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)

# --- CLI аргументы ---
parser = argparse.ArgumentParser(
    description="Генерация zip-архива отчётов по GitHub Projects для одного/всех assignee"
)
parser.add_argument("-a", "--assignee", help="GitHub-логин assignee")
parser.add_argument("-s", "--start",    required=True, help="Дата начала YYYY-MM-DD")
parser.add_argument("-e", "--end",      required=True, help="Дата окончания YYYY-MM-DD")
parser.add_argument("-o", "--output",   help="Имя выходного ZIP-файла (по умолчанию reports_<start>_<end>.zip)")
args = parser.parse_args()

# --- Валидация дат ---
try:
    start_date = datetime.fromisoformat(args.start)
    end_date   = datetime.fromisoformat(args.end)
except ValueError:
    parser.error("Даты должны быть в формате YYYY-MM-DD")

# --- Имя выходного архива ---
output_name = args.output or f"reports_{start_date.date()}_{end_date.date()}.zip"

# --- Загрузка GitHub токена ---
load_dotenv()
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
if not GITHUB_TOKEN:
    logger.error("Не найден GITHUB_TOKEN в окружении.")
    sys.exit(1)
headers = {"Authorization": f"Bearer {GITHUB_TOKEN}"}

# --- Настройки проекта ---
ORG = "Maxinum"
PAGE_SIZE = 100
ACTUAL_NAMES   = {"actual","actual hours","acutal hours"}
ESTIMATE_NAMES = {"estimate","planned hours","hours","estimate hours","estimates"}

# --- GraphQL-запросы ---
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
          content{...on Issue{
            number,title,
            repository{name},
            assignees(first:10){nodes{login}},
            url,createdAt
          }}
          fieldValues(first:20){nodes{
            ...on ProjectV2ItemFieldNumberValue{
              field{...on ProjectV2FieldCommon{name}},number
            }
          }}
        }
      }
    }
  }
}'''

def graphql_paginate(query, variables, path):
    items, cursor = [], None
    while True:
        vars_ = {**variables, "first": PAGE_SIZE, "after": cursor}
        r = requests.post(
            "https://api.github.com/graphql",
            json={"query": query, "variables": vars_},
            headers=headers, timeout=10
        )
        r.raise_for_status()
        data = r.json().get("data", {})
        for part in path.split('.'):
            data = data.get(part, {})
        nodes = data.get("nodes", [])
        items.extend(nodes)
        pi = data.get("pageInfo", {})
        if not pi.get("hasNextPage"):
            break
        cursor = pi.get("endCursor")
    return items

# --- Сбор проектов после start_date ---
all_projects = graphql_paginate(proj_query, {"org": ORG}, "organization.projectsV2")
projects = []
for p in all_projects:
    try:
        dt = datetime.fromisoformat(p["updatedAt"].rstrip("Z"))
        if dt >= start_date:
            projects.append(p)
    except Exception:
        logger.warning(f"Invalid updatedAt в проекте #{p.get('number')}")
if not projects:
    logger.info(f"Нет проектов после {start_date.date()}. Выход.")
    sys.exit(0)
logger.info(f"Найдено {len(projects)} проектов после {start_date.date()}")

# --- Сбор задач из каждого проекта ---
def fetch_proj_all(proj):
    num = proj["number"]
    title_safe = re.sub(r"[\\\/\?\*\[\]:]", "_", proj.get("title",""))[:25]
    rows, cursor = [], None
    while True:
        r = requests.post(
            "https://api.github.com/graphql",
            json={
                "query": items_query,
                "variables": {"org": ORG, "projNum": num, "first": PAGE_SIZE, "after": cursor}
            },
            headers=headers, timeout=10
        )
        r.raise_for_status()
        items = (r.json().get("data",{})
                    .get("organization",{})
                    .get("projectV2",{})
                    .get("items",{}))
        for it in items.get("nodes", []):
            issue = it.get("content")
            if not issue:
                continue
            try:
                created = datetime.fromisoformat(issue["createdAt"].rstrip("Z"))
            except Exception:
                continue
            if not (start_date <= created <= end_date):
                continue
            assignees = [a["login"] for a in issue.get("assignees",{}).get("nodes",[])]
            actual = estimate = 0
            for fv in it.get("fieldValues",{}).get("nodes",[]):
                fld, val = fv.get("field"), fv.get("number")
                if fld and val is not None:
                    nm = fld["name"].strip().lower()
                    if nm in ACTUAL_NAMES: actual = val
                    if nm in ESTIMATE_NAMES: estimate = val
            rows.append({
                "project":   f"{num}_{title_safe}",
                "number":    issue["number"],
                "title":     issue["title"],
                "repo":      issue["repository"]["name"],
                "url":       issue["url"],
                "createdAt": created,
                "assignees": assignees,
                "actual":    actual,
                "estimate":  estimate
            })
        pi = items.get("pageInfo",{})
        if not pi.get("hasNextPage"):
            break
        cursor = pi.get("endCursor")
    return pd.DataFrame(rows)

# --- Параллельный сбор всех задач ---
with ThreadPoolExecutor(max_workers=4) as executor:
    dfs = list(executor.map(fetch_proj_all, projects))
all_tasks = pd.concat(dfs, ignore_index=True)

# --- Определяем assignees ---
users = [args.assignee] if args.assignee else sorted(
    {u for lst in all_tasks["assignees"] for u in lst}
)
if not users:
    logger.info("Нет assignee для отчёта. Выход.")
    sys.exit(0)

# --- Генерация Excel + диаграммы и упаковка в ZIP ---
with tempfile.TemporaryDirectory() as tmpdir:
    generated = []
    for user in users:
        df_user = all_tasks[all_tasks["assignees"].apply(lambda lst: user in lst)]
        if df_user.empty:
            logger.info(f"Нет задач для {user}, пропускаем")
            continue

        fname = f"GitHub_Report_{user}_{start_date.date()}_{end_date.date()}.xlsx"
        fpath = os.path.join(tmpdir, fname)
        logger.info(f"Создаём {fname} ...")

        # используем xlsxwriter, чтобы рисовать графики
        with pd.ExcelWriter(fpath, engine="xlsxwriter") as writer:
            workbook  = writer.book

            # 1) Листы по проектам
            for proj, grp in df_user.groupby("project"):
                grp2 = grp.drop(columns=["project"], errors="ignore")
                sheet = proj[:31]  # имя листа ≤31 символ
                grp2.to_excel(writer, sheet_name=sheet, index=False)
                ws = writer.sheets[sheet]

                # вычисляем строки/столбцы для summary
                nrows     = len(grp2)
                label_row = nrows + 1    # 0-based: сразу под данными
                data_row  = nrows + 2
                cols      = list(grp2.columns)
                if "estimate" in cols and "actual" in cols:
                    i_est = cols.index("estimate")
                    i_act = cols.index("actual")

                    # пишем заголовки и суммы
                    ws.write(label_row, i_est, "estimate")
                    ws.write(label_row, i_act, "actual")
                    ws.write(data_row,  i_est, grp2["estimate"].sum())
                    ws.write(data_row,  i_act, grp2["actual"].sum())

                    # строим график
                    chart = workbook.add_chart({"type": "column"})
                    chart.add_series({
                        "name":       "Estimate",
                        "categories": [sheet, label_row, i_est, label_row, i_act],
                        "values":     [sheet, data_row, i_est,   data_row, i_est],
                    })
                    chart.add_series({
                        "name":       "Actual",
                        "categories": [sheet, label_row, i_est, label_row, i_act],
                        "values":     [sheet, data_row, i_act,   data_row, i_act],
                    })
                    chart.set_title({"name": "Сумма часов"})
                    chart.set_legend({"position": "bottom"})
                    ws.insert_chart(data_row + 2, 0, chart, {"x_scale": 1.5, "y_scale": 1.5})

            # 2) Summary-лист
            df_user.to_excel(writer, sheet_name="Summary", index=False)
            ws_sum = writer.sheets["Summary"]
            nrows  = len(df_user)
            lbl_r  = nrows + 1
            dat_r  = nrows + 2
            cols   = list(df_user.columns)
            if "estimate" in cols and "actual" in cols:
                ie = cols.index("estimate")
                ia = cols.index("actual")
                ws_sum.write(lbl_r, ie, "estimate")
                ws_sum.write(lbl_r, ia, "actual")
                ws_sum.write(dat_r, ie, df_user["estimate"].sum())
                ws_sum.write(dat_r, ia, df_user["actual"].sum())

                chart = workbook.add_chart({"type": "column"})
                chart.add_series({
                    "name":       "Estimate",
                    "categories": ["Summary", lbl_r, ie, lbl_r, ia],
                    "values":     ["Summary", dat_r, ie, dat_r, ie],
                })
                chart.add_series({
                    "name":       "Actual",
                    "categories": ["Summary", lbl_r, ie, lbl_r, ia],
                    "values":     ["Summary", dat_r, ia, dat_r, ia],
                })
                chart.set_title({"name": "Сумма часов (Summary)"})
                chart.set_legend({"position": "bottom"})
                ws_sum.insert_chart(dat_r + 2, 0, chart, {"x_scale": 1.5, "y_scale": 1.5})

        generated.append((fpath, fname))

    if not generated:
        logger.error("Ни одного отчёта не было сгенерировано. Выход.")
        sys.exit(1)

    # пакуем всё в ZIP
    zip_path = os.path.abspath(output_name)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for full, arc in generated:
            zf.write(full, arcname=arc)
    logger.info(f"ZIP-архив готов: {zip_path}")

print(f"✅ Готово! Ваш архив: {zip_path}")
