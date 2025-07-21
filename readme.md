# GitHub Projects Export

Скрипт `script.py` позволяет автоматически генерировать отчёты по задачам (issues) из GitHub Projects, сохраняя данные в Excel. Отчёты создаются для одного или всех assignee внутри заданного временного интервала.

## ✨ Возможности

* Постраничная выгрузка проектов и задач через GitHub GraphQL API
* Фильтрация по дате создания задач и по assignee
* Генерация отдельного Excel для каждого assignee (или одного для указанного пользователя)
* Вставка листов с данными и листа `Summary` с общей сводкой
* Построение диаграмм (столбчатая диаграмма «Сумма часов») в каждом листе

## 📋 Требования

* Python 3.9+
* GitHub Personal Access Token с правом `read:org` и `repo`

## 🔧 Установка

1. Клонируйте репозиторий:

   ```bash
   git clone https://github.com/beko-o/github-project-issues-parser.git
   cd github-project-issues-parser
   ```

2. Создайте и активируйте виртуальное окружение:

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Linux/macOS
   .\.venv\Scripts\activate   # Windows
   ```

3. Установите зависимости:

   ```bash
   pip install -r requirements.txt
   ```

4. Создайте файл `.env` на основе примера `.env.example`:

   ```ini
   GITHUB_TOKEN=your_github_pat_here
   ```

## 🚀 Использование

### Запуск для одного assignee

```bash
python script.py \
  --assignee username \
  --start 2025-05-01 \
  --end   2025-05-31 \
```

* `--assignee` (`-a`) — GitHub логин (опционально).
* `--start` (`-s`) — дата начала (обязательно).
* `--end` (`-e`) — дата окончания (обязательно).

### Запуск для всех assignee

```bash
python script.py -s 2025-05-01 -e 2025-05-31
```

## 📈 Структура отчёта

* **Каждый Excel** содержит:

  * Листы для каждого проекта, где assignee имел задачи.
  * Лист `Summary` со сводным списком всех задач.
  * Под каждым листом — диаграмма с суммарным временем (estimate vs actual).

## 🐞 Отладка и ограничения

* **Rate limits**: при большом числе запросов к GitHub может превышаться квота. Рекомендуется:

  * Уменьшить `max_workers` до `3`.

* **Ошибки авторизации**: проверьте `GITHUB_TOKEN`

## 🚧 Лицензия

MIT © Maxinum Consulting
