cd d:\github\jd_scrapy
$env:PLAYWRIGHT_BROWSERS_PATH="d:\github\jd_scrapy\.playwright"

& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="d:\github\jd_scrapy\.user_data"

# 方案A（推荐）：BOSS 用 CDP（更稳），智联照常；全量按 conf/config.yml 的 cities/keywords/pages_per_city 跑
python script\crawl_salary.py --config conf\config.yml --platform both --boss-cdp "http://127.0.0.1:9222" --out "output\salary_full.xlsx" --verbose true