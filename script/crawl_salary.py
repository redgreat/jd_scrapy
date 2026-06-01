import argparse
import dataclasses
import datetime as dt
import hashlib
import json
import os
import random
import re
import statistics
import time
import urllib.parse
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pymysql
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright
import yaml


@dataclass(frozen=True)
class SalaryInfo:
    min_month: Optional[float]
    max_month: Optional[float]
    avg_month: Optional[float]
    pay_months: Optional[int]
    annualized_avg: Optional[float]
    unit: Optional[str]


@dataclass(frozen=True)
class JobRecord:
    platform: str
    city: str
    keyword: str
    page_no: int
    job_title: Optional[str]
    company_name: Optional[str]
    company_industry: Optional[str]
    company_industry_mapped: Optional[str]
    salary_text: Optional[str]
    salary_min_month: Optional[float]
    salary_max_month: Optional[float]
    salary_avg_month: Optional[float]
    pay_months: Optional[int]
    salary_annualized_avg: Optional[float]
    location_text: Optional[str]
    education: Optional[str]
    experience_text: Optional[str]
    experience_year_min: Optional[float]
    experience_year_max: Optional[float]
    skills_text: Optional[str]
    job_url: Optional[str]
    crawled_at: dt.datetime
    raw: Dict[str, Any]

    def job_hash(self) -> str:
        """生成用于去重的稳定哈希，优先使用链接，否则使用关键字段拼接。"""
        base = (self.job_url or "").strip()
        if base:
            key = f"{self.platform}|{base}"
        else:
            key = "|".join(
                [
                    self.platform,
                    self.city,
                    self.keyword,
                    (self.job_title or "").strip(),
                    (self.company_name or "").strip(),
                    (self.salary_text or "").strip(),
                ]
            )
        return hashlib.sha1(key.encode("utf-8")).hexdigest()


def parse_salary_text(text: Optional[str]) -> SalaryInfo:
    """解析常见薪资字符串，统一输出为“人民币/月”的区间与平均值，并计算年化平均。"""
    if not text:
        return SalaryInfo(None, None, None, None, None, None)

    s = str(text).strip()
    if not s:
        return SalaryInfo(None, None, None, None, None, None)

    s = s.replace(" ", "")
    pay_months = 12

    m_pay = re.search(r"·(\d{2})薪", s)
    if m_pay:
        try:
            pay_months = int(m_pay.group(1))
        except ValueError:
            pay_months = 12

    s_main = s.split("·", 1)[0]

    if any(x in s_main for x in ["面议", "暂无", "N/A"]):
        return SalaryInfo(None, None, None, pay_months, None, None)

    def to_float(num: str) -> Optional[float]:
        try:
            return float(num)
        except ValueError:
            return None

    def normalize_range(a: Optional[float], b: Optional[float]) -> Tuple[Optional[float], Optional[float]]:
        if a is None and b is None:
            return None, None
        if a is None:
            return b, b
        if b is None:
            return a, a
        return (a, b) if a <= b else (b, a)

    unit = None
    min_month = None
    max_month = None

    m_k = re.match(r"(?P<a>\d+(\.\d+)?)-(?P<b>\d+(\.\d+)?)(?P<u>[Kk])$", s_main)
    if m_k:
        unit = "K/月"
        a = to_float(m_k.group("a"))
        b = to_float(m_k.group("b"))
        if a is not None:
            a *= 1000.0
        if b is not None:
            b *= 1000.0
        min_month, max_month = normalize_range(a, b)

    if min_month is None and max_month is None:
        m_wan = re.match(r"(?P<a>\d+(\.\d+)?)-(?P<b>\d+(\.\d+)?)(?P<u>万)$", s_main)
        if m_wan:
            unit = "万/月"
            a = to_float(m_wan.group("a"))
            b = to_float(m_wan.group("b"))
            if a is not None:
                a *= 10000.0
            if b is not None:
                b *= 10000.0
            min_month, max_month = normalize_range(a, b)

    if min_month is None and max_month is None:
        m_day = re.match(r"(?P<a>\d+(\.\d+)?)-(?P<b>\d+(\.\d+)?)(?P<u>元/天)$", s_main)
        if m_day:
            unit = "元/天"
            a = to_float(m_day.group("a"))
            b = to_float(m_day.group("b"))
            if a is not None:
                a *= 21.75
            if b is not None:
                b *= 21.75
            min_month, max_month = normalize_range(a, b)

    if min_month is None and max_month is None:
        m_year = re.match(r"(?P<a>\d+(\.\d+)?)-(?P<b>\d+(\.\d+)?)(?P<u>万/年)$", s_main)
        if m_year:
            unit = "万/年"
            a = to_float(m_year.group("a"))
            b = to_float(m_year.group("b"))
            if a is not None:
                a = a * 10000.0 / 12.0
            if b is not None:
                b = b * 10000.0 / 12.0
            min_month, max_month = normalize_range(a, b)

    if min_month is None and max_month is None:
        m_rmb_range = re.match(r"(?P<a>\d+(\.\d+)?)-(?P<b>\d+(\.\d+)?)(?P<u>元(/月)?)$", s_main)
        if m_rmb_range:
            unit = "元/月"
            a = to_float(m_rmb_range.group("a"))
            b = to_float(m_rmb_range.group("b"))
            min_month, max_month = normalize_range(a, b)

    if min_month is None and max_month is None:
        m_single_k = re.match(r"(?P<a>\d+(\.\d+)?)(?P<u>[Kk])$", s_main)
        if m_single_k:
            unit = "K/月"
            a = to_float(m_single_k.group("a"))
            if a is not None:
                a *= 1000.0
            min_month, max_month = normalize_range(a, a)

    if min_month is None and max_month is None:
        m_single_w = re.match(r"(?P<a>\d+(\.\d+)?)(?P<u>万)$", s_main)
        if m_single_w:
            unit = "万/月"
            a = to_float(m_single_w.group("a"))
            if a is not None:
                a *= 10000.0
            min_month, max_month = normalize_range(a, a)

    if min_month is None and max_month is None:
        m_single_rmb = re.match(r"(?P<a>\d+(\.\d+)?)(?P<u>元(/月)?)$", s_main)
        if m_single_rmb:
            unit = "元/月"
            a = to_float(m_single_rmb.group("a"))
            min_month, max_month = normalize_range(a, a)

    avg_month = None
    if min_month is not None and max_month is not None:
        avg_month = (min_month + max_month) / 2.0

    annualized_avg = None
    if avg_month is not None and pay_months:
        annualized_avg = avg_month * float(pay_months)

    return SalaryInfo(min_month, max_month, avg_month, pay_months, annualized_avg, unit)


def map_industry(industry: Optional[str]) -> Optional[str]:
    """按业务归类映射行业文本，未命中规则时返回None。"""
    if not industry:
        return None
    s = str(industry).strip()
    if not s:
        return None

    if any(k in s for k in ["金融外包"]):
        return "汽车服务"
    if "金融" in s:
        return "金融科技"
    if any(k in s for k in ["计算机服务", "系统", "数据服务", "汽车后市场", "互联网", "IT", "电子"]):
        return "物联网/车联网"
    if any(k in s for k in ["软件", "信息技术", "汽车电子", "物流科技", "汽车", "摩托车"]):
        return "汽车服务"
    return None


def parse_experience_years(text: Optional[str]) -> Tuple[Optional[float], Optional[float]]:
    """从经验要求文本中提取年限区间。"""
    if not text:
        return None, None
    s = str(text).strip()
    if not s:
        return None, None
    if any(k in s for k in ["经验不限", "不限", "无经验"]):
        return None, None

    s = s.replace("年以上", "+年").replace("年及以上", "+年")
    m_range = re.search(r"(\d+)\s*-\s*(\d+)\s*年", s)
    if m_range:
        return float(m_range.group(1)), float(m_range.group(2))
    m_plus = re.search(r"(\d+)\s*\\+\\s*年", s)
    if m_plus:
        v = float(m_plus.group(1))
        return v, None
    m_single = re.search(r"(\\d+)\\s*年", s)
    if m_single:
        v = float(m_single.group(1))
        return v, v
    return None, None


def normalize_education(text: Optional[str]) -> Optional[str]:
    """规范化学历字段，无法识别时返回原文或None。"""
    if not text:
        return None
    s = str(text).strip()
    if not s:
        return None
    for k in ["博士", "硕士", "本科", "大专", "中专", "高中", "初中"]:
        if k in s:
            return k
    if "不限" in s:
        return "不限"
    return s


def pick_company_industry(tags: Sequence[str]) -> Optional[str]:
    """从公司标签中推断行业。"""
    if not tags:
        return None
    cleaned = [str(t).strip() for t in tags if str(t).strip()]
    if not cleaned:
        return None
    ignore = {
        "民营",
        "国企",
        "外商独资",
        "合资",
        "上市公司",
        "股份制",
        "事业单位",
        "社会团体",
        "港澳台公司",
        "外企代表处",
    }
    cand: List[str] = []
    for t in cleaned:
        if t in ignore:
            continue
        if re.search(r"\\d+\\s*-\\s*\\d+\\s*人", t) or re.search(r"\\d+\\s*人", t):
            continue
        if "人以上" in t or "人以下" in t:
            continue
        if re.search(r"\\d+\\s*-\\s*\\d+\\s*万", t):
            continue
        cand.append(t)
    if not cand:
        return None
    return cand[-1]

def expand_keywords(keywords: Sequence[str]) -> List[str]:
    """将配置中的岗位关键词做规则展开（例如初/中/高级、经理/主管/专员），并去重保持顺序。"""
    expanded: List[str] = []
    seen = set()
    for kw in keywords:
        for item in expand_one_keyword(str(kw).strip()):
            s = item.strip()
            if not s or s in seen:
                continue
            seen.add(s)
            expanded.append(s)
    return expanded


def expand_one_keyword(keyword: str) -> List[str]:
    """展开单个岗位关键词，支持括号级别与&/、分隔的后缀角色组合。"""
    s = str(keyword).strip()
    if not s:
        return []

    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("，", ",").replace("、", ",").replace("/", ",")

    role_suffixes = ["专员", "主管", "经理", "总监"]
    level_prefixes = ["高级", "中级", "初级", "资深"]

    m = re.search(r"\(([^)]{1,50})\)", s)
    if m:
        inside = m.group(1)
        parts = [p.strip() for p in inside.split(",") if p.strip()]
        base = (s[: m.start()] + s[m.end() :]).strip()
        base = base.strip(" -_")
        out: List[str] = []
        for p in parts:
            if p in level_prefixes:
                out.append(f"{p}{base}")
            else:
                out.append(f"{base}{p}")
        return out or [s]

    tokens = [t.strip() for t in re.split(r"[&,+]", s) if t.strip()]
    if len(tokens) <= 1:
        return [s]

    def is_suffix_only(t: str) -> bool:
        return t in role_suffixes

    def strip_suffix(text: str) -> str:
        for suf in role_suffixes:
            if text.endswith(suf):
                return text[: -len(suf)]
        return text

    out: List[str] = []
    last_full: Optional[str] = None
    for t in tokens:
        if is_suffix_only(t) and last_full:
            base = strip_suffix(last_full)
            out.append(base + t)
        else:
            out.append(t)
            last_full = t
    return out or [s]


def classify_role(text: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """从岗位标题中识别统一岗位与级别（初/中/高/资深/专员/主管/经理/总监等）。"""
    if not text:
        return None, None
    s = str(text).strip()
    if not s:
        return None, None

    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", "", s)

    level_prefixes = ["资深", "高级", "中级", "初级"]
    role_suffixes = ["总监", "经理", "主管", "专员", "助理"]

    level = None
    for p in level_prefixes:
        if p in s:
            level = p
            s = s.replace(p, "")
            break

    suffix_level = None
    for suf in role_suffixes:
        if s.endswith(suf):
            suffix_level = suf
            s = s[: -len(suf)]
            break

    s = re.sub(r"\([^)]*\)", "", s).strip("-_ ")
    base = s or None
    final_level = suffix_level or level
    return base, final_level


def filter_records_by_keyword(records: Sequence[JobRecord], keyword: str, city: Optional[str] = None) -> List[JobRecord]:
    """按关键词推断约束对结果做轻量过滤，降低模糊搜索带来的噪音。"""
    base, level = classify_role(keyword)
    city_text = str(city).strip() if city else ""

    base = base or ""
    core_base = base
    for suf in ["工程师", "专员", "主管", "经理", "总监", "助理", "顾问"]:
        if core_base.endswith(suf):
            core_base = core_base[: -len(suf)]
            break
    core_base = core_base.strip()

    out: List[JobRecord] = []
    for r in records:
        title = (r.job_title or "").strip()
        if not title:
            continue
        if city_text:
            loc = (r.location_text or "").strip()
            if loc and city_text not in loc:
                continue
        ok = True
        if core_base and len(core_base) >= 2 and core_base not in title:
            ok = False
        if ok and level and level not in title:
            ok = False
        if ok:
            out.append(r)
    return out


def build_stealth_context_kwargs() -> Dict[str, Any]:
    """生成较温和的浏览器上下文参数，降低被识别为自动化的概率。"""
    viewport_candidates = [(1920, 1080), (1366, 768), (1440, 900), (1536, 864)]
    w, h = random.choice(viewport_candidates)
    ua = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    )
    return {
        "locale": "zh-CN",
        "timezone_id": "Asia/Shanghai",
        "viewport": {"width": w, "height": h},
        "user_agent": ua,
    }


def add_stealth_init_script(context) -> None:
    """向页面注入初始化脚本，尽量减少自动化特征。"""
    context.add_init_script(
        """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(navigator, 'languages', {get: () => ['zh-CN', 'zh']});
        Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
        window.chrome = window.chrome || { runtime: {} };
        """
    )


def normalize_url(base: str, href: Optional[str]) -> Optional[str]:
    """将相对链接或协议相对链接规范化为绝对URL。"""
    if not href:
        return None
    u = href.strip()
    if not u:
        return None
    if u.startswith("//"):
        return "https:" + u
    if u.startswith("http://") or u.startswith("https://"):
        return u
    return urllib.parse.urljoin(base, u)


def safe_text(el) -> Optional[str]:
    """从元素中提取文本，失败时返回None。"""
    try:
        t = el.inner_text()
        t = t.strip()
        return t or None
    except Exception:
        return None


def query_first_text(root, selectors: Sequence[str]) -> Optional[str]:
    """按候选选择器顺序取第一个命中的文本。"""
    for sel in selectors:
        try:
            el = root.query_selector(sel)
            if el:
                t = safe_text(el)
                if t:
                    return t
        except Exception:
            continue
    return None


def detect_challenge(page) -> Optional[str]:
    """检测常见安全校验/验证码页面特征，命中时返回原因。"""
    try:
        title = (page.title() or "").strip()
    except Exception:
        title = ""
    try:
        content = page.content()
    except Exception:
        content = ""

    for kw in ["安全验证", "滑动验证", "验证码", "人机识别", "访问受限", "验证中心"]:
        if kw in title or kw in content:
            return kw
    return None


def boss_normalize_city(city_code_or_name: str) -> str:
    """将常见城市中文名映射为BOSS直聘city参数，未命中时原样返回。"""
    s = str(city_code_or_name).strip()
    if not s:
        return s
    if re.fullmatch(r"\d{6,12}", s):
        return s
    mapping = {
        "北京": "101010100",
        "上海": "101020100",
        "广州": "101280100",
        "深圳": "101280600",
        "杭州": "101210100",
        "南京": "101190100",
        "苏州": "101190400",
        "成都": "101270100",
        "重庆": "101040100",
        "武汉": "101200100",
        "西安": "101110100",
        "天津": "101030100",
        "郑州": "101180100",
        "长沙": "101250100",
        "合肥": "101220100",
        "厦门": "101230200",
        "福州": "101230100",
        "济南": "101120100",
        "青岛": "101120200",
        "沈阳": "101070100",
        "大连": "101070200",
        "昆明": "101290100",
        "南宁": "101300100",
        "南昌": "101240100",
        "石家庄": "101090100",
        "太原": "101100100",
        "无锡": "101190200",
        "宁波": "101210400",
        "东莞": "101281600",
        "佛山": "101280800",
    }
    return mapping.get(s, s)


_ZL_CITY_CODE_CACHE: Optional[Dict[str, str]] = None


def zhilian_normalize_city(city_code_or_name: str) -> str:
    """将城市中文名映射为智联城市code（如703），未命中时原样返回。"""
    s = str(city_code_or_name).strip()
    if not s:
        return s
    if re.fullmatch(r"\\d{2,6}", s):
        return s
    global _ZL_CITY_CODE_CACHE
    if _ZL_CITY_CODE_CACHE is None:
        _ZL_CITY_CODE_CACHE = load_zhilian_city_code_map()
    return _ZL_CITY_CODE_CACHE.get(s, s)


def load_zhilian_city_code_map() -> Dict[str, str]:
    """从智联基础数据接口加载城市code映射表。"""
    import urllib.request

    url = "https://fe-api.zhaopin.com/c/i/search/base/data"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://sou.zhaopin.com/"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read())
    root = data.get("data", {}) if isinstance(data, dict) else {}
    out: Dict[str, str] = {}
    for key in ["hotCity", "allCity"]:
        items = root.get(key)
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            code = item.get("code")
            name = item.get("name")
            if code and name:
                out[str(name)] = str(code)
            sub = item.get("sublist")
            if isinstance(sub, list):
                for sub_item in sub:
                    if not isinstance(sub_item, dict):
                        continue
                    code2 = sub_item.get("code")
                    name2 = sub_item.get("name")
                    if code2 and name2:
                        out[str(name2)] = str(code2)
    return out


class BossCrawler:
    def __init__(
        self, headless: bool, proxy: Optional[str], slow_mo_ms: int, min_sleep: float, max_sleep: float, verbose: bool
    ):
        self.headless = headless
        self.proxy = proxy
        self.slow_mo_ms = slow_mo_ms
        self.min_sleep = min_sleep
        self.max_sleep = max_sleep
        self.verbose = verbose

    def build_url(self, keyword: str, city_code_or_name: str, page_no: int) -> str:
        """构造BOSS直聘搜索URL，city建议传城市code（如101010100）。"""
        q = urllib.parse.quote(keyword)
        city = urllib.parse.quote(boss_normalize_city(city_code_or_name))
        return f"https://www.zhipin.com/web/geek/job?query={q}&city={city}&page={page_no}"

    def extract_jobs(self, page, city: str, keyword: str, page_no: int) -> List[JobRecord]:
        """从BOSS结果页抽取岗位记录，尽量做选择器兜底。"""
        cards = []
        for sel in [
            "div.job-card-wrapper",
            "li.job-card-wrapper",
            "div.job-card",
            "div.search-job-result div.job-card-wrapper",
        ]:
            try:
                cards = page.query_selector_all(sel)
                if cards:
                    break
            except Exception:
                continue

        records: List[JobRecord] = []
        for card in cards:
            title = query_first_text(card, [".job-name", ".job-title", "a .job-name"])
            salary = query_first_text(card, [".salary", ".job-salary", "span.salary"])
            company = query_first_text(card, [".company-name", ".job-card-company-name", ".company-text"])
            href = None
            try:
                a = card.query_selector("a")
                if a:
                    href = a.get_attribute("href")
            except Exception:
                href = None
            url = normalize_url("https://www.zhipin.com", href)

            salary_info = parse_salary_text(salary)
            raw = {
                "query_keyword": keyword,
                "title": title,
                "salary": salary,
                "company": company,
            }
            records.append(
                JobRecord(
                    platform="boss",
                    city=city,
                    keyword=keyword,
                    page_no=page_no,
                    job_title=title,
                    company_name=company,
                    company_industry=None,
                    company_industry_mapped=None,
                    salary_text=salary,
                    salary_min_month=salary_info.min_month,
                    salary_max_month=salary_info.max_month,
                    salary_avg_month=salary_info.avg_month,
                    pay_months=salary_info.pay_months,
                    salary_annualized_avg=salary_info.annualized_avg,
                    location_text=None,
                    education=None,
                    experience_text=None,
                    experience_year_min=None,
                    experience_year_max=None,
                    skills_text=None,
                    job_url=url,
                    crawled_at=dt.datetime.now(),
                    raw=raw,
                )
            )
        return records

    def crawl(self, cities: Sequence[str], keyword: str, pages_per_city: int) -> List[JobRecord]:
        """执行BOSS直聘采集。"""
        results: List[JobRecord] = []
        with sync_playwright() as p:
            launch_kwargs: Dict[str, Any] = {"headless": self.headless, "slow_mo": self.slow_mo_ms}
            if self.proxy:
                launch_kwargs["proxy"] = {"server": self.proxy}

            browser = p.chromium.launch(**launch_kwargs)
            context = browser.new_context(**build_stealth_context_kwargs())
            add_stealth_init_script(context)
            page = context.new_page()

            for city in cities:
                for page_no in range(1, pages_per_city + 1):
                    url = self.build_url(keyword, city, page_no)
                    if self.verbose:
                        print(f"[boss] city={city} kw={keyword} page={page_no}/{pages_per_city} url={url}")
                    page.goto(url, wait_until="domcontentloaded", timeout=60000)
                    reason = detect_challenge(page)
                    if reason:
                        if self.headless:
                            raise RuntimeError(f"BOSS触发安全校验：{reason}，建议使用 --headless false 运行后手动过验证")
                        while True:
                            time.sleep(2)
                            reason2 = detect_challenge(page)
                            if not reason2:
                                break
                    page.wait_for_timeout(random.randint(800, 1600))
                    page_records = self.extract_jobs(page, city=city, keyword=keyword, page_no=page_no)
                    kept = filter_records_by_keyword(page_records, keyword=keyword, city=city)
                    if self.verbose:
                        print(f"[boss] extracted={len(page_records)} kept={len(kept)}")
                    results.extend(kept)
                    time.sleep(random.uniform(self.min_sleep, self.max_sleep))

            context.close()
            browser.close()
        return results


class ZhilianCrawler:
    def __init__(
        self, headless: bool, proxy: Optional[str], slow_mo_ms: int, min_sleep: float, max_sleep: float, verbose: bool
    ):
        self.headless = headless
        self.proxy = proxy
        self.slow_mo_ms = slow_mo_ms
        self.min_sleep = min_sleep
        self.max_sleep = max_sleep
        self.verbose = verbose

    def build_url(self, keyword: str, city_name: str, page_no: int) -> str:
        """构造智联招聘搜索URL，city建议使用城市code以确保筛选生效。"""
        q = urllib.parse.quote(keyword)
        c = urllib.parse.quote(zhilian_normalize_city(city_name))
        return f"https://sou.zhaopin.com/?jl={c}&kw={q}&p={page_no}"

    def extract_jobs(self, page, city: str, keyword: str, page_no: int) -> List[JobRecord]:
        """从智联结果页抽取岗位记录，尽量做选择器兜底。"""
        cards = []
        for sel in [
            "div.joblist-box__iteminfo",
            "div.joblist-box__item",
            "div.joblist-box__item-content",
            "div.joblist-box__itembox",
            "div.joblist div.item",
        ]:
            try:
                cards = page.query_selector_all(sel)
                if cards:
                    break
            except Exception:
                continue

        records: List[JobRecord] = []
        for card in cards:
            title = query_first_text(card, ["a.jobinfo__name", ".jobinfo__name", "a[href*='jobdetail']"])
            salary = query_first_text(card, [".jobinfo__salary", ".joblist-box__salary", ".salary", ".job-salary"])
            company = query_first_text(card, [".companyinfo__name", ".companyinfo__company", ".joblist-box__company"])
            location_text = query_first_text(
                card,
                [
                    ".jobinfo__other-info-location-image + span",
                    ".jobinfo__other-info-item span",
                ],
            )

            experience_text = None
            education_text = None
            try:
                other_items = card.query_selector_all(".jobinfo__other-info-item")
            except Exception:
                other_items = []
            for el in other_items:
                t = safe_text(el)
                if not t:
                    continue
                if "·" in t:
                    continue
                if any(k in t for k in ["博士", "硕士", "本科", "大专", "中专", "高中", "学历不限", "不限"]):
                    if education_text is None:
                        education_text = t
                    continue
                if "年" in t or "经验" in t or "不限" in t:
                    if experience_text is None:
                        experience_text = t
                    continue

            experience_year_min, experience_year_max = parse_experience_years(experience_text)
            education = normalize_education(education_text)

            skills: List[str] = []
            try:
                for el in card.query_selector_all(".jobinfo__tag .joblist-box__item-tag"):
                    t = safe_text(el)
                    if t:
                        skills.append(t)
            except Exception:
                skills = []
            skills_text = ";".join(dict.fromkeys(skills)) if skills else None

            company_tags: List[str] = []
            try:
                for el in card.query_selector_all(".companyinfo__tag .joblist-box__item-tag"):
                    t = safe_text(el)
                    if t:
                        company_tags.append(t)
            except Exception:
                company_tags = []
            company_industry = pick_company_industry(company_tags)
            company_industry_mapped = map_industry(company_industry)
            href = None
            try:
                a = card.query_selector("a.jobinfo__name, a[href*='jobdetail']")
                if a:
                    href = a.get_attribute("href")
            except Exception:
                href = None
            url = normalize_url("https://www.zhaopin.com", href)

            salary_info = parse_salary_text(salary)
            raw = {
                "query_keyword": keyword,
                "title": title,
                "salary": salary,
                "company": company,
                "location": location_text,
                "education": education,
                "experience": experience_text,
                "skills": skills,
                "company_tags": company_tags,
                "company_industry": company_industry,
                "company_industry_mapped": company_industry_mapped,
            }
            records.append(
                JobRecord(
                    platform="zhilian",
                    city=city,
                    keyword=keyword,
                    page_no=page_no,
                    job_title=title,
                    company_name=company,
                    company_industry=company_industry,
                    company_industry_mapped=company_industry_mapped,
                    salary_text=salary,
                    salary_min_month=salary_info.min_month,
                    salary_max_month=salary_info.max_month,
                    salary_avg_month=salary_info.avg_month,
                    pay_months=salary_info.pay_months,
                    salary_annualized_avg=salary_info.annualized_avg,
                    location_text=location_text,
                    education=education,
                    experience_text=experience_text,
                    experience_year_min=experience_year_min,
                    experience_year_max=experience_year_max,
                    skills_text=skills_text,
                    job_url=url,
                    crawled_at=dt.datetime.now(),
                    raw=raw,
                )
            )
        return records

    def crawl(self, cities: Sequence[str], keyword: str, pages_per_city: int) -> List[JobRecord]:
        """执行智联招聘采集。"""
        results: List[JobRecord] = []
        with sync_playwright() as p:
            launch_kwargs: Dict[str, Any] = {"headless": self.headless, "slow_mo": self.slow_mo_ms}
            if self.proxy:
                launch_kwargs["proxy"] = {"server": self.proxy}

            browser = p.chromium.launch(**launch_kwargs)
            context = browser.new_context(**build_stealth_context_kwargs())
            add_stealth_init_script(context)
            page = context.new_page()

            for city in cities:
                for page_no in range(1, pages_per_city + 1):
                    url = self.build_url(keyword, city, page_no)
                    if self.verbose:
                        print(f"[zhilian] city={city} kw={keyword} page={page_no}/{pages_per_city} url={url}")
                    page.goto(url, wait_until="domcontentloaded", timeout=60000)
                    reason = detect_challenge(page)
                    if reason and self.headless:
                        raise RuntimeError(f"智联触发安全校验：{reason}，建议使用 --headless false 运行后手动过验证")
                    try:
                        page.wait_for_selector("div.joblist-box__iteminfo, div.jobinfo__top", timeout=15000)
                    except PlaywrightTimeoutError:
                        pass
                    page.wait_for_timeout(random.randint(800, 1600))
                    page_records = self.extract_jobs(page, city=city, keyword=keyword, page_no=page_no)
                    kept = filter_records_by_keyword(page_records, keyword=keyword, city=city)
                    if self.verbose:
                        print(f"[zhilian] extracted={len(page_records)} kept={len(kept)}")
                    results.extend(kept)
                    time.sleep(random.uniform(self.min_sleep, self.max_sleep))

            context.close()
            browser.close()
        return results


def ensure_dir(path: str) -> None:
    """确保目录存在，不存在则创建。"""
    os.makedirs(path, exist_ok=True)


def write_excel(records: Sequence[JobRecord], out_path: str) -> str:
    """将明细与汇总写入Excel。"""
    wb = Workbook()
    ws_raw = wb.active
    ws_raw.title = "raw"

    raw_headers = [
        "平台",
        "城市",
        "岗位关键词",
        "统一岗位",
        "级别",
        "页码",
        "岗位名称",
        "公司",
        "行业(原始)",
        "行业(归类)",
        "工作地点",
        "学历",
        "工作经验",
        "经验下限(年)",
        "经验上限(年)",
        "技能要求",
        "薪资文本",
        "月薪下限(元)",
        "月薪上限(元)",
        "月薪均值(元)",
        "薪资月数",
        "年化均值(元)",
        "岗位链接",
        "抓取时间",
    ]
    ws_raw.append(raw_headers)
    for c in range(1, len(raw_headers) + 1):
        ws_raw.cell(row=1, column=c).font = Font(bold=True)

    for r in records:
        base_role, level = classify_role(r.job_title) if r.job_title else (None, None)
        if not base_role:
            base_role, level = classify_role(r.keyword)
        ws_raw.append(
            [
                r.platform,
                r.city,
                r.keyword,
                base_role,
                level,
                r.page_no,
                r.job_title,
                r.company_name,
                r.company_industry,
                r.company_industry_mapped,
                r.location_text,
                r.education,
                r.experience_text,
                r.experience_year_min,
                r.experience_year_max,
                r.skills_text,
                r.salary_text,
                r.salary_min_month,
                r.salary_max_month,
                r.salary_avg_month,
                r.pay_months,
                r.salary_annualized_avg,
                r.job_url,
                r.crawled_at.strftime("%Y-%m-%d %H:%M:%S"),
            ]
        )

    for idx, h in enumerate(raw_headers, start=1):
        max_len = len(h)
        for row in ws_raw.iter_rows(min_row=2, min_col=idx, max_col=idx, max_row=min(ws_raw.max_row, 300)):
            v = row[0].value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws_raw.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 60)

    ws_sum = wb.create_sheet("summary")
    sum_headers = ["平台", "城市", "统一岗位", "样本数", "月薪均值(元)", "月薪中位数(元)", "月薪最小(元)", "月薪最大(元)"]
    ws_sum.append(sum_headers)
    for c in range(1, len(sum_headers) + 1):
        ws_sum.cell(row=1, column=c).font = Font(bold=True)

    grouped: Dict[Tuple[str, str, str], List[float]] = {}
    for r in records:
        if r.salary_avg_month is None:
            continue
        base_role, _ = classify_role(r.job_title) if r.job_title else (None, None)
        if not base_role:
            base_role, _ = classify_role(r.keyword)
        if not base_role:
            continue
        key = (r.platform, r.city, base_role)
        grouped.setdefault(key, []).append(float(r.salary_avg_month))

    for (platform, city, base_role), values in sorted(grouped.items()):
        values_sorted = sorted(values)
        ws_sum.append(
            [
                platform,
                city,
                base_role,
                len(values_sorted),
                round(statistics.mean(values_sorted), 2) if values_sorted else None,
                round(statistics.median(values_sorted), 2) if values_sorted else None,
                round(values_sorted[0], 2) if values_sorted else None,
                round(values_sorted[-1], 2) if values_sorted else None,
            ]
        )

    for idx, h in enumerate(sum_headers, start=1):
        ws_sum.column_dimensions[get_column_letter(idx)].width = max(len(h) + 2, 16)

    ensure_dir(os.path.dirname(out_path) or ".")
    try:
        wb.save(out_path)
        return out_path
    except PermissionError:
        base, ext = os.path.splitext(out_path)
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = f"{base}_{ts}{ext or '.xlsx'}"
        wb.save(alt)
        return alt


def mysql_connect_from_env() -> Optional[pymysql.Connection]:
    """从环境变量读取MySQL连接信息并建立连接，缺失关键信息则返回None。"""
    host = os.getenv("MYSQL_HOST")
    user = os.getenv("MYSQL_USER")
    password = os.getenv("MYSQL_PASSWORD")
    database = os.getenv("MYSQL_DATABASE")
    port = int(os.getenv("MYSQL_PORT", "3306"))
    if not host or not user or password is None or not database:
        return None
    return pymysql.connect(
        host=host,
        user=user,
        password=password,
        database=database,
        port=port,
        charset="utf8mb4",
        autocommit=True,
    )


def mysql_connect_from_config(mysql_cfg: Dict[str, Any]) -> Optional[pymysql.Connection]:
    """从配置字典读取MySQL连接信息并建立连接，缺失关键信息则返回None。"""
    if not mysql_cfg:
        return None
    enabled = mysql_cfg.get("enabled", False)
    if not enabled:
        return None
    host = mysql_cfg.get("host")
    user = mysql_cfg.get("user")
    password = mysql_cfg.get("password")
    database = mysql_cfg.get("database")
    port = int(mysql_cfg.get("port", 3306))
    if not host or not user or password is None or not database:
        return None
    return pymysql.connect(
        host=str(host),
        user=str(user),
        password=str(password),
        database=str(database),
        port=port,
        charset="utf8mb4",
        autocommit=True,
    )


def load_yaml_config(path: str) -> Dict[str, Any]:
    """加载YAML配置文件，返回字典结构。"""
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        raise ValueError("config.yml 顶层必须是字典结构")
    return data



def mysql_ensure_table(conn: pymysql.Connection, table: str) -> None:
    """创建用于落库的明细表（不存在则创建），并补齐新增字段列。"""
    ddl = f"""
    CREATE TABLE IF NOT EXISTS `{table}` (
      `id` BIGINT NOT NULL AUTO_INCREMENT,
      `job_hash` CHAR(40) NOT NULL,
      `platform` VARCHAR(20) NOT NULL,
      `city` VARCHAR(50) NOT NULL,
      `keyword` VARCHAR(100) NOT NULL,
      `page_no` INT NOT NULL,
      `job_title` VARCHAR(200) NULL,
      `company_name` VARCHAR(200) NULL,
      `company_industry` VARCHAR(100) NULL,
      `company_industry_mapped` VARCHAR(100) NULL,
      `location_text` VARCHAR(100) NULL,
      `education` VARCHAR(50) NULL,
      `experience_text` VARCHAR(50) NULL,
      `experience_year_min` DECIMAL(6,2) NULL,
      `experience_year_max` DECIMAL(6,2) NULL,
      `skills_text` TEXT NULL,
      `salary_text` VARCHAR(100) NULL,
      `salary_min_month` DECIMAL(12,2) NULL,
      `salary_max_month` DECIMAL(12,2) NULL,
      `salary_avg_month` DECIMAL(12,2) NULL,
      `pay_months` INT NULL,
      `salary_annualized_avg` DECIMAL(14,2) NULL,
      `job_url` VARCHAR(500) NULL,
      `crawled_at` DATETIME NOT NULL,
      `raw_json` LONGTEXT NULL,
      PRIMARY KEY (`id`),
      UNIQUE KEY `uk_job_hash` (`job_hash`),
      KEY `idx_platform_city_kw` (`platform`, `city`, `keyword`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
    with conn.cursor() as cur:
        cur.execute(ddl)
    mysql_ensure_columns(conn, table=table)


def mysql_ensure_columns(conn: pymysql.Connection, table: str) -> None:
    """检查并补齐新增字段列，避免历史表结构缺列导致写入失败。"""
    desired = {
        "company_industry": "VARCHAR(100) NULL",
        "company_industry_mapped": "VARCHAR(100) NULL",
        "location_text": "VARCHAR(100) NULL",
        "education": "VARCHAR(50) NULL",
        "experience_text": "VARCHAR(50) NULL",
        "experience_year_min": "DECIMAL(6,2) NULL",
        "experience_year_max": "DECIMAL(6,2) NULL",
        "skills_text": "TEXT NULL",
    }
    sql_cols = """
    SELECT COLUMN_NAME
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = %s;
    """
    with conn.cursor() as cur:
        cur.execute(sql_cols, (table,))
        existing = {row[0] for row in cur.fetchall()}
        for col, typ in desired.items():
            if col in existing:
                continue
            cur.execute(f"ALTER TABLE `{table}` ADD COLUMN `{col}` {typ};")


def mysql_insert_records(conn: pymysql.Connection, table: str, records: Sequence[JobRecord]) -> int:
    """批量写入MySQL，使用job_hash去重。"""
    if not records:
        return 0
    sql = f"""
    INSERT INTO `{table}` (
      job_hash, platform, city, keyword, page_no,
      job_title, company_name, company_industry, company_industry_mapped,
      location_text, education, experience_text, experience_year_min, experience_year_max, skills_text,
      salary_text,
      salary_min_month, salary_max_month, salary_avg_month, pay_months, salary_annualized_avg,
      job_url, crawled_at, raw_json
    ) VALUES (
      %s, %s, %s, %s, %s,
      %s, %s, %s, %s,
      %s, %s, %s, %s, %s, %s,
      %s,
      %s, %s, %s, %s, %s,
      %s, %s, %s
    )
    ON DUPLICATE KEY UPDATE
      crawled_at = VALUES(crawled_at),
      company_industry = VALUES(company_industry),
      company_industry_mapped = VALUES(company_industry_mapped),
      location_text = VALUES(location_text),
      education = VALUES(education),
      experience_text = VALUES(experience_text),
      experience_year_min = VALUES(experience_year_min),
      experience_year_max = VALUES(experience_year_max),
      skills_text = VALUES(skills_text),
      salary_text = VALUES(salary_text),
      salary_min_month = VALUES(salary_min_month),
      salary_max_month = VALUES(salary_max_month),
      salary_avg_month = VALUES(salary_avg_month),
      pay_months = VALUES(pay_months),
      salary_annualized_avg = VALUES(salary_annualized_avg),
      raw_json = VALUES(raw_json);
    """
    rows = 0
    with conn.cursor() as cur:
        for r in records:
            payload = (
                r.job_hash(),
                r.platform,
                r.city,
                r.keyword,
                r.page_no,
                r.job_title,
                r.company_name,
                r.company_industry,
                r.company_industry_mapped,
                r.location_text,
                r.education,
                r.experience_text,
                r.experience_year_min,
                r.experience_year_max,
                r.skills_text,
                r.salary_text,
                r.salary_min_month,
                r.salary_max_month,
                r.salary_avg_month,
                r.pay_months,
                r.salary_annualized_avg,
                r.job_url,
                r.crawled_at.strftime("%Y-%m-%d %H:%M:%S"),
                json.dumps(r.raw, ensure_ascii=False),
            )
            cur.execute(sql, payload)
            rows += 1
    return rows


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """解析命令行参数。"""
    p = argparse.ArgumentParser()
    p.add_argument("--config", default=None, help="YAML配置文件路径，例如：conf/config.yml")
    p.add_argument("--keyword", required=False, default=None, help="岗位关键词，例如：Python工程师")
    p.add_argument("--cities", required=False, default=None, help="城市列表，逗号分隔，例如：北京,上海,深圳 或 BOSS城市code")
    p.add_argument("--pages", type=int, default=3, help="每个城市抓取页数")
    p.add_argument("--platform", choices=["boss", "zhilian", "both"], default="both", help="抓取平台")
    p.add_argument("--headless", type=str, default="true", help="true/false，遇验证码建议false")
    p.add_argument("--proxy", default=None, help="代理，例如：http://user:pass@host:port")
    p.add_argument("--slowmo", type=int, default=0, help="Playwright slow_mo 毫秒")
    p.add_argument("--sleep-min", type=float, default=2.0, help="翻页间隔最小秒数")
    p.add_argument("--sleep-max", type=float, default=5.0, help="翻页间隔最大秒数")
    p.add_argument("--out", default=None, help="输出Excel路径，默认 output/salary_时间戳.xlsx")
    p.add_argument("--mysql-table", default="job_salary_raw", help="MySQL表名")
    p.add_argument("--save-mysql", type=str, default="false", help="true/false，从环境变量读取MySQL连接并落库")
    p.add_argument("--verbose", type=str, default="false", help="true/false，打印抓取进度日志")
    return p.parse_args(argv)


def str_to_bool(v: str) -> bool:
    """将常见字符串转换为布尔值。"""
    return str(v).strip().lower() in {"1", "true", "yes", "y", "on"}


def main(argv: Optional[Sequence[str]] = None) -> int:
    """主入口：按平台抓取数据，导出Excel，可选落库到MySQL。"""
    args = parse_args(argv)
    cfg: Dict[str, Any] = {}
    if args.config:
        cfg = load_yaml_config(str(args.config))

    crawl_cfg = cfg.get("crawl", {}) if isinstance(cfg.get("crawl", {}), dict) else {}
    data_cfg = cfg.get("data", {}) if isinstance(cfg.get("data", {}), dict) else {}
    mysql_cfg = cfg.get("mysql", {}) if isinstance(cfg.get("mysql", {}), dict) else {}

    cities_arg = str(args.cities).strip() if args.cities else ""
    cities = [c.strip() for c in cities_arg.split(",") if c.strip()] if cities_arg else []
    if not cities:
        cities = [str(x).strip() for x in data_cfg.get("cities", []) if str(x).strip()]
    if not cities:
        raise SystemExit("cities 不能为空（请通过 --cities 或 config.yml 的 data.cities 提供）")

    keywords: List[str] = []
    if args.keyword:
        k = str(args.keyword).strip()
        if k:
            keywords = [k]
    if not keywords:
        keywords = [str(x).strip() for x in data_cfg.get("keywords", []) if str(x).strip()]

    if not keywords:
        raise SystemExit("keyword 不能为空（请通过 --keyword 或 config.yml 的 data.keywords 提供）")

    keywords = expand_keywords(keywords)

    pages = int(crawl_cfg.get("pages_per_city", args.pages))
    if pages <= 0:
        raise SystemExit("pages 必须大于0")

    max_keywords = crawl_cfg.get("max_keywords", None)
    if max_keywords is not None and str(max_keywords).strip() != "":
        try:
            mk = int(max_keywords)
        except ValueError:
            mk = 0
        if mk > 0:
            keywords = keywords[:mk]

    platform = str(crawl_cfg.get("platform", args.platform)).strip() or str(args.platform)
    if platform not in {"boss", "zhilian", "both"}:
        raise SystemExit("platform 仅支持 boss/zhilian/both")

    headless = str_to_bool(str(crawl_cfg.get("headless", args.headless)))
    verbose = str_to_bool(str(crawl_cfg.get("verbose", args.verbose)))
    proxy = crawl_cfg.get("proxy", args.proxy)
    proxy = str(proxy).strip() if proxy else None
    slowmo = int(crawl_cfg.get("slowmo_ms", args.slowmo))
    sleep_min = float(crawl_cfg.get("sleep_min", args.sleep_min))
    sleep_max = float(crawl_cfg.get("sleep_max", args.sleep_max))
    out_path = str(args.out).strip() if args.out else str(crawl_cfg.get("output_excel", "")).strip() or None
    if not out_path:
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join("output", f"salary_{ts}.xlsx")

    all_records: List[JobRecord] = []
    if verbose:
        print(
            f"[run] platform={platform} cities={len(cities)} pages_per_city={pages} keywords={len(keywords)} "
            f"headless={headless} proxy={'Y' if proxy else 'N'} out={out_path}"
        )
    for kw in keywords:
        if platform in {"boss", "both"}:
            all_records.extend(
                BossCrawler(
                    headless=headless,
                    proxy=proxy,
                    slow_mo_ms=slowmo,
                    min_sleep=sleep_min,
                    max_sleep=sleep_max,
                    verbose=verbose,
                ).crawl(cities=cities, keyword=kw, pages_per_city=pages)
            )

        if platform in {"zhilian", "both"}:
            all_records.extend(
                ZhilianCrawler(
                    headless=headless,
                    proxy=proxy,
                    slow_mo_ms=slowmo,
                    min_sleep=sleep_min,
                    max_sleep=sleep_max,
                    verbose=verbose,
                ).crawl(cities=cities, keyword=kw, pages_per_city=pages)
            )

    out_path = write_excel(all_records, out_path=out_path)
    if verbose:
        print(f"[run] excel_saved={out_path} records={len(all_records)}")

    save_mysql = str_to_bool(str(mysql_cfg.get("enabled", args.save_mysql)))
    if save_mysql:
        table = str(mysql_cfg.get("table", args.mysql_table)).strip() or str(args.mysql_table)
        conn = mysql_connect_from_config(mysql_cfg) or mysql_connect_from_env()
        if not conn:
            raise SystemExit(
                "未检测到MySQL配置，请通过 config.yml 的 mysql 配置或设置 MYSQL_HOST/MYSQL_PORT/MYSQL_USER/MYSQL_PASSWORD/MYSQL_DATABASE"
            )
        try:
            mysql_ensure_table(conn, table=table)
            rows = mysql_insert_records(conn, table=table, records=all_records)
            if verbose:
                print(f"[mysql] table={table} rows_written={rows}")
        finally:
            conn.close()

    print(f"keywords={len(keywords)} records={len(all_records)} out={out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
