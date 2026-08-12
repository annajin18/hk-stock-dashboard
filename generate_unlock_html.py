#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
港股解禁概览 — 静态网页生成脚本
================================
功能：读取基石投资者Excel + 港交所公告搜索 + 实时行情 + CCASS校验 → 生成 hk_unlock_overview.html

数据流：
  基石投资者0810.xlsx ──┐
  港交所披露易公告 ─────┤
  腾讯财经实时行情 ─────┼──→ 合并计算 → 模板注入 → hk_unlock_overview.html
  CCASS托管量校验 ──────┘

输出：
  - hk_unlock_overview.html  最终静态网页（双击浏览器打开）
  - unlock_data.json         中间结构化数据（调试用）

依赖：pip install openpyxl requests beautifulsoup4 pdfplumber pymupdf
"""

import os
import sys
import json
import logging
import time
import random
import re
from datetime import datetime, date, timedelta
from collections import defaultdict

# ============================
# 配置
# ============================
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "基石投资者0810.xlsx")
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.html")
OUTPUT_HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hk_unlock_overview.html")
OUTPUT_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unlock_data.json")

TODAY = date.today()
USD_TO_HKD = 7.8
LARGE_UNLOCK_THRESHOLD = 5_0000_0000  # 5亿 HKD
CCASS_TIMEOUT = 15  # CCASS请求超时（秒）

# 解禁类型颜色/样式映射
UNLOCK_TYPE_STYLE = {
    "基石解禁": 'bg-appleBlue/10 text-appleBlue',
    "国际配售": 'bg-applePurple/10 text-applePurple',
    "股东配售": 'bg-gray-100 dark:bg-gray-700 text-appleGrayTag',
}

# ============================
# 日志配置
# ============================
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    datefmt='%H:%M:%S',
    handlers=[
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)

# ============================
# 工具函数
# ============================

def hk_code_raw(code):
    """提取纯数字代码（去掉.HK后缀），补零到5位"""
    code = str(code).strip().replace(".HK", "").replace(".hk", "")
    return code.zfill(5)


def hk_code_display(code):
    """港股代码展示格式: 5位.HK"""
    return hk_code_raw(code) + ".HK"


def safe_int(val, default=0):
    """安全转整数"""
    try:
        return int(float(val)) if val else default
    except (ValueError, TypeError):
        return default


def fmt_value_yi(val):
    """格式化金额（亿）"""
    if abs(val) >= 1e8:
        return f"{val/1e8:.1f}"
    return f"{val/1e8:.2f}"


def fmt_shares(val):
    """格式化股数"""
    if val >= 1e8:
        return f"{val/1e8:.1f}亿股"
    if val >= 1e4:
        return f"{val/1e4:.0f}万股"
    return f"{int(val)}股"


def get_week_range(d):
    """获取d所在周的周一~周日"""
    weekday = d.weekday()  # 0=周一
    monday = d - timedelta(days=weekday)
    sunday = monday + timedelta(days=6)
    return monday, sunday


def random_delay(min_s=2.0, max_s=5.0):
    """港交所请求延时2~5秒"""
    delay = random.uniform(min_s, max_s)
    logger.debug(f"  延时 {delay:.1f}s ...")
    time.sleep(delay)


# ============================
# Step 1: 加载 Excel 基石解禁数据
# ============================

def load_excel_data(excel_path):
    """
    从基石投资者Excel加载数据，按股票聚合。
    返回: list[dict] 每只股票的聚合解禁信息
    """
    logger.info("=" * 60)
    logger.info("Step 1: 加载 Excel 基石投资者解禁数据")
    logger.info("=" * 60)

    try:
        import openpyxl
    except ImportError:
        logger.error("缺少 openpyxl 库，请运行: pip install openpyxl")
        return []

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb["基石投资者"]

    raw_records = []
    skipped_gem = 0
    skipped_date = 0

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2:
            continue  # 跳过前2行表头

        code = hk_code_raw(row[0])
        name = str(row[1] or "").strip()

        if not code or code == "00000" or not name:
            continue

        # 过滤创业板（GEM: 08xxx）
        if code.startswith("08"):
            skipped_gem += 1
            continue

        # 解析解禁日
        lift_date = row[15]
        if isinstance(lift_date, str):
            try:
                d = datetime.strptime(lift_date[:10], "%Y-%m-%d").date()
            except ValueError:
                skipped_date += 1
                logger.debug(f"  解禁日解析失败: {code} {name} -> {lift_date}")
                continue
        elif lift_date and hasattr(lift_date, 'strftime'):
            d = lift_date.date() if hasattr(lift_date, 'date') else lift_date
            if hasattr(d, 'strftime'):
                pass  # it's a date
            else:
                skipped_date += 1
                continue
        else:
            skipped_date += 1
            continue

        # 上市日期
        listing_date = row[2]
        if isinstance(listing_date, str):
            ld = listing_date[:10]
        elif listing_date and hasattr(listing_date, 'strftime'):
            ld = listing_date.strftime("%Y-%m-%d")
        else:
            ld = ""

        # 金额处理
        amount = float(row[11] or 0)
        currency = str(row[12] or "HKD").strip().upper()
        amount_hkd = amount if currency == "HKD" else amount * USD_TO_HKD

        investor = str(row[6] or "").strip()
        shares = safe_int(row[10])
        percentage = float(row[13] or 0)
        lockup_months = safe_int(row[14])

        raw_records.append({
            "code_raw": code,
            "name": name,
            "listingDate": ld,
            "liftDate": d.strftime("%Y-%m-%d"),
            "liftDateObj": d,
            "investor": investor,
            "shares": shares,
            "amount": amount,
            "currency": currency,
            "amountHKD": round(amount_hkd),
            "percentage": percentage,
            "lockupMonths": lockup_months,
            "industry": str(row[16] or "").strip(),
            "unlock_type": "基石解禁",
        })

    logger.info(f"  读取记录: {len(raw_records)} 条")
    logger.info(f"  过滤创业板: {skipped_gem} 条")
    logger.info(f"  过滤无效日期: {skipped_date} 条")

    # 按股票代码聚合
    stock_map = defaultdict(list)
    for r in raw_records:
        stock_map[r["code_raw"]].append(r)

    stocks = []
    for code, recs in stock_map.items():
        first = recs[0]
        total_shares = sum(r["shares"] for r in recs)
        total_amount_hkd = sum(r["amountHKD"] for r in recs)

        # 取最早解禁日
        earliest_lift = min(r["liftDateObj"] for r in recs)

        stocks.append({
            "code": hk_code_display(code),
            "code_raw": code,
            "name": first["name"],
            "listingDate": first["listingDate"],
            "liftDate": earliest_lift.strftime("%Y-%m-%d"),
            "liftDateObj": earliest_lift,
            "totalShares": int(total_shares),
            "totalAmountHKD": int(total_amount_hkd),
            "unlockType": "基石解禁",
            "industry": first["industry"],
            "investors": [{
                "name": r["investor"],
                "shares": int(r["shares"]),
                "amount": int(r["amount"]),
                "currency": r["currency"],
                "percentage": round(r["percentage"], 2),
                "lockupMonths": r["lockupMonths"],
            } for r in recs],
        })

    logger.info(f"  聚合股票: {len(stocks)} 只")
    return stocks


# ============================
# Step 2: 搜索港交所配售/股东承诺公告
# ============================

def search_hkex_placement():
    """
    搜索港交所近6个月配售/锁仓相关公告，补充「国际配售」和「股东配售」类型。
    返回: list[dict] 补充的解禁条目
    """
    logger.info("=" * 60)
    logger.info("Step 2: 搜索港交所配售/股东承诺公告")
    logger.info("=" * 60)

    supplementary = []

    # 关键词列表
    keywords = [
        ("placing", "国际配售"),
        ("placement", "国际配售"),
        ("top-up placing", "国际配售"),
        ("lock-up undertaking", "股东配售"),
        ("voluntary lock-up", "股东配售"),
    ]

    session = None
    try:
        import requests
        from bs4 import BeautifulSoup

        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        })

        # 获取 ViewState
        search_url = "https://www1.hkexnews.hk/search/titlesearch.xhtml"
        logger.info(f"  获取搜索页面 ViewState ...")

        try:
            resp = session.get(search_url, timeout=30)
            soup = BeautifulSoup(resp.text, 'html.parser')
            viewstate_input = soup.find("input", {"name": "javax.faces.ViewState"})
            if not viewstate_input:
                logger.warning("  未找到 ViewState，跳过港交所搜索")
                return supplementary
            viewstate = viewstate_input.get("value", "")
        except Exception as e:
            logger.warning(f"  获取 ViewState 失败: {e}，跳过港交所搜索")
            return supplementary

        random_delay(2, 4)

        # 设置日期范围：近6个月
        end_date = TODAY
        start_date = TODAY - timedelta(days=180)

        # POST 设置日期
        post_data = {
            "j_idt10": "j_idt10",
            "j_idt10:loadMoreRange": "500",
            "javax.faces.ViewState": viewstate,
            "from": start_date.strftime("%Y%m%d"),
            "to": end_date.strftime("%Y%m%d"),
        }

        try:
            resp = session.post(search_url, data=post_data, timeout=30)
            logger.info(f"  POST设置日期范围: {start_date.strftime('%Y%m%d')} ~ {end_date.strftime('%Y%m%d')}")
        except Exception as e:
            logger.warning(f"  POST设置日期失败: {e}")
            return supplementary

        random_delay(2, 4)

        # 搜索每个关键词
        for keyword, unlock_type in keywords:
            logger.info(f"  搜索关键词: \"{keyword}\" → {unlock_type}")

            api_url = (
                "https://www1.hkexnews.hk/search/titleSearchServlet.do"
                f"?sortDir=0&sortByOptions=DateTime&category=0&market=SEHK"
                f"&stockId=-1&documentType=-1"
                f"&fromDate={start_date.strftime('%Y%m%d')}"
                f"&toDate={end_date.strftime('%Y%m%d')}"
                f"&title={keyword}&searchType=0&t1code=-2&t2Gcode=-2&t2code=-2"
                f"&rowRange=100&lang=E"
            )

            try:
                resp = session.get(api_url, headers={
                    "Accept": "application/json",
                    "X-Requested-With": "XMLHttpRequest",
                    "Referer": "https://www1.hkexnews.hk/search/titlesearch.xhtml",
                }, timeout=30)

                data = resp.json()
                result_str = data.get("result", "[]")
                items = json.loads(result_str)

                logger.info(f"    匹配 {len(items)} 条公告")

                for item in items[:20]:  # 限前20条
                    stock_code = item.get("STOCK_CODE", "")
                    stock_name = item.get("STOCK_NAME", "")
                    title = item.get("TITLE", "")
                    file_link = item.get("FILE_LINK", "")
                    date_time = item.get("DATE_TIME", "")

                    if not stock_code or stock_code.zfill(5).startswith("08"):
                        continue

                    # PDF链接
                    pdf_url = ""
                    if file_link:
                        pdf_url = "https://www1.hkexnews.hk" + file_link

                    supplementary.append({
                        "code_raw": stock_code.zfill(5),
                        "name": stock_name,
                        "unlockType": unlock_type,
                        "title": title,
                        "announcementUrl": pdf_url,
                        "announcementDate": date_time,
                        "source": "hkex_search",
                    })

            except Exception as e:
                logger.warning(f"    搜索 \"{keyword}\" 失败: {e}")

            random_delay(2, 5)

    except ImportError:
        logger.warning("缺少 requests/beautifulsoup4 库，跳过港交所搜索")
    except Exception as e:
        logger.warning(f"港交所搜索异常: {e}")

    logger.info(f"  补充条目: {len(supplementary)} 条（含重复股票）")

    # 标记：这些条目来自搜索，数据完整度有限
    for s in supplementary:
        s["data_quality"] = "partial"

    return supplementary


# ============================
# Step 3: 拉取实时行情
# ============================

def parse_cn_shares(text):
    """解析中文股数字符串：'430.74万' → 4307400，'1.10亿' → 110000000"""
    text = str(text or "").strip().replace(",", "")
    if not text:
        return 0
    try:
        if text.endswith("亿"):
            return int(float(text[:-1]) * 1e8)
        if text.endswith("万"):
            return int(float(text[:-1]) * 1e4)
        return int(float(text))
    except ValueError:
        return 0


def fetch_zhitong_limit_sale():
    """
    从智通财经获取基石投资者限售解禁列表（主要数据源，接口实时更新）。
    返回: list[dict] 每条为一只股票的限售解禁概要
    """
    logger.info("=" * 60)
    logger.info("Step 2.5: 智通财经限售解禁数据（补充数据源）")
    logger.info("=" * 60)

    url = "https://betaapi.zhitongcaijing.com/v12/ctoneinvestor/limit-sale-list.html"
    params = {"page": "1", "pageSize": "500"}
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://www.zhitongcaijing.com/data.html",
    }
    try:
        import requests
        resp = requests.get(url, params=params, headers=headers, timeout=25)
        resp.raise_for_status()
        payload = resp.json()
        rows = (payload.get("data") or {}).get("list") or []
    except Exception as e:
        logger.warning(f"  智通财经数据获取失败，跳过补充: {e}")
        return []

    records = []
    for r in rows:
        code_raw = str(r.get("secuCode") or "").replace(".HK", "").replace(".hk", "").zfill(5)
        name = str(r.get("secuAbbr") or "").strip()
        lift_date = str(r.get("limitSaleDate") or "").strip()[:10]
        if not code_raw or not name or not lift_date:
            continue
        if code_raw.startswith("08"):  # 过滤创业板
            continue
        records.append({
            "code_raw": code_raw,
            "name": name,
            "liftDate": lift_date,
            "listingDate": str(r.get("listedDate") or "").strip()[:10],
            "totalShares": parse_cn_shares(r.get("totalNumSubActualSum")),
            "lockupPeriod": str(r.get("limitSalePeriod") or "").strip(),
            "investorNum": int(r.get("num") or 0),
            "investorSample": str(r.get("cStoneInvName") or "").strip(),
        })

    logger.info(f"  获取智通财经限售解禁记录: {len(records)} 条")
    return records


def merge_zhitong_primary(excel_stocks, zhitong_records):
    """
    以智通财经数据为主、Excel 为辅合并股票列表。
    - 智通有、Excel 也有：以智通日期/股数为准，用 Excel 补充投资者明细（差异写日志）
    - 智通有、Excel 无：直接用智通数据
    - 智通无、Excel 有：保留 Excel 数据作为辅助
    """
    excel_by_code = {s["code_raw"]: s for s in excel_stocks}
    zhitong_by_code = {}
    for rec in zhitong_records:
        zhitong_by_code.setdefault(rec["code_raw"], rec)

    merged = []
    for code_raw, rec in sorted(zhitong_by_code.items()):
        ex = excel_by_code.get(code_raw)
        try:
            lift_obj = datetime.strptime(rec["liftDate"], "%Y-%m-%d").date()
        except ValueError:
            continue

        if ex:
            if ex["liftDate"] != rec["liftDate"]:
                logger.warning(
                    f"  智通与Excel解禁日不一致 [{code_raw}] {rec['name']}: "
                    f"Excel={ex['liftDate']} 智通={rec['liftDate']}（以智通为准）")
            if ex["totalShares"] != rec["totalShares"]:
                logger.warning(
                    f"  智通与Excel股数不一致 [{code_raw}] {rec['name']}: "
                    f"Excel={ex['totalShares']} 智通={rec['totalShares']}（以智通为准）")

        lockup_months = 0
        m = re.search(r"(\d+)\s*个月", rec.get("lockupPeriod", ""))
        if m:
            lockup_months = int(m.group(1))

        # Excel 有投资者明细则优先用明细；否则用智通的代表投资者占位
        investors = ex["investors"] if ex and ex.get("investors") else [{
            "name": rec.get("investorSample") or "基石投资者",
            "shares": rec.get("totalShares", 0),
            "amount": 0,
            "currency": "HKD",
            "percentage": 0,
            "lockupMonths": lockup_months,
        }]

        merged.append({
            "code": hk_code_display(code_raw),
            "code_raw": code_raw,
            "name": rec["name"] or (ex["name"] if ex else ""),
            "listingDate": rec.get("listingDate") or (ex.get("listingDate", "") if ex else ""),
            "liftDate": rec["liftDate"],
            "liftDateObj": lift_obj,
            "totalShares": rec.get("totalShares", 0),
            "totalAmountHKD": ex.get("totalAmountHKD", 0) if ex else 0,
            "unlockType": "基石解禁",
            "industry": ex.get("industry", "") if ex else "",
            "source": "智通财经",
            "investors": investors,
        })

    zhitong_codes = set(zhitong_by_code)
    excel_only = 0
    for s in excel_stocks:
        if s["code_raw"] not in zhitong_codes:
            s["source"] = "Excel补充"
            merged.append(s)
            excel_only += 1

    logger.info(
        f"  合并结果: 智通为主 {len(zhitong_by_code)} 只 + Excel补充 {excel_only} 只 = {len(merged)} 只")
    return merged


def fetch_market_prices(codes_raw):
    """
    批量拉取港股实时行情（腾讯财经接口）。
    参数: codes_raw - 纯数字代码列表，如 ['00700', '09988']
    返回: dict {code_raw: {"price": float, "prevClose": float, "changePct": float, "name": str}}
    """
    logger.info("=" * 60)
    logger.info("Step 3: 拉取港股实时行情")
    logger.info("=" * 60)

    import requests

    prices = {}
    batch_size = 50
    all_codes = list(set(codes_raw))

    for i in range(0, len(all_codes), batch_size):
        batch = all_codes[i:i + batch_size]
        codes_param = ",".join(f"hk{c}" for c in batch)

        try:
            url = f"https://qt.gtimg.cn/q={codes_param}"
            resp = requests.get(url, timeout=15,
                headers={"User-Agent": "Mozilla/5.0"})

            if resp.status_code != 200:
                logger.warning(f"  腾讯行情API返回 {resp.status_code}")
                continue

            # 解析返回文本: v_hk00700="1~腾讯控股~00700~428.60~..."
            # 字段索引（~分隔）：
            # [1]=名称, [3]=现价, [4]=昨收, [32]=涨跌幅
            # 腾讯接口返回GBK编码，直接使用content.decode
            text = resp.content.decode('gbk', errors='replace')

            for line in text.split('\n'):
                line = line.strip()
                if not line or '=' not in line:
                    continue
                match = re.match(r'v_hk(\d+)="(.+)"', line)
                if not match:
                    continue
                code = match.group(1)
                fields = match.group(2).split('~')
                if len(fields) < 33:
                    continue

                try:
                    price = float(fields[3]) if fields[3] else 0
                    prev_close = float(fields[4]) if fields[4] else 0
                    change_pct = float(fields[32]) if fields[32] else 0
                    name = fields[1].strip()

                    prices[code] = {
                        "price": price,
                        "prevClose": prev_close,
                        "changePct": change_pct,
                        "name": name,
                    }
                except (ValueError, IndexError) as e:
                    logger.debug(f"    解析失败 {code}: {e}")

            logger.info(f"  批次 {i//batch_size+1}: 获取 {len(batch)} 只 → 成功 {sum(1 for c in batch if c in prices)} 只")

        except requests.exceptions.Timeout:
            logger.warning(f"  批次 {i//batch_size+1} 超时")
        except Exception as e:
            logger.warning(f"  批次 {i//batch_size+1} 异常: {e}")

        # 批次间延时
        if i + batch_size < len(all_codes):
            time.sleep(0.3)

    logger.info(f"  行情总数: {len(prices)} 只")
    return prices


# ============================
# Step 4: CCASS 交叉校验
# ============================

def validate_ccass_single(code_raw, session=None):
    """
    查询单只股票的CCASS托管量。
    返回: dict {"success": bool, "totalShares": int, "ccassShares": int, "ratio": float, "note": str}
    """
    import requests
    from bs4 import BeautifulSoup

    result = {
        "success": False,
        "totalShares": 0,
        "ccassShares": 0,
        "ratio": 0,
        "note": "",
    }

    own_session = False
    if session is None:
        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        })
        own_session = True

    try:
        ccass_url = "https://www.hkexnews.hk/sdw/search/searchsdw.aspx"

        # Step 1: GET 获取 ViewState
        resp = session.get(ccass_url, timeout=CCASS_TIMEOUT)
        soup = BeautifulSoup(resp.text, 'html.parser')

        # 提取隐藏字段
        form_data = {}
        for inp in soup.select('input[type="hidden"]'):
            name = inp.get('name', '')
            value = inp.get('value', '')
            if name:
                form_data[name] = value

        if '__VIEWSTATE' not in form_data:
            result["note"] = "CCASS: 无法获取ViewState"
            return result

        # Step 2: POST 查询
        # 使用最近的交易日（通常T+2延迟）
        query_date = TODAY
        form_data['ddlShareholdingDay'] = f"{query_date.day:02d}"
        form_data['ddlShareholdingMonth'] = f"{query_date.month:02d}"
        form_data['ddlShareholdingYear'] = str(query_date.year)
        form_data['txt_stock_code'] = code_raw
        form_data['txt_stock_name'] = ''
        form_data['txt_ParticipantID'] = ''
        form_data['txt_Participant_name'] = ''
        form_data['btnSearch.x'] = '23'
        form_data['btnSearch.y'] = '12'

        resp = session.post(ccass_url, data=form_data, timeout=CCASS_TIMEOUT)

        # Step 3: 解析汇总表格
        soup = BeautifulSoup(resp.text, 'html.parser')

        # 查找 "已发行股份/权证总数" 行
        # CCASS 页面结构：汇总表在上部
        summary_table = None
        for table in soup.select('table'):
            text = table.get_text()
            if '参与者数目' in text or '已发行股份' in text:
                summary_table = table
                break

        if not summary_table:
            # 尝试从页面文本中提取
            page_text = soup.get_text()
            # 正则匹配已发行股份
            issued_match = re.search(r'已发行股[份票].*?[：:]\s*([\d,]+)', page_text)
            if issued_match:
                result["totalShares"] = int(issued_match.group(1).replace(',', ''))
                result["success"] = True
                result["note"] = "CCASS: 部分数据（仅已发行股份）"
                return result
            result["note"] = "CCASS: 未找到汇总数据"
            return result

        # 解析汇总表
        rows = summary_table.select('tr')
        for row in rows:
            cells = row.select('td, th')
            cell_texts = [c.get_text(strip=True) for c in cells]

            for ct in cell_texts:
                # 匹配已发行总数
                if '已发行' in ct and '总数' in ct:
                    for ct2 in cell_texts:
                        num_match = re.search(r'([\d,]+)', ct2)
                        if num_match:
                            result["totalShares"] = int(num_match.group(1).replace(',', ''))

        if result["totalShares"] > 0:
            result["success"] = True

            # CCASS总托管量 = 市场中介者 + 投资者户口持有人（愿意披露）
            # 通常在所有行中
            for row in rows:
                cells = row.select('td, th')
                cell_texts = [c.get_text(strip=True) for c in cells]
                for ct in cell_texts:
                    if '于中央结算系统' in ct or 'CCASS' in ct:
                        for ct2 in cell_texts:
                            num_match = re.search(r'([\d,]+)', ct2)
                            if num_match and int(num_match.group(1).replace(',', '')) > 0:
                                val = int(num_match.group(1).replace(',', ''))
                                if val > result["ccassShares"]:
                                    result["ccassShares"] = val

            if result["ccassShares"] > 0 and result["totalShares"] > 0:
                result["ratio"] = round(result["ccassShares"] / result["totalShares"] * 100, 2)
                result["note"] = f"CCASS: 托管{result['ccassShares']:,}股 / 已发行{result['totalShares']:,}股 = {result['ratio']}%"

    except requests.exceptions.Timeout:
        result["note"] = "CCASS: 请求超时"
    except Exception as e:
        result["note"] = f"CCASS: 异常 - {str(e)[:80]}"

    return result


def validate_ccass_batch(stocks_this_month):
    """
    对「本月解禁」的股票做CCASS批量校验。
    返回: dict {code_raw: ccass_result}
    """
    logger.info("=" * 60)
    logger.info("Step 4: CCASS 托管量交叉校验")
    logger.info("=" * 60)

    if not stocks_this_month:
        logger.info("  本月无解禁股票，跳过CCASS校验")
        return {}

    ccass_results = {}

    try:
        import requests
        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        })
    except ImportError:
        logger.warning("缺少 requests 库，CCASS校验全部降级")
        return {s["code_raw"]: {"success": False, "note": "CCASS: 缺少依赖"} for s in stocks_this_month}

    success_count = 0
    fail_count = 0

    for i, stock in enumerate(stocks_this_month):
        code = stock["code_raw"]
        logger.info(f"  [{i+1}/{len(stocks_this_month)}] {code} {stock['name']}")

        try:
            result = validate_ccass_single(code, session)
            ccass_results[code] = result

            if result["success"]:
                success_count += 1
                logger.info(f"    [OK] {result['note']}")
            else:
                fail_count += 1
                logger.warning(f"    [FAIL] {result['note']}")

        except Exception as e:
            fail_count += 1
            logger.error(f"    [ERR] 校验异常: {e}")
            ccass_results[code] = {"success": False, "note": f"CCASS: 异常 - {str(e)[:80]}"}

        # 港交所请求延时
        if i < len(stocks_this_month) - 1:
            random_delay(2, 4)

    logger.info(f"  CCASS校验完成: 成功 {success_count} / 失败 {fail_count} / 总计 {len(stocks_this_month)}")
    return ccass_results


# ============================
# Step 5: 合并计算
# ============================

def merge_and_enrich(stocks, prices, ccass_results, supplementary):
    """
    合并所有数据源，计算解禁市值，分类筛选。
    返回: dict 完整的结构化数据
    """
    logger.info("=" * 60)
    logger.info("Step 5: 合并数据与计算")
    logger.info("=" * 60)

    enriched = []
    dropped = []
    deviation_warnings = []

    for stock in stocks:
        code_raw = stock["code_raw"]
        code_display = stock["code"]

        # 注入实时价格
        price_info = prices.get(code_raw, {})
        current_price = price_info.get("price", 0)
        stock["price"] = round(current_price, 2)
        stock["prevClose"] = price_info.get("prevClose", 0)
        stock["changePct"] = price_info.get("changePct", 0)

        # 本地计算解禁市值
        unlock_value = stock["totalShares"] * current_price
        stock["unlockValue"] = round(unlock_value)
        stock["unlockValueYi"] = unlock_value / 1e8

        # 发行价：从基石投资者取第一个的金额/股数
        if stock["investors"] and stock["investors"][0]["shares"] > 0:
            issue_price = stock["investors"][0]["amount"] / stock["investors"][0]["shares"]
            # 如果币种是USD，转换为HKD
            if stock["investors"][0]["currency"] == "USD":
                issue_price *= USD_TO_HKD
            stock["issuePrice"] = round(issue_price, 2)
        else:
            stock["issuePrice"] = 0

        # CCASS 校验结果
        ccass = ccass_results.get(code_raw, {})
        stock["ccassValidated"] = ccass.get("success", False)
        stock["ccassRatio"] = ccass.get("ratio", 0)
        stock["ccassNote"] = ccass.get("note", "未经过CCASS交叉校验，仅供参考")

        # 解禁股数交叉校验：公告披露股数 VS CCASS 托管可解禁股数
        # 偏差 = |公告股数 - CCASS托管股数| / 公告股数
        # 注意：不能用「解禁市值 vs 认购金额」当偏差，那是股价涨跌，不是数据错误。
        ccass_shares = ccass.get("ccassShares", 0)
        if stock["totalShares"] > 0 and ccass_shares > 0 and ccass.get("success", False):
            deviation = abs(stock["totalShares"] - ccass_shares) / stock["totalShares"]
            stock["deviation"] = round(deviation * 100, 2)

            if deviation > 0.15:
                dropped.append({
                    "code": code_display,
                    "name": stock["name"],
                    "reason": f"CCASS股数偏差 {deviation*100:.1f}% > 15%",
                    "announcedShares": stock["totalShares"],
                    "ccassShares": ccass_shares,
                })
                logger.warning(f"  丢弃 [{code_display}] {stock['name']}: CCASS股数偏差 {deviation*100:.1f}% > 15%")
                continue
            elif deviation > 0.05:
                stock["deviationFlag"] = "warning"
                deviation_warnings.append({
                    "code": code_display,
                    "name": stock["name"],
                    "deviation": f"{deviation*100:.1f}%",
                })
                logger.warning(f"  偏差告警 [{code_display}] {stock['name']}: CCASS股数偏差 {deviation*100:.1f}%")
            else:
                stock["deviationFlag"] = "ok"
        else:
            # CCASS 未成功校验 → 降级：不丢弃条目（页面不展示校验标签）
            stock["deviation"] = 0
            stock["deviationFlag"] = "ok"

        # 大额解禁判定
        stock["isLarge"] = unlock_value >= LARGE_UNLOCK_THRESHOLD

        # 风险提示文本
        stock["riskText"] = generate_risk_text(stock)

        enriched.append(stock)

    # 补充配售/股东锁仓条目
    # 如果 supplementary 中有不在 enriched 中的股票，作为补充条目添加
    existing_codes = {s["code_raw"] for s in enriched}
    for sup in supplementary:
        if sup["code_raw"] not in existing_codes:
            sup_code = sup["code_raw"]
            price_info = prices.get(sup_code, {})
            current_price = price_info.get("price", 0)

            sup["code"] = hk_code_display(sup_code)
            sup["price"] = round(current_price, 2)
            sup["totalShares"] = 0
            sup["unlockValue"] = 0
            sup["unlockValueYi"] = 0
            sup["totalAmountHKD"] = 0
            sup["issuePrice"] = 0
            sup["isLarge"] = False
            sup["deviationFlag"] = "ok"
            sup["deviation"] = 0
            sup["ccassValidated"] = False
            sup["ccassNote"] = "未经过CCASS交叉校验，仅供参考"
            sup["investors"] = []
            sup["industry"] = ""
            sup["listingDate"] = sup.get("announcementDate", "")
            sup["liftDate"] = sup.get("announcementDate", "")
            sup["liftDateObj"] = TODAY
            sup["riskText"] = "此为港交所公告搜索结果，数据完整度有限，请查阅原始公告。"
            enriched.append(sup)

    # 按解禁日期排序
    enriched.sort(key=lambda x: x.get("liftDateObj", date(2099, 12, 31)))

    logger.info(f"  有效条目: {len(enriched)}")
    logger.info(f"  丢弃条目: {len(dropped)}")
    for d in dropped:
        logger.info(f"    - {d['code']} {d['name']}: {d['reason']}")

    logger.info(f"  偏差告警: {len(deviation_warnings)} 条")
    for w in deviation_warnings:
        logger.info(f"    [!] {w['code']} {w['name']}: 偏差 {w['deviation']}")

    return {
        "stocks": enriched,
        "dropped": dropped,
        "warnings": deviation_warnings,
        "today": TODAY.strftime("%Y-%m-%d"),
    }


def generate_risk_text(stock):
    """根据数据动态生成风险提示文本"""
    parts = []

    # 现价 vs 发行价
    if stock["price"] > 0 and stock.get("issuePrice", 0) > 0:
        change = (stock["price"] - stock["issuePrice"]) / stock["issuePrice"] * 100
        if change > 50:
            parts.append(f"现价较发行价涨幅{change:.0f}%，基石投资者浮盈丰厚，解禁后减持意愿较强")
        elif change > 20:
            parts.append(f"现价较发行价涨幅{change:.0f}%，存在一定减持动力")
        elif change < -20:
            parts.append(f"现价较发行价跌幅{abs(change):.0f}%，持仓处于亏损状态，解禁抛压相对可控")
        elif change < 0:
            parts.append(f"现价低于发行价{abs(change):.0f}%，基石投资者处于浮亏，被动减持可能性较低")
        else:
            parts.append(f"现价较发行价变动{change:.0f}%，浮动空间有限")

    # 解禁市值
    if stock.get("unlockValueYi", 0) >= 5:
        parts.append(f"解禁市值{stock['unlockValueYi']:.1f}亿港元，属于大额解禁，关注集中抛售风险")
    elif stock.get("unlockValueYi", 0) >= 1:
        parts.append(f"解禁市值{stock['unlockValueYi']:.1f}亿港元，对短期盘面有一定影响")

    # 偏差告警
    if stock.get("deviationFlag") == "warning":
        parts.append("数据存在偏差，建议查阅原始公告确认解禁细节")

    if not parts:
        parts.append("解禁到期≠必然减持，需持续跟踪交易所CCASS持仓变动")

    parts.append("本页面数据仅供资讯参考，不构成任何投资建议")
    return "。".join(parts) + "。"


# ============================
# Step 6: 生成 HTML
# ============================

def generate_calendar_html(stocks, today):
    """
    生成解禁日历 HTML：本周 + 下周，周一~周五共 10 个格子。
    每个格子展示该工作日解禁的股票（每只一行，含市值/股数/类型标签）。
    返回: str HTML代码
    """
    weekday_names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

    # 本周一 / 下周一
    monday = today - timedelta(days=today.weekday())
    next_monday = monday + timedelta(days=7)
    window_end = next_monday + timedelta(days=5)  # 下周五（含）

    # 按解禁日分组（仅保留落在两周工作日内的股票）
    by_date = defaultdict(list)
    for s in stocks:
        lift_date_str = s.get("liftDate", "")
        try:
            lift_d = datetime.strptime(lift_date_str, "%Y-%m-%d").date()
        except (ValueError, TypeError):
            continue
        if monday <= lift_d < window_end and lift_d.weekday() < 5:
            by_date[lift_date_str].append(s)

    cells = []
    for week_offset in (0, 7):
        for wd in range(5):  # 周一~周五
            d = monday + timedelta(days=week_offset + wd)
            date_str = d.strftime("%Y-%m-%d")
            month_day = d.strftime("%m.%d")
            weekday = weekday_names[d.weekday()]
            is_today = (d == today)

            day_label = f"{month_day} {weekday}"
            if is_today:
                day_label += "（今日）"

            day_class = "day-today" if is_today else ""
            day_stocks = by_date.get(date_str, [])

            if not day_stocks:
                items_html = '<div class="text-xs text-appleTextGray py-2">无解禁</div>'
            else:
                # 单日最多展示5只，超出折叠
                visible_stocks = day_stocks[:5]
                hidden_stocks = day_stocks[5:]

                items_html = ""
                for s in visible_stocks:
                    value_str = f"{s.get('unlockValueYi', 0):.1f}亿｜{fmt_shares(s.get('totalShares', 0))}"
                    type_style = UNLOCK_TYPE_STYLE.get(s.get("unlockType", "基石解禁"), UNLOCK_TYPE_STYLE["基石解禁"])

                    flags_html = ""
                    if s.get("deviationFlag") == "warning":
                        flags_html += '<div class="text-xs text-appleOrange mt-1">数据存在偏差，请查阅原始公告</div>'

                    items_html += f'''
                        <div class="border-b border-appleLine dark:border-gray-700 pb-2 stock-row" data-code="{s['code']}" data-name="{s['name']}">
                            <div>{s['code']} {s['name']}</div>
                            <div class="text-appleTextGray">{value_str}</div>
                            <span class="px-1 py-0 rounded-[4px] {type_style} text-[10px]">{s.get('unlockType', '基石解禁')}</span>
                            {flags_html}
                        </div>'''

                if hidden_stocks:
                    items_html += f'''
                        <details class="arrow-wrap text-xs">
                            <summary class="pr-6 text-appleBlue cursor-pointer">+{len(hidden_stocks)}只更多</summary>
                            <div class="pt-2 space-y-2">'''
                    for s in hidden_stocks:
                        value_str = f"{s.get('unlockValueYi', 0):.1f}亿｜{fmt_shares(s.get('totalShares', 0))}"
                        type_style = UNLOCK_TYPE_STYLE.get(s.get("unlockType", "基石解禁"), UNLOCK_TYPE_STYLE["基石解禁"])
                        items_html += f'''
                                <div class="stock-row pb-1" data-code="{s['code']}" data-name="{s['name']}">
                                    <div>{s['code']} {s['name']}</div>
                                    <div class="text-appleTextGray">{value_str}</div>
                                    <span class="px-1 py-0 rounded-[4px] {type_style} text-[10px]">{s.get('unlockType', '基石解禁')}</span>
                                </div>'''
                    items_html += '''
                            </div>
                        </details>'''

            cells.append(f'''
                <div class="border border-appleLine dark:border-gray-700 rounded-appleSmall p-3 day-item {day_class}">
                    <div class="font-medium mb-2 text-sm sticky top-0 bg-inherit pb-1">{day_label}</div>
                    <div class="space-y-2 text-xs">
                        {items_html}
                    </div>
                </div>''')

    return ''.join(cells)


def generate_large_cards_html(large_stocks):
    """
    生成本月大额解禁卡片 HTML。
    返回: str HTML代码
    """
    if not large_stocks:
        return '<div class="col-span-4 text-center py-8 text-appleTextGray">本月无5亿港元以上大额解禁标的</div>'

    cards = []
    for s in large_stocks:
        value_str = f"{s.get('unlockValueYi', 0):.1f}"
        date_str = s.get("liftDate", "")[-5:]  # MM-DD
        type_style = UNLOCK_TYPE_STYLE.get(s.get("unlockType", "基石解禁"), UNLOCK_TYPE_STYLE["基石解禁"])

        deviation_html = ""
        if s.get("deviationFlag") == "warning":
            deviation_html = '<div class="mt-2 text-xs text-appleOrange bg-orange-50 dark:bg-orange-900/20 px-2 py-1 rounded-appleSmall">⚠ 数据存在偏差，请查阅原始公告</div>'

        cards.append(f'''
                <div class="border border-appleLine dark:border-gray-700 rounded-appleMid p-4 stock-row card-hover" data-code="{s['code']}" data-name="{s['name']}">
                    <div class="font-semibold">{s['code']} {s['name']}</div>
                    <div class="mt-2 text-sm space-y-1">
                        <div>解禁市值：{value_str} 亿 HKD</div>
                        <div>解禁日期：{date_str}</div>
                        <span class="px-1 py-0 rounded-[4px] {type_style} text-xs">{s.get('unlockType', '基石解禁')}</span>
                    </div>
                    {deviation_html}
                </div>''')

    return ''.join(cards)


def generate_html(template_path, data, output_path):
    """
    读取模板，注入数据，输出最终HTML。
    """
    logger.info("=" * 60)
    logger.info("Step 6: 生成 HTML")
    logger.info("=" * 60)

    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    stocks = data["stocks"]
    today = TODAY

    # 当前月份
    current_month = today.strftime("%Y-%m")

    # 本月解禁股票
    month_stocks = [s for s in stocks if s.get("liftDate", "")[:7] == current_month]

    # 本周（ISO周：周一~周日）解禁市值，用于「本周解禁总市值」
    week_monday = today - timedelta(days=today.weekday())
    week_sunday = week_monday + timedelta(days=6)
    week_stocks = [
        s for s in stocks
        if week_monday.strftime("%Y-%m-%d") <= s.get("liftDate", "") <= week_sunday.strftime("%Y-%m-%d")
    ]

    # 上周（用于环比）
    prev_monday = week_monday - timedelta(days=7)
    prev_sunday = week_monday - timedelta(days=1)
    prev_stocks = [
        s for s in stocks
        if prev_monday.strftime("%Y-%m-%d") <= s.get("liftDate", "") <= prev_sunday.strftime("%Y-%m-%d")
    ]

    # 本周解禁总市值
    week_total_value = sum(s.get("unlockValue", 0) for s in week_stocks) / 1e8
    prev_week_total_value = sum(s.get("unlockValue", 0) for s in prev_stocks) / 1e8

    # 环比
    if prev_week_total_value > 0:
        week_compare_pct = (week_total_value - prev_week_total_value) / prev_week_total_value * 100
        if week_compare_pct > 0:
            week_compare = f"↑{week_compare_pct:.1f}%"
        elif week_compare_pct < 0:
            week_compare = f"↓{abs(week_compare_pct):.1f}%"
        else:
            week_compare = "持平"
    else:
        week_compare = "--"

    # 本月解禁个股总数（去重）
    month_codes = set(s["code"] for s in month_stocks)
    month_total_count = len(month_codes)

    # 大额解禁
    large_stocks = [s for s in month_stocks if s.get("isLarge", False)]

    # 本周解禁：取本周+下周最紧迫的
    # 实际展示本月为重点

    # 生成日历 HTML
    calendar_html = generate_calendar_html(stocks, today)

    # 大额解禁卡片
    large_unlock_html = generate_large_cards_html(large_stocks)

    # 日期文本
    weekday_names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    render_date = f"{today.strftime('%Y-%m-%d')} {weekday_names[today.weekday()]}"

    # 数据集 JSON（精简版，仅包含弹窗需要的字段）
    dataset = []
    for s in stocks:
        dataset.append({
            "code": s.get("code", ""),
            "name": s.get("name", ""),
            "price": f"{s.get('price', 0):.2f}",
            "issuePrice": f"{s.get('issuePrice', 0):.2f}",
            "listDate": s.get("listingDate", ""),
            "unlockDate": s.get("liftDate", ""),
            "unlockShares": s.get("totalShares", 0),
            "unlockValue": round(s.get("unlockValueYi", 0), 1),
            "unlockType": s.get("unlockType", "基石解禁"),
            "riskText": s.get("riskText", ""),
            "investors": s.get("investors", []),
            "lockStartDate": s.get("listingDate", ""),
        })

    dataset_json = json.dumps(dataset, ensure_ascii=False, indent=8)

    # 替换占位符
    replacements = {
        "{{render_date}}": render_date,
        "{{week_total_value}}": f"{week_total_value:.1f}",
        "{{week_compare}}": week_compare,
        "{{month_total_count}}": str(month_total_count),
        "{{month_large_count}}": str(len(large_stocks)),
        "{{calendar_html}}": calendar_html,
        "{{large_unlock_html}}": large_unlock_html,
        "{{dataset_json}}": dataset_json,
    }

    for placeholder, value in replacements.items():
        template = template.replace(placeholder, value)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(template)

    logger.info(f"  输出: {output_path}")
    logger.info(f"  本周解禁总市值: {week_total_value:.1f} 亿 HKD")
    logger.info(f"  本月解禁个股: {month_total_count} 只")
    logger.info(f"  大额解禁(≥5亿): {len(large_stocks)} 只")
    logger.info(f"  数据集条目: {len(dataset)} 条")

    # 保存中间JSON
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    logger.info(f"  中间JSON: {OUTPUT_JSON}")


# ============================
# Main
# ============================

def main():
    import argparse

    parser = argparse.ArgumentParser(description="港股解禁概览 - 静态网页生成工具")
    parser.add_argument("--no-ccass", action="store_true", help="跳过CCASS交叉校验（快速生成）")
    parser.add_argument("--date", default=None, help="基准日期 YYYY-MM-DD（默认使用今天）")
    args = parser.parse_args()

    global TODAY
    if args.date:
        try:
            TODAY = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            logger.error("--date 格式应为 YYYY-MM-DD，使用默认今天")

    logger.info("=" * 60)
    logger.info("港股解禁概览 - 静态网页生成工具")
    logger.info(f"运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"基准日期: {TODAY}")
    logger.info("=" * 60)

    # Step 1: 加载 Excel
    stocks = load_excel_data(EXCEL_PATH)
    if not stocks:
        logger.error("未加载到任何股票数据，退出")
        return 1

# Step 1.5: 智通财经为主、Excel 为辅合并数据
    try:
        zhitong_records = fetch_zhitong_limit_sale()
    except Exception as e:
        logger.warning(f"智通财经数据获取异常，仅使用Excel: {e}")
        zhitong_records = []
    stocks = merge_zhitong_primary(stocks, zhitong_records)

    # Step 2: 港交所公告搜索（补充数据源，可降级）
    supplementary = []
    try:
        supplementary = search_hkex_placement()
    except Exception as e:
        logger.warning(f"港交所搜索阶段异常，跳过: {e}")

    # Step 3: 拉取行情
    all_codes = [s["code_raw"] for s in stocks]
    for s in supplementary:
        if s["code_raw"] not in all_codes:
            all_codes.append(s["code_raw"])

    prices = {}
    try:
        prices = fetch_market_prices(all_codes)
    except Exception as e:
        logger.error(f"行情拉取阶段异常: {e}")
        # 降级：使用0价格继续
        prices = {}

    # Step 4: CCASS 校验（本月解禁股票）
    current_month = TODAY.strftime("%Y-%m")
    month_stocks = [s for s in stocks if s.get("liftDate", "")[:7] == current_month]

    ccass_results = {}
    if args.no_ccass:
        logger.info("Step 4: 已通过 --no-ccass 跳过 CCASS 校验")
        ccass_results = {}
    else:
        try:
            ccass_results = validate_ccass_batch(month_stocks)
        except Exception as e:
            logger.warning(f"CCASS校验阶段异常，全部降级: {e}")

    # Step 5: 合并计算
    data = merge_and_enrich(stocks, prices, ccass_results, supplementary)

    # Step 6: 生成 HTML
    generate_html(TEMPLATE_PATH, data, OUTPUT_HTML)

    # 最终统计
    final_stocks = data["stocks"]
    # 本月特定统计
    current_month = TODAY.strftime("%Y-%m")
    month_final = [s for s in final_stocks if s.get("liftDate", "")[:7] == current_month]
    large_month = sum(1 for s in month_final if s.get("isLarge", False))
    large_all = sum(1 for s in final_stocks if s.get("isLarge", False))
    deviation_count = sum(1 for s in final_stocks if s.get("deviationFlag") == "warning")
    no_ccass = sum(1 for s in final_stocks if not s.get("ccassValidated", False))

    logger.info("=" * 60)
    logger.info("=== 生成完成! ===")
    logger.info(f"  有效解禁条目: {len(final_stocks)}")
    logger.info(f"  本月解禁: {len(month_final)}只 (大额{large_month}只)")
    logger.info(f"  全部大额(≥5亿HKD): {large_all}")
    logger.info(f"  偏差告警(5-15%): {deviation_count}")
    logger.info(f"  未CCASS校验: {no_ccass}")
    logger.info(f"  丢弃条目: {len(data['dropped'])}")
    logger.info(f"  输出文件: {OUTPUT_HTML}")
    logger.info("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())
