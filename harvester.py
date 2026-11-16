# -*- coding: utf-8 -*-
"""
网页表格采集器 · 核心库
=====================
把公开网页里的表格抓下来，清洗、去重，导出成 CSV / Excel。

设计原则：
1. **只抓公开数据**。不做登录态、不绕验证码、不破反爬。拿不到就老实报错。
2. **不猜**。表格结构靠 DOM 判断，判断不出来就返回原始结果让用户自己看。
3. **可当库用**。CLI 和 Web 界面都只是薄薄一层，核心逻辑全在这里。

典型用法：
    from harvester import fetch_html, find_tables, dedupe, to_csv

    html = fetch_html("https://example.com/list")
    tables = find_tables(html)
    headers, rows = tables[0]["headers"], tables[0]["rows"]
    rows = dedupe(rows, key_cols=[0])
    to_csv("out.csv", headers, rows)
"""
import csv
import io
import re
import time
import urllib.robotparser
from urllib.parse import urljoin, urlparse

import requests
from lxml import html as lxml_html

DEFAULT_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0 Safari/537.36 TableHarvester/1.0"
)

MAX_HTML_BYTES = 8 * 1024 * 1024  # 8MB，再大基本不是普通网页


class HarvestError(Exception):
    """抓取/解析过程中的可预期错误，消息直接面向用户。"""


# --------------------------------------------------------------------------
# 抓取
# --------------------------------------------------------------------------
def fetch_html(url, timeout=20, headers=None, sleep=0.0):
    """
    抓一个页面，返回解码好的 HTML 文本。

    编码处理：很多中文站 Content-Type 不带 charset，requests 会默认猜 ISO-8859-1，
    直接用会全是乱码，所以优先用 apparent_encoding。
    """
    if not re.match(r"^https?://", url or ""):
        raise HarvestError("地址要以 http:// 或 https:// 开头")

    if sleep:
        time.sleep(sleep)

    hdrs = {"User-Agent": DEFAULT_UA, "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
    if headers:
        hdrs.update(headers)

    try:
        resp = requests.get(url, headers=hdrs, timeout=timeout, stream=True, allow_redirects=True)
    except requests.exceptions.Timeout:
        raise HarvestError("请求超时（%d 秒）：%s" % (timeout, url))
    except requests.exceptions.SSLError:
        raise HarvestError("SSL 证书校验失败：%s" % url)
    except requests.exceptions.ConnectionError as e:
        raise HarvestError("连不上：%s（%s）" % (url, e.__class__.__name__))
    except requests.exceptions.RequestException as e:
        raise HarvestError("请求失败：%s" % e)

    if resp.status_code >= 400:
        raise HarvestError("服务器返回 HTTP %d：%s" % (resp.status_code, resp.url))

    ctype = resp.headers.get("Content-Type", "")
    if "html" not in ctype and "xml" not in ctype and ctype:
        raise HarvestError("这不是网页（Content-Type: %s），本工具只处理 HTML" % ctype)

    chunks = []
    total = 0
    for chunk in resp.iter_content(65536):
        total += len(chunk)
        if total > MAX_HTML_BYTES:
            break
        chunks.append(chunk)
    raw = b"".join(chunks)

    # 优先按 meta/头里声明的编码，不行再靠猜
    resp._content = raw
    if resp.encoding and resp.encoding.lower() not in ("iso-8859-1", "ascii"):
        try:
            text = raw.decode(resp.encoding, errors="replace")
        except LookupError:
            text = raw.decode(resp.apparent_encoding or "utf-8", errors="replace")
    else:
        text = raw.decode(resp.apparent_encoding or "utf-8", errors="replace")

    return text


def robots_allows(url, user_agent="TableHarvester"):
    """粗略检查 robots.txt 是否禁止抓这个地址。失败时按「允许」处理并返回 True。"""
    try:
        parsed = urlparse(url)
        robots_url = "%s://%s/robots.txt" % (parsed.scheme, parsed.netloc)
        rp = urllib.robotparser.RobotFileParser()
        rp.set_url(robots_url)
        rp.read()
        return rp.can_fetch(user_agent, url)
    except Exception:
        return True


# --------------------------------------------------------------------------
# 解析
# --------------------------------------------------------------------------
def _text(el):
    """取元素文本，压缩所有空白。"""
    return " ".join(el.text_content().split())


def _row_values(tr):
    """取一行的单元格值；colspan 会按跨度重复，保证列对齐。"""
    values = []
    for cell in tr.xpath("./th|./td"):
        text = _text(cell)
        try:
            span = max(1, int(cell.get("colspan") or 1))
        except ValueError:
            span = 1
        values.extend([text] * span)
    return values


def parse_table(table_el):
    """
    解析单个 <table>，返回 (headers, rows)。

    表头判定顺序：
      1. 有 <thead> 就用 thead 第一行
      2. 否则第一行的单元格全是 <th>，就把它当表头
      3. 都不满足 → headers 为空，第一行也算数据
    """
    headers = []
    trs = table_el.xpath("./thead/tr")
    body_rows = table_el.xpath("./tbody/tr") or table_el.xpath("./tr")

    if trs:
        headers = _row_values(trs[0])
    elif body_rows:
        first_cells = body_rows[0].xpath("./th|./td")
        if first_cells and all(c.tag == "th" for c in first_cells):
            headers = _row_values(body_rows[0])
            body_rows = body_rows[1:]

    rows = []
    for tr in body_rows:
        vals = _row_values(tr)
        if not any(vals):
            continue
        if headers and vals == headers:
            continue  # 有些站会在 tbody 里重复一遍表头
        rows.append(vals)

    # 列数对齐：短行补空，长行截断到最宽
    width = max([len(headers)] + [len(r) for r in rows]) if (headers or rows) else 0
    headers = headers + [""] * (width - len(headers))
    rows = [r + [""] * (width - len(r)) for r in rows]

    return headers, rows


def find_tables(html, min_rows=1):
    """
    找出页面里所有有数据的表格，按「行数×列数」从大到小排序（最大的通常是主表）。
    返回 [{"index": 原序, "headers": [...], "rows": [...], "score": n}]。
    """
    try:
        doc = lxml_html.fromstring(html)
    except Exception as e:
        raise HarvestError("HTML 解析失败：%s" % e)

    found = []
    for i, table in enumerate(doc.xpath("//table")):
        headers, rows = parse_table(table)
        if len(rows) < min_rows:
            continue
        width = len(headers) or (len(rows[0]) if rows else 0)
        found.append(
            {
                "index": i,
                "headers": headers,
                "rows": rows,
                "score": len(rows) * max(width, 1),
            }
        )

    found.sort(key=lambda t: t["score"], reverse=True)
    return found


def has_layout_table(html):
    """有些老站用 table 做页面布局，会产生 1 行 N 列或 N 行 1 列的无意义表格。"""
    tables = find_tables(html)
    for t in tables:
        width = len(t["headers"]) or (len(t["rows"][0]) if t["rows"] else 0)
        if len(t["rows"]) >= 2 and width >= 2:
            return False
    return bool(tables)


# --------------------------------------------------------------------------
# 清洗 / 去重
# --------------------------------------------------------------------------
def clean_rows(rows):
    """去掉完全空白的行。"""
    return [r for r in rows if any(c.strip() for c in r)]


def dedupe(rows, key_cols=None):
    """
    保序去重。
    key_cols=None  → 整行完全一样才算重复
    key_cols=[0]   → 只按第 1 列判重（保留第一次出现的那行）
    """
    seen = set()
    out = []
    for r in rows:
        if key_cols:
            key = tuple((r[i] if i < len(r) else "") for i in key_cols)
        else:
            key = tuple(r)
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def merge_tables(tables):
    """
    把多个表格合并成一张（要求列数一致，否则跳过不一致的）。
    返回 (headers, rows, 被跳过的表数量)。
    """
    if not tables:
        return [], [], 0
    headers = tables[0]["headers"]
    width = len(headers) or len(tables[0]["rows"][0])
    rows = list(tables[0]["rows"])
    skipped = 0
    for t in tables[1:]:
        w = len(t["headers"]) or (len(t["rows"][0]) if t["rows"] else 0)
        if w != width:
            skipped += 1
            continue
        rows.extend(t["rows"])
    return headers, rows, skipped


# --------------------------------------------------------------------------
# 导出
# --------------------------------------------------------------------------
def to_csv_text(headers, rows):
    """生成 CSV 文本（Excel 打开不乱码：带 UTF-8 BOM）。"""
    buf = io.StringIO()
    writer = csv.writer(buf)
    if headers and any(headers):
        writer.writerow(headers)
    for r in rows:
        writer.writerow(r)
    return "\ufeff" + buf.getvalue()


def to_csv(path, headers, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        if headers and any(headers):
            writer.writerow(headers)
        writer.writerows(rows)
    return path


def to_excel(path, headers, rows, sheet_name="数据"):
    """
    导出 xlsx：表头加粗、冻结首行、自动列宽、加自动筛选。
    需要 openpyxl。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31] or "数据"

    has_header = bool(headers) and any(headers)
    start = 1
    if has_header:
        ws.append(headers)
        thin = Side(style="thin", color="D9DEE7")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="3B5BFD")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        ws.freeze_panes = "A2"
        start = 2

    for r in rows:
        ws.append(r)

    # 列宽按内容估算（中文按 2 个字符宽）
    width = max([len(headers)] + [len(r) for r in rows]) if (headers or rows) else 0
    for col in range(1, width + 1):
        longest = 0
        for row_idx in range(start, ws.max_row + 1):
            v = ws.cell(row=row_idx, column=col).value
            if v is None:
                continue
            s = str(v)
            size = sum(2 if ord(ch) > 127 else 1 for ch in s)
            longest = max(longest, size)
        ws.column_dimensions[get_column_letter(col)].width = min(max(longest + 2, 10), 60)

    if has_header and ws.max_row > 1:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(width), ws.max_row)

    for row in ws.iter_rows(min_row=start, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=False)

    wb.save(path)
    return path


# --------------------------------------------------------------------------
# 一步到位
# --------------------------------------------------------------------------
def harvest(url, table_index=0, key_cols=None, do_dedupe=True, respect_robots=False):
    """
    抓取 + 解析 + 去重，返回 dict：
        {"url", "headers", "rows", "table_count", "raw_rows", "dropped"}
    异常：HarvestError
    """
    if respect_robots and not robots_allows(url):
        raise HarvestError("robots.txt 不允许抓取该地址")

    html = fetch_html(url)
    tables = find_tables(html)
    if not tables:
        raise HarvestError("这个页面里没有找到带数据的表格（可能是 JS 动态渲染的，本工具不执行 JS）")

    if table_index >= len(tables):
        raise HarvestError(
            "页面里只有 %d 个表格，选不了第 %d 个" % (len(tables), table_index + 1)
        )

    picked = tables[table_index]
    raw_rows = picked["rows"]
    rows = clean_rows(raw_rows)
    if do_dedupe:
        rows = dedupe(rows, key_cols)

    return {
        "url": url,
        "headers": picked["headers"],
        "rows": rows,
        "raw_rows": raw_rows,
        "dropped": len(raw_rows) - len(rows),
        "table_count": len(tables),
        "all_tables": tables,
    }
