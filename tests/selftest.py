# -*- coding: utf-8 -*-
"""
自动化测试（全部离线，不联网）
==============================
    python tests/selftest.py
"""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)

import harvester  # noqa: E402
from harvester import (  # noqa: E402
    HarvestError,
    clean_rows,
    dedupe,
    find_tables,
    merge_tables,
    parse_table,
    to_csv,
    to_csv_text,
    to_excel,
)

passed, failed = [], []


def check(name, cond, extra=""):
    if cond:
        passed.append(name)
        print("  PASS  %s" % name)
    else:
        failed.append(name)
        print("  FAIL  %s  %s" % (name, extra))


HTML_STANDARD = """
<html><body>
<h1>员工名单</h1>
<table id="t1">
  <thead><tr><th>姓名</th><th>部门</th><th>工号</th></tr></thead>
  <tbody>
    <tr><td>张三</td><td>技术部</td><td>A001</td></tr>
    <tr><td>李四</td><td>市场部</td><td>A002</td></tr>
    <tr><td>张三</td><td>技术部</td><td>A001</td></tr>
    <tr><td></td><td></td><td></td></tr>
  </tbody>
</table>
</body></html>
"""

HTML_TH_FIRST = """
<table>
  <tr><th>城市</th><th>人口</th></tr>
  <tr><td>上海</td><td>2487 万</td></tr>
  <tr><td>北京</td><td>2185 万</td></tr>
</table>
"""

HTML_NO_HEADER = """
<table>
  <tr><td>a</td><td>b</td></tr>
  <tr><td>c</td><td>d</td></tr>
</table>
"""

HTML_COLSPAN = """
<table>
  <thead><tr><th colspan="2">联系方式</th><th>备注</th></tr></thead>
  <tbody>
    <tr><td>电话</td><td>123</td><td>x</td></tr>
  </tbody>
</table>
"""

HTML_REPEAT_HEADER = """
<table>
  <thead><tr><th>名</th><th>值</th></tr></thead>
  <tbody>
    <tr><th>名</th><th>值</th></tr>
    <tr><td>a</td><td>1</td></tr>
  </tbody>
</table>
"""

HTML_TWO_TABLES = """
<html><body>
<table><tr><td>1</td></tr><tr><td>2</td></tr></table>
<table>
  <tr><th>a</th><th>b</th><th>c</th></tr>
  <tr><td>1</td><td>2</td><td>3</td></tr>
  <tr><td>4</td><td>5</td><td>6</td></tr>
  <tr><td>7</td><td>8</td><td>9</td></tr>
  <tr><td>10</td><td>11</td><td>12</td></tr>
</table>
</body></html>
"""

HTML_NO_TABLE = "<html><body><div>这里只有 div，没有表格</div></body></html>"


def el_of(html):
    from lxml import html as lx

    return lx.fromstring(html).xpath("//table")[0]


def run():
    # ---- 1. 标准表格 ----
    headers, rows = parse_table(el_of(HTML_STANDARD))
    check("标准表头解析", headers == ["姓名", "部门", "工号"], str(headers))
    check("标准数据行数（解析阶段已剔除全空行）", len(rows) == 3, str(len(rows)))
    rows_clean = clean_rows(rows)
    check("clean_rows 去掉全空行", len(rows_clean) == 3)

    # ---- 2. 首行 th 当表头 ----
    headers, rows = parse_table(el_of(HTML_TH_FIRST))
    check("首行 th 识别为表头", headers == ["城市", "人口"], str(headers))
    check("th 表头不算数据行", len(rows) == 2)

    # ---- 3. 无表头 ----
    headers, rows = parse_table(el_of(HTML_NO_HEADER))
    check("无表头时 headers 为空", headers == [""] * 2 or all(h == "" for h in headers), str(headers))
    check("无表头时首行算数据", len(rows) == 2 and rows[0] == ["a", "b"])

    # ---- 4. colspan 对齐 ----
    headers, rows = parse_table(el_of(HTML_COLSPAN))
    check("colspan 展开表头", headers == ["联系方式", "联系方式", "备注"], str(headers))
    check("colspan 行列数对齐", len(rows[0]) == 3, str(rows[0]))

    # ---- 5. tbody 重复表头被去掉 ----
    headers, rows = parse_table(el_of(HTML_REPEAT_HEADER))
    check("重复表头行被剔除", len(rows) == 1 and rows[0] == ["a", "1"], str(rows))

    # ---- 6. find_tables 排序 ----
    tables = find_tables(HTML_TWO_TABLES)
    check("找到两个表格", len(tables) == 2)
    check("大表排第一", len(tables[0]["rows"]) == 4 and len(tables[1]["rows"]) == 2,
          "%d vs %d" % (len(tables[0]["rows"]), len(tables[1]["rows"])))

    # ---- 7. 没有表格 ----
    check("无表格时返回空列表", find_tables(HTML_NO_TABLE) == [])

    # ---- 8. 去重 ----
    rows = [["a", "1"], ["b", "2"], ["a", "1"], ["a", "3"]]
    check("整行去重", dedupe(rows) == [["a", "1"], ["b", "2"], ["a", "3"]])
    check("按第一列去重", dedupe(rows, key_cols=[0]) == [["a", "1"], ["b", "2"]])
    check("去重保持原顺序", dedupe([["z"], ["a"], ["z"], ["m"]]) == [["z"], ["a"], ["m"]])

    # ---- 9. 合并多表 ----
    t1 = {"headers": ["x", "y"], "rows": [["1", "2"]]}
    t2 = {"headers": ["x", "y"], "rows": [["3", "4"]]}
    t3 = {"headers": ["x", "y", "z"], "rows": [["5", "6", "7"]]}
    h, r, skipped = merge_tables([t1, t2, t3])
    check("合并列数一致的表格", h == ["x", "y"] and r == [["1", "2"], ["3", "4"]])
    check("列数不一致的表被跳过", skipped == 1)

    # ---- 10. CSV ----
    txt = to_csv_text(["名", "值"], [["a", "1"], ["b", "2"]])
    check("CSV 带 UTF-8 BOM", txt.startswith("\ufeff"))
    check("CSV 有表头和数据", "名,值" in txt and "a,1" in txt)

    tmpdir = tempfile.mkdtemp()
    csv_path = os.path.join(tmpdir, "t.csv")
    to_csv(csv_path, ["名", "值"], [["a", "1"]])
    with open(csv_path, encoding="utf-8-sig") as f:
        content = f.read()
    check("CSV 落盘正确", "名,值" in content and "a,1" in content)

    # ---- 11. Excel ----
    xlsx_path = os.path.join(tmpdir, "t.xlsx")
    to_excel(xlsx_path, ["姓名", "部门"], [["张三", "技术部"], ["李四", "市场部"]])
    check("Excel 文件已生成", os.path.exists(xlsx_path) and os.path.getsize(xlsx_path) > 0)

    from openpyxl import load_workbook

    wb = load_workbook(xlsx_path)
    ws = wb.active
    check("Excel 表头正确", ws.cell(1, 1).value == "姓名" and ws.cell(1, 2).value == "部门")
    check("Excel 数据正确", ws.cell(2, 1).value == "张三" and ws.cell(3, 2).value == "市场部")
    check("Excel 冻结首行", ws.freeze_panes == "A2")
    check("Excel 有自动筛选", ws.auto_filter.ref is not None)
    check("Excel 表头加粗", ws.cell(1, 1).font.bold is True)
    check("Excel 列宽已设置", ws.column_dimensions["A"].width >= 10)

    # ---- 12. harvest 端到端（替换掉网络请求） ----
    real_fetch = harvester.fetch_html
    harvester.fetch_html = lambda url, **kw: HTML_STANDARD
    try:
        res = harvester.harvest("https://example.com/x")
        check("harvest 返回表头", res["headers"] == ["姓名", "部门", "工号"])
        check("harvest 去掉了重复行和空行", len(res["rows"]) == 2, str(len(res["rows"])))
        check("harvest 记录丢弃行数", res["dropped"] == 1, str(res["dropped"]))
        check("harvest 带上表格总数", res["table_count"] == 1)

        res2 = harvester.harvest("https://example.com/x", do_dedupe=False)
        check("harvest 可以关闭去重", len(res2["rows"]) == 3)

        harvester.fetch_html = lambda url, **kw: HTML_NO_TABLE
        try:
            harvester.harvest("https://example.com/x")
            check("没有表格时抛 HarvestError", False)
        except HarvestError as e:
            check("没有表格时抛 HarvestError", "没有找到" in str(e))

        harvester.fetch_html = lambda url, **kw: HTML_TWO_TABLES
        try:
            harvester.harvest("https://example.com/x", table_index=5)
            check("表序号越界时报错", False)
        except HarvestError as e:
            check("表序号越界时报错", "选不了" in str(e))
    finally:
        harvester.fetch_html = real_fetch

    # ---- 13. 参数校验 ----
    try:
        harvester.fetch_html("ftp://example.com")
        check("非 http 地址被拒", False)
    except HarvestError as e:
        check("非 http 地址被拒", "http" in str(e))

    print("\n%d passed, %d failed" % (len(passed), len(failed)))
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(run())
