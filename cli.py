# -*- coding: utf-8 -*-
"""
网页表格采集器 · 命令行版
========================
    # 抓第一个（数据量最大的）表格，存 CSV
    python cli.py https://example.com/list

    # 存 Excel，并同时列出页面上所有表格
    python cli.py https://example.com/list --list-tables --excel -o 名单.xlsx

    # 抓 3 个分页，合并后按第 1 列去重
    python cli.py "https://example.com/p1" "https://example.com/p2" "https://example.com/p3" --key 0 -o 合并.csv
"""
import argparse
import os
import sys

from harvester import HarvestError, dedupe, fetch_html, find_tables, merge_tables, to_csv, to_excel


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="把公开网页里的表格抓下来，去重后导出 CSV / Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("urls", nargs="+", help="一个或多个网址")
    ap.add_argument("--table", type=int, default=0, help="取第几个表格，从 0 开始（按数据量从大到小），默认 0")
    ap.add_argument("-o", "--out", default="data.csv", help="输出文件，默认 data.csv")
    ap.add_argument("--excel", action="store_true", help="导出 Excel（.xlsx）而不是 CSV")
    ap.add_argument("--no-dedupe", action="store_true", help="不去重")
    ap.add_argument("--key", type=int, nargs="+", default=None, help="按第几列判重，从 0 开始，可多列")
    ap.add_argument("--list-tables", action="store_true", help="只列出页面里的表格，不导出")
    ap.add_argument("--robots", action="store_true", help="检查 robots.txt，禁止抓的地址就不抓")
    ap.add_argument("--timeout", type=int, default=20, help="单页超时秒数，默认 20")
    args = ap.parse_args(argv)

    all_tables = []
    total_raw = 0

    for url in args.urls:
        print("→ 抓取 %s" % url)
        try:
            html = fetch_html(url, timeout=args.timeout)
        except HarvestError as e:
            print("  ✗ %s" % e)
            continue

        tables = find_tables(html)
        if not tables:
            print("  ✗ 页面里没有找到带数据的表格（可能是 JS 渲染的）")
            continue

        if args.list_tables:
            print("  页面里共 %d 个表格：" % len(tables))
            for i, t in enumerate(tables):
                width = len(t["headers"]) or (len(t["rows"][0]) if t["rows"] else 0)
                head = " / ".join(t["headers"][:4]) if any(t["headers"]) else "（无表头）"
                print("    [%d] %d 行 × %d 列   %s" % (i, len(t["rows"]), width, head[:60]))
            continue

        if args.table >= len(tables):
            print("  ✗ 只有 %d 个表格，选不了第 %d 个" % (len(tables), args.table + 1))
            continue

        picked = tables[args.table]
        print("  取第 %d 个表格：%d 行" % (args.table, len(picked["rows"])))
        all_tables.append(picked)
        total_raw += len(picked["rows"])

    if args.list_tables:
        return 0

    if not all_tables:
        print("没有抓到任何表格。")
        return 1

    if len(all_tables) == 1:
        headers = all_tables[0]["headers"]
        rows = list(all_tables[0]["rows"])
    else:
        headers, rows, skipped = merge_tables(all_tables)
        if skipped:
            print("  注意：%d 个表格列数不一致，已跳过" % skipped)
        print("  合并 %d 个表格" % (len(all_tables) - skipped))

    rows = [r for r in rows if any(c.strip() for c in r)]
    if not args.no_dedupe:
        before = len(rows)
        rows = dedupe(rows, args.key)
        print("  去重：%d → %d 行" % (before, len(rows)))

    out = os.path.abspath(args.out)
    try:
        if args.excel or out.lower().endswith((".xlsx", ".xlsm")):
            if not out.lower().endswith((".xlsx", ".xlsm")):
                out += ".xlsx"
            to_excel(out, headers, rows)
        else:
            if not out.lower().endswith(".csv"):
                out += ".csv"
            to_csv(out, headers, rows)
    except ImportError:
        print("✗ 导出 Excel 需要 openpyxl：pip install openpyxl")
        return 1

    width = len(headers) or (len(rows[0]) if rows else 0)
    print("\n✓ 完成：%d 行 × %d 列" % (len(rows), width))
    print("  文件：%s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
