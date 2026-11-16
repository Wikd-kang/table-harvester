# -*- coding: utf-8 -*-
"""
网页表格采集器 · 网页版
======================
两阶段：
  1. POST /preview  抓取页面，把找到的表格缓存到内存，渲染预览让人挑
  2. POST /export   按选的表格导出 CSV / Excel 并下载

缓存只在内存里，10 分钟过期，最多 50 个任务，重启即清空。
"""
import io
import time
import uuid
from datetime import datetime

from flask import Flask, flash, redirect, render_template, request, send_file, url_for

from harvester import (
    HarvestError,
    dedupe,
    fetch_html,
    find_tables,
    merge_tables,
    to_csv_text,
    to_excel,
)

app = Flask(__name__)
app.secret_key = "table-harvester-local-only"

JOBS = {}
JOB_TTL = 600          # 秒
MAX_JOBS = 50
MAX_URLS = 20


def _gc():
    now = time.time()
    for k in [k for k, v in JOBS.items() if now - v["created"] > JOB_TTL]:
        JOBS.pop(k, None)
    while len(JOBS) > MAX_JOBS:
        JOBS.pop(next(iter(JOBS)))


def _parse_urls(raw):
    urls = []
    for line in (raw or "").splitlines():
        line = line.strip()
        if not line:
            continue
        if not line.startswith(("http://", "https://")):
            raise HarvestError("这一行不是网址：%s" % line[:60])
        urls.append(line)
    if not urls:
        raise HarvestError("请至少填一个网址")
    if len(urls) > MAX_URLS:
        raise HarvestError("一次最多 %d 个网址" % MAX_URLS)
    return urls


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/preview", methods=["POST"])
def preview():
    raw = request.form.get("urls", "")
    try:
        urls = _parse_urls(raw)
    except HarvestError as e:
        flash(str(e), "err")
        return redirect(url_for("index"))

    _gc()
    job_id = uuid.uuid4().hex
    pages = []   # [{"url":..., "tables":[...] 或 "error":...}]

    for url in urls:
        try:
            html = fetch_html(url, timeout=20)
            tables = find_tables(html)
            if not tables:
                pages.append({"url": url, "error": "没有找到带数据的表格（可能是 JS 渲染的）"})
            else:
                pages.append({"url": url, "tables": tables})
        except HarvestError as e:
            pages.append({"url": url, "error": str(e)})

    ok_pages = [p for p in pages if p.get("tables")]
    if not ok_pages:
        JOBS[job_id] = {"created": time.time(), "pages": pages}
        flash("一个表格都没抓到，看看下面的报错。", "err")
        return render_template("preview.html", job_id=job_id, pages=pages, empty=True)

    JOBS[job_id] = {"created": time.time(), "pages": pages}
    return render_template("preview.html", job_id=job_id, pages=pages, empty=False)


@app.route("/export", methods=["POST"])
def export():
    job_id = request.form.get("job_id", "")
    job = JOBS.get(job_id)
    if not job:
        flash("数据已过期（缓存 10 分钟），请重新抓一次。", "err")
        return redirect(url_for("index"))

    try:
        table_index = int(request.form.get("table_index", "0"))
    except ValueError:
        table_index = 0

    do_dedupe = bool(request.form.get("dedupe"))
    fmt = request.form.get("fmt", "csv")

    key_cols = None
    key_raw = (request.form.get("key") or "").strip()
    if key_raw:
        try:
            key_cols = [int(x.strip()) for x in key_raw.split(",") if x.strip()]
        except ValueError:
            flash("去重列要填列号，比如 0 或 0,2", "err")
            return redirect(url_for("index"))

    picked = []
    for p in job["pages"]:
        tables = p.get("tables") or []
        if table_index < len(tables):
            picked.append(tables[table_index])

    if not picked:
        flash("选的那个表格序号在页面里不存在。", "err")
        return redirect(url_for("index"))

    if len(picked) == 1:
        headers = picked[0]["headers"]
        rows = list(picked[0]["rows"])
    else:
        headers, rows, _skipped = merge_tables(picked)

    rows = [r for r in rows if any(c.strip() for c in r)]
    if do_dedupe:
        rows = dedupe(rows, key_cols)
    if not rows:
        flash("这个表格没有数据行。", "err")
        return redirect(url_for("index"))

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    if fmt == "excel":
        import tempfile, os
        fd, tmp = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        to_excel(tmp, headers, rows)
        with open(tmp, "rb") as f:
            data = f.read()
        os.remove(tmp)
        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="table-%s.xlsx" % stamp,
        )

    data = to_csv_text(headers, rows).encode("utf-8")
    return send_file(
        io.BytesIO(data),
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name="table-%s.csv" % stamp,
    )


if __name__ == "__main__":
    app.run(debug=True, port=5002)
