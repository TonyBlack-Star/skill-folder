# -*- coding: utf-8 -*-
"""
Shopify 流量表 + 销售表整合 —— 共享 ETL 模块

供 integrate.py（全量重算）与 update.py（增量更新）复用，避免两份解析逻辑分歧。

指标口径（与 Shopify 报表对齐）：
- 计数列 COUNTS：同天同产品求和
- DERIVED 比率：由合计后的分子/分母精确重算，**不沿用原表行比率，也不对比率取算术平均**
- Bounce rate：源表无跳出 Session 计数，按 Sessions 加权平均（近似）
- Average session duration：按 Sessions 加权平均

本模块不含任何硬编码目录路径，工作目录由调用方传入。
"""
import csv
import os
from collections import defaultdict

# 语言/市场前缀（多语言店铺的 URL 会出现这些前缀，需剥离后才能取到 handle）
LOCALES = {"fr-fr", "de-de", "en-de", "en-gb", "it-it", "es-es", "nl-nl"}

COUNTS = ["Online store visitors", "Sessions", "Sessions with cart additions",
          "Sessions that reached checkout", "Sessions that completed checkout"]
DERIVED = {  # 输出比率 -> (分子计数列, 分母计数列)
    "Added to cart rate": ("Sessions with cart additions", "Sessions"),
    "Reached checkout rate": ("Sessions that reached checkout", "Sessions"),
    "Checkout conversion rate": ("Sessions that completed checkout", "Sessions that reached checkout"),
    "Conversion rate": ("Sessions that completed checkout", "Sessions"),
}
BOUNCE = "Bounce rate"
DURATION = "Average session duration"
RATE_COLS = list(DERIVED.keys()) + [BOUNCE]   # 需以百分比显示的列
OUT_COLS = (["Day", "Product ID", "handle"] + COUNTS + list(DERIVED.keys())
            + [BOUNCE, DURATION, "Net items sold", "Total sales"])

DEFAULT_MAP_NAME = "路径与ID匹配规则.txt"


def load_mapping(mapfile):
    """解析 handle -> Product ID 映射文件。

    支持两种写法（宽松解析，取第一个 | 前为 handle，含 'Product ID:' 的段为 ID）：
        auto-heat-press-a100 | SALES | Product ID: 9870162985208
        auto-heat-press-a100,9870162985208
    """
    mapping = {}
    with open(mapfile, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.replace("\t", "|").split("|")]
            if len(parts) < 2:
                parts = [p.strip() for p in line.split(",")]
            handle = parts[0].strip("/")
            pid = ""
            for seg in parts[1:]:
                if "Product ID" in seg:
                    pid = seg.split(":", 1)[1].strip()
                    break
            if not pid and len(parts) > 1:
                pid = parts[1]
            if handle and pid:
                mapping[handle] = pid
    return mapping


def extract_handle(raw):
    """从 Landing page path 提取 handle；非产品路径返回 None。

    /fr-fr/products/auto-heat-press-a100?x=1 -> auto-heat-press-a100
    """
    if not raw:
        return None
    p = raw.strip().split("?")[0].split("#")[0].strip().strip("/")
    segs = [s for s in p.split("/") if s]
    i = 0
    while i < len(segs) and segs[i].lower() in LOCALES:
        i += 1
    if i < len(segs) and segs[i] == "products" and i + 1 < len(segs):
        return segs[i + 1]
    return None


def to_float(v):
    v = (v or "").strip().replace(",", "")
    if v == "":
        return None
    try:
        return float(v)
    except ValueError:
        return None


def detect_kind(header):
    """按表头判定文件类型：traffic / sales / None。**严禁按文件名判断。**"""
    h = ",".join(header)
    if "Landing page path" in h:
        return "traffic"
    if "Product ID" in h:
        return "sales"
    return None


def aggregate_files(file_paths, mapping):
    """解析一批 CSV 并聚合到 (Day, Product ID) 粒度。

    返回 (traffic_agg, sales_agg, anomalies, counters, src_files, traffic_keys, sales_keys)
    - traffic_agg[(day,pid)]: 计数列求和 + __bw/__bs（跳出加权）+ __dw/__ds（时长加权）
    - sales_agg[(day,pid)]:   Net items sold / Total sales（同一表内 ID 前向填充）
    - 同一文件内 (day, 完整path) 重复：完全一致则去重，指标冲突则排除并上报
    """
    txt_ids = set(mapping.values())
    traffic_agg = defaultdict(lambda: {c: 0.0 for c in COUNTS})
    sales_agg = defaultdict(lambda: {"Net items sold": 0.0, "Total sales": 0.0})
    seen_keys = {}
    anomalies = []
    counters = {"nonproduct": 0, "unmatched": 0}
    src_files = []

    for fp in file_paths:
        name = os.path.basename(fp)
        with open(fp, encoding="utf-8-sig", errors="replace") as f:
            r = csv.reader(f)
            header = next(r, [])
            data = list(r)

        kind = detect_kind(header)
        if kind is None:
            continue

        if kind == "traffic":
            src_files.append(("traffic", name))
            pidx = header.index("Landing page path")
            cidx = {c: header.index(c) for c in COUNTS if c in header}
            bidx = header.index(BOUNCE) if BOUNCE in header else None
            didx = header.index(DURATION) if DURATION in header else None
            sidx = header.index("Sessions") if "Sessions" in header else None
            for row in data:
                if not row or not row[0].strip():
                    continue
                day = row[0].strip()
                rawpath = row[pidx].strip() if len(row) > pidx else ""
                handle = extract_handle(rawpath)
                if handle is None:
                    counters["nonproduct"] += 1
                    continue
                pid = mapping.get(handle)
                if not pid:
                    counters["unmatched"] += 1
                    continue
                dkey = (name, day, rawpath)
                if dkey in seen_keys:
                    if seen_keys[dkey] == tuple(row):
                        anomalies.append((name, day, rawpath, "EXACT DUPLICATE（已去重，仅计一次）"))
                    else:
                        anomalies.append((name, day, rawpath, "CONFLICT 同 path 指标不一致（已排除，待确认）"))
                    continue
                seen_keys[dkey] = tuple(row)
                a = traffic_agg[(day, pid)]
                for c, idx in cidx.items():
                    val = to_float(row[idx]) if idx < len(row) else None
                    if val is not None:
                        a[c] += val
                if bidx is not None and bidx < len(row) and sidx is not None:
                    bv, sv = to_float(row[bidx]), to_float(row[sidx])
                    if bv is not None and sv is not None:
                        a["__bw"] = a.get("__bw", 0) + bv * sv
                        a["__bs"] = a.get("__bs", 0) + sv
                if didx is not None and didx < len(row) and sidx is not None:
                    dv, sv = to_float(row[didx]), to_float(row[sidx])
                    if dv is not None and sv is not None:
                        a["__dw"] = a.get("__dw", 0) + dv * sv
                        a["__ds"] = a.get("__ds", 0) + sv

        else:  # sales
            src_files.append(("sales", name))
            pidx = header.index("Product ID")
            nidx = header.index("Net items sold") if "Net items sold" in header else None
            tidx = header.index("Total sales") if "Total sales" in header else None
            ids = {row[pidx].strip() for row in data
                   if len(row) > pidx and row[pidx].strip()}
            if len(ids) > 1:
                anomalies.append((name, "-", "-", "销售表含多个 ID: %s" % sorted(ids)))
            fid = next(iter(ids)) if ids else None
            if not fid:
                continue
            for row in data:
                if not row or not row[0].strip():
                    continue
                day = row[0].strip()
                # 同一销售表内 Product ID 一致：空白单元格前向填充该表唯一 ID
                pid = row[pidx].strip() if len(row) > pidx and row[pidx].strip() else fid
                if pid not in txt_ids:
                    counters["unmatched"] += 1
                    continue
                a = sales_agg[(day, pid)]
                if nidx is not None and nidx < len(row):
                    v = to_float(row[nidx])
                    a["Net items sold"] += v if v is not None else 0
                if tidx is not None and tidx < len(row):
                    v = to_float(row[tidx])
                    a["Total sales"] += v if v is not None else 0

    return (traffic_agg, sales_agg, anomalies, counters, src_files,
            set(traffic_agg.keys()), set(sales_agg.keys()))


def build_records(traffic_agg, sales_agg, mapping):
    """聚合结果 -> 行记录列表（比率底层为小数，由 write_outputs 统一格式化）。"""
    txt_ids = set(mapping.values())
    id_to_handle = {v: k for k, v in mapping.items()}
    all_keys = {(d, p) for (d, p) in (set(traffic_agg) | set(sales_agg)) if p in txt_ids}
    rows = []
    for (day, pid) in sorted(all_keys):
        t = traffic_agg.get((day, pid), {})
        s = sales_agg.get((day, pid), {"Net items sold": 0.0, "Total sales": 0.0})
        out = {"Day": day, "Product ID": pid, "handle": id_to_handle.get(pid, "")}
        for c in COUNTS:
            out[c] = t.get(c, 0.0)
        for rate, (num, den) in DERIVED.items():
            n, d = t.get(num, 0.0), t.get(den, 0.0)
            out[rate] = round(n / d, 6) if d else ""
        out[BOUNCE] = round(t["__bw"] / t["__bs"], 6) if t.get("__bs", 0) else ""
        out[DURATION] = round(t["__dw"] / t["__ds"], 6) if t.get("__ds", 0) else ""
        out["Net items sold"] = round(s["Net items sold"], 2)
        out["Total sales"] = round(s["Total sales"], 2)
        rows.append(out)
    return rows


def write_xlsx(records, mapping, xlsx_path):
    """写多子表 XLSX：每个产品一个 sheet，以清洗后的 handle 命名（Excel 表名不允许 /）。

    第 1 行标题（完整 landing path + Product ID），第 2 行表头（冻结），第 3 行起日度明细。
    比率列设 0.00% 显示格式，底层保留小数以便继续计算。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    sheet_cols = (["Day", "Product ID"] + COUNTS + list(DERIVED.keys())
                  + [BOUNCE, DURATION, "Net items sold", "Total sales"])
    wb = Workbook()
    wb.remove(wb.active)
    for handle, pid in mapping.items():
        ws = wb.create_sheet(title=handle[:31])
        ws.append(["Landing path: /products/%s" % handle, "Product ID: %s" % pid])
        ws["A1"].font = Font(bold=True)
        ws.append(sheet_cols)
        for c in range(1, len(sheet_cols) + 1):
            cell = ws.cell(row=2, column=c)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
        for r in sorted([r for r in records if r["Product ID"] == pid], key=lambda r: r["Day"]):
            ws.append([r.get(c, "") for c in sheet_cols])
        rate_idx = [sheet_cols.index(rc) + 1 for rc in RATE_COLS if rc in sheet_cols]
        for ridx in range(3, ws.max_row + 1):
            for ci in rate_idx:
                ws.cell(row=ridx, column=ci).number_format = "0.00%"
        ws.freeze_panes = "A3"
    wb.save(xlsx_path)


def read_xlsx(path, cols=None):
    """把 XLSX 读回为记录列表（子表名即 handle；比率底层为小数）。"""
    from openpyxl import load_workbook
    cols = cols or OUT_COLS
    wb = load_workbook(path)
    records = []
    for ws in wb.worksheets:
        header = [c.value for c in ws[2]]
        idx = {name: i for i, name in enumerate(header) if name}
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row or not row[0]:
                continue
            records.append({name: (row[idx[name]] if name in idx else "") for name in cols})
    return records
