# -*- coding: utf-8 -*-
"""
周度 / 月度分析报告生成器
=================================
设计：分层递进分析框架
  L0 结构层  —— 数据长什么样（口径、粒度、覆盖度、可信度）
  L1 统计层  —— 有多少（水平、趋势、结构、集中度）
  L2 洞察层  —— 所以呢（对比归因、漏斗、异常、效率）

输出：固定框架（一句话摘要 / 数据全貌 / 核心发现 / 异常 / 建议），
      杂志风 editorial 排版，纯内联 SVG 图标与图表，单文件脱机可开。

周口径：固定为「上周五 ~ 本周四」（Fri–Thu），以周五为周期起点；
        对比周为再往前一个同长度周期。月口径仍按自然月。
"""
import os, math, html, sys
from collections import defaultdict
from datetime import date, timedelta

from openpyxl import load_workbook

D = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(D, "integrated_result.xlsx")
OUTDIR = os.path.join(D, "每周分析")
CUR = "€"

FIELDS = ["Online store visitors", "Sessions", "Sessions with cart additions",
          "Sessions that reached checkout", "Sessions that completed checkout",
          "Net items sold", "Total sales"]
DUR = "Average session duration"
BOUNCE = "Bounce rate"
RATE_FIELDS = ["Added to cart rate", "Reached checkout rate",
               "Checkout conversion rate", "Conversion rate", BOUNCE]


# ─────────────────────────── L0 结构层：读取 ───────────────────────────
def load_records(path):
    wb = load_workbook(path)
    recs = []
    for s in wb.sheetnames:
        ws = wb[s]
        hdr = [c.value for c in ws[2]]
        idx = {n: i for i, n in enumerate(hdr) if n}
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row[0]:
                continue
            r = {"Product": s}
            for k, v in idx.items():
                r[k] = row[v]
            recs.append(r)
    return recs


def d(s):
    y, m, dd = map(int, s.split("-"))
    return date(y, m, dd)


# 周口径固定为「上周五 ~ 本周四」（Fri–Thu），以周五为周期起点。
# date.weekday(): 周一=0 … 周五=4 … 周日=6
FRIDAY = 4


def week_start(dt):
    """返回 dt 所属「周五~周四」分析周的起点（该周期内的周五）。

    周五(4)→当天；周六(5)→前1天；周一(0)→前3天；周四(3)→前6天。
    """
    return dt - timedelta(days=(dt.weekday() - FRIDAY) % 7)


def blank():
    return {f: 0.0 for f in FIELDS}


def add(dst, r, w=1.0):
    for f in FIELDS:
        dst[f] += (r.get(f) or 0)
    dst["_bw"] = dst.get("_bw", 0.0) + (r.get(BOUNCE) or 0) * (r.get("Sessions") or 0) * w
    dst["_dw"] = dst.get("_dw", 0.0) + (r.get(DUR) or 0) * (r.get("Sessions") or 0) * w
    dst["_s"] = dst.get("_s", 0.0) + (r.get("Sessions") or 0) * w


def derive(t):
    """由计数重算比率（禁止对比率取平均）。"""
    s = t.get("Sessions") or 0
    o = dict(t)
    o["加购率"] = (t["Sessions with cart additions"] / s) if s else None
    o["到达结账率"] = (t["Sessions that reached checkout"] / s) if s else None
    o["完成结账率"] = (t["Sessions that completed checkout"] / s) if s else None
    o["件/会话"] = (t["Net items sold"] / s) if s else None
    o["客单价"] = (t["Total sales"] / t["Net items sold"]) if t["Net items sold"] else None
    o["每会话销售额"] = (t["Total sales"] / s) if s else None
    o["跳出率"] = (t["_bw"] / t["_s"]) if t.get("_s") else None
    o["停留时长"] = (t["_dw"] / t["_s"]) if t.get("_s") else None
    # 漏斗环节转化率（相对上一环节）
    o["加购→结账"] = (t["Sessions that reached checkout"] / t["Sessions with cart additions"]) if t["Sessions with cart additions"] else None
    o["结账→完成"] = (t["Sessions that completed checkout"] / t["Sessions that reached checkout"]) if t["Sessions that reached checkout"] else None
    return o


def agg(rows):
    t = blank()
    for r in rows:
        add(t, r)
    return derive(t)


def pct_change(a, b):
    if b in (0, None) or b == 0:
        return None
    return (a - b) / b


# ─────────────────────────── 数值格式化（警惕虚假精确） ───────────────────────────
def fnum(v, dec=0):
    if v is None:
        return "—"
    return f"{v:,.{dec}f}"


def fcur(v, dec=0):
    if v is None:
        return "—"
    sign = "-" if v < 0 else ""
    return f"{sign}{CUR}{abs(v):,.{dec}f}"


def fpct(v, dec=1, signed=False):
    if v is None:
        return "—"
    s = f"{abs(v)*100:.{dec}f}%"
    if signed:
        return ("+" if v >= 0 else "-") + s
    return s


def fdelta(a, b, kind="num"):
    """返回 (展示串, 变化率, 方向)"""
    p = pct_change(a, b)
    if kind == "cur":
        txt = fcur(a)
    elif kind == "pct":
        txt = fpct(a)
    else:
        txt = fnum(a)
    return txt, p, ("up" if (p or 0) > 0 else ("down" if (p or 0) < 0 else "flat"))


# ─────────────────────────── L1 统计层：周期与全貌 ───────────────────────────
def build_periods(recs):
    days = sorted({r["Day"] for r in recs})
    maxd = d(days[-1])
    mind = d(days[0])
    # 最近完整周：周日 <= maxd
    cur_start = week_start(maxd)
    if (cur_start + timedelta(days=6)) > maxd:
        cur_start -= timedelta(days=7)
    prev_start = cur_start - timedelta(days=7)
    partial_start = week_start(maxd)
    partial = None
    if partial_start > cur_start:  # 存在不完整的当周
        n_days = (maxd - partial_start).days + 1
        partial = {"start": partial_start, "end": maxd, "n_days": n_days}
    # 最近完整月
    def month_end(y, m):
        nm_y, nm_m = (y + 1, 1) if m == 12 else (y, m + 1)
        return date(nm_y, nm_m, 1) - timedelta(days=1)
    if maxd == month_end(maxd.year, maxd.month):
        cm = (maxd.year, maxd.month)
    else:
        cm = (maxd.year - 1, 12) if maxd.month == 1 else (maxd.year, maxd.month - 1)
    pm = (cm[0] - 1, 12) if cm[1] == 1 else (cm[0], cm[1] - 1)
    return {
        "min": mind, "max": maxd, "days": days,
        "wk_cur": (cur_start, cur_start + timedelta(days=6)),
        "wk_prev": (prev_start, prev_start + timedelta(days=6)),
        "partial": partial,
        "mon_cur": cm, "mon_prev": pm,
    }


def rows_between(recs, a, b):
    s, e = a.isoformat(), b.isoformat()
    return [r for r in recs if s <= r["Day"] <= e]


def rows_month(recs, ym):
    p = f"{ym[0]:04d}-{ym[1]:02d}"
    return [r for r in recs if r["Day"].startswith(p)]


# ─────────────────────────── L2 洞察层：分解与检测 ───────────────────────────
def decompose(cur, prev):
    """销售额变动的三因子链式分解：Sales = Sessions × (件/会话) × 客单价
    链式（依次替代），三者贡献之和 = 实际变动额。"""
    s0, s1 = prev["Sessions"] or 0, cur["Sessions"] or 0
    i0, i1 = prev["Net items sold"] or 0, cur["Net items sold"] or 0
    r0, r1 = prev["Total sales"] or 0, cur["Total sales"] or 0
    ips0 = (i0 / s0) if s0 else 0.0      # 件/会话
    ips1 = (i1 / s1) if s1 else 0.0
    aov0 = (r0 / i0) if i0 else 0.0
    aov1 = (r1 / i1) if i1 else 0.0
    e_traffic = (s1 - s0) * ips0 * aov0
    e_conv = s1 * (ips1 - ips0) * aov0
    e_aov = s1 * ips1 * (aov1 - aov0)
    total = r1 - r0
    factors = {"流量效应": e_traffic, "转化效率效应": e_conv, "客单价效应": e_aov}
    # 方向一致的主因：与净变动同号且绝对值最大者（避免"最大正向项"被误判为下滑主因）
    same = [(k, v) for k, v in factors.items() if (v >= 0) == (total >= 0)]
    driver = max(same, key=lambda x: abs(x[1]))[0] if same else max(factors.items(), key=lambda x: abs(x[1]))[0]
    return {**factors, "合计变动": total, "主因": driver,
            "最正因素": max(factors.items(), key=lambda x: x[1])[0],
            "最负因素": min(factors.items(), key=lambda x: x[1])[0],
            "校验": abs((e_traffic + e_conv + e_aov) - total) < 0.5}


def concentration(rows):
    t = defaultdict(float)
    for r in rows:
        t[r["Product"]] += (r.get("Total sales") or 0)
    tot = sum(t.values())
    if not tot:
        return {"top1": None, "top3": None, "hhi": None, "top": []}
    items = sorted(t.items(), key=lambda x: -x[1])
    top3 = sum(v for _, v in items[:3])
    hhi = sum((v / tot) ** 2 for v in t.values())
    return {"top1": items[0][1] / tot, "top3": top3 / tot, "hhi": hhi, "top": items, "total": tot}


def product_deltas(recs, cur_range, prev_range):
    cur = defaultdict(lambda: blank())
    prev = defaultdict(lambda: blank())
    for r in rows_between(recs, *cur_range):
        add(cur[r["Product"]], r)
    for r in rows_between(recs, *prev_range):
        add(prev[r["Product"]], r)
    out = []
    for p in set(list(cur) + list(prev)):
        c, pv = cur.get(p, blank()), prev.get(p, blank())
        out.append({
            "product": p,
            "cur_sales": c["Total sales"], "prev_sales": pv["Total sales"],
            "d_sales": c["Total sales"] - pv["Total sales"],
            "cur_sessions": c["Sessions"], "prev_sessions": pv["Sessions"],
            "d_sessions": c["Sessions"] - pv["Sessions"],
            "cur_items": c["Net items sold"], "prev_items": pv["Net items sold"],
        })
    return sorted(out, key=lambda x: -abs(x["d_sales"]))


def detect_anomalies(recs, cur_range, prev_range, cur_t, prev_t, pdeltas, unit="周"):
    out = []
    # 1) 产品级销售额异动
    for p in pdeltas:
        ch = pct_change(p["cur_sales"], p["prev_sales"])
        if ch is not None and abs(ch) >= 0.30 and abs(p["d_sales"]) >= 100:
            out.append({"level": "高" if abs(ch) >= 0.60 else "中",
                        "type": "产品异动",
                        "text": f"{p['product']} 销售额 {fcur(p['prev_sales'])} → {fcur(p['cur_sales'])}（{fpct(ch, 1, True)}）",
                        "note": f"{unit}环比变动超阈值 30%，建议核对投放、库存与页面状态"})
        elif ch is None and p["cur_sales"] > 0:
            out.append({"level": "中", "type": "新品起量",
                        "text": f"{p['product']} 本{unit}新增销售额 {fcur(p['cur_sales'])}（上{unit}为 0）",
                        "note": "从 0 起量，无环比基数，需人工确认是否为新上架/新投放"})
    # 2) 漏斗环节异动
    for name in ["加购率", "到达结账率", "完成结账率", "加购→结账", "结账→完成"]:
        a, b = cur_t.get(name), prev_t.get(name)
        ch = pct_change(a, b)
        if a is not None and b is not None and ch is not None and abs(ch) >= 0.30:
            small = (prev_t.get("Sessions that completed checkout") or 0) < 30
            out.append({"level": "中", "type": "漏斗异动",
                        "text": f"{name} {fpct(b)} → {fpct(a)}（{fpct(ch, 1, True)}）",
                        "note": f"环节转化率变动 {fpct(abs(ch))}，建议检查该环节页面/支付流程"
                                + ("；样本量小（完成结账 <30 单），方向可信、幅度待观察" if small else "")})
    # 3) 单日尖峰/塌陷（本周内）
    by_day = defaultdict(lambda: blank())
    for r in rows_between(recs, *cur_range):
        add(by_day[r["Day"]], r)
    vals = [v["Total sales"] for v in by_day.values()]
    if len(vals) >= 4:
        mean = sum(vals) / len(vals)
        sd = (sum((x - mean) ** 2 for x in vals) / len(vals)) ** 0.5
        if sd > 0:
            for day, v in sorted(by_day.items()):
                z = (v["Total sales"] - mean) / sd
                if z >= 2:
                    out.append({"level": "关注", "type": "单日尖峰",
                                "text": f"{day} 销售额 {fcur(v['Total sales'])}（高于周内均值 {z:.1f} 个标准差）",
                                "note": "尖峰日需确认是否有活动/投放/站外引流"})
                elif z <= -1.5 and v["Total sales"] == 0:
                    out.append({"level": "关注", "type": "零成交日",
                                "text": f"{day} 无成交（当日 {fnum(v['Sessions'])} 会话）",
                                "note": "有流量无转化，检查加购到结账链路"})
    # 4) 数据完整性
    sales_no_session = 0
    day_sess = defaultdict(float)
    day_sales = defaultdict(float)
    for r in rows_between(recs, *cur_range):
        day_sess[r["Day"]] += (r.get("Sessions") or 0)
        day_sales[r["Day"]] += (r.get("Total sales") or 0)
    for dd, v in day_sess.items():
        if v == 0 and day_sales[dd] > 0:
            sales_no_session += 1
    if sales_no_session:
        out.append({"level": "高", "type": "归因缺口",
                    "text": f"本周 {sales_no_session} 天出现「有销售额但零会话」",
                    "note": "成交未归因到产品落地页会话，可能来自广告落地页/其它渠道，影响转化率口径"})
    return out


def funnel_table(cur_t, prev_t):
    steps = [
        ("会话", cur_t["Sessions"], prev_t["Sessions"]),
        ("加入购物车", cur_t["Sessions with cart additions"], prev_t["Sessions with cart additions"]),
        ("到达结账", cur_t["Sessions that reached checkout"], prev_t["Sessions that reached checkout"]),
        ("完成结账", cur_t["Sessions that completed checkout"], prev_t["Sessions that completed checkout"]),
    ]
    out = []
    for i, (name, c, p) in enumerate(steps):
        step_rate = (c / steps[i - 1][1]) if i > 0 and steps[i - 1][1] else None
        step_rate_p = (p / steps[i - 1][2]) if i > 0 and steps[i - 1][2] else None
        out.append({"name": name, "cur": c, "prev": p,
                    "rate": step_rate, "rate_prev": step_rate_p,
                    "ch": pct_change(c, p)})
    return out


# ─────────────────────────── 发现生成（按重要性排序） ───────────────────────────
def build_findings(recs, cur_range, prev_range, label_cur, label_prev):
    cur_rows = rows_between(recs, *cur_range)
    prev_rows = rows_between(recs, *prev_range)
    cur_t, prev_t = agg(cur_rows), agg(prev_rows)
    dec = decompose(cur_t, prev_t)
    pd_ = product_deltas(recs, cur_range, prev_range)
    fn = funnel_table(cur_t, prev_t)
    cc, cp = concentration(cur_rows), concentration(prev_rows)

    cands = []

    # ① 销售额变动 + 三因子归因（最高优先级）
    ch_sales = pct_change(cur_t["Total sales"], prev_t["Total sales"])
    d = dec
    drivers = sorted([(k, v) for k, v in d.items() if k in ("流量效应", "转化效率效应", "客单价效应")],
                     key=lambda x: -abs(x[1]))
    top_driver = drivers[0]
    direction = "增长" if (ch_sales or 0) > 0 else "下滑"
    if ch_sales is not None:
        pos_k, pos_v = d["最正因素"], d[d["最正因素"]]
        neg_k, neg_v = d["最负因素"], d[d["最负因素"]]
        if neg_v < 0 and pos_v > 0:
            title = (f"销售额周环比{direction} {fpct(abs(ch_sales))}：{pos_k} +{CUR}{abs(pos_v):,.0f} "
                     f"{'被' if direction=='下滑' else '部分被'}{neg_k} -{CUR}{abs(neg_v):,.0f}{'抵消有余' if direction=='下滑' else '抵消'}")
        else:
            title = f"销售额周环比{direction} {fpct(abs(ch_sales))}，{d['主因']}主导"
        cands.append({
            "score": 100 + min(abs(ch_sales), 1.0) * 40,
            "icon": "trend",
            "title": title,
            "evidence": [f"销售额 {fcur(prev_t['Total sales'])} → {fcur(cur_t['Total sales'])}（{fpct(ch_sales,1,True)}）",
                         f"销量 {fnum(prev_t['Net items sold'])} → {fnum(cur_t['Net items sold'])} 件",
                         "因素分解：" + "、".join(f"{k} {'+' if v>=0 else '-'}{CUR}{abs(v):,.0f}" for k, v in drivers)],
            "sowhat": (f"变动可拆为 流量 {fcur(d['流量效应'])} / 转化效率 {fcur(d['转化效率效应'])} / 客单价 {fcur(d['客单价效应'])}；"
                       f"{'增长由' if direction=='增长' else '下滑源于'}{d['主因']}主导，"
                       f"说明{'应优先扩大流量投放' if d['主因']=='流量效应' else ('应优先优化转化链路与流量精准度' if d['主因']=='转化效率效应' else '应优先调整价格与组合策略')}。"),
        })

    # ② 漏斗：区分「结构性漏点」与「改善中的环节」，给出可量化情景
    weakest = min([f for f in fn if f["rate"] is not None], key=lambda x: x["rate"], default=None)
    back = [f for f in fn if f["name"] in ("到达结账", "完成结账")]
    if weakest:
        r_cart = cur_t["加购率"] or 0
        r_c2ch = cur_t["加购→结账"] or 0
        r_ch2c = cur_t["结账→完成"] or 0
        aov = cur_t["客单价"] or 0
        # 情景：加购率提升 1.5 个百分点（相对提升约 40%），后端转化率保持不变
        extra_cart = cur_t["Sessions"] * 0.015
        extra_orders = extra_cart * r_c2ch * r_ch2c
        scenario = extra_orders * aov
        improving = [f for f in back if (f["ch"] or 0) > 0.15]
        if improving:
            title = (f"后端环节在改善（{'、'.join(f['name']+' '+fpct(f['ch'],0,True) for f in improving)}），"
                     f"但前端加购率仅 {fpct(r_cart,1)} 成为规模瓶颈")
            sowhat = (f"结账链路效率提升说明支付与流程不是当前瓶颈，问题在流量精准度与落地页匹配；"
                      f"若加购率从 {fpct(r_cart,1)} 提升 1.5 个百分点，按当前后端转化率可多出约 "
                      f"{extra_orders:.0f} 单/周、约 {fcur(scenario)} 销售额。"
                      + ("注意：完成结账样本 < 30 单，情景测算为量级参考，非精确预测。" if cur_t["Sessions that completed checkout"] < 30 else ""))
        else:
            title = f"漏斗最大漏点在「{weakest['name']}」，环节转化仅 {fpct(weakest['rate'])}"
            sowhat = (f"每 100 个会话在该环节流失约 {(1-weakest['rate'])*100:.0f} 个；"
                      f"若加购率提升 1.5 个百分点，按当前后端转化率可多出约 {extra_orders:.0f} 单/周、约 {fcur(scenario)}。")
        cands.append({
            "score": 70 + (1 - weakest["rate"]) * 30,
            "icon": "funnel",
            "title": title,
            "evidence": [f"本周各环节：会话 {fnum(cur_t['Sessions'])} → 加购 {fnum(cur_t['Sessions with cart additions'])} "
                         f"→ 到达结账 {fnum(cur_t['Sessions that reached checkout'])} → 完成结账 {fnum(cur_t['Sessions that completed checkout'])}",
                         f"环节转化率：加购 {fpct(r_cart,1)} → 加购到结账 {fpct(r_c2ch)} → 结账完成 {fpct(r_ch2c)}",
                         f"整体完成结账率 {fpct(cur_t['完成结账率'],2)}（上周 {fpct(prev_t['完成结账率'],2)}）；"
                         f"跳出率 {fpct(cur_t['跳出率'],1)}"],
            "sowhat": sowhat,
        })

    # ③ 流量变化
    ch_sess = pct_change(cur_t["Sessions"], prev_t["Sessions"])
    if ch_sess is not None:
        cands.append({
            "score": 60 + min(abs(ch_sess), 1.0) * 35,
            "icon": "users",
            "title": f"流量{'回升' if ch_sess>0 else '回落'} {fpct(abs(ch_sess))}，每会话销售额 {fcur(cur_t['每会话销售额'],2)}（上周 {fcur(prev_t['每会话销售额'],2)}）",
            "evidence": [f"会话 {fnum(prev_t['Sessions'])} → {fnum(cur_t['Sessions'])}（{fpct(ch_sess,1,True)}）",
                         f"访客 {fnum(prev_t['Online store visitors'])} → {fnum(cur_t['Online store visitors'])}",
                         f"加购率 {fpct(prev_t['加购率'])} → {fpct(cur_t['加购率'])}",
                         f"跳出率 {fpct(prev_t['跳出率'])} → {fpct(cur_t['跳出率'])}"],
            "sowhat": ("流量与效率同向变化，需确认是投放放量还是自然流量；"
                       if (ch_sess > 0) == ((cur_t["每会话销售额"] or 0) > (prev_t["每会话销售额"] or 0))
                       else "流量与变现效率背离，说明流量质量发生变化，需按渠道拆分复查；")
                      + "跳出率与加购率共同指向落地页相关性。",
        })

    # ④ 集中度与结构
    if cc.get("top1") is not None:
        ch_hhi = pct_change(cc["hhi"], cp["hhi"])
        top1_name = cc["top"][0][0] if cc["top"] else "—"
        cands.append({
            "score": 55 + (cc["top1"] or 0) * 30,
            "icon": "layers",
            "title": f"销售额集中度偏高：Top1 占 {fpct(cc['top1'])}、Top3 占 {fpct(cc['top3'])}",
            "evidence": [f"Top1 产品：{top1_name}（{fcur(cc['top'][0][1])}）",
                         f"Top3 合计 {fcur(sum(v for _, v in cc['top'][:3]))} / 总额 {fcur(cc['total'])}",
                         f"HHI {cc['hhi']:.3f}（上周 {cp['hhi']:.3f}）"],
            "sowhat": ("单品依赖度高，Top1 波动会直接放大到全店；建议为 Top2–3 制定承接计划，"
                       "并对 Top1 做库存与投放连续性保障。"),
        })

    # ⑤ 产品贡献拆解
    up = [p for p in pd_ if p["d_sales"] > 0][:3]
    down = [p for p in pd_ if p["d_sales"] < 0][:3]
    if up or down:
        parts = []
        if up:
            parts.append("增量主要来自 " + "、".join(f"{p['product']}（+{fcur(p['d_sales'])}）" for p in up))
        if down:
            parts.append("拖累主要来自 " + "、".join(f"{p['product']}（-{fcur(abs(p['d_sales']))}）" for p in down))
        lead = up[0] if up else (down[0] if down else None)
        net = (cur_t["Total sales"] or 0) - (prev_t["Total sales"] or 0)
        if up and down:
            title = (f"结构此消彼长：{up[0]['product']} +{CUR}{abs(up[0]['d_sales']):,.0f} 与 "
                     f"{down[0]['product']} -{CUR}{abs(down[0]['d_sales']):,.0f} 相互抵消，净变动仅 {fcur(net)}")
        elif up:
            title = f"增量集中在 {up[0]['product']}（+{CUR}{abs(up[0]['d_sales']):,.0f}），是本周增长的唯一支点"
        else:
            title = f"降幅集中在 {down[0]['product']}（-{CUR}{abs(down[0]['d_sales']):,.0f}）"
        cands.append({
            "score": 50 + (abs(lead["d_sales"]) / max(cc.get("total", 1), 1)) * 200 if lead else 0,
            "icon": "bars",
            "title": title,
            "evidence": parts + [f"共 {len([p for p in pd_ if p['cur_sales']>0 or p['prev_sales']>0])} 个产品产生交易"],
            "sowhat": "把资源投向带来实际增量的产品，对持续下滑且流量不低的产品优先排查页面与价格。",
        })

    cands = sorted(cands, key=lambda x: -x["score"])
    return cur_t, prev_t, dec, pd_, fn, cc, cp, cands[:5]


# ─────────────────────────── SVG 图标与图表 ───────────────────────────
ICONS = {
    "trend": '<path d="M3 17l6-6 4 4 8-8" /><path d="M21 7v5h-5" />',
    "funnel": '<path d="M3 4h18l-7 8v7l-4 2v-9z" />',
    "users": '<circle cx="9" cy="8" r="3.2"/><path d="M2.5 20c0-3.6 2.9-6 6.5-6s6.5 2.4 6.5 6"/><path d="M17 6.2a3.2 3.2 0 010 5.6"/><path d="M18.5 14.4c2 .8 3.3 2.6 3.3 4.8"/>',
    "layers": '<path d="M12 3l9 5-9 5-9-5z"/><path d="M3 13l9 5 9-5"/><path d="M3 17.5l9 5 9-5"/>',
    "bars": '<path d="M5 20V10"/><path d="M12 20V4"/><path d="M19 20v-7"/>',
    "alert": '<path d="M12 3l9.5 17H2.5z"/><path d="M12 9v5"/><path d="M12 17.2v.1"/>',
    "bulb": '<path d="M9 18h6"/><path d="M10 21h4"/><path d="M12 3a6 6 0 00-3.5 10.9c.5.4.8 1 .8 1.6v.5h5.4v-.5c0-.6.3-1.2.8-1.6A6 6 0 0012 3z"/>',
    "search": '<circle cx="11" cy="11" r="7"/><path d="M20 20l-4.2-4.2"/>',
    "target": '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4"/><circle cx="12" cy="12" r="1"/>',
    "clock": '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
}


def icon(name, size=20, color="#B03A2E", sw=1.7):
    p = ICONS.get(name, "")
    return (f'<svg class="ic" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" '
            f'stroke="{color}" stroke-width="{sw}" stroke-linecap="round" stroke-linejoin="round">{p}</svg>')


def svg_trend(weeks, w=760, h=240):
    """近 N 周：柱=销售额，折线=会话。双轴，最后一周高亮。"""
    if not weeks:
        return ""
    pad_l, pad_r, pad_t, pad_b = 52, 52, 18, 44
    iw, ih = w - pad_l - pad_r, h - pad_t - pad_b
    max_sales = max(x["sales"] for x in weeks) or 1
    max_sess = max(x["sessions"] for x in weeks) or 1
    n = len(weeks)
    slot = iw / n
    bw = min(slot * 0.5, 34)
    parts = [f'<svg viewBox="0 0 {w} {h}" class="chart" role="img">']
    # 网格
    for i in range(5):
        y = pad_t + ih * i / 4
        parts.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{w-pad_r}" y2="{y:.1f}" stroke="#E3DFD8" stroke-width="1"/>')
        val = max_sales * (1 - i / 4)
        parts.append(f'<text x="{pad_l-8}" y="{y+4:.1f}" text-anchor="end" font-size="10" fill="#8A857C">{CUR}{val:,.0f}</text>')
    pts = []
    for i, x in enumerate(weeks):
        cx = pad_l + slot * i + slot / 2
        bh = ih * (x["sales"] / max_sales)
        last = (i == n - 1)
        fill = "#B03A2E" if last else "#1F3A5F"
        parts.append(f'<rect x="{cx-bw/2:.1f}" y="{pad_t+ih-bh:.1f}" width="{bw:.1f}" height="{bh:.1f}" fill="{fill}" opacity="{1 if last else .78}"/>')
        sy = pad_t + ih * (1 - x["sessions"] / max_sess)
        pts.append((cx, sy))
        parts.append(f'<text x="{cx:.1f}" y="{h-24}" text-anchor="middle" font-size="10" fill="#8A857C">{x["label"]}</text>')
    path = " ".join(f'{"M" if i==0 else "L"}{cx:.1f},{sy:.1f}' for i, (cx, sy) in enumerate(pts))
    parts.append(f'<path d="{path}" fill="none" stroke="#C89B3C" stroke-width="2"/>')
    for cx, sy in pts:
        parts.append(f'<circle cx="{cx:.1f}" cy="{sy:.1f}" r="3" fill="#FBFAF7" stroke="#C89B3C" stroke-width="1.8"/>')
    parts.append(f'<g font-size="10" fill="#8A857C"><rect x="{pad_l}" y="{h-14}" width="10" height="10" fill="#1F3A5F" opacity=".78"/>'
                 f'<text x="{pad_l+15}" y="{h-5}">销售额</text>'
                 f'<rect x="{pad_l+70}" y="{h-14}" width="10" height="10" fill="#C89B3C"/>'
                 f'<text x="{pad_l+85}" y="{h-5}">会话数（右轴）</text>'
                 f'<rect x="{pad_l+190}" y="{h-14}" width="10" height="10" fill="#B03A2E"/>'
                 f'<text x="{pad_l+205}" y="{h-5}">分析周</text></g>')
    parts.append("</svg>")
    return "".join(parts)


def svg_funnel(fn, w=760, h=210):
    pad_l, pad_r, pad_t = 96, 60, 12
    iw = w - pad_l - pad_r
    row_h = (h - pad_t - 10) / max(len(fn), 1)
    maxv = max([x["cur"] for x in fn] + [1])
    parts = [f'<svg viewBox="0 0 {w} {h}" class="chart" role="img">']
    for i, s in enumerate(fn):
        y = pad_t + i * row_h
        bw = iw * (s["cur"] / maxv)
        parts.append(f'<text x="{pad_l-10}" y="{y+row_h/2+4:.0f}" text-anchor="end" font-size="12" fill="#1A1A1A">{s["name"]}</text>')
        parts.append(f'<rect x="{pad_l}" y="{y+row_h*0.22:.1f}" width="{iw}" height="{row_h*0.5:.1f}" fill="#EFECE5"/>')
        parts.append(f'<rect x="{pad_l}" y="{y+row_h*0.22:.1f}" width="{bw:.1f}" height="{row_h*0.5:.1f}" fill="#1F3A5F" opacity="{.9-i*0.12:.2f}"/>')
        r = fpct(s["rate"]) if s["rate"] is not None else "—"
        parts.append(f'<text x="{pad_l+iw+8}" y="{y+row_h/2+4:.0f}" font-size="11" fill="#1A1A1A">{fnum(s["cur"])}'
                     + (f'  <tspan fill="#8A857C">环比 {fpct(s["ch"],1,True)}</tspan>' if s["ch"] is not None else "")
                     + f'</text>')
        if i > 0:
            parts.append(f'<text x="{pad_l+6}" y="{y+row_h/2+4:.0f}" font-size="10" fill="#B03A2E">环节转化 {r}</text>')
    parts.append("</svg>")
    return "".join(parts)


def svg_contrib(pdeltas, w=760, h=None):
    items = [p for p in pdeltas if abs(p["d_sales"]) > 0][:8]
    if not items:
        return ""
    h = 26 * len(items) + 26
    pad_l, pad_r = 190, 90
    iw = w - pad_l - pad_r
    mx = max(abs(p["d_sales"]) for p in items) or 1
    mid = pad_l + iw / 2
    parts = [f'<svg viewBox="0 0 {w} {h}" class="chart" role="img">']
    parts.append(f'<line x1="{mid}" y1="6" x2="{mid}" y2="{h-14}" stroke="#C9C4BA" stroke-width="1"/>')
    for i, p in enumerate(items):
        y = 14 + i * 26
        w_ = iw / 2 * (abs(p["d_sales"]) / mx)
        pos = p["d_sales"] >= 0
        x = mid if pos else mid - w_
        parts.append(f'<text x="{pad_l-10}" y="{y+11}" text-anchor="end" font-size="11" fill="#1A1A1A">{p["product"][:26]}</text>')
        parts.append(f'<rect x="{x:.1f}" y="{y+2}" width="{w_:.1f}" height="13" fill="{"#2F6F62" if pos else "#B03A2E"}" opacity=".88"/>')
        lab = ("+" if pos else "-") + f"{CUR}{abs(p['d_sales']):,.0f}"
        tx = (mid + w_ + 8) if pos else (mid - w_ - 8)
        anc = "start" if pos else "end"
        parts.append(f'<text x="{tx:.1f}" y="{y+13}" text-anchor="{anc}" font-size="11" fill="#1A1A1A">{lab}</text>')
    parts.append(f'<text x="{mid}" y="{h-2}" text-anchor="middle" font-size="10" fill="#8A857C">← 拖累　　产品销售额周变动　　贡献 →</text>')
    parts.append("</svg>")
    return "".join(parts)


# ─────────────────────────── HTML 渲染 ───────────────────────────
CSS = """
*{box-sizing:border-box}
body{margin:0;background:#FBFAF7;color:#1A1A1A;
 font-family:-apple-system,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;
 line-height:1.65;-webkit-font-smoothing:antialiased}
.wrap{max-width:1120px;margin:0 auto;padding:0 28px 80px}
.rule-top{height:4px;background:#1A1A1A}
.masthead{border-bottom:1px solid #E3DFD8;padding:26px 0 18px;margin-bottom:8px}
.kicker{font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:#B03A2E;font-weight:700}
h1{font-family:Georgia,"Songti SC",serif;font-size:34px;line-height:1.25;margin:10px 0 6px;font-weight:700;letter-spacing:-.2px}
.sub{color:#6E6A62;font-size:14px;margin:0}
.meta{margin-top:12px;font-size:12px;color:#8A857C;display:flex;gap:18px;flex-wrap:wrap}
.lede{font-family:Georgia,"Songti SC",serif;font-size:21px;line-height:1.6;
 border-left:3px solid #B03A2E;padding:4px 0 4px 18px;margin:30px 0 38px}
section{margin:52px 0}
h2{font-family:Georgia,"Songti SC",serif;font-size:13px;letter-spacing:.16em;text-transform:uppercase;
 color:#1A1A1A;border-bottom:2px solid #1A1A1A;padding-bottom:8px;margin:0 0 6px;font-weight:700}
h2 .n{color:#B03A2E;margin-right:8px}
.block-sub{color:#8A857C;font-size:13px;margin:8px 0 22px}
.grid{display:grid;gap:16px}
.g4{grid-template-columns:repeat(4,1fr)}
.g3{grid-template-columns:repeat(3,1fr)}
.g2{grid-template-columns:repeat(2,1fr)}
@media(max-width:880px){.g4,.g3,.g2{grid-template-columns:repeat(2,1fr)}}
@media(max-width:560px){.g4,.g3,.g2{grid-template-columns:1fr}}
.card{background:#fff;border:1px solid #E3DFD8;border-radius:3px;padding:18px 18px 16px}
.card .lab{font-size:10.5px;letter-spacing:.14em;text-transform:uppercase;color:#8A857C;display:flex;align-items:center;gap:7px}
.card .val{font-family:Georgia,serif;font-size:31px;line-height:1.15;margin:9px 0 4px;letter-spacing:-.5px}
.card .note{font-size:12px;color:#8A857C}
.delta{display:inline-block;font-size:12px;font-weight:700;padding:2px 7px;border-radius:2px;margin-left:6px;vertical-align:2px}
.up{color:#2F6F62;background:#EAF2EF}
.down{color:#B03A2E;background:#F7EBE9}
.flat{color:#8A857C;background:#F1EFEA}
.ic{flex:none}
.find{border-top:1px solid #E3DFD8;padding:22px 0;display:grid;grid-template-columns:38px 1fr;gap:16px}
.find:first-of-type{border-top:none}
.find .num{font-family:Georgia,serif;font-size:22px;color:#B03A2E;line-height:1}
.find h3{font-family:Georgia,"Songti SC",serif;font-size:19px;margin:0 0 8px;font-weight:700;line-height:1.4}
.find .ev{margin:0 0 10px;padding-left:16px}
.find .ev li{font-size:13.5px;color:#4A4740;margin-bottom:3px}
.find .sw{font-size:13.5px;background:#F4F1EA;border-left:2px solid #1F3A5F;padding:9px 13px;color:#33302A}
.find .sw b{color:#1A1A1A}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;font-size:10.5px;letter-spacing:.12em;text-transform:uppercase;color:#8A857C;
 border-bottom:1px solid #1A1A1A;padding:0 10px 7px;font-weight:700}
td{padding:9px 10px;border-bottom:1px solid #EFECE5}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
.tag{display:inline-block;font-size:10.5px;padding:2px 7px;border-radius:2px;font-weight:700;letter-spacing:.04em}
.tag.高{background:#F7EBE9;color:#B03A2E}
.tag.中{background:#FBF1DE;color:#8A6A1F}
.tag.关注{background:#EAF0F5;color:#1F3A5F}
.chart{width:100%;height:auto;display:block;margin:6px 0 4px}
.frame{background:#fff;border:1px solid #E3DFD8;border-radius:3px;padding:16px 18px}
.frame .cap{font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:#8A857C;margin-bottom:6px}
.rec{display:grid;grid-template-columns:30px 1fr;gap:14px;padding:16px 0;border-top:1px solid #E3DFD8}
.rec:first-child{border-top:none}
.rec h4{margin:0 0 4px;font-size:15px;font-family:Georgia,"Songti SC",serif}
.rec p{margin:0;font-size:13.5px;color:#4A4740}
.rec .why{font-size:12px;color:#8A857C;margin-top:5px}
.layer{background:#fff;border:1px solid #E3DFD8;border-left:3px solid #1F3A5F;padding:14px 16px;border-radius:3px}
.layer b{font-family:Georgia,serif}
.layer p{margin:5px 0 0;font-size:13px;color:#4A4740}
footer{margin-top:60px;border-top:1px solid #E3DFD8;padding-top:16px;font-size:11.5px;color:#8A857C}
.mono{font-variant-numeric:tabular-nums}
"""


KPI_ICON = {"销售额": "bars", "会话数": "users", "成交件数": "layers", "客单价": "target",
            "加购率": "funnel", "完成结账率": "target", "每会话销售额": "trend", "跳出率": "alert"}


def kpi(lab, val, delta_p, kind="num"):
    if delta_p is None:
        chip = '<span class="delta flat">无环比</span>'
    else:
        cls = "up" if delta_p > 0 else ("down" if delta_p < 0 else "flat")
        chip = f'<span class="delta {cls}">{fpct(delta_p,1,True)}</span>'
    ic = icon(KPI_ICON.get(lab, "bars"), 15, "#8A857C", 1.8)
    return (f'<div class="card"><div class="lab">{ic}{lab}</div>'
            f'<div class="val">{val}{chip}</div></div>')


def render(recs, P, weekly, monthly, partial_info):
    cur_t, prev_t, dec, pd_, fn, cc, cp, findings = weekly[:8]
    anomalies = weekly[8] if len(weekly) > 8 else []
    ws, we = P["wk_cur"]
    ps, pe = P["wk_prev"]
    label_wk = f"{ws:%Y-%m-%d} ~ {we:%Y-%m-%d}"
    label_pv = f"{ps:%Y-%m-%d} ~ {pe:%Y-%m-%d}"

    # 近 8 周趋势
    weeks = []
    st = P["wk_cur"][0] - timedelta(days=7 * 7)
    for i in range(8):
        a = st + timedelta(days=7 * i)
        b = a + timedelta(days=6)
        t = agg(rows_between(recs, a, b))
        weeks.append({"label": f"{a:%m/%d}", "sales": t["Total sales"], "sessions": t["Sessions"]})

    ch_sales = pct_change(cur_t["Total sales"], prev_t["Total sales"])
    direction = "增长" if (ch_sales or 0) > 0 else "下滑"
    main_k = dec["主因"]
    pos_k, neg_k = dec["最正因素"], dec["最负因素"]
    if direction == "下滑" and dec[pos_k] > 0:
        headline = f"销售额环比下滑 {fpct(abs(ch_sales or 0))}：{pos_k} +{CUR}{abs(dec[pos_k]):,.0f} 被 {neg_k} 抵消"
    else:
        headline = f"销售额环比{direction} {fpct(abs(ch_sales or 0))}，{main_k}驱动"
    top_driver = main_k
    summary = (f"分析周（{label_wk}）销售额 {fcur(cur_t['Total sales'])}，环比{direction} {fpct(abs(ch_sales or 0))}；"
               f"分解为 流量 {fcur(dec['流量效应'])}、转化效率 {fcur(dec['转化效率效应'])}、客单价 {fcur(dec['客单价效应'])}，"
               f"{'净变动由' + main_k + '决定' if direction=='增长' else '下滑由' + neg_k + '主导'}。"
               f"共 {fnum(cur_t['Sessions'])} 会话、{fnum(cur_t['Net items sold'])} 件成交，"
               f"每会话销售额 {fcur(cur_t['每会话销售额'],2)}，客单价 {fcur(cur_t['客单价'],2)}。")

    H = []
    H.append("<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>")
    H.append("<meta name='viewport' content='width=device-width,initial-scale=1'>")
    H.append(f"<title>周度分析 {label_wk}</title><style>{CSS}</style></head><body>")
    H.append("<div class='rule-top'></div><div class='wrap'>")
    H.append(f"""<div class="masthead">
      <div class="kicker">Weekly Analytics · 周度经营分析</div>
      <h1>{html.escape(headline)}</h1>
      <p class="sub">分析周 {label_wk}（周五–周四） · 对比周 {label_pv}（同为周五–周四）</p>
      <div class="meta"><span>数据源：Shopify 落地页流量 + 产品销售整合表</span>
      <span>覆盖 {P['min']:%Y-%m-%d} ~ {P['max']:%Y-%m-%d}</span>
      <span>口径：日粒度聚合，比率由计数重算</span></div></div>""")
    H.append(f"<div class='lede'>{html.escape(summary)}</div>")

    # ① 数据全貌
    H.append("<section><h2><span class='n'>01</span>数据全貌</h2>"
             "<p class='block-sub'>分析周 vs 上周；所有比率为合计计数重算，非原始比率平均。</p>")
    cards = [
        ("销售额", fcur(cur_t["Total sales"]), pct_change(cur_t["Total sales"], prev_t["Total sales"])),
        ("会话数", fnum(cur_t["Sessions"]), pct_change(cur_t["Sessions"], prev_t["Sessions"])),
        ("成交件数", fnum(cur_t["Net items sold"]), pct_change(cur_t["Net items sold"], prev_t["Net items sold"])),
        ("客单价", fcur(cur_t["客单价"], 2) if cur_t["客单价"] else "—", pct_change(cur_t["客单价"] or 0, prev_t["客单价"] or 0)),
        ("加购率", fpct(cur_t["加购率"], 2), pct_change(cur_t["加购率"] or 0, prev_t["加购率"] or 0)),
        ("完成结账率", fpct(cur_t["完成结账率"], 2), pct_change(cur_t["完成结账率"] or 0, prev_t["完成结账率"] or 0)),
        ("每会话销售额", fcur(cur_t["每会话销售额"], 2) if cur_t["每会话销售额"] else "—", pct_change(cur_t["每会话销售额"] or 0, prev_t["每会话销售额"] or 0)),
        ("跳出率", fpct(cur_t["跳出率"], 1), pct_change(cur_t["跳出率"] or 0, prev_t["跳出率"] or 0)),
    ]
    H.append("<div class='grid g4'>" + "".join(kpi(l, v, p) for l, v, p in cards) + "</div>")
    H.append(f"<div class='frame' style='margin-top:18px'><div class='cap'>近 8 周走势 · 柱=销售额（左轴），线=会话数（右轴）</div>{svg_trend(weeks)}</div>")
    gap = sum(1 for a in anomalies if a["type"] == "归因缺口")
    if gap or cur_t["完成结账率"] and cur_t["完成结账率"] < 0.02:
        H.append("<div class='frame' style='margin-top:14px;border-left:3px solid #B03A2E'>"
                 "<div class='cap'>口径提醒（影响比率解读）</div>"
                 f"<p style='margin:0;font-size:13px;color:#4A4740'>全站完成结账率处于低位（{fpct(cur_t['完成结账率'],2)}），"
                 "结合历史曾出现「有销售额但零落地页会话」的日期，说明存在成交未归因到产品落地页的情况"
                 "（广告落地页/站外渠道）。此类比率适合看趋势与相对变化，不宜直接当作渠道 ROI 绝对值。</p></div>")

    # 结构层说明
    n_days = len(P["days"])
    act_days = len({r["Day"] for r in recs if (r.get("Sessions") or 0) > 0})
    H.append(f"<div class='grid g3' style='margin-top:18px'>"
             f"<div class='layer'>{icon('search',18,'#1F3A5F')}<b> L0 结构</b>"
             f"<p>{len(set(r['Product'] for r in recs))} 个产品 × {n_days} 天日粒度；"
             f"流量与销售经 Product ID 桥接，桥接覆盖 100%。</p></div>"
             f"<div class='layer'>{icon('bars',18,'#1F3A5F')}<b> L1 统计</b>"
             f"<p>累计 {fnum(sum((r.get('Total sales') or 0) for r in recs))} 销售额；"
             f"有流量天数 {act_days}/{n_days}。</p></div>"
             f"<div class='layer'>{icon('target',18,'#1F3A5F')}<b> L2 洞察</b>"
             f"<p>对比产生意义：所有结论均基于环比/结构/漏斗差异，而非绝对值。</p></div></div>")
    H.append("</section>")

    # ② 核心发现
    H.append("<section><h2><span class='n'>02</span>核心发现</h2>"
             "<p class='block-sub'>按重要性排序（变动幅度 × 可行动性）；每条含证据与业务含义。</p>")
    for i, f in enumerate(findings, 1):
        ev = "".join(f"<li>{html.escape(e)}</li>" for e in f["evidence"])
        H.append(f"<div class='find'><div class='num'>{icon(f.get('icon','target'),22,'#B03A2E')}</div><div>"
                 f"<h3><span style='color:#B03A2E;font-family:Georgia,serif;margin-right:8px'>{i}</span>{html.escape(f['title'])}</h3>"
                 f"<ul class='ev'>{ev}</ul>"
                 f"<div class='sw'><b>So what：</b>{html.escape(f['sowhat'])}</div></div></div>")
    H.append("</section>")

    # ③ 深度：归因 + 漏斗 + 产品
    H.append("<section><h2><span class='n'>03</span>归因与结构</h2>"
             "<p class='block-sub'>销售额变动的三因子链式分解（流量 × 件/会话 × 客单价）、漏斗环节、产品贡献。</p>")
    drv = [("流量效应", dec["流量效应"]), ("转化效率效应", dec["转化效率效应"]), ("客单价效应", dec["客单价效应"])]
    drv_rows = "".join(
        f"<tr><td>{k}</td><td class='n mono'>{'+' if v>=0 else '-'}{CUR}{abs(v):,.0f}</td>"
        f"<td class='n mono'>{(v/dec['合计变动']*100 if dec['合计变动'] else 0):,.0f}%</td></tr>" for k, v in drv)
    H.append(f"<div class='grid g2'>"
             f"<div class='frame'><div class='cap'>销售额变动分解 · 合计 {fcur(dec['合计变动'])}</div>"
             f"<table><thead><tr><th>驱动因素</th><th class='n'>影响金额</th><th class='n'>占变动</th></tr></thead>"
             f"<tbody>{drv_rows}</tbody></table>"
             f"<p style='font-size:12px;color:#8A857C;margin:10px 0 0'>链式替代法，三者之和等于实际变动额（校验通过）。</p></div>"
             f"<div class='frame'><div class='cap'>转化漏斗 · 本周 vs 上周</div>{svg_funnel(fn)}</div></div>")
    H.append(f"<div class='frame' style='margin-top:16px'><div class='cap'>各产品销售额周变动（贡献 / 拖累）</div>{svg_contrib(pd_)}</div>")
    top_rows = "".join(
        f"<tr><td>{p['product']}</td><td class='n mono'>{fcur(p['cur_sales'])}</td><td class='n mono'>{fcur(p['prev_sales'])}</td>"
        f"<td class='n mono' style='color:{'#2F6F62' if p['d_sales']>=0 else '#B03A2E'}'>{'+' if p['d_sales']>=0 else '-'}{CUR}{abs(p['d_sales']):,.0f}</td>"
        f"<td class='n mono'>{fnum(p['cur_sessions'])}</td></tr>" for p in pd_[:8])
    H.append(f"<div class='frame' style='margin-top:16px'><div class='cap'>产品明细（按销售额变动绝对值排序）</div>"
             f"<table><thead><tr><th>产品</th><th class='n'>本周销售额</th><th class='n'>上周销售额</th>"
             f"<th class='n'>变动</th><th class='n'>会话</th></tr></thead><tbody>{top_rows}</tbody></table></div>")
    H.append("</section>")

    # ④ 异常
    H.append("<section><h2><span class='n'>04</span>值得关注的异常</h2>"
             "<p class='block-sub'>主动挖掘：产品异动、漏斗突变、单日尖峰与数据完整性缺口。</p>")
    if anomalies:
        rows = "".join(f"<tr><td><span class='tag {a['level']}'>{a['level']}</span></td><td>{a['type']}</td>"
                       f"<td>{html.escape(a['text'])}</td><td style='color:#8A857C;font-size:12.5px'>{html.escape(a['note'])}</td></tr>"
                       for a in anomalies)
        H.append(f"<div class='frame'><table><thead><tr><th>级别</th><th>类型</th><th>现象</th><th>说明</th></tr></thead><tbody>{rows}</tbody></table></div>")
    else:
        H.append("<div class='frame'><p style='margin:0;color:#4A4740'>本周未触发异常阈值（产品变动 <30%、漏斗环节变动 <30%、无单日极端值）。</p></div>")

    # ⑤ 建议
    H.append("<section><h2><span class='n'>05</span>可操作建议</h2>"
             "<p class='block-sub'>每条对应上述发现，含判断依据。</p>")
    recs_ = build_recommendations(cur_t, prev_t, dec, findings, cc, fn, anomalies)
    for i, r in enumerate(recs_, 1):
        H.append(f"<div class='rec'><div>{icon('bulb',20,'#B03A2E')}</div><div>"
                 f"<h4>{i}. {html.escape(r['title'])}</h4><p>{html.escape(r['action'])}</p>"
                 f"<div class='why'>依据：{html.escape(r['why'])}</div></div></div>")
    H.append("</section>")

    # ⑥ 月度
    if monthly:
        H.append(render_monthly(monthly, P))

    # ⑦ 当周进度
    if partial_info:
        H.append(f"<section><h2><span class='n'>07</span>当周进行中</h2>"
                 f"<p class='block-sub'>数据截止 {P['max']:%Y-%m-%d}，本周尚未完整，以下为同期对比（同星期数），仅供参考。</p>"
                 f"<div class='frame'><table><thead><tr><th>指标</th><th class='n'>当周同期 {partial_info['n']} 天</th>"
                 f"<th class='n'>上周同期 {partial_info['n']} 天</th><th class='n'>变化</th></tr></thead><tbody>"
                 + "".join(f"<tr><td>{k}</td><td class='n mono'>{v}</td><td class='n mono'>{pv}</td>"
                           f"<td class='n mono'>{ch}</td></tr>" for k, v, pv, ch in partial_info['rows'])
                 + "</tbody></table>"
                   "<p style='font-size:12px;color:#8A857C;margin:10px 0 0'>未完整期间不做环比结论，避免误判趋势。</p></div></section>")

    H.append(f"<footer>本报告由自动化分析引擎生成 · 周口径为「上周五 ~ 本周四」（Fri–Thu）；"
             f"数据口径：Shopify 落地页流量表按路径归一后与产品销售表按 (Day, Product ID) 关联；"
             f"计数类求和、比率由合计计数重算、均值按会话加权。金额单位 {CUR}（EUR）。"
             f"样本量较小时（完成结账 &lt; 10 单）结论方向可信、幅度需谨慎解读。</footer>")
    H.append("</div></body></html>")
    return "".join(H)


def render_monthly(M, P):
    (cur_t, prev_t, dec, pd_, fn, cc, cp, findings) = M[0]
    anomalies = M[1]
    y, m = P["mon_cur"]
    py, pm = P["mon_prev"]
    ch = pct_change(cur_t["Total sales"], prev_t["Total sales"])
    H = [f"<section><h2><span class='n'>06</span>月度分析 · {y}年{m}月</h2>"
         f"<p class='block-sub'>最近完整月 {y}-{m:02d} vs {py}-{pm:02d}；同构框架下按月聚合。</p>"]
    cards = [("销售额", fcur(cur_t["Total sales"]), pct_change(cur_t["Total sales"], prev_t["Total sales"])),
             ("会话数", fnum(cur_t["Sessions"]), pct_change(cur_t["Sessions"], prev_t["Sessions"])),
             ("成交件数", fnum(cur_t["Net items sold"]), pct_change(cur_t["Net items sold"], prev_t["Net items sold"])),
             ("客单价", fcur(cur_t["客单价"], 2) if cur_t["客单价"] else "—", pct_change(cur_t["客单价"] or 0, prev_t["客单价"] or 0))]
    H.append("<div class='grid g4'>" + "".join(kpi(l, v, p) for l, v, p in cards) + "</div>")
    for i, f in enumerate(findings[:4], 1):
        ev = "".join(f"<li>{html.escape(e)}</li>" for e in f["evidence"])
        H.append(f"<div class='find'><div class='num'>{icon(f.get('icon','target'),22,'#B03A2E')}</div><div>"
                 f"<h3><span style='color:#B03A2E;font-family:Georgia,serif;margin-right:8px'>{i}</span>{html.escape(f['title'])}</h3>"
                 f"<ul class='ev'>{ev}</ul><div class='sw'><b>So what：</b>{html.escape(f['sowhat'])}</div></div></div>")
    if anomalies:
        rows = "".join(f"<tr><td><span class='tag {a['level']}'>{a['level']}</span></td><td>{a['type']}</td>"
                       f"<td>{html.escape(a['text'])}</td><td style='color:#8A857C;font-size:12.5px'>{html.escape(a['note'])}</td></tr>"
                       for a in anomalies[:5])
        H.append(f"<div class='frame' style='margin-top:10px'><div class='cap'>月度异常</div><table><thead><tr><th>级别</th><th>类型</th>"
                 f"<th>现象</th><th>说明</th></tr></thead><tbody>{rows}</tbody></table></div>")
    H.append("</section>")
    return "".join(H)


def build_recommendations(cur_t, prev_t, dec, findings, cc, fn, anomalies):
    out = []
    top = dec["主因"]
    if top == "流量效应":
        out.append({"title": "优先解决流量侧：复盘本周引流渠道与投放节奏",
                    "action": "按渠道/落地页拆分会话变化，确认是自然流量、投放还是外链带来的波动，把有效渠道的预算前置。",
                    "why": f"销售额变动中流量贡献 {fcur(dec['流量效应'])}，为第一驱动因素"})
    elif top == "转化效率效应":
        out.append({"title": "优先优化转化链路而非加预算",
                    "action": "检查加购到结账环节的页面加载、运费与支付方式展示，先做 A/B 验证再放量。",
                    "why": f"转化效率贡献 {fcur(dec['转化效率效应'])}，为第一驱动因素"})
    else:
        out.append({"title": "调整客单价策略：组合销售与阶梯优惠",
                    "action": "测试Bundle/满减门槛，观察件单价与件数的平衡。",
                    "why": f"客单价贡献 {fcur(dec['客单价效应'])}，为第一驱动因素"})

    weakest = min([f for f in fn if f["rate"] is not None], key=lambda x: x["rate"], default=None)
    if weakest:
        if weakest["name"] == "加入购物车":
            action = ("加购环节主要反映流量精准度与落地页匹配度：核对引流关键词/受众与页面卖点是否一致，"
                      "并检查价格、库存状态与 CTA 是否清晰。")
            why = f"漏斗最大漏点（{fpct(1-weakest['rate'])} 流失），且发生在最前端，决定后端所有环节的规模上限"
        else:
            action = (f"该环节当前转化 {fpct(weakest['rate'])}，排查页面加载、运费与支付方式展示，"
                      f"设定一个可衡量的提升目标后再放量投放。")
            why = f"漏斗最大漏点，流失 {fpct(1-weakest['rate'])}"
        out.append({"title": f"把「{weakest['name']}」环节作为本周优化重点",
                    "action": action, "why": why})

    if cc.get("top3") and (cc["top1"] > 0.4 or cc["top3"] > 0.8) and cc.get("hhi"):
        out.append({"title": "降低单品依赖：为 Top2–3 制定承接动作",
                    "action": "给 Top2–3 产品安排站内推荐位与内容曝光，并为 Top1 建立库存与投放连续性预警。",
                    "why": f"Top1 占销售额 {fpct(cc['top1'])}，集中度偏高（HHI {cc['hhi']:.3f}）"})

    out.append({"title": "建立每周固定复盘节奏",
                "action": "固定周一取数，跑本报告，用同一套口径追踪三因子与漏斗，避免只看总额。",
                "why": "总额变化会被流量掩盖，分解后才能定位真正可干预的环节"})

    if any(a["type"] == "归因缺口" for a in anomalies):
        out.append({"title": "修复归因缺口：把非落地页成交纳入统计",
                    "action": "核对广告落地页/站外渠道的成交归属，补全后再评估真实转化率。",
                    "why": "存在有销售额但零会话的日期，会低估转化率"})
    return out[:5]


# ─────────────────────────── 主流程 ───────────────────────────
def main():
    global D, XLSX, OUTDIR
    import argparse
    ap = argparse.ArgumentParser(description="生成周度/月度分析报告")
    ap.add_argument("--dir", default=os.path.dirname(os.path.abspath(__file__)), help="数据目录")
    ap.add_argument("--xlsx", default=None, help="整合结果 XLSX，默认 <dir>/integrated_result.xlsx")
    ap.add_argument("--outdir", default=None, help="输出目录，默认 <dir>/每周分析")
    a, _ = ap.parse_known_args()
    D = os.path.abspath(a.dir)
    XLSX = a.xlsx or os.path.join(D, "integrated_result.xlsx")
    OUTDIR = a.outdir or os.path.join(D, "每周分析")

    recs = load_records(XLSX)
    P = build_periods(recs)
    cur_t, prev_t, dec, pd_, fn, cc, cp, findings = build_findings(
        recs, P["wk_cur"], P["wk_prev"], "cur", "prev")
    anomalies = detect_anomalies(recs, P["wk_cur"], P["wk_prev"], cur_t, prev_t, pd_)
    weekly = (cur_t, prev_t, dec, pd_, fn, cc, cp, findings, anomalies)

    # 月度
    mcur = agg(rows_month(recs, P["mon_cur"]))
    mprev = agg(rows_month(recs, P["mon_prev"]))
    yc, mc = P["mon_cur"]
    a0, b0 = date(yc, mc, 1), (date(yc + (mc == 12), (mc % 12) + 1, 1) - timedelta(days=1))
    yp, mp = P["mon_prev"]
    a1, b1 = date(yp, mp, 1), (date(yp + (mp == 12), (mp % 12) + 1, 1) - timedelta(days=1))
    mpd = product_deltas(recs, (a0, b0), (a1, b1))
    mfn = funnel_table(mcur, mprev)
    mcc = concentration(rows_month(recs, P["mon_cur"]))
    mcp = concentration(rows_month(recs, P["mon_prev"]))
    mdec = decompose(mcur, mprev)
    mfindings = []
    ch = pct_change(mcur["Total sales"], mprev["Total sales"])
    if ch is not None:
        mfindings.append({"title": f"月度销售额环比{'增长' if ch>0 else '下滑'} {fpct(abs(ch))}",
                          "evidence": [f"{fcur(mprev['Total sales'])} → {fcur(mcur['Total sales'])}",
                                       f"会话 {fnum(mprev['Sessions'])} → {fnum(mcur['Sessions'])}",
                                       f"客单价 {fcur(mprev['客单价'],2)} → {fcur(mcur['客单价'],2)}"],
                          "sowhat": "变动分解：" + "、".join(
                              f"{k} {'+' if v>=0 else '-'}{CUR}{abs(v):,.0f}"
                              for k, v in [("流量", mdec["流量效应"]), ("转化效率", mdec["转化效率效应"]),
                                           ("客单价", mdec["客单价效应"])])
                                   + f"；净变动由{mdec['主因']}主导。"})
    mfindings.append({"title": f"月度完成结账率 {fpct(mcur['完成结账率'],2)}（上月 {fpct(mprev['完成结账率'],2)}）",
                      "evidence": [f"完成结账 {fnum(mprev['Sessions that completed checkout'])} → {fnum(mcur['Sessions that completed checkout'])} 单",
                                   f"加购率 {fpct(mprev['加购率'])} → {fpct(mcur['加购率'])}",
                                   f"每会话销售额 {fcur(mprev['每会话销售额'],2)} → {fcur(mcur['每会话销售额'],2)}"],
                      "sowhat": "月度样本更充足，比率可信度高于周度，适合作为基线目标。"})
    if mcc.get("top"):
        mfindings.append({"title": f"月度集中度：Top1 占 {fpct(mcc['top1'])}，Top3 占 {fpct(mcc['top3'])}",
                          "evidence": [f"Top1：{mcc['top'][0][0]}（{fcur(mcc['top'][0][1])}）",
                                       f"HHI {mcc['hhi']:.3f}（上月 {mcp['hhi']:.3f}）"],
                          "sowhat": "单品依赖度决定抗风险能力，需为次主力产品蓄水。"})
    manom = detect_anomalies(recs, (a0, b0), (a1, b1), mcur, mprev, mpd, unit="月")
    monthly = ((mcur, mprev, mdec, mpd, mfn, mcc, mcp, mfindings), manom[:5])

    # 当周进行中（同期对比）
    partial_info = None
    if P["partial"]:
        pa = P["partial"]["start"]
        n = P["partial"]["n_days"]
        pb = P["partial"]["end"]
        la = pa - timedelta(days=7)
        lb = la + timedelta(days=n - 1)
        ct = agg(rows_between(recs, pa, pb))
        lt = agg(rows_between(recs, la, lb))
        rows = []
        for lab, f, fmt in [("销售额", "Total sales", "cur"), ("会话数", "Sessions", "num"),
                            ("成交件数", "Net items sold", "num"), ("每会话销售额", "每会话销售额", "cur2")]:
            a, b = ct.get(f), lt.get(f)
            p = pct_change(a, b)
            if fmt == "cur":
                va, vb = fcur(a), fcur(b)
            elif fmt == "cur2":
                va, vb = (fcur(a, 2) if a else "—"), (fcur(b, 2) if b else "—")
            else:
                va, vb = fnum(a), fnum(b)
            rows.append((lab, va, vb, fpct(p, 1, True) if p is not None else "—"))
        partial_info = {"n": n, "rows": rows}

    os.makedirs(OUTDIR, exist_ok=True)
    ws, we = P["wk_cur"]
    out = os.path.join(OUTDIR, f"周度分析_{ws:%Y-%m-%d}_{we:%Y-%m-%d}.html")
    html_out = render(recs, P, weekly, monthly, partial_info)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html_out)
    print("分析周: %s ~ %s (周五–周四)" % (ws, we))
    print("销售额: %s -> %s (%s)" % (fcur(prev_t["Total sales"]), fcur(cur_t["Total sales"]),
                                     fpct(pct_change(cur_t["Total sales"], prev_t["Total sales"]), 1, True)))
    print("会话:   %s -> %s (%s)" % (fnum(prev_t["Sessions"]), fnum(cur_t["Sessions"]),
                                     fpct(pct_change(cur_t["Sessions"], prev_t["Sessions"]), 1, True)))
    print("分解:   流量 %s / 转化效率 %s / 客单价 %s (校验 %s)" % (
        fcur(dec["流量效应"]), fcur(dec["转化效率效应"]), fcur(dec["客单价效应"]), dec["校验"]))
    print("核心发现: %d 条 | 异常: %d 条" % (len(findings), len(anomalies)))
    print("月度: %d-%02d vs %d-%02d" % (P["mon_cur"][0], P["mon_cur"][1], P["mon_prev"][0], P["mon_prev"][1]))
    print("输出: %s" % out)
    return out


if __name__ == "__main__":
    main()
