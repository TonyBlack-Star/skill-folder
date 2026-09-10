# -*- coding: utf-8 -*-
"""
增量更新：以现有 XLSX 为基准，合并 incoming/ 目录下的新日期 CSV。

设计要点：
- 基准即真相来源：历史数据以 XLSX 为准，无需保留历史原始 CSV
- 防重复累加：任何与基准重叠的 (Day, Product ID) 都显式报警，绝不静默求和
- 字段级合并：增量只提供流量时，不动基准的销售字段（否则会把真实销量清零）

用法：
    python update.py                          # 默认读/写 ./integrated_result.xlsx
    python update.py --incoming ./incoming    # 指定增量目录
    python update.py --base a.xlsx --out b.xlsx --incoming ./inc   # 测试（不碰真实基准）
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import etl_common as E


def archive_files(files, archive_dir):
    """把已处理的文件移入 archive/，文件名加日期前缀防重名，绝不静默覆盖。"""
    os.makedirs(archive_dir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d")
    moved = []
    for f in files:
        base = os.path.basename(f)
        dest = os.path.join(archive_dir, "%s_%s" % (stamp, base))
        if os.path.exists(dest):
            dest = os.path.join(archive_dir, "%s_%s" % (
                datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), base))
        shutil.move(f, dest)
        moved.append(os.path.basename(dest))
    return moved


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=os.getcwd())
    ap.add_argument("--map", default=None)
    ap.add_argument("--base", default=None, help="基准 XLSX（默认 <dir>/integrated_result.xlsx）")
    ap.add_argument("--out", default=None, help="输出 XLSX（默认与 --base 相同）")
    ap.add_argument("--incoming", default=None, help="增量目录（默认 <dir>/incoming）")
    ap.add_argument("--dashboard", default=None, help="可选：看板生成脚本，存在则自动重生成")
    ap.add_argument("--no-archive", action="store_true", help="处理完不自动归档")
    args = ap.parse_args()

    D = os.path.abspath(args.dir)
    mapfile = args.map or os.path.join(D, E.DEFAULT_MAP_NAME)
    base_path = args.base or os.path.join(D, "integrated_result.xlsx")
    out_path = args.out or base_path
    incoming_dir = args.incoming or os.path.join(D, "incoming")

    if not os.path.exists(base_path):
        raise SystemExit("基准 XLSX 不存在: %s（先跑 integrate.py 全量重算）" % base_path)

    mapping = E.load_mapping(mapfile)
    base = E.read_xlsx(base_path)
    incoming = sorted(glob.glob(os.path.join(incoming_dir, "*.csv")))
    anomalies = []

    if not incoming:
        print("[update] %s 无 CSV，跳过合并。" % incoming_dir)
    else:
        t_agg, s_agg, inc_anomalies, counters, src_files, traffic_keys, sales_keys = \
            E.aggregate_files(incoming, mapping)
        anomalies.extend(inc_anomalies)
        new_records = E.build_records(t_agg, s_agg, mapping)

        base_map = {(r["Day"], r["Product ID"]): dict(r) for r in base}
        new_map = {(r["Day"], r["Product ID"]): dict(r) for r in new_records}
        overlaps = set(base_map) & set(new_map)
        new_only = set(new_map) - set(base_map)

        TRAFFIC_FIELDS = E.COUNTS + list(E.DERIVED.keys()) + [E.BOUNCE, E.DURATION]
        for k in sorted(new_map):
            if k in base_map:
                rec, new = base_map[k], new_map[k]
                if k in traffic_keys:                 # 仅当增量确实含流量时才覆盖流量
                    for f in TRAFFIC_FIELDS:
                        rec[f] = new[f]
                if k in sales_keys:                   # 仅当增量确实含销售时才覆盖销售
                    rec["Net items sold"] = new["Net items sold"]
                    rec["Total sales"] = new["Total sales"]
                anomalies.append((k[1], k[0], "-", "增量与基准重叠日，已按字段级合并（请核对）"))
            else:
                base_map[k] = new_map[k]

        merged = sorted(base_map.values(), key=lambda r: (r["Day"], r["Product ID"]))
        E.write_xlsx(merged, mapping, out_path)
        moved = [] if args.no_archive else archive_files(incoming, os.path.join(incoming_dir, "archive"))

    # 可选：重生成看板
    dash = args.dashboard or os.path.join(D, "make_dashboard.py")
    if os.path.exists(dash):
        subprocess.run([sys.executable, dash], cwd=D)

    print("=== 增量更新报告 ===")
    if incoming:
        print("基准记录数: %d" % len(base))
        print("增量文件: %s" % [os.path.basename(f) for f in incoming])
        print("纯新增(天,产品)数: %d" % len(new_only))
        print("重叠数: %d" % len(overlaps))
        print("丢弃-非产品路径: %d | 未匹配handle: %d" % (counters["nonproduct"], counters["unmatched"]))
        print("合并后记录数: %d" % len(merged))
        if not args.no_archive:
            print("已归档: %d 个 -> %s" % (len(moved), os.path.join(incoming_dir, "archive")))
    print("异常数: %d" % len(anomalies))
    for a in anomalies[:50]:
        print("  异常:", a)
    print("输出: %s" % out_path)


if __name__ == "__main__":
    main()
