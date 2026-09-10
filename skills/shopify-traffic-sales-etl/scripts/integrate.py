# -*- coding: utf-8 -*-
"""
全量重算：扫描数据目录内所有 CSV，整合成多子表 XLSX。

用法：
    python integrate.py                       # 以当前目录为数据目录
    python integrate.py --dir D:/data --out result.xlsx

约定：
- 按表头自动识别流量表（含 Landing page path）与销售表（含 Product ID）
- 输出 XLSX 即「单一真相来源」，后续增量更新以它为基准
"""
import argparse
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import etl_common as E

# 输入扫描排除：映射文件、本脚本产物、脚本自身（避免自污染重复累加）
SKIP_NAMES = {"integrated_result.xlsx", "_scan.py", "etl_common.py", "integrate.py",
              "update.py", "make_dashboard.py"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=os.getcwd(), help="数据目录（默认当前目录）")
    ap.add_argument("--map", default=None, help="handle->ID 映射文件（默认 <dir>/路径与ID匹配规则.txt）")
    ap.add_argument("--out", default=None, help="输出 XLSX（默认 <dir>/integrated_result.xlsx）")
    args = ap.parse_args()

    D = os.path.abspath(args.dir)
    mapfile = args.map or os.path.join(D, E.DEFAULT_MAP_NAME)
    out = args.out or os.path.join(D, "integrated_result.xlsx")

    if not os.path.exists(mapfile):
        raise SystemExit("未找到映射文件: %s（用 --map 指定）" % mapfile)

    mapping = E.load_mapping(mapfile)
    skip = set(SKIP_NAMES) | {os.path.basename(mapfile)}
    files = [fp for fp in sorted(glob.glob(os.path.join(D, "*.csv")))
             if os.path.basename(fp) not in skip]

    traffic_agg, sales_agg, anomalies, counters, src_files, _, _ = E.aggregate_files(files, mapping)
    records = E.build_records(traffic_agg, sales_agg, mapping)
    E.write_xlsx(records, mapping, out)

    print("=== 全量整合报告 ===")
    print("数据目录: %s" % D)
    print("流量表(%d): %s" % (len([n for t, n in src_files if t == "traffic"]),
                             [n for t, n in src_files if t == "traffic"]))
    print("销售表(%d): %s" % (len([n for t, n in src_files if t == "sales"]),
                             [n for t, n in src_files if t == "sales"]))
    print("丢弃-非产品路径行: %d" % counters["nonproduct"])
    print("丢弃-未匹配 handle 行: %d" % counters["unmatched"])
    print("输出记录数: %d" % len(records))
    print("产品数: %d" % len(mapping))
    print("异常数: %d" % len(anomalies))
    for a in anomalies[:50]:
        print("  异常:", a)
    print("XLSX: %s" % out)


if __name__ == "__main__":
    main()
