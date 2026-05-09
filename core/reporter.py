"""
输出模块：终端表格打印 + CSV 导出。
"""

import csv
import os
from tabulate import tabulate


def export_results(results, output_path="match_result.csv"):
    """导出结果到 CSV 并打印终端表格。

    Args:
        results: match_files 返回的结果列表
        output_path: CSV 输出路径
    """
    fieldnames = ["原文件名", "建议新文件名", "状态", "匹配分数", "匹配方式"]

    # 导出 CSV（UTF-8-BOM 编码，确保 Excel 正常打开）
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(results)

    print(f"\n结果已导出至：{os.path.abspath(output_path)}")

    # 终端表格输出
    print(
        tabulate(
            results,
            headers="keys",
            tablefmt="grid",
            showindex=False,
            stralign="left",
        )
    )

    # 统计摘要
    matched = sum(1 for r in results if r["状态"] == "✅ 已匹配")
    unmatched = sum(1 for r in results if r["状态"] == "❌ 未匹配")
    multi = sum(1 for r in results if r["状态"] == "⚠️ 多重匹配")
    print(f"\n统计：共 {len(results)} 个文件 | ✅ 已匹配 {matched} | ❌ 未匹配 {unmatched} | ⚠️ 多重匹配 {multi}")
