"""
重命名模块：对已匹配的文件执行实际重命名操作，并记录日志以支持撤销。
"""

import csv
import os
from datetime import datetime


def apply_renaming(results, folder_path, log_path=None):
    """对已匹配的文件执行实际重命名操作。

    安全校验：
    1. 仅处理状态为「✅ 已匹配」的文件
    2. 目标文件已存在时跳过
    3. 原文件不存在时跳过
    4. 新旧文件名相同时跳过
    5. 所有异常均捕获，不中断程序
    6. 若提供 log_path，每次操作写入日志 CSV（UTF-8-BOM）

    Args:
        results: match_files 返回的结果列表
        folder_path: 文件所在文件夹路径
        log_path: 日志文件路径（可选），用于记录操作以支持撤销
    """
    print("\n开始执行重命名...\n")

    renamed = 0
    skipped = 0
    log_entries = []

    for r in results:
        if r["状态"] != "✅ 已匹配":
            continue

        old_name = r["原文件名"]
        new_name = r["建议新文件名"]

        old_path = os.path.join(folder_path, old_name)
        new_path = os.path.join(folder_path, new_name)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status = ""
        reason = ""

        try:
            if old_name == new_name:
                status, reason = "跳过", "新旧文件名相同"
                print(f"  [!]跳过（新旧文件名相同）: {old_name}")
                skipped += 1
                continue

            if not os.path.isfile(old_path):
                status, reason = "跳过", "原文件不存在"
                print(f"  [!]跳过（原文件不存在）: {old_name}")
                skipped += 1
                continue

            if os.path.exists(new_path):
                status, reason = "跳过", "目标已存在"
                print(f"  [!]跳过（目标已存在）: {new_name}")
                skipped += 1
                continue

            os.rename(old_path, new_path)
            status, reason = "成功", ""
            print(f"  [OK]已重命名: {old_name} -> {new_name}")
            renamed += 1

        except Exception as e:
            status, reason = "跳过", str(e)
            print(f"  [!]跳过（异常）: {old_name} —— {e}")
            skipped += 1

        finally:
            if log_path:
                log_entries.append({
                    "时间":   timestamp,
                    "原文件名": old_name,
                    "新文件名": new_name,
                    "状态":   status,
                    "备注":   reason,
                })

    # 保存日志
    if log_path and log_entries:
        _write_log(log_path, log_entries)
        print(f"\n[Log] 操作日志已保存至：{os.path.abspath(log_path)}")

    print(f"\n重命名完成：[OK]{renamed} 个成功 | [!]{skipped} 个跳过")


def undo_renaming(log_path, folder_path):
    """根据日志文件撤销重命名操作（将文件从新名称恢复为旧名称）。

    Args:
        log_path: 操作日志 CSV 路径
        folder_path: 文件所在文件夹路径
    """
    if not os.path.exists(log_path):
        print(f"错误：日志文件不存在 —— {log_path}")
        return

    # 读取日志
    entries = []
    try:
        with open(log_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("状态") == "成功":
                    entries.append(row)
    except Exception as e:
        print(f"错误：无法读取日志文件 —— {e}")
        return

    if not entries:
        print("日志中没有可撤销的成功操作。")
        return

    print(f"\n开始撤销重命名（共 {len(entries)} 条记录）...\n")

    undone = 0
    failed = 0

    for entry in entries:
        old_name = entry["原文件名"]
        new_name = entry["新文件名"]

        # 撤销：当前文件（新名称）→ 旧名称
        current_path = os.path.join(folder_path, new_name)
        restore_path = os.path.join(folder_path, old_name)

        try:
            if not os.path.isfile(current_path):
                print(f"  [!]跳过（文件不存在）: {new_name}")
                failed += 1
                continue

            if os.path.exists(restore_path):
                print(f"  [!]跳过（目标已存在）: {old_name}")
                failed += 1
                continue

            os.rename(current_path, restore_path)
            print(f"  [Undo]已撤销: {new_name} -> {old_name}")
            undone += 1

        except Exception as e:
            print(f"  [!]跳过（异常）: {new_name} —— {e}")
            failed += 1

    print(f"\n撤销完成：[Undo]{undone} 个成功 | [!]{failed} 个失败")


def _write_log(log_path, entries):
    """将操作记录写入 CSV 日志（UTF-8-BOM 编码）。

    Args:
        log_path: 日志文件路径
        entries: list[dict]，每条包含 时间/原文件名/新文件名/状态/备注
    """
    # 确保日志目录存在
    log_dir = os.path.dirname(log_path)
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)

    fieldnames = ["时间", "原文件名", "新文件名", "状态", "备注"]
    with open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(entries)
