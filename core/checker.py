"""
提交检查模块：根据匹配结果与名单比对，检测多交 / 缺交 / 正常。
"""


def check_submissions(results, name_id_map):
    """根据匹配结果检查每位学生的提交情况。

    对 Excel 名单中的每位学生：
        - 匹配到 0 个文件 → 缺交
        - 匹配到 1 个文件 → 正常
        - 匹配到 2+ 个文件 → 多交

    Args:
        results: match_files 返回的结果列表
        name_id_map: {姓名: 学号} 映射字典

    Returns:
        list[dict]: 每条包含 姓名、学号、提交数、状态、匹配文件
    """
    # 统计每个学生的匹配文件
    submission = {}  # {姓名: [原文件名列表]}
    for r in results:
        if r["状态"] == "✅ 已匹配":
            # 从建议新文件名中反查姓名（也可以从 results 中取）
            # 实际上 results 里没有直接存姓名，需要从 name_id_map 匹配
            # 换个思路：在 load_data 阶段 already has name→id mapping
            # 这里我们用另一种方式：遍历 name_id_map 找匹配
            pass

    # 更可靠的做法：遍历 name_id_map，对每个姓名在 results 中搜索
    report = []

    for name, student_id in name_id_map.items():
        matched_files = []
        for r in results:
            if r["状态"] == "✅ 已匹配":
                # 检查建议新文件名中是否含该姓名
                new_name = r.get("建议新文件名", "")
                if name in new_name:
                    matched_files.append(r["原文件名"])

        count = len(matched_files)
        if count == 0:
            status = "❌ 缺交"
        elif count == 1:
            status = "✅ 正常"
        else:
            status = "⚠️ 多交"

        report.append({
            "姓名":   name,
            "学号":   student_id,
            "提交数": count,
            "状态":   status,
            "匹配文件": " / ".join(matched_files) if matched_files else "",
        })

    return report


def print_check_report(report):
    """在终端打印提交检查摘要。

    Args:
        report: check_submissions 返回的报告列表
    """
    normal  = [r for r in report if r["状态"] == "✅ 正常"]
    missing = [r for r in report if r["状态"] == "❌ 缺交"]
    extra   = [r for r in report if r["状态"] == "⚠️ 多交"]

    print("\n" + "=" * 50)
    print("📋 提交情况检查")
    print("=" * 50)

    print(f"\n✅ 正常提交：{len(normal)} 人")
    if normal:
        names = "、".join(r["姓名"] for r in normal)
        print(f"   {names}")

    print(f"\n❌ 缺交：{len(missing)} 人")
    if missing:
        for r in missing:
            print(f"   {r['姓名']}（{r['学号']}）")

    print(f"\n⚠️ 多交：{len(extra)} 人")
    if extra:
        for r in extra:
            print(f"   {r['姓名']}（{r['学号']}）—— {r['提交数']} 个文件")
            for f in r["匹配文件"].split(" / "):
                print(f"      · {f}")

    total = len(report)
    print(f"\n总计 {total} 人 | 正常 {len(normal)} | 缺交 {len(missing)} | 多交 {len(extra)}")
