"""
匹配模块：中文姓名 + 拼音（全拼/首字母）+ rapidfuzz 模糊匹配。
"""

import os
import re
from rapidfuzz import fuzz
from pypinyin import lazy_pinyin, Style

# 中文数字 ↔ 阿拉伯数字互转（支持 0-10）
_CN2ARAB = {'零':'0','一':'1','二':'2','三':'3','四':'4','五':'5','六':'6','七':'7','八':'8','九':'9','十':'10'}
_ARAB2CN = {v: k for k, v in _CN2ARAB.items()}


def _extract_digits(keyword):
    """从关键词中提取数字部分，返回 (阿拉伯数字字符串, 中文数字字符串)。

    同时支持阿拉伯数字（"实验2"→"2"）和中文数字（"实验二"→"一"）输入。
    若关键词不含数字，返回 (None, None)。
    """
    # 尝试阿拉伯数字
    m = re.findall(r'\d+', keyword)
    if m:
        arabic = m[-1]
        chinese = ''.join(_ARAB2CN.get(d, d) for d in arabic)
        return arabic, chinese
    # 尝试中文数字
    m = re.findall(r'[零一二三四五六七八九十]+', keyword)
    if m:
        chinese = m[-1]
        arabic = ''.join(_CN2ARAB.get(c, c) for c in chinese)
        return arabic, chinese
    return None, None


def _build_variants(arabic, chinese, keyword):
    """根据数字生成中英文搜索变体列表。

    例如 arabic='3', chinese='三' 产生：
        [实验3, part3, lab3, 作业3, 实验三, part三, lab三, 作业三]
    """
    prefixes = ["实验", "part", "lab", "作业", "Part", "PART"]
    seen = {keyword}
    variants = [keyword]
    for p in prefixes:
        for num in (arabic, chinese):
            v = f"{p}{num}"
            if v not in seen:
                seen.add(v)
                variants.append(v)
    return variants


def build_pinyin_map(name_id_map):
    """为每个学生姓名生成拼音全拼和首字母简写。

    Args:
        name_id_map: {姓名: 学号} 映射字典

    Returns:
        dict: {姓名: {'full': 'zhangsan', 'initial': 'zs'}}
    """
    pinyin_map = {}
    for name in name_id_map:
        # 全拼（无空格）：张三 → zhangsan
        full = "".join(lazy_pinyin(name, style=Style.NORMAL))
        # 首字母简写：张三 → zs
        initial = "".join(lazy_pinyin(name, style=Style.FIRST_LETTER))
        pinyin_map[name] = {"full": full, "initial": initial}
    return pinyin_map


def _keyword_matches(base, extract_keyword):
    """检查文件名是否包含用户指定关键词的中英文变体（含中文数字互转）。

    Args:
        base: 文件名（不含扩展名）
        extract_keyword: 用户输入的关键词（如"实验3"或"实验二"）

    Returns:
        bool: 文件名中是否匹配到关键词或其变体
    """
    arabic, chinese = _extract_digits(extract_keyword)
    if not arabic:
        return True  # 关键词无数字时放行
    variants = _build_variants(arabic, chinese, extract_keyword)
    base_lower = base.lower()
    return any(v.lower() in base_lower for v in variants)


def generate_new_name(template, student_id, name, base, ext, extract_keyword=""):
    """根据命名模板生成新文件名。

    支持的占位符：
        {学号}    —— 学生学号
        {姓名}    —— 学生姓名
        {原文件名} —— 原文件名（不含扩展名）
        {扩展名}   —— 原扩展名（含点号，如 .docx）
        {实验号}   —— 从原文件名自动提取"实验N/partN/作业N"等
        {匹配项}   —— 根据 extract_keyword 在文件名中查找对应文本
                      （如输入"实验3"可匹配"part3"并统一输出"实验3"）
    若模板未包含 {扩展名}，则自动追加扩展名。

    Args:
        template: 命名模板字符串
        student_id: 学号
        name: 学生姓名
        base: 原文件名（不含扩展名）
        ext: 原扩展名（含点号）
        extract_keyword: 用户指定的提取关键词（如"实验3"），
                         工具自动生成中英文变体在文件名中查找

    Returns:
        str: 生成的新文件名
    """
    # 提取实验号（实验1 / 实验二 / part2 / 作业3 / lab4 等，含中文数字）
    lab_match = re.search(
        r'(?:实验|part|作业|lab)[.\s]*(\d+|[零一二三四五六七八九十]+)', base, re.IGNORECASE
    )
    lab_number = lab_match.group(0) if lab_match else ""

    # 处理 {匹配项}：根据用户输入的关键词，在文件名中查找中英文变体（含中文数字互转）
    match_text = ""
    if extract_keyword:
        arabic, chinese = _extract_digits(extract_keyword)
        if arabic:
            variants = _build_variants(arabic, chinese, extract_keyword)
            base_lower = base.lower()
            for v in variants:
                if v.lower() in base_lower:
                    match_text = extract_keyword
                    break

    new_name = template.replace("{学号}", student_id)
    new_name = new_name.replace("{姓名}", name)
    new_name = new_name.replace("{原文件名}", base)
    new_name = new_name.replace("{实验号}", lab_number)
    new_name = new_name.replace("{匹配项}", match_text)

    if "{扩展名}" in template:
        new_name = new_name.replace("{扩展名}", ext)
    else:
        new_name += ext

    return new_name


def match_files(filenames, name_id_map, pinyin_map, template, threshold=80, enable_pinyin=True, extract_keyword=""):
    """使用「中文→拼音→简写」三层匹配 + rapidfuzz 回退，将文件名与学生姓名进行比对。

    匹配规则（按优先级）：
    1. 文件名包含中文姓名         → score = 100，匹配方式 = 中文
    2. 文件名包含拼音全拼（大小写不敏感）→ score = 90，匹配方式 = 拼音
    3. 文件名包含拼音首字母简写   → score = 80，匹配方式 = 简写
    4. 以上均不满足时，使用 rapidfuzz 对中文姓名及拼音变体模糊匹配
    5. 若存在多个学生分数与最高分差值 < 5，标记为多重匹配

    Args:
        filenames: 文件名列表（含扩展名）
        name_id_map: {姓名: 学号} 映射字典
        pinyin_map: {姓名: {'full': 'zhangsan', 'initial': 'zs'}}
        template: 命名模板字符串
        threshold: 模糊匹配相似度阈值（默认 80）
        enable_pinyin: 是否启用拼音匹配
        extract_keyword: 用户指定的提取关键词，用于 {匹配项} 占位符

    Returns:
        list[dict]: 每条结果包含 原文件名、建议新文件名、状态、匹配分数、匹配方式
    """
    results = []
    name_list = list(name_id_map.keys())

    for full_name in filenames:
        base, ext = os.path.splitext(full_name)
        base_lower = base.lower()

        candidates = []  # [(name, score, match_type), ...]

        # 第一轮：精确子串匹配（中文 / 拼音全拼 / 拼音首字母）
        for name in name_list:
            # 中文姓名匹配
            if name in base:
                candidates.append((name, 100, "中文"))
                continue

            if not enable_pinyin:
                continue

            pinyin = pinyin_map[name]

            # 拼音全拼匹配（大小写不敏感）
            if pinyin["full"] and pinyin["full"].lower() in base_lower:
                candidates.append((name, 90, "拼音"))
                continue

            # 拼音首字母简写匹配（大小写不敏感）
            if pinyin["initial"] and pinyin["initial"].lower() in base_lower:
                candidates.append((name, 80, "简写"))

        # 第二轮：精确匹配未命中时，使用 rapidfuzz 模糊匹配回退
        if not candidates:
            for name in name_list:
                # 中文姓名模糊匹配
                score_chinese = fuzz.partial_ratio(name, base)

                best_fuzzy = score_chinese
                match_type = "中文"

                if enable_pinyin:
                    pinyin = pinyin_map[name]
                    score_full = fuzz.partial_ratio(pinyin["full"], base_lower)
                    score_initial = fuzz.partial_ratio(pinyin["initial"], base_lower)

                    if score_full > best_fuzzy:
                        best_fuzzy = score_full
                        match_type = "拼音"
                    if score_initial > best_fuzzy:
                        best_fuzzy = score_initial
                        match_type = "简写"

                if best_fuzzy >= threshold:
                    candidates.append((name, best_fuzzy, match_type))

        # 无任何匹配 → 未匹配
        if not candidates:
            results.append({
                "原文件名": full_name,
                "建议新文件名": "",
                "状态": "❌ 未匹配",
                "匹配分数": 0,
                "匹配方式": "",
            })
            continue

        # 按分数降序排列
        candidates.sort(key=lambda x: x[1], reverse=True)
        best_score = candidates[0][1]

        # 检查是否存在与最高分接近的其他匹配（差值 < 5）
        close_matches = [c for c in candidates if best_score - c[1] < 5]
        if len(close_matches) > 1:
            results.append({
                "原文件名": full_name,
                "建议新文件名": "",
                "状态": "⚠️ 多重匹配",
                "匹配分数": best_score,
                "匹配方式": "",
            })
            continue

        best_name, best_score, match_type = candidates[0]
        student_id = name_id_map[best_name]

        # 关键词过滤：若设置了 extract_keyword，文件名必须包含其变体才允许重命名
        if extract_keyword:
            if not _keyword_matches(base, extract_keyword):
                results.append({
                    "原文件名": full_name,
                    "建议新文件名": "",
                    "状态": "❌ 未匹配",
                    "匹配分数": 0,
                    "匹配方式": "",
                })
                continue

        new_name = generate_new_name(template, student_id, best_name, base, ext, extract_keyword)
        results.append({
            "原文件名": full_name,
            "建议新文件名": new_name,
            "状态": "✅ 已匹配",
            "匹配分数": best_score,
            "匹配方式": match_type,
        })

    return results
