import re

from format_qa import (
    check_format,
    format_qa_summary,
)


# =========================================================
# 韩中游戏本地化 Agent
# 最终译文 QA + 安全标点规范化
#
# 功能：
#
# 1. 真实换行保护 / 恢复
# 2. 韩文残留检测
# 3. 中文标点QA
# 4. 安全标点自动规范化
#
#
# 标点自动规范化原则：
#
# 可以确定安全的情况才自动修改。
#
# 保护：
# - URL
# - Email
# - HTML/XML标签
# - 花括号变量
# - 方括号标识符
# - printf占位符
# - $变量
# - 数字中的逗号/冒号
#
# =========================================================


# =========================================================
# 韩文
# =========================================================

KOREAN_PATTERN = re.compile(
    r"[\u1100-\u11FF"
    r"\u3130-\u318F"
    r"\uA960-\uA97F"
    r"\uAC00-\uD7AF"
    r"\uD7B0-\uD7FF]+"
)


# =========================================================
# 换行Token
# =========================================================

NEWLINE_TOKEN_PREFIX = "⟦NL_"
NEWLINE_TOKEN_SUFFIX = "⟧"

NEWLINE_TOKEN_PATTERN = re.compile(
    r"⟦NL_\d+⟧"
)


# =========================================================
# 受保护结构
# =========================================================

PROTECTED_PATTERNS = [

    # URL
    re.compile(
        r"https?://[^\s]+",
        re.IGNORECASE,
    ),

    # Email
    re.compile(
        r"[A-Za-z0-9._%+-]+"
        r"@[A-Za-z0-9.-]+"
        r"\.[A-Za-z]{2,}"
    ),

    # HTML/XML标签
    re.compile(
        r"<[^<>]+>"
    ),

    # 花括号变量
    re.compile(
        r"\{[^{}\r\n]*\}"
    ),

    # 方括号标识符
    re.compile(
        r"\[[^\[\]\r\n]+\]"
    ),

    # printf
    re.compile(
        r"%(?:\d+\$)?[sdif]"
    ),

    # $变量
    re.compile(
        r"\$[A-Za-z_][A-Za-z0-9_]*"
    ),

    # 换行Token
    NEWLINE_TOKEN_PATTERN,
]


# =========================================================
# 换行保护
# =========================================================

def protect_newlines(
    text
):

    text = str(
        text
        or
        ""
    )


    if not text:

        return (
            "",
            [],
        )


    newline_types = []

    output = []

    index = 0

    token_index = 1


    while index < len(
        text
    ):

        # CRLF
        if text.startswith(
            "\r\n",
            index,
        ):

            newline_types.append(
                "\r\n"
            )


            output.append(
                (
                    f"{NEWLINE_TOKEN_PREFIX}"
                    f"{token_index}"
                    f"{NEWLINE_TOKEN_SUFFIX}"
                )
            )


            token_index += 1

            index += 2

            continue


        # LF
        if text[
            index
        ] == "\n":

            newline_types.append(
                "\n"
            )


            output.append(
                (
                    f"{NEWLINE_TOKEN_PREFIX}"
                    f"{token_index}"
                    f"{NEWLINE_TOKEN_SUFFIX}"
                )
            )


            token_index += 1

            index += 1

            continue


        # CR
        if text[
            index
        ] == "\r":

            newline_types.append(
                "\r"
            )


            output.append(
                (
                    f"{NEWLINE_TOKEN_PREFIX}"
                    f"{token_index}"
                    f"{NEWLINE_TOKEN_SUFFIX}"
                )
            )


            token_index += 1

            index += 1

            continue


        output.append(
            text[
                index
            ]
        )


        index += 1


    return (
        "".join(
            output
        ),
        newline_types,
    )


# =========================================================
# 换行Token检查
# =========================================================

def check_newline_tokens(
    translated_text,
    newline_types,
):

    translated_text = str(
        translated_text
        or
        ""
    )


    newline_types = (
        newline_types
        or
        []
    )


    missing = []


    for index in range(
        1,
        len(
            newline_types
        )
        + 1,
    ):

        token = (
            f"{NEWLINE_TOKEN_PREFIX}"
            f"{index}"
            f"{NEWLINE_TOKEN_SUFFIX}"
        )


        if token not in translated_text:

            missing.append(
                token
            )


    return {
        "通过":
            len(
                missing
            )
            ==
            0,

        "缺失Token":
            missing,
    }


# =========================================================
# 恢复换行
# =========================================================

def restore_newlines(
    text,
    newline_types,
):

    text = str(
        text
        or
        ""
    )


    newline_types = (
        newline_types
        or
        []
    )


    for index, newline in enumerate(
        newline_types,
        start=1,
    ):

        token = (
            f"{NEWLINE_TOKEN_PREFIX}"
            f"{index}"
            f"{NEWLINE_TOKEN_SUFFIX}"
        )


        text = text.replace(
            token,
            newline,
        )


    return text


# =========================================================
# 查找受保护区间
# =========================================================

def get_protected_spans(
    text
):

    text = str(
        text
        or
        ""
    )


    spans = []


    for pattern in PROTECTED_PATTERNS:

        for match in pattern.finditer(
            text
        ):

            spans.append(
                (
                    match.start(),
                    match.end(),
                )
            )


    if not spans:

        return []


    spans.sort()


    merged = []


    for start, end in spans:

        if not merged:

            merged.append(
                [
                    start,
                    end,
                ]
            )

            continue


        previous = merged[
            -1
        ]


        if start <= previous[
            1
        ]:

            previous[
                1
            ] = max(
                previous[
                    1
                ],
                end,
            )

        else:

            merged.append(
                [
                    start,
                    end,
                ]
            )


    return [
        tuple(
            item
        )
        for item in merged
    ]


# =========================================================
# 是否处于受保护区域
# =========================================================

def position_is_protected(
    position,
    spans
):

    for start, end in spans:

        if (
            start
            <=
            position
            <
            end
        ):

            return True


    return False


# =========================================================
# 移除受保护区域
#
# 只用于扫描。
# 不修改真正译文。
# =========================================================

def remove_protected_regions(
    text
):

    text = str(
        text
        or
        ""
    )


    spans = get_protected_spans(
        text
    )


    if not spans:

        return text


    characters = list(
        text
    )


    for start, end in spans:

        for index in range(
            start,
            end,
        ):

            characters[
                index
            ] = " "


    return "".join(
        characters
    )


# =========================================================
# 韩文残留
# =========================================================

def find_korean_residue(
    chinese_text
):

    scan_text = (
        remove_protected_regions(
            chinese_text
        )
    )


    matches = (
        KOREAN_PATTERN.findall(
            scan_text
        )
    )


    result = []


    for match in matches:

        value = str(
            match
            or
            ""
        ).strip()


        if (
            value
            and
            value not in result
        ):

            result.append(
                value
            )


    return result


# =========================================================
# 是否是数字内部字符
#
# 例如：
#
# 1,000
# 12:30
#
# 这里不自动转换。
# =========================================================

def is_between_digits(
    text,
    index
):

    if (
        index <= 0
        or
        index >=
        len(
            text
        )
        - 1
    ):

        return False


    return (
        text[
            index - 1
        ].isdigit()
        and
        text[
            index + 1
        ].isdigit()
    )


# =========================================================
# 安全标点规范化
#
# 返回：
#
# {
#   "文本": "...",
#   "是否修改": True,
#   "修改项": [...]
# }
#
# =========================================================

def normalize_chinese_punctuation(
    text
):

    text = str(
        text
        or
        ""
    )


    if not text:

        return {
            "文本":
                text,

            "是否修改":
                False,

            "修改项":
                [],
        }


    spans = get_protected_spans(
        text
    )


    output = []

    changes = []

    index = 0


    while index < len(
        text
    ):

        # =================================================
        # 受保护区域直接复制
        # =================================================

        protected_span = None


        for start, end in spans:

            if start == index:

                protected_span = (
                    start,
                    end,
                )

                break


        if protected_span:

            start, end = (
                protected_span
            )


            output.append(
                text[
                    start:end
                ]
            )


            index = end

            continue


        # =================================================
        # ?!
        # =================================================

        if text.startswith(
            "?!",
            index,
        ):

            output.append(
                "？！"
            )


            changes.append(
                "?! → ？！"
            )


            index += 2

            continue


        # =================================================
        # !?
        # =================================================

        if text.startswith(
            "!?",
            index,
        ):

            output.append(
                "！？"
            )


            changes.append(
                "!? → ！？"
            )


            index += 2

            continue


        # =================================================
        # 三点及以上英文省略号
        # =================================================

        if text.startswith(
            "...",
            index,
        ):

            end = index


            while (
                end
                <
                len(
                    text
                )
                and
                text[
                    end
                ]
                ==
                "."
            ):

                end += 1


            count = (
                end
                -
                index
            )


            if count >= 3:

                output.append(
                    "……"
                )


                changes.append(
                    (
                        f"{'.' * count}"
                        " → ……"
                    )
                )


                index = end

                continue


        char = text[
            index
        ]


        # =================================================
        # 单个中文省略号
        #
        # 已经是“……”时保持不变。
        # =================================================

        if char == "…":

            next_char = (
                text[
                    index + 1
                ]
                if
                index + 1
                <
                len(
                    text
                )
                else
                ""
            )


            previous_char = (
                text[
                    index - 1
                ]
                if index > 0
                else
                ""
            )


            if (
                previous_char
                !=
                "…"
                and
                next_char
                !=
                "…"
            ):

                output.append(
                    "……"
                )


                changes.append(
                    "… → ……"
                )


                index += 1

                continue


        # =================================================
        # 问号
        # =================================================

        if char == "?":

            output.append(
                "？"
            )


            changes.append(
                "? → ？"
            )


            index += 1

            continue


        # =================================================
        # 感叹号
        # =================================================

        if char == "!":

            output.append(
                "！"
            )


            changes.append(
                "! → ！"
            )


            index += 1

            continue


        # =================================================
        # ASCII逗号
        #
        # 数字：
        # 1,000
        #
        # 不修改。
        # =================================================

        if char == ",":

            if is_between_digits(
                text,
                index,
            ):

                output.append(
                    char
                )

            else:

                output.append(
                    "，"
                )


                changes.append(
                    ", → ，"
                )


            index += 1

            continue


        # =================================================
        # 分号
        # =================================================

        if char == ";":

            output.append(
                "；"
            )


            changes.append(
                "; → ；"
            )


            index += 1

            continue


        # =================================================
        # 冒号
        #
        # 12:30不修改。
        # =================================================

        if char == ":":

            if is_between_digits(
                text,
                index,
            ):

                output.append(
                    char
                )

            else:

                output.append(
                    "："
                )


                changes.append(
                    ": → ："
                )


            index += 1

            continue


        # =================================================
        # 其他字符原样
        # =================================================

        output.append(
            char
        )


        index += 1


    normalized = "".join(
        output
    )


    return {
        "文本":
            normalized,

        "是否修改":
            normalized
            !=
            text,

        "修改项":
            list(
                dict.fromkeys(
                    changes
                )
            ),
    }


# =========================================================
# 标点QA
# =========================================================

def check_chinese_punctuation(
    source_text,
    translated_text
):

    translated_text = str(
        translated_text
        or
        ""
    )


    scan_text = (
        remove_protected_regions(
            translated_text
        )
    )


    issues = []


    if "?!" in scan_text:

        issues.append(
            "译文含 ?!，中文应使用？！"
        )


    if "!?" in scan_text:

        issues.append(
            "译文含 !?，中文应使用！？"
        )


    if "?" in scan_text:

        issues.append(
            "译文仍含ASCII问号 ?，通常应使用？"
        )


    if "!" in scan_text:

        issues.append(
            "译文仍含ASCII感叹号 !，通常应使用！"
        )


    # =====================================================
    # ASCII逗号
    #
    # 排除数字内部：
    # 1,000
    # =====================================================

    for index, char in enumerate(
        scan_text
    ):

        if (
            char == ","
            and
            not is_between_digits(
                scan_text,
                index,
            )
        ):

            issues.append(
                "译文正文仍含ASCII逗号 ,，通常应使用，"
            )

            break


    if ";" in scan_text:

        issues.append(
            "译文正文仍含ASCII分号 ;，通常应使用；"
        )


    # =====================================================
    # ASCII冒号
    #
    # 排除：
    # 12:30
    # =====================================================

    for index, char in enumerate(
        scan_text
    ):

        if (
            char == ":"
            and
            not is_between_digits(
                scan_text,
                index,
            )
        ):

            issues.append(
                "译文正文仍含ASCII冒号 :，通常应使用："
            )

            break


    if "..." in scan_text:

        issues.append(
            "译文正文仍含英文省略号 ...，通常应使用……"
        )


    single_ellipsis_pattern = re.compile(
        r"(?<!…)…(?!…)"
    )


    if single_ellipsis_pattern.search(
        scan_text
    ):

        issues.append(
            "译文存在单个省略号 …，通常应规范为……"
        )


    return {
        "通过":
            len(
                issues
            )
            ==
            0,

        "问题":
            issues,
    }


# =========================================================
# 最终QA
# =========================================================

def check_final_translation(
    source_text,
    final_translation,
):

    source_text = str(
        source_text
        or
        ""
    )


    final_translation = str(
        final_translation
        or
        ""
    )


    issues = []


    # =====================================================
    # 原format_qa
    # =====================================================

    base_result = check_format(
        source_text,
        final_translation,
    )


    base_summary = (
        format_qa_summary(
            source_text,
            final_translation,
        )
    )


    if not base_result.get(
        "通过",
        False,
    ):

        issues.append(
            base_summary
        )


    # =====================================================
    # 韩文残留
    # =====================================================

    korean_residue = (
        find_korean_residue(
            final_translation
        )
    )


    if korean_residue:

        issues.append(
            (
                "译文存在韩文残留："
                +
                " / ".join(
                    korean_residue
                )
            )
        )


    # =====================================================
    # 标点
    # =====================================================

    punctuation_result = (
        check_chinese_punctuation(
            source_text,
            final_translation,
        )
    )


    if not punctuation_result.get(
        "通过",
        False,
    ):

        issues.extend(
            punctuation_result.get(
                "问题",
                [],
            )
        )


    return {
        "通过":
            len(
                issues
            )
            ==
            0,

        "问题":
            issues,

        "基础格式QA":
            base_result,

        "基础格式摘要":
            base_summary,

        "韩文残留":
            korean_residue,

        "标点QA":
            punctuation_result,
    }


# =========================================================
# QA摘要
# =========================================================

def final_qa_summary(
    source_text,
    final_translation,
):

    result = (
        check_final_translation(
            source_text,
            final_translation,
        )
    )


    if result.get(
        "通过",
        False,
    ):

        return "通过"


    return (
        "不通过："
        +
        "；".join(
            result.get(
                "问题",
                [],
            )
        )
    )


# =========================================================
# 是否存在风险
# =========================================================

def has_final_qa_risk(
    source_text,
    final_translation,
):

    return not (
        check_final_translation(
            source_text,
            final_translation,
        ).get(
            "通过",
            False,
        )
    )


# =========================================================
# 独立测试
# =========================================================

if __name__ == "__main__":

    print(
        "最终译文QA + 安全标点规范化模块加载成功。"
    )


    # =====================================================
    # 1 单省略号
    # =====================================================

    result = (
        normalize_chinese_punctuation(
            "好，只要熟悉了雪橇…"
        )
    )


    print(
        "单省略号自动规范化：",
        (
            "PASS"
            if result[
                "文本"
            ]
            ==
            "好，只要熟悉了雪橇……"
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 2 英文省略号
    # =====================================================

    result = (
        normalize_chinese_punctuation(
            "好吧...继续。"
        )
    )


    print(
        "英文省略号自动规范化：",
        (
            "PASS"
            if result[
                "文本"
            ]
            ==
            "好吧……继续。"
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 3 ?!
    # =====================================================

    result = (
        normalize_chinese_punctuation(
            "真的吗?!"
        )
    )


    print(
        "?!自动规范化：",
        (
            "PASS"
            if result[
                "文本"
            ]
            ==
            "真的吗？！"
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 4 !?
    # =====================================================

    result = (
        normalize_chinese_punctuation(
            "什么!?"
        )
    )


    print(
        "!?自动规范化：",
        (
            "PASS"
            if result[
                "文本"
            ]
            ==
            "什么！？"
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 5 普通ASCII标点
    # =====================================================

    result = (
        normalize_chinese_punctuation(
            "你好,世界!说明:测试;"
        )
    )


    print(
        "普通ASCII标点规范化：",
        (
            "PASS"
            if result[
                "文本"
            ]
            ==
            "你好，世界！说明：测试；"
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 6 数字逗号保护
    # =====================================================

    result = (
        normalize_chinese_punctuation(
            "获得1,000金币"
        )
    )


    print(
        "数字逗号保护：",
        (
            "PASS"
            if result[
                "文本"
            ]
            ==
            "获得1,000金币"
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 7 时间冒号保护
    # =====================================================

    result = (
        normalize_chinese_punctuation(
            "开放时间12:30"
        )
    )


    print(
        "时间冒号保护：",
        (
            "PASS"
            if result[
                "文本"
            ]
            ==
            "开放时间12:30"
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 8 URL保护
    # =====================================================

    url_text = (
        "访问https://example.com/test?a=1"
    )


    result = (
        normalize_chinese_punctuation(
            url_text
        )
    )


    print(
        "URL保护：",
        (
            "PASS"
            if result[
                "文本"
            ]
            ==
            url_text
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 9 HTML保护
    # =====================================================

    tag_text = (
        "<color=#FF0000>危险!</color>"
    )


    result = (
        normalize_chinese_punctuation(
            tag_text
        )
    )


    print(
        "HTML标签保护：",
        (
            "PASS"
            if
            "<color=#FF0000>"
            in
            result[
                "文本"
            ]
            and
            "</color>"
            in
            result[
                "文本"
            ]
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 10 韩文残留
    # =====================================================

    result = check_final_translation(
        "루디가 온다.",
        "루디来了。",
    )


    print(
        "韩文残留检测：",
        (
            "PASS"
            if not result[
                "通过"
            ]
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 11 正常中文
    # =====================================================

    result = check_final_translation(
        "루디가 온다.",
        "鲁迪来了。",
    )


    print(
        "正常中文QA：",
        (
            "PASS"
            if result[
                "通过"
            ]
            else
            "FAIL"
        ),
    )


    # =====================================================
    # 12 换行
    # =====================================================

    original = (
        "좋아...\n"
        "가능성은 있어."
    )


    protected, newline_types = (
        protect_newlines(
            original
        )
    )


    simulated_ai = (
        "好吧……"
        "⟦NL_1⟧"
        "还是有可能的。"
    )


    token_result = (
        check_newline_tokens(
            simulated_ai,
            newline_types,
        )
    )


    restored = (
        restore_newlines(
            simulated_ai,
            newline_types,
        )
    )


    print(
        "换行Token保护：",
        (
            "PASS"
            if token_result[
                "通过"
            ]
            else
            "FAIL"
        ),
    )


    print(
        "换行恢复：",
        (
            "PASS"
            if "\n" in restored
            else
            "FAIL"
        ),
    )