from agent_rules import (
    build_system_prompt,
)

from model_gateway import (
    MODEL_API_KEY,
    MODEL_BASE_URL,
    MODEL_NAME,
    LOG_FILE,
    logger,
    call_chat_completion,
    create_client as gateway_create_client,
    get_gateway_status,
)


# =========================================================
# 韩中游戏本地化 Agent
# AI翻译业务层
#
# 职责：
#
# 1. 构建翻译知识库证据
# 2. 构建翻译Prompt
# 3. 调用统一model_gateway
# 4. 将统一网关结果转换成旧业务接口
#
#
# 注意：
#
# 本文件不再自己管理：
#
# - API Key
# - Base URL
# - Timeout
# - Retry
# - HTTP连接
# - API异常分类
# - Model API日志
#
# 这些全部由：
#
# model_gateway.py
#
# 统一负责。
# =========================================================


# =========================================================
# 兼容旧模块
#
# strict_reviewer.py等旧模块目前可能还在：
#
# from ai_translator import create_client
#
# 所以暂时保留这个代理函数。
#
# 下一阶段再把strict_reviewer彻底迁移到model_gateway。
# =========================================================

def create_client():

    return gateway_create_client()


# =========================================================
# 通用文本去重
# =========================================================

def unique_texts(
    values
):

    result = []


    for value in values or []:

        text = str(
            value
            or
            ""
        ).strip()


        if (
            text
            and
            text not in result
        ):

            result.append(
                text
            )


    return result


# =========================================================
# 角色名证据
# =========================================================

def format_roles(
    roles
):

    if not roles:

        return "无"


    lines = []


    for item in roles:

        korean = str(
            item.get(
                "韩文角色名",
                "",
            )
            or
            ""
        ).strip()


        chinese = str(
            item.get(
                "正式中文名",
                "",
            )
            or
            ""
        ).strip()


        if korean and chinese:

            lines.append(
                f"{korean} → {chinese}"
            )


    return (
        "\n".join(
            lines
        )
        if lines
        else
        "无"
    )


# =========================================================
# 正式完整句记录
# =========================================================

def format_exact_records(
    records,
    source_name,
):

    if not records:

        return "无"


    lines = []


    for index, record in enumerate(
        records,
        start=1,
    ):

        korean = str(
            record.get(
                "msgid",
                "",
            )
            or
            ""
        ).strip()


        chinese = str(
            record.get(
                "msgstr[0]",
                "",
            )
            or
            ""
        ).strip()


        references = str(
            record.get(
                "references",
                "",
            )
            or
            ""
        ).strip()


        release = str(
            record.get(
                "Release",
                "",
            )
            or
            ""
        ).strip()


        content = str(
            record.get(
                "Content",
                "",
            )
            or
            ""
        ).strip()


        script = str(
            record.get(
                "script",
                "",
            )
            or
            ""
        ).strip()


        block = [
            f"[{source_name}完整句 {index}]",
            f"韩文：{korean}",
            f"中文：{chinese}",
        ]


        if references:

            block.append(
                f"references：{references}"
            )


        if release:

            block.append(
                f"Release：{release}"
            )


        if content:

            block.append(
                f"Content：{content}"
            )


        if script:

            block.append(
                f"script：{script}"
            )


        lines.append(
            "\n".join(
                block
            )
        )


    return "\n\n".join(
        lines
    )


# =========================================================
# 长术语 / 短语
# =========================================================

def format_long_terms(
    terms,
    source_name,
):

    if not terms:

        return "无"


    lines = []


    for term in terms:

        korean_term = str(
            term.get(
                "韩文术语",
                "",
            )
            or
            ""
        ).strip()


        translations = []


        for record in term.get(
            "历史记录",
            [],
        ):

            chinese = str(
                record.get(
                    "msgstr[0]",
                    "",
                )
                or
                ""
            ).strip()


            if (
                chinese
                and
                chinese not in translations
            ):

                translations.append(
                    chinese
                )


        if not korean_term:

            continue


        if translations:

            lines.append(
                (
                    f"{source_name}："
                    f"{korean_term} → "
                    +
                    " / ".join(
                        translations
                    )
                )
            )

        else:

            lines.append(
                (
                    f"{source_name}："
                    f"{korean_term}"
                )
            )


    return (
        "\n".join(
            lines
        )
        if lines
        else
        "无"
    )


# =========================================================
# 历史上下文
# =========================================================

def format_context_matches(
    context_result,
    source_name,
):

    if not context_result:

        return "无"


    results = context_result.get(
        "结果",
        [],
    )


    if not results:

        return "无"


    lines = []


    for index, item in enumerate(
        results,
        start=1,
    ):

        historical_korean = str(
            item.get(
                "历史韩文",
                "",
            )
            or
            ""
        ).strip()


        translations = []


        for record in item.get(
            "历史记录",
            [],
        ):

            chinese = str(
                record.get(
                    "msgstr[0]",
                    "",
                )
                or
                ""
            ).strip()


            if (
                chinese
                and
                chinese not in translations
            ):

                translations.append(
                    chinese
                )


        block = [
            (
                f"[{source_name}"
                f"历史上下文 {index}]"
            ),
            (
                f"韩文："
                f"{historical_korean}"
            ),
        ]


        if translations:

            block.append(
                (
                    "历史中文："
                    +
                    " / ".join(
                        translations
                    )
                )
            )


        lines.append(
            "\n".join(
                block
            )
        )


    return "\n\n".join(
        lines
    )


# =========================================================
# 构建完整知识库证据
# =========================================================

def build_knowledge_context(
    search_result
):

    search_result = (
        search_result
        or
        {}
    )


    roles = search_result.get(
        "角色名",
        [],
    )


    uwo_exact = search_result.get(
        "UWO完整句",
        [],
    )


    quest_exact = search_result.get(
        "Quest完整句",
        [],
    )


    uwo_terms = search_result.get(
        "UWO长术语",
        [],
    )


    quest_terms = search_result.get(
        "Quest长术语",
        [],
    )


    uwo_context = search_result.get(
        "UWO包含匹配",
        {},
    )


    quest_context = search_result.get(
        "Quest包含匹配",
        {},
    )


    return f"""
【正式角色名】
{format_roles(
    roles
)}

【UWO正式主库完整句】
{format_exact_records(
    uwo_exact,
    "UWO",
)}

【Quest正式主库完整句】
{format_exact_records(
    quest_exact,
    "Quest",
)}

【UWO历史术语 / 短语】
{format_long_terms(
    uwo_terms,
    "UWO",
)}

【Quest历史术语 / 短语】
{format_long_terms(
    quest_terms,
    "Quest",
)}

【UWO历史上下文】
{format_context_matches(
    uwo_context,
    "UWO",
)}

【Quest历史上下文】
{format_context_matches(
    quest_context,
    "Quest",
)}
""".strip()


# =========================================================
# 构建翻译User Prompt
# =========================================================

def build_user_prompt(
    korean_text,
    search_result,
    extra_context="",
    mode="快速翻译",
):

    korean_text = str(
        korean_text
        or
        ""
    )


    extra_context = str(
        extra_context
        or
        ""
    ).strip()


    mode = str(
        mode
        or
        "快速翻译"
    ).strip()


    knowledge_context = (
        build_knowledge_context(
            search_result
        )
    )


    if not extra_context:

        extra_context = (
            "未提供额外上下文。"
        )


    return f"""
当前处理模式：{mode}

请将下面韩文游戏文本翻译成简体中文。

【固定知识库优先级】

角色名库
＞
UWO正式主库
＞
Quest正式主库
＞
当前上下文
＞
游戏内同类术语与命名方式
＞
保守翻译


【必须遵守】

1. 如果存在唯一可靠的正式完整句译文，应直接使用正式译文，不得重新润色。

2. 正式角色名必须严格使用角色名库中的正式中文名。

3. 正式术语必须保持一致，不得自行改写、缩写、换同义词或重新音译。

4. 不得增译、漏译、猜测。

5. 必须保留：
   - 信息
   - 否定
   - 条件
   - 因果
   - 时间
   - 数量
   - 范围
   - 强弱程度
   - 敬语
   - 情绪
   - 逻辑关系

6. 必须完整保留：
   - {{0}}
   - {{1}}
   - %s
   - %d
   - %1$s
   - \\n
   - \\t
   - <br>
   - HTML/XML标签
   - 颜色标签
   - 变量
   - 代码
   - 数字
   - 百分比
   - 换行

7. 允许进行必要的中文语序调整，
   但禁止自由润色。

8. 如果原文是残句、短语或片段，
   不得擅自补全不存在的信息。

9. 如果存在多个历史译法、
   UWO与Quest发生冲突、
   或当前上下文不足，
   必须保守处理。

10. 如果确实不能可靠确定，
    可以在译文末尾追加：

【需人工确认：原因】

11. 除上述人工确认标记外，
    不要输出解释、分析、标题或其他说明。


【韩文原文】

{korean_text}


【当前上下文】

{extra_context}


【本地知识库证据】

{knowledge_context}
""".strip()


# =========================================================
# 旧接口兼容：
# 统一错误返回
#
# 同时保留：
#
# 错误信息
# 错误
#
# 两个字段。
#
# 这样旧模块无论读取哪个都不会出错。
# =========================================================

def make_error_result(
    error_type,
    error_message,
    fatal=True,
    error_stage="",
    elapsed=None,
    model=None,
):

    return {

        "成功":
            False,

        "译文":
            "",

        "错误类型":
            str(
                error_type
                or
                ""
            ),

        "错误阶段":
            str(
                error_stage
                or
                ""
            ),

        "错误信息":
            str(
                error_message
                or
                ""
            ),

        "错误":
            str(
                error_message
                or
                ""
            ),

        "致命错误":
            bool(
                fatal
            ),

        "耗时秒":
            elapsed,

        "模型":
            model
            or
            MODEL_NAME,
    }


# =========================================================
# 统一成功返回
# =========================================================

def make_success_result(
    translation,
    elapsed=None,
    model=None,
):

    return {

        "成功":
            True,

        "译文":
            str(
                translation
                or
                ""
            ).strip(),

        "错误类型":
            "",

        "错误阶段":
            "",

        "错误信息":
            "",

        "错误":
            "",

        "致命错误":
            False,

        "耗时秒":
            elapsed,

        "模型":
            model
            or
            MODEL_NAME,
    }


# =========================================================
# 真正翻译入口
#
# 现在不再：
#
# OpenAI(...)
# client.chat.completions.create(...)
#
# 而是统一：
#
# model_gateway.call_chat_completion(...)
# =========================================================

def translate_with_ai(
    korean_text,
    search_result,
    extra_context="",
    mode="快速翻译",
):

    korean_text = str(
        korean_text
        or
        ""
    )


    mode = str(
        mode
        or
        "快速翻译"
    ).strip()


    if not korean_text.strip():

        return make_error_result(
            error_type=
                "EmptyInput",

            error_message=
                "韩文原文为空。",

            fatal=
                False,

            error_stage=
                "输入检查阶段",
        )


    # =====================================================
    # System Prompt
    # =====================================================

    try:

        system_prompt = (
            build_system_prompt(
                mode=
                    mode
            )
        )


        user_prompt = (
            build_user_prompt(
                korean_text=
                    korean_text,

                search_result=
                    search_result,

                extra_context=
                    extra_context,

                mode=
                    mode,
            )
        )


    except Exception as error:

        logger.exception(
            "翻译Prompt构建失败"
        )


        return make_error_result(
            error_type=
                type(
                    error
                ).__name__,

            error_message=
                str(
                    error
                ),

            fatal=
                True,

            error_stage=
                "Prompt构建阶段",
        )


    # =====================================================
    # 统一调用Gateway
    # =====================================================

    gateway_result = (
        call_chat_completion(
            messages=[
                {
                    "role":
                        "system",

                    "content":
                        system_prompt,
                },
                {
                    "role":
                        "user",

                    "content":
                        user_prompt,
                },
            ],

            mode=
                mode,

            request_label=
                "翻译",
        )
    )


    # =====================================================
    # 成功
    # =====================================================

    if gateway_result.get(
        "成功",
        False,
    ):

        content = str(
            gateway_result.get(
                "内容",
                "",
            )
            or
            ""
        ).strip()


        if not content:

            return make_error_result(
                error_type=
                    "EmptyResponse",

                error_message=
                    "AI模型返回了空译文。",

                fatal=
                    False,

                error_stage=
                    "响应解析阶段",

                elapsed=
                    gateway_result.get(
                        "耗时秒"
                    ),

                model=
                    gateway_result.get(
                        "模型"
                    ),
            )


        return make_success_result(
            translation=
                content,

            elapsed=
                gateway_result.get(
                    "耗时秒"
                ),

            model=
                gateway_result.get(
                    "模型"
                ),
        )


    # =====================================================
    # Gateway失败
    #
    # 转换成旧业务接口，
    # 保证document_processor/xlsx_translator仍然兼容。
    # =====================================================

    error_message = str(
        gateway_result.get(
            "错误信息",
            ""
        )
        or
        ""
    )


    return make_error_result(
        error_type=
            gateway_result.get(
                "错误类型",
                "GatewayError",
            ),

        error_message=
            error_message,

        fatal=
            gateway_result.get(
                "致命错误",
                True,
            ),

        error_stage=
            gateway_result.get(
                "错误阶段",
                "",
            ),

        elapsed=
            gateway_result.get(
                "耗时秒"
            ),

        model=
            gateway_result.get(
                "模型"
            ),
    )


# =========================================================
# Prompt预览
# =========================================================

def preview_prompt(
    korean_text,
    search_result,
    extra_context="",
    mode="快速翻译",
):

    user_prompt = (
        build_user_prompt(
            korean_text=
                korean_text,

            search_result=
                search_result,

            extra_context=
                extra_context,

            mode=
                mode,
        )
    )


    gateway_status = (
        get_gateway_status()
    )


    return {

        "规则来源":
            "agent_rules.py",

        "模型调用层":
            "model_gateway.py",

        "处理模式":
            mode,

        "固定知识库优先级":
            (
                "角色名库 ＞ "
                "UWO正式主库 ＞ "
                "Quest正式主库 ＞ "
                "当前上下文 ＞ "
                "游戏内同类术语与命名方式 ＞ "
                "保守翻译"
            ),

        "正式完整句策略":
            (
                "存在唯一正式完整句时直接复用，"
                "AI不得重新润色。"
            ),

        "AI自然化策略":
            (
                "只允许必要的中文语序调整，"
                "禁止自由润色。"
            ),

        "模型":
            gateway_status.get(
                "模型",
                "",
            ),

        "连接超时":
            gateway_status.get(
                "连接超时",
                "",
            ),

        "响应超时":
            gateway_status.get(
                "响应超时",
                "",
            ),

        "自动重试":
            gateway_status.get(
                "自动重试",
                0,
            ),

        "API Key":
            gateway_status.get(
                "API Key",
                "",
            ),

        "Base URL":
            gateway_status.get(
                "Base URL",
                "",
            ),

        "日志文件":
            gateway_status.get(
                "日志文件",
                str(
                    LOG_FILE
                ),
            ),

        "User Prompt":
            user_prompt,
    }


# =========================================================
# 独立测试
#
# 不真正调用Model API。
# =========================================================

if __name__ == "__main__":

    print(
        "AI翻译业务层加载成功。"
    )


    print(
        "模型调用层：model_gateway.py"
    )


    print(
        f"模型：{MODEL_NAME or '未配置'}"
    )


    print(
        f"日志：{LOG_FILE}"
    )