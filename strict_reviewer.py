import json
import re

from agent_rules import (
    build_system_prompt,
)

from ai_translator import (
    build_knowledge_context,
)

from model_gateway import (
    MODEL_NAME,
    call_chat_completion,
    logger,
)

from document_processor import (
    choose_exact_translation,
)

from format_qa import (
    check_format,
    format_qa_summary,
)


# =========================================================
# 韩中游戏本地化 Agent
# 严格审校业务层
#
# 模型调用统一经过：
#
# model_gateway.py
#
# 本模块负责：
#
# 1. 正式完整句判断
# 2. 角色名检查
# 3. 术语证据整理
# 4. 格式QA
# 5. AI严格审校
# 6. AI结果二次程序校验
# 7. 人工确认升级
#
# 不再自行：
#
# OpenAI(...)
# client.chat.completions.create(...)
#
# =========================================================


VALID_VERDICTS = {
    "通过",
    "建议修改",
    "人工确认",
}


VALID_MATCH_STATUS = {
    "精确匹配",
    "相似匹配",
    "候选",
    "未确认",
}


# =========================================================
# 基础工具
# =========================================================

def clean_text(
    value
):

    if value is None:
        return ""

    return str(
        value
    ).strip()


def unique_texts(
    values
):

    result = []


    for value in values or []:

        text = clean_text(
            value
        )


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
# 角色名命中摘要
# =========================================================

def summarize_role_hits(
    search_result
):

    roles = (
        search_result.get(
            "角色名",
            [],
        )
        if search_result
        else []
    )


    lines = []


    for item in roles:

        korean = clean_text(
            item.get(
                "韩文角色名",
                "",
            )
        )


        chinese = clean_text(
            item.get(
                "正式中文名",
                "",
            )
        )


        if korean and chinese:

            lines.append(
                f"{korean} → {chinese}"
            )


    return "；".join(
        unique_texts(
            lines
        )
    )


# =========================================================
# 历史术语摘要
# =========================================================

def summarize_term_hits(
    search_result
):

    if not search_result:
        return ""


    results = []


    for (
        source_name,
        key,
    ) in [
        (
            "UWO",
            "UWO长术语",
        ),
        (
            "Quest",
            "Quest长术语",
        ),
    ]:

        for item in search_result.get(
            key,
            [],
        ):

            korean_term = clean_text(
                item.get(
                    "韩文术语",
                    "",
                )
            )


            translations = []


            for record in item.get(
                "历史记录",
                [],
            ):

                chinese = clean_text(
                    record.get(
                        "msgstr[0]",
                        "",
                    )
                )


                if (
                    chinese
                    and
                    chinese not in translations
                ):

                    translations.append(
                        chinese
                    )


            if korean_term:

                if translations:

                    results.append(
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

                    results.append(
                        (
                            f"{source_name}："
                            f"{korean_term}"
                        )
                    )


    return "；".join(
        unique_texts(
            results
        )
    )


# =========================================================
# 知识库来源摘要
# =========================================================

def summarize_reference_sources(
    search_result
):

    if not search_result:
        return ""


    sources = []


    if search_result.get(
        "角色名",
        [],
    ):

        sources.append(
            "角色名库"
        )


    if (
        search_result.get(
            "UWO完整句",
            [],
        )
        or
        search_result.get(
            "UWO长术语",
            [],
        )
        or
        search_result.get(
            "UWO包含匹配",
            {},
        ).get(
            "总数",
            0,
        )
        > 0
    ):

        sources.append(
            "UWO正式主库"
        )


    if (
        search_result.get(
            "Quest完整句",
            [],
        )
        or
        search_result.get(
            "Quest长术语",
            [],
        )
        or
        search_result.get(
            "Quest包含匹配",
            {},
        ).get(
            "总数",
            0,
        )
        > 0
    ):

        sources.append(
            "Quest正式主库"
        )


    return " ＞ ".join(
        sources
    )


# =========================================================
# 判断知识库匹配状态
# =========================================================

def determine_match_status(
    search_result,
    exact_decision=None,
):

    search_result = (
        search_result
        or
        {}
    )


    if (
        exact_decision
        and
        exact_decision.get(
            "可直接复用",
            False,
        )
    ):

        return "精确匹配"


    if (
        search_result.get(
            "UWO完整句",
            []
        )
        or
        search_result.get(
            "Quest完整句",
            []
        )
    ):

        # 有完整句证据，但没有唯一结果，
        # 属于冲突/未确认。
        return "未确认"


    if (
        search_result.get(
            "UWO长术语",
            []
        )
        or
        search_result.get(
            "Quest长术语",
            []
        )
        or
        search_result.get(
            "UWO包含匹配",
            {},
        ).get(
            "总数",
            0,
        )
        > 0
        or
        search_result.get(
            "Quest包含匹配",
            {},
        ).get(
            "总数",
            0,
        )
        > 0
    ):

        return "相似匹配"


    return "候选"


# =========================================================
# 检查角色名是否正确使用
# =========================================================

def check_role_usage(
    chinese_text,
    search_result
):

    chinese_text = clean_text(
        chinese_text
    )


    issues = []


    for item in (
        search_result.get(
            "角色名",
            [],
        )
        if search_result
        else []
    ):

        korean_name = clean_text(
            item.get(
                "韩文角色名",
                "",
            )
        )


        official_chinese = clean_text(
            item.get(
                "正式中文名",
                "",
            )
        )


        if (
            not korean_name
            or
            not official_chinese
        ):

            continue


        if official_chinese not in chinese_text:

            issues.append(
                (
                    f"角色名“{korean_name}”"
                    f"必须使用正式译名"
                    f"“{official_chinese}”"
                )
            )


    return issues


# =========================================================
# 合并原因
# =========================================================

def join_reasons(
    *values
):

    result = []


    for value in values:

        if isinstance(
            value,
            list,
        ):

            candidates = value

        else:

            candidates = [
                value
            ]


        for candidate in candidates:

            text = clean_text(
                candidate
            )


            if (
                text
                and
                text not in result
            ):

                result.append(
                    text
                )


    return "；".join(
        result
    )


# =========================================================
# 规范审校结论
# =========================================================

def normalize_verdict(
    value
):

    value = clean_text(
        value
    )


    if value in VALID_VERDICTS:

        return value


    if "人工" in value:

        return "人工确认"


    if (
        "修改" in value
        or
        "修订" in value
    ):

        return "建议修改"


    if (
        "通过" in value
        or
        "正确" in value
    ):

        return "通过"


    return "人工确认"


# =========================================================
# 规范匹配状态
# =========================================================

def normalize_match_status(
    value
):

    value = clean_text(
        value
    )


    if value in VALID_MATCH_STATUS:

        return value


    if "精确" in value:

        return "精确匹配"


    if (
        "相似" in value
        or
        "历史" in value
    ):

        return "相似匹配"


    if "候选" in value:

        return "候选"


    return "未确认"


# =========================================================
# JSON代码块清理
# =========================================================

def strip_code_fence(
    text
):

    text = clean_text(
        text
    )


    if not text:

        return ""


    text = re.sub(
        r"^\s*```(?:json)?\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )


    text = re.sub(
        r"\s*```\s*$",
        "",
        text,
    )


    return text.strip()


# =========================================================
# 从AI响应中提取JSON
# =========================================================

def parse_json_response(
    text
):

    text = strip_code_fence(
        text
    )


    if not text:

        raise ValueError(
            "AI返回内容为空。"
        )


    # =====================================================
    # 第一尝试：完整JSON
    # =====================================================

    try:

        data = json.loads(
            text
        )


        if isinstance(
            data,
            dict,
        ):

            return data


    except Exception:

        pass


    # =====================================================
    # 第二尝试：截取第一个 { 到最后一个 }
    # =====================================================

    start = text.find(
        "{"
    )


    end = text.rfind(
        "}"
    )


    if (
        start >= 0
        and
        end > start
    ):

        fragment = text[
            start:
            end + 1
        ]


        data = json.loads(
            fragment
        )


        if isinstance(
            data,
            dict,
        ):

            return data


    raise ValueError(
        "AI返回内容不是有效JSON对象。"
    )


# =========================================================
# 严格审校Prompt
# =========================================================

def build_strict_review_prompt(
    korean_text,
    existing_chinese,
    search_result,
    extra_context="",
    record_id="",
):

    knowledge_context = (
        build_knowledge_context(
            search_result
        )
    )


    korean_text = clean_text(
        korean_text
    )


    existing_chinese = clean_text(
        existing_chinese
    )


    extra_context = clean_text(
        extra_context
    )


    record_id = clean_text(
        record_id
    )


    if not extra_context:

        extra_context = (
            "未提供额外上下文。"
        )


    if not existing_chinese:

        existing_chinese = (
            "【当前译文为空】"
        )


    if not record_id:

        record_id = (
            "未提供"
        )


    return f"""
你正在执行韩文游戏文本的“严格审校”。

不是自由翻译，也不是文学润色。

【固定优先级】

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
保守判断


【必须检查】

1. 原文信息是否完整保留。

2. 是否存在增译、漏译、猜测。

3. 否定、条件、因果、时间、数量、范围、
   强弱程度、敬语、情绪、逻辑关系是否一致。

4. 正式角色名是否严格使用角色名库译法。

5. 已确认正式术语是否被擅自改写、
   缩写、换同义词或重新音译。

6. 如果知识库存在唯一可靠正式完整句，
   应以正式完整句为准。

7. 多个历史译法发生冲突时，
   不得假装已经确认。

8. 上下文不足时，
   不得编造人物身份、性别、地点、
   行为、原因、结果、系统效果。

9. 必须保护：
   {{0}}
   {{1}}
   %s
   %d
   %1$s
   \\n
   \\t
   <br>
   HTML/XML标签
   颜色标签
   变量
   代码
   数字
   百分比
   换行

10. 仅允许进行必要的中文自然语序调整，
    禁止自由润色。

11. 如果现有中文本身正确，
    必须判定“通过”，不要为了风格偏好强行修改。

12. 如果能够明确修正，
    判定“建议修改”。

13. 如果存在知识库冲突、信息不足、
    身份不明、语义歧义或无法可靠判断，
    判定“人工确认”。


【输出要求】

只能输出一个JSON对象。

不要输出Markdown。
不要输出```json。
不要在JSON前后添加任何解释。

字段必须如下：

{{
  "建议中文": "",
  "审校结论": "通过/建议修改/人工确认",
  "匹配状态": "精确匹配/相似匹配/候选/未确认",
  "问题说明": "",
  "修改原因": "",
  "需人工确认": false,
  "人工确认原因": ""
}}


【ID】

{record_id}


【韩文原文】

{korean_text}


【现有中文】

{existing_chinese}


【当前上下文】

{extra_context}


【本地知识库证据】

{knowledge_context}
""".strip()


# =========================================================
# 统一审校结果
# =========================================================

def make_review_result(
    record_id="",
    korean_text="",
    existing_chinese="",
    suggested_chinese="",
    verdict="人工确认",
    role_hits="",
    term_hits="",
    reference_kb="",
    match_status="未确认",
    issue="",
    reason="",
    existing_format="",
    suggested_format="",
    need_manual=True,
    manual_reason="",
    review_method="",
    ai_success=False,
    error_type="",
    error_message="",
    fatal=False,
    error_stage="",
    elapsed=None,
    model=None,
):

    return {

        "ID":
            clean_text(
                record_id
            ),

        "韩文原文":
            clean_text(
                korean_text
            ),

        "现有中文":
            clean_text(
                existing_chinese
            ),

        "建议中文":
            clean_text(
                suggested_chinese
            ),

        "审校结论":
            normalize_verdict(
                verdict
            ),

        "角色名命中":
            clean_text(
                role_hits
            ),

        "术语命中":
            clean_text(
                term_hits
            ),

        "参考知识库":
            clean_text(
                reference_kb
            ),

        "匹配状态":
            normalize_match_status(
                match_status
            ),

        "问题说明":
            clean_text(
                issue
            ),

        "修改原因":
            clean_text(
                reason
            ),

        "现有译文格式检查":
            clean_text(
                existing_format
            ),

        "建议译文格式检查":
            clean_text(
                suggested_format
            ),

        "需人工确认":
            bool(
                need_manual
            ),

        "人工确认原因":
            clean_text(
                manual_reason
            ),

        "审校方式":
            clean_text(
                review_method
            ),

        "AI成功":
            bool(
                ai_success
            ),

        "错误类型":
            clean_text(
                error_type
            ),

        "错误阶段":
            clean_text(
                error_stage
            ),

        # 保留两个字段兼容旧代码
        "错误":
            clean_text(
                error_message
            ),

        "错误信息":
            clean_text(
                error_message
            ),

        "致命错误":
            bool(
                fatal
            ),

        "耗时秒":
            elapsed,

        "模型":
            clean_text(
                model
                or
                MODEL_NAME
            ),
    }


# =========================================================
# AI严格审校
#
# ★ 现在统一走 model_gateway ★
# =========================================================

def strict_review_with_ai(
    korean_text,
    existing_chinese,
    search_result,
    extra_context="",
    record_id="",
):

    try:

        system_prompt = (
            build_system_prompt(
                mode=
                    "严格审校"
            )
        )


        user_prompt = (
            build_strict_review_prompt(
                korean_text=
                    korean_text,

                existing_chinese=
                    existing_chinese,

                search_result=
                    search_result,

                extra_context=
                    extra_context,

                record_id=
                    record_id,
            )
        )


    except Exception as error:

        logger.exception(
            "严格审校Prompt构建失败"
        )


        return {
            "成功":
                False,

            "数据":
                None,

            "错误类型":
                type(
                    error
                ).__name__,

            "错误阶段":
                "Prompt构建阶段",

            "错误信息":
                str(
                    error
                ),

            "致命错误":
                True,

            "耗时秒":
                None,

            "模型":
                MODEL_NAME,
        }


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
                "严格审校",

            request_label=
                "严格审校",
        )
    )


    # =====================================================
    # Gateway失败
    # =====================================================

    if not gateway_result.get(
        "成功",
        False,
    ):

        return {
            "成功":
                False,

            "数据":
                None,

            "错误类型":
                gateway_result.get(
                    "错误类型",
                    "GatewayError",
                ),

            "错误阶段":
                gateway_result.get(
                    "错误阶段",
                    "",
                ),

            "错误信息":
                gateway_result.get(
                    "错误信息",
                    "",
                ),

            "致命错误":
                gateway_result.get(
                    "致命错误",
                    True,
                ),

            "耗时秒":
                gateway_result.get(
                    "耗时秒"
                ),

            "模型":
                gateway_result.get(
                    "模型",
                    MODEL_NAME,
                ),
        }


    # =====================================================
    # JSON解析
    # =====================================================

    raw_content = clean_text(
        gateway_result.get(
            "内容",
            "",
        )
    )


    try:

        parsed = parse_json_response(
            raw_content
        )


    except Exception as error:

        logger.error(
            (
                "严格审校AI返回JSON解析失败 | "
                "elapsed=%s"
            ),
            gateway_result.get(
                "耗时秒"
            ),
        )


        logger.exception(
            "严格审校JSON解析异常"
        )


        return {
            "成功":
                False,

            "数据":
                None,

            "错误类型":
                "InvalidJSON",

            "错误阶段":
                "响应解析阶段",

            "错误信息":
                str(
                    error
                ),

            # 模型成功响应，只是格式不合规。
            # 不属于网络致命故障，
            # 批量可以记录人工确认，而不是整批退出。
            "致命错误":
                False,

            "耗时秒":
                gateway_result.get(
                    "耗时秒"
                ),

            "模型":
                gateway_result.get(
                    "模型",
                    MODEL_NAME,
                ),
        }


    return {
        "成功":
            True,

        "数据":
            parsed,

        "错误类型":
            "",

        "错误阶段":
            "",

        "错误信息":
            "",

        "致命错误":
            False,

        "耗时秒":
            gateway_result.get(
                "耗时秒"
            ),

        "模型":
            gateway_result.get(
                "模型",
                MODEL_NAME,
            ),
    }


# =========================================================
# 正式完整句本地审校
#
# 唯一正式完整句：
#
# 不调用AI。
# =========================================================

def review_by_unique_exact(
    record_id,
    korean_text,
    existing_chinese,
    official_chinese,
    search_result,
    match_status,
):

    role_hits = (
        summarize_role_hits(
            search_result
        )
    )


    term_hits = (
        summarize_term_hits(
            search_result
        )
    )


    reference_kb = (
        summarize_reference_sources(
            search_result
        )
    )


    existing_format = (
        format_qa_summary(
            korean_text,
            existing_chinese,
        )
        if existing_chinese
        else
        "现有中文为空"
    )


    official_format = (
        format_qa_summary(
            korean_text,
            official_chinese,
        )
    )


    official_qa = check_format(
        korean_text,
        official_chinese,
    )


    official_role_issues = (
        check_role_usage(
            official_chinese,
            search_result,
        )
    )


    # =====================================================
    # 理论上正式知识库也可能存在历史格式异常
    #
    # 这种情况下不能盲目覆盖。
    # =====================================================

    if (
        not official_qa.get(
            "通过",
            False,
        )
        or
        official_role_issues
    ):

        manual_reason = join_reasons(
            (
                ""
                if official_qa.get(
                    "通过",
                    False,
                )
                else
                (
                    "正式完整句自身未通过"
                    "程序级格式QA"
                )
            ),
            official_role_issues,
        )


        return make_review_result(
            record_id=
                record_id,

            korean_text=
                korean_text,

            existing_chinese=
                existing_chinese,

            suggested_chinese=
                official_chinese,

            verdict=
                "人工确认",

            role_hits=
                role_hits,

            term_hits=
                term_hits,

            reference_kb=
                reference_kb,

            match_status=
                match_status,

            issue=
                "正式知识库记录存在程序级风险。",

            reason=
                "",

            existing_format=
                existing_format,

            suggested_format=
                official_format,

            need_manual=
                True,

            manual_reason=
                manual_reason,

            review_method=
                "程序审校-正式完整句",

            ai_success=
                False,
        )


    # =====================================================
    # 当前中文为空
    # =====================================================

    if not existing_chinese:

        return make_review_result(
            record_id=
                record_id,

            korean_text=
                korean_text,

            existing_chinese=
                existing_chinese,

            suggested_chinese=
                official_chinese,

            verdict=
                "建议修改",

            role_hits=
                role_hits,

            term_hits=
                term_hits,

            reference_kb=
                reference_kb,

            match_status=
                match_status,

            issue=
                "现有中文为空。",

            reason=
                "直接采用唯一正式完整句译文。",

            existing_format=
                existing_format,

            suggested_format=
                official_format,

            need_manual=
                False,

            manual_reason=
                "",

            review_method=
                "程序审校-正式完整句",

            ai_success=
                False,
        )


    # =====================================================
    # 完全一致
    # =====================================================

    if clean_text(
        existing_chinese
    ) == clean_text(
        official_chinese
    ):

        existing_qa = check_format(
            korean_text,
            existing_chinese,
        )


        existing_role_issues = (
            check_role_usage(
                existing_chinese,
                search_result,
            )
        )


        if (
            existing_qa.get(
                "通过",
                False,
            )
            and
            not existing_role_issues
        ):

            return make_review_result(
                record_id=
                    record_id,

                korean_text=
                    korean_text,

                existing_chinese=
                    existing_chinese,

                suggested_chinese=
                    existing_chinese,

                verdict=
                    "通过",

                role_hits=
                    role_hits,

                term_hits=
                    term_hits,

                reference_kb=
                    reference_kb,

                match_status=
                    match_status,

                issue=
                    "",

                reason=
                    "现有中文与唯一正式完整句一致。",

                existing_format=
                    existing_format,

                suggested_format=
                    official_format,

                need_manual=
                    False,

                manual_reason=
                    "",

                review_method=
                    "程序审校-正式完整句",

                ai_success=
                    False,
            )


    # =====================================================
    # 和唯一正式译文不同
    # =====================================================

    return make_review_result(
        record_id=
            record_id,

        korean_text=
            korean_text,

        existing_chinese=
            existing_chinese,

        suggested_chinese=
            official_chinese,

        verdict=
            "建议修改",

        role_hits=
            role_hits,

        term_hits=
            term_hits,

        reference_kb=
            reference_kb,

        match_status=
            match_status,

        issue=
            "现有中文与唯一正式完整句不一致。",

        reason=
            "建议统一为正式知识库译文。",

        existing_format=
            existing_format,

        suggested_format=
            official_format,

        need_manual=
            False,

        manual_reason=
            "",

        review_method=
            "程序审校-正式完整句",

        ai_success=
            False,
    )


# =========================================================
# 严格审校总入口
# =========================================================

def review_text(
    korean_text,
    existing_chinese,
    knowledge_base,
    extra_context="",
    record_id="",
):

    korean_text = clean_text(
        korean_text
    )


    existing_chinese = clean_text(
        existing_chinese
    )


    extra_context = clean_text(
        extra_context
    )


    record_id = clean_text(
        record_id
    )


    # =====================================================
    # 韩文为空
    # =====================================================

    if not korean_text:

        return make_review_result(
            record_id=
                record_id,

            korean_text=
                korean_text,

            existing_chinese=
                existing_chinese,

            suggested_chinese=
                existing_chinese,

            verdict=
                "人工确认",

            match_status=
                "未确认",

            issue=
                "韩文原文为空。",

            reason=
                "",

            existing_format=
                "未执行",

            suggested_format=
                "未执行",

            need_manual=
                True,

            manual_reason=
                "缺少韩文原文，无法执行严格审校。",

            review_method=
                "输入检查",

            ai_success=
                False,

            fatal=
                False,
        )


    # =====================================================
    # 查询知识库
    # =====================================================

    try:

        search_result = (
            knowledge_base.search(
                korean_text
            )
        )


    except Exception as error:

        logger.exception(
            "严格审校知识库查询失败"
        )


        return make_review_result(
            record_id=
                record_id,

            korean_text=
                korean_text,

            existing_chinese=
                existing_chinese,

            suggested_chinese=
                existing_chinese,

            verdict=
                "人工确认",

            match_status=
                "未确认",

            issue=
                "知识库查询失败。",

            reason=
                "",

            existing_format=
                (
                    format_qa_summary(
                        korean_text,
                        existing_chinese,
                    )
                    if existing_chinese
                    else
                    "现有中文为空"
                ),

            suggested_format=
                "未执行",

            need_manual=
                True,

            manual_reason=
                str(
                    error
                ),

            review_method=
                "知识库查询失败",

            ai_success=
                False,

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
                "知识库查询阶段",
        )


    # =====================================================
    # 正式完整句判断
    # =====================================================

    exact_decision = (
        choose_exact_translation(
            search_result
        )
    )


    match_status = (
        determine_match_status(
            search_result,
            exact_decision,
        )
    )


    # =====================================================
    # 唯一正式完整句：
    # 完全本地审校，不调用AI
    # =====================================================

    if exact_decision.get(
        "可直接复用",
        False,
    ):

        official_chinese = clean_text(
            exact_decision.get(
                "译文",
                "",
            )
        )


        if official_chinese:

            return review_by_unique_exact(
                record_id=
                    record_id,

                korean_text=
                    korean_text,

                existing_chinese=
                    existing_chinese,

                official_chinese=
                    official_chinese,

                search_result=
                    search_result,

                match_status=
                    match_status,
            )


    # =====================================================
    # 无唯一正式完整句
    #
    # 需要AI严格审校。
    # =====================================================

    ai_result = (
        strict_review_with_ai(
            korean_text=
                korean_text,

            existing_chinese=
                existing_chinese,

            search_result=
                search_result,

            extra_context=
                extra_context,

            record_id=
                record_id,
        )
    )


    role_hits = (
        summarize_role_hits(
            search_result
        )
    )


    term_hits = (
        summarize_term_hits(
            search_result
        )
    )


    reference_kb = (
        summarize_reference_sources(
            search_result
        )
    )


    existing_format = (
        format_qa_summary(
            korean_text,
            existing_chinese,
        )
        if existing_chinese
        else
        "现有中文为空"
    )


    existing_qa = (
        check_format(
            korean_text,
            existing_chinese,
        )
        if existing_chinese
        else
        {
            "通过":
                False
        }
    )


    existing_role_issues = (
        check_role_usage(
            existing_chinese,
            search_result,
        )
        if existing_chinese
        else []
    )


    # =====================================================
    # AI失败
    # =====================================================

    if not ai_result.get(
        "成功",
        False,
    ):

        error_message = clean_text(
            ai_result.get(
                "错误信息",
                "",
            )
        )


        return make_review_result(
            record_id=
                record_id,

            korean_text=
                korean_text,

            existing_chinese=
                existing_chinese,

            suggested_chinese=
                existing_chinese,

            verdict=
                "人工确认",

            role_hits=
                role_hits,

            term_hits=
                term_hits,

            reference_kb=
                reference_kb,

            match_status=
                match_status,

            issue=
                "AI严格审校未能正常完成。",

            reason=
                "",

            existing_format=
                existing_format,

            suggested_format=
                existing_format,

            need_manual=
                True,

            manual_reason=
                error_message
                or
                "AI严格审校失败。",

            review_method=
                "AI严格审校失败",

            ai_success=
                False,

            error_type=
                ai_result.get(
                    "错误类型",
                    "",
                ),

            error_message=
                error_message,

            fatal=
                ai_result.get(
                    "致命错误",
                    False,
                ),

            error_stage=
                ai_result.get(
                    "错误阶段",
                    "",
                ),

            elapsed=
                ai_result.get(
                    "耗时秒"
                ),

            model=
                ai_result.get(
                    "模型"
                ),
        )


    # =====================================================
    # AI结果
    # =====================================================

    data = (
        ai_result.get(
            "数据",
            {}
        )
        or
        {}
    )


    suggested_chinese = clean_text(
        data.get(
            "建议中文",
            "",
        )
    )


    verdict = normalize_verdict(
        data.get(
            "审校结论",
            "人工确认",
        )
    )


    ai_match_status = (
        normalize_match_status(
            data.get(
                "匹配状态",
                match_status,
            )
        )
    )


    issue = clean_text(
        data.get(
            "问题说明",
            "",
        )
    )


    reason = clean_text(
        data.get(
            "修改原因",
            "",
        )
    )


    ai_need_manual = bool(
        data.get(
            "需人工确认",
            False,
        )
    )


    ai_manual_reason = clean_text(
        data.get(
            "人工确认原因",
            "",
        )
    )


    # =====================================================
    # AI没有提供建议译文
    # =====================================================

    if not suggested_chinese:

        if existing_chinese:

            suggested_chinese = (
                existing_chinese
            )

        else:

            verdict = (
                "人工确认"
            )

            ai_need_manual = True

            ai_manual_reason = join_reasons(
                ai_manual_reason,
                "AI没有返回可用的建议中文。",
            )


    # =====================================================
    # 建议译文程序QA
    # =====================================================

    suggested_qa = check_format(
        korean_text,
        suggested_chinese,
    )


    suggested_format = (
        format_qa_summary(
            korean_text,
            suggested_chinese,
        )
    )


    suggested_role_issues = (
        check_role_usage(
            suggested_chinese,
            search_result,
        )
    )


    manual_reasons = []


    # =====================================================
    # 正式完整句冲突
    #
    # 如果完整句存在，但没有唯一结果，
    # 不能让AI自行宣布已确认。
    # =====================================================

    if (
        exact_decision.get(
            "存在风险",
            False,
        )
        or
        (
            (
                search_result.get(
                    "UWO完整句",
                    []
                )
                or
                search_result.get(
                    "Quest完整句",
                    []
                )
            )
            and
            not exact_decision.get(
                "可直接复用",
                False,
            )
        )
    ):

        manual_reasons.append(
            (
                "正式完整句存在多个历史译法或知识库冲突，"
                "无法自动确认唯一译文。"
            )
        )


    # =====================================================
    # AI建议格式异常
    # =====================================================

    if not suggested_qa.get(
        "通过",
        False,
    ):

        manual_reasons.append(
            "AI建议译文未通过程序级格式QA。"
        )


    # =====================================================
    # 正式角色名异常
    # =====================================================

    if suggested_role_issues:

        manual_reasons.extend(
            suggested_role_issues
        )


    # =====================================================
    # AI主动要求人工确认
    # =====================================================

    if ai_need_manual:

        manual_reasons.append(
            ai_manual_reason
            or
            "AI判断需要人工确认。"
        )


    # =====================================================
    # 最终人工确认
    # =====================================================

    if manual_reasons:

        verdict = (
            "人工确认"
        )


        ai_need_manual = (
            True
        )


        ai_manual_reason = (
            join_reasons(
                manual_reasons
            )
        )


        # 完整句冲突时匹配状态必须未确认
        if (
            search_result.get(
                "UWO完整句",
                []
            )
            or
            search_result.get(
                "Quest完整句",
                []
            )
        ):

            ai_match_status = (
                "未确认"
            )


    # =====================================================
    # 如果AI说“通过”，
    # 但原译文本身程序QA失败，
    # 不能判通过。
    # =====================================================

    if verdict == "通过":

        pass_problems = []


        if not existing_chinese:

            pass_problems.append(
                "现有中文为空。"
            )


        if (
            existing_chinese
            and
            not existing_qa.get(
                "通过",
                False,
            )
        ):

            pass_problems.append(
                "现有中文未通过程序级格式QA。"
            )


        if existing_role_issues:

            pass_problems.extend(
                existing_role_issues
            )


        if pass_problems:

            # 如果AI建议译文可用，
            # 转为建议修改。
            if (
                suggested_chinese
                and
                suggested_qa.get(
                    "通过",
                    False,
                )
                and
                not suggested_role_issues
            ):

                verdict = (
                    "建议修改"
                )


                reason = join_reasons(
                    reason,
                    pass_problems,
                )


            else:

                verdict = (
                    "人工确认"
                )


                ai_need_manual = True


                ai_manual_reason = (
                    join_reasons(
                        ai_manual_reason,
                        pass_problems,
                    )
                )


    # =====================================================
    # “通过”时建议中文必须保持现有中文
    #
    # 防止AI一边说通过一边偷偷重写。
    # =====================================================

    if (
        verdict == "通过"
        and
        existing_chinese
    ):

        suggested_chinese = (
            existing_chinese
        )


        suggested_format = (
            existing_format
        )


    # =====================================================
    # 无现有译文却判通过
    # =====================================================

    if (
        verdict == "通过"
        and
        not existing_chinese
    ):

        verdict = (
            "建议修改"
        )


        reason = join_reasons(
            reason,
            "现有中文为空，不能判定通过。",
        )


    # =====================================================
    # 最终结果
    # =====================================================

    return make_review_result(
        record_id=
            record_id,

        korean_text=
            korean_text,

        existing_chinese=
            existing_chinese,

        suggested_chinese=
            suggested_chinese,

        verdict=
            verdict,

        role_hits=
            role_hits,

        term_hits=
            term_hits,

        reference_kb=
            reference_kb,

        match_status=
            ai_match_status,

        issue=
            issue,

        reason=
            reason,

        existing_format=
            existing_format,

        suggested_format=
            suggested_format,

        need_manual=
            ai_need_manual,

        manual_reason=
            ai_manual_reason,

        review_method=
            "AI严格审校",

        ai_success=
            True,

        error_type=
            "",

        error_message=
            "",

        fatal=
            False,

        error_stage=
            "",

        elapsed=
            ai_result.get(
                "耗时秒"
            ),

        model=
            ai_result.get(
                "模型"
            ),
    )


# =========================================================
# 独立加载测试
#
# 不实际调用Model API。
# =========================================================

if __name__ == "__main__":

    print(
        "严格审校业务层加载成功。"
    )


    print(
        "模型调用层：model_gateway.py"
    )


    print(
        f"模型：{MODEL_NAME or '未配置'}"
    )