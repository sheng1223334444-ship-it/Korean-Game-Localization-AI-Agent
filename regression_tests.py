import sys

from format_qa import (
    check_format,
)

from final_qa import (
    protect_newlines,
    restore_newlines,
    check_newline_tokens,
    normalize_chinese_punctuation,
    check_final_translation,
    find_korean_residue,
)

from document_processor import (
    contains_korean,
    choose_exact_translation,
    remove_manual_confirmation_marker,
    raise_if_fatal_ai_error,
    BatchAIUnavailableError,
    translate_text_unit,
)

from xlsx_translator import (
    strip_manual_marker,
)

from model_gateway import (
    classify_api_status_error,
)


# =========================================================
# 韩中游戏本地化 Agent
# 回归测试
#
# 每次修改核心程序后运行：
#
#     python regression_tests.py
#
# 必须：
#
#     通过：31
#     失败：0
#     总计：31
#
#
# 特点：
#
# - 不调用Model API
# - 不消耗Token
# - 不修改知识库
# - 不写Excel文件
# - 只验证程序核心规则
# =========================================================


PASSED = 0
FAILED = 0


# =========================================================
# 测试执行器
# =========================================================

def run_test(
    name,
    test_function,
):

    global PASSED
    global FAILED


    try:

        result = (
            test_function()
        )


        if result:

            PASSED += 1

            print(
                f"[PASS] {name}"
            )


        else:

            FAILED += 1

            print(
                f"[FAIL] {name}"
            )


    except Exception as error:

        FAILED += 1

        print(
            f"[FAIL] {name}"
        )

        print(
            f"       {type(error).__name__}: {error}"
        )


# =========================================================
# 工具
# =========================================================

def is_pass(
    result
):

    return bool(
        result.get(
            "通过",
            False,
        )
    )


# =========================================================
# 01
# 韩文检测
# =========================================================

def test_contains_korean():

    return (
        contains_korean(
            "루디가 온다."
        )
        is True
    )


# =========================================================
# 02
# 中文不应被判断为韩文
# =========================================================

def test_no_korean():

    return (
        contains_korean(
            "鲁迪来了。"
        )
        is False
    )


# =========================================================
# 03
# 数字保持
# =========================================================

def test_number_preserved():

    result = check_format(
        "100골드를 획득했다.",
        "获得100金币。",
    )


    return is_pass(
        result
    )


# =========================================================
# 04
# 数字变化必须抓到
# =========================================================

def test_number_changed():

    result = check_format(
        "5회 사용할 수 있다.",
        "可以使用6次。",
    )


    return not is_pass(
        result
    )


# =========================================================
# 05
# 花括号占位符保持
# =========================================================

def test_curly_placeholder_preserved():

    result = check_format(
        "{0}골드를 획득했습니다.",
        "获得了{0}金币。",
    )


    return is_pass(
        result
    )


# =========================================================
# 06
# 花括号占位符丢失
# =========================================================

def test_curly_placeholder_missing():

    result = check_format(
        "{0}골드를 획득했습니다.",
        "获得了金币。",
    )


    return not is_pass(
        result
    )


# =========================================================
# 07
# printf占位符
# =========================================================

def test_printf_preserved():

    result = check_format(
        "%s님이 접속했습니다.",
        "%s已登录。",
    )


    return is_pass(
        result
    )


# =========================================================
# 08
# HTML / Color标签
# =========================================================

def test_html_tag_preserved():

    result = check_format(
        "<color=#FF0000>위험</color>",
        "<color=#FF0000>危险</color>",
    )


    return is_pass(
        result
    )


# =========================================================
# 09
# 百分号和数字
# =========================================================

def test_percent_preserved():

    result = check_format(
        "성공률 50%",
        "成功率50%",
    )


    return is_pass(
        result
    )


# =========================================================
# 10
# 真实换行数量一致
# =========================================================

def test_real_newline_preserved():

    result = check_format(
        "첫째 줄\n둘째 줄",
        "第一行\n第二行",
    )


    return is_pass(
        result
    )


# =========================================================
# 11
# 换行丢失必须失败
# =========================================================

def test_real_newline_missing():

    result = check_format(
        "첫째 줄\n둘째 줄",
        "第一行第二行",
    )


    return not is_pass(
        result
    )


# =========================================================
# 12
# 最终中文存在韩文残留
# =========================================================

def test_korean_residue_detected():

    result = check_final_translation(
        "루디, 시작하자.",
        "루디，开始吧。",
    )


    return (
        result.get(
            "通过"
        )
        is False
        and
        "루디"
        in
        result.get(
            "韩文残留",
            [],
        )
    )


# =========================================================
# 13
# 正常中文不能误报韩文
# =========================================================

def test_normal_chinese_final_qa():

    result = check_final_translation(
        "루디가 온다.",
        "鲁迪来了。",
    )


    return (
        result.get(
            "通过"
        )
        is True
    )


# =========================================================
# 14
# 受保护变量中的韩文不误报
# =========================================================

def test_protected_korean_not_false_positive():

    residue = find_korean_residue(
        "角色变量：{루디}"
    )


    return residue == []


# =========================================================
# 15
# 单省略号自动变成中文双省略号
# =========================================================

def test_single_ellipsis_normalized():

    result = normalize_chinese_punctuation(
        "好吧…继续。"
    )


    return (
        result.get(
            "文本"
        )
        ==
        "好吧……继续。"
    )


# =========================================================
# 16
# 英文三个句点变中文省略号
# =========================================================

def test_three_dots_normalized():

    result = normalize_chinese_punctuation(
        "好吧...继续。"
    )


    return (
        result.get(
            "文本"
        )
        ==
        "好吧……继续。"
    )


# =========================================================
# 17
# ?! → ？！
# =========================================================

def test_question_exclamation_normalized():

    result = normalize_chinese_punctuation(
        "真的吗?!"
    )


    return (
        result.get(
            "文本"
        )
        ==
        "真的吗？！"
    )


# =========================================================
# 18
# !? → ！？
# =========================================================

def test_exclamation_question_normalized():

    result = normalize_chinese_punctuation(
        "什么!?"
    )


    return (
        result.get(
            "文本"
        )
        ==
        "什么！？"
    )


# =========================================================
# 19
# 普通ASCII中文标点自动规范化
# =========================================================

def test_common_ascii_punctuation_normalized():

    result = normalize_chinese_punctuation(
        "你好,世界!说明:测试;"
    )


    return (
        result.get(
            "文本"
        )
        ==
        "你好，世界！说明：测试；"
    )


# =========================================================
# 20
# 千位数字逗号不能改
# =========================================================

def test_number_comma_protected():

    result = normalize_chinese_punctuation(
        "获得1,000金币"
    )


    return (
        result.get(
            "文本"
        )
        ==
        "获得1,000金币"
    )


# =========================================================
# 21
# 时间冒号不能改
# =========================================================

def test_time_colon_protected():

    result = normalize_chinese_punctuation(
        "开放时间12:30"
    )


    return (
        result.get(
            "文本"
        )
        ==
        "开放时间12:30"
    )


# =========================================================
# 22
# URL必须原样保护
# =========================================================

def test_url_protected():

    original = (
        "https://example.com/test?a=1"
    )


    result = normalize_chinese_punctuation(
        original
    )


    return (
        result.get(
            "文本"
        )
        ==
        original
    )


# =========================================================
# 23
# 真实换行保护 + 恢复
# =========================================================

def test_newline_protect_and_restore():

    original = (
        "좋아...\n"
        "가능성은 있어."
    )


    protected, newline_types = (
        protect_newlines(
            original
        )
    )


    if (
        "⟦NL_1⟧"
        not in
        protected
    ):

        return False


    if len(
        newline_types
    ) != 1:

        return False


    simulated_ai = (
        "好吧……"
        "⟦NL_1⟧"
        "还是有可能的。"
    )


    restored = restore_newlines(
        simulated_ai,
        newline_types,
    )


    return (
        restored
        ==
        "好吧……\n还是有可能的。"
    )


# =========================================================
# 24
# AI删除换行Token必须被发现
# =========================================================

def test_missing_newline_token_detected():

    newline_types = [
        "\n"
    ]


    result = check_newline_tokens(
        "第一行第二行",
        newline_types,
    )


    return (
        result.get(
            "通过"
        )
        is False
        and
        "⟦NL_1⟧"
        in
        result.get(
            "缺失Token",
            [],
        )
    )


# =========================================================
# 25
# document_processor人工确认标记
#
# 正文与标记必须先分离。
# =========================================================

def test_document_manual_marker_cleanup():

    raw = (
        "鲁迪来了。\n"
        "【需人工确认：角色名需要确认】"
    )


    (
        cleaned,
        has_marker,
        reason,
    ) = remove_manual_confirmation_marker(
        raw
    )


    return (
        cleaned
        ==
        "鲁迪来了。"
        and
        has_marker
        is True
        and
        "角色名需要确认"
        in
        reason
    )


# =========================================================
# 26
# XLSX层防御性人工确认标记清理
# =========================================================

def test_xlsx_manual_marker_cleanup():

    raw = (
        "鲁迪来了。"
        "【需人工确认：角色名不确定】"
    )


    cleaned, reason = (
        strip_manual_marker(
            raw
        )
    )


    return (
        cleaned
        ==
        "鲁迪来了。"
        and
        "角色名不确定"
        in
        reason
    )


# =========================================================
# 27
# 完整句优先级：
#
# UWO ＞ Quest
# =========================================================

def test_exact_uwo_priority():

    search_result = {
        "角色名":
            [],

        "UWO完整句":
            [
                {
                    "msgstr[0]":
                        "UWO正式译文"
                }
            ],

        "Quest完整句":
            [
                {
                    "msgstr[0]":
                        "Quest正式译文"
                }
            ],
    }


    decision = choose_exact_translation(
        search_result
    )


    return (
        decision.get(
            "可直接复用"
        )
        is True
        and
        decision.get(
            "译文"
        )
        ==
        "UWO正式译文"
        and
        decision.get(
            "来源"
        )
        ==
        "UWO正式主库"
    )


# =========================================================
# 28
# UWO完整句多历史译文
#
# 不能随便选一个。
# =========================================================

def test_exact_uwo_conflict():

    search_result = {
        "角色名":
            [],

        "UWO完整句":
            [
                {
                    "msgstr[0]":
                        "译文A"
                },
                {
                    "msgstr[0]":
                        "译文B"
                },
            ],

        "Quest完整句":
            [
                {
                    "msgstr[0]":
                        "Quest译文"
                }
            ],
    }


    decision = choose_exact_translation(
        search_result
    )


    candidates = (
        decision.get(
            "候选译文",
            [],
        )
    )


    return (
        decision.get(
            "可直接复用"
        )
        is False
        and
        decision.get(
            "存在风险"
        )
        is True
        and
        "译文A"
        in
        candidates
        and
        "译文B"
        in
        candidates
    )


# =========================================================
# 29
# 批量AI错误安全停止
# =========================================================

def test_batch_fatal_error_guard():

    # -----------------------------------------------------
    # 非致命错误不得触发整批异常
    # -----------------------------------------------------

    try:

        raise_if_fatal_ai_error(
            {
                "致命错误":
                    False,

                "错误":
                    "普通非致命错误",
            },

            location=
                "测试位置",
        )


    except Exception:

        return False


    # -----------------------------------------------------
    # 致命错误必须触发
    # -----------------------------------------------------

    try:

        raise_if_fatal_ai_error(
            {
                "致命错误":
                    True,

                "错误类型":
                    "ConnectTimeout",

                "错误":
                    "测试致命错误",
            },

            location=
                "Excel第10行",
        )


    except BatchAIUnavailableError as error:

        return (
            error.location
            ==
            "Excel第10行"
            and
            error.error_type
            ==
            "ConnectTimeout"
        )


    except Exception:

        return False


    return False


# =========================================================
# 30
# 正式完整句必须直接复用
#
# - 不调用Model API
# - 外层空格保留
# - UWO完整句直接复用
# - 最终QA通过
# =========================================================

class DummyKnowledgeBase:

    def search(
        self,
        text,
    ):

        if text != "안녕하세요":

            raise AssertionError(
                (
                    "translate_text_unit传给知识库的"
                    f"文本不正确：{repr(text)}"
                )
            )


        return {
            "角色名":
                [],

            "UWO完整句":
                [
                    {
                        "msgstr[0]":
                            "你好"
                    }
                ],

            "Quest完整句":
                [],

            "UWO长术语":
                [],

            "Quest长术语":
                [],

            "UWO包含匹配":
                [],

            "Quest包含匹配":
                [],
        }


def test_exact_reuse_without_ai():

    knowledge_base = (
        DummyKnowledgeBase()
    )


    result = translate_text_unit(
        text=
            "  안녕하세요  ",

        knowledge_base=
            knowledge_base,

        extra_context=
            "",
    )


    return (
        result.get(
            "译文"
        )
        ==
        "  你好  "
        and
        result.get(
            "处理方式"
        )
        ==
        "正式完整句直接复用"
        and
        result.get(
            "来源"
        )
        ==
        "UWO正式主库"
        and
        result.get(
            "AI成功"
        )
        is False
        and
        result.get(
            "最终QA通过"
        )
        is True
    )


# =========================================================
# 31
# Model API额度不足识别
#
# 不真正请求Model API。
#
# 模拟公司网关返回：
#
# HTTP 403
# +
# token quota is not enough
# +
# insufficient_quota
#
# 必须识别成：
#
# QuotaInsufficient
# 模型额度检查
# =========================================================

class DummyQuotaError(
    Exception
):

    def __init__(
        self
    ):

        super().__init__(
            (
                "Error code: 403 - "
                "token quota is not enough"
            )
        )


        self.status_code = 403


        self.body = {
            "error": {
                "message":
                    "token quota is not enough",

                "code":
                    "insufficient_quota",
            }
        }


        self.code = (
            "insufficient_quota"
        )


def test_model_quota_insufficient_classification():

    error = (
        DummyQuotaError()
    )


    result = (
        classify_api_status_error(
            error
        )
    )


    return (
        result.get(
            "错误类型"
        )
        ==
        "QuotaInsufficient"
        and
        result.get(
            "错误阶段"
        )
        ==
        "模型额度检查"
        and
        result.get(
            "状态码"
        )
        ==
        403
        and
        "额度不足"
        in
        result.get(
            "用户说明",
            "",
        )
    )


# =========================================================
# 主测试列表
# =========================================================

TESTS = [

    (
        "01 韩文检测",
        test_contains_korean,
    ),

    (
        "02 非韩文检测",
        test_no_korean,
    ),

    (
        "03 数字保持",
        test_number_preserved,
    ),

    (
        "04 数字变化检测",
        test_number_changed,
    ),

    (
        "05 花括号占位符保持",
        test_curly_placeholder_preserved,
    ),

    (
        "06 花括号占位符丢失检测",
        test_curly_placeholder_missing,
    ),

    (
        "07 printf占位符保持",
        test_printf_preserved,
    ),

    (
        "08 HTML标签保持",
        test_html_tag_preserved,
    ),

    (
        "09 百分号保持",
        test_percent_preserved,
    ),

    (
        "10 真实换行保持",
        test_real_newline_preserved,
    ),

    (
        "11 真实换行丢失检测",
        test_real_newline_missing,
    ),

    (
        "12 韩文残留检测",
        test_korean_residue_detected,
    ),

    (
        "13 正常中文最终QA",
        test_normal_chinese_final_qa,
    ),

    (
        "14 受保护韩文变量不误报",
        test_protected_korean_not_false_positive,
    ),

    (
        "15 单省略号规范化",
        test_single_ellipsis_normalized,
    ),

    (
        "16 英文省略号规范化",
        test_three_dots_normalized,
    ),

    (
        "17 ?!规范化",
        test_question_exclamation_normalized,
    ),

    (
        "18 !?规范化",
        test_exclamation_question_normalized,
    ),

    (
        "19 ASCII标点规范化",
        test_common_ascii_punctuation_normalized,
    ),

    (
        "20 数字千位逗号保护",
        test_number_comma_protected,
    ),

    (
        "21 时间冒号保护",
        test_time_colon_protected,
    ),

    (
        "22 URL保护",
        test_url_protected,
    ),

    (
        "23 换行保护与恢复",
        test_newline_protect_and_restore,
    ),

    (
        "24 换行Token缺失检测",
        test_missing_newline_token_detected,
    ),

    (
        "25 主流程人工确认标记清理",
        test_document_manual_marker_cleanup,
    ),

    (
        "26 XLSX人工确认标记清理",
        test_xlsx_manual_marker_cleanup,
    ),

    (
        "27 UWO完整句优先级",
        test_exact_uwo_priority,
    ),

    (
        "28 UWO多历史译文冲突",
        test_exact_uwo_conflict,
    ),

    (
        "29 批量致命错误安全停止",
        test_batch_fatal_error_guard,
    ),

    (
        "30 正式完整句无AI直接复用",
        test_exact_reuse_without_ai,
    ),

    (
        "31 Model API额度不足识别",
        test_model_quota_insufficient_classification,
    ),
]


# =========================================================
# MAIN
# =========================================================

def main():

    global PASSED
    global FAILED


    PASSED = 0
    FAILED = 0


    print(
        "="
        *
        70
    )

    print(
        "KoreanTranslatorAgent 回归测试"
    )

    print(
        "本测试不会调用Model API，不消耗Token。"
    )

    print(
        "="
        *
        70
    )


    for name, function in TESTS:

        run_test(
            name,
            function,
        )


    total = (
        PASSED
        +
        FAILED
    )


    print()

    print(
        "="
        *
        70
    )

    print(
        f"通过：{PASSED}"
    )

    print(
        f"失败：{FAILED}"
    )

    print(
        f"总计：{total}"
    )

    print(
        "="
        *
        70
    )


    if FAILED == 0:

        print(
            "✅ 全部回归测试通过。"
        )

        print(
            "当前核心翻译链路基线正常。"
        )


        return 0


    print(
        "❌ 存在回归测试失败。"
    )

    print(
        "在修复失败项之前，不建议继续修改其他核心模块。"
    )


    return 1


if __name__ == "__main__":

    sys.exit(
        main()
    )