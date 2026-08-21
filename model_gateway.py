import logging
import os
import time

from logging.handlers import RotatingFileHandler
from pathlib import Path

import httpx
from dotenv import load_dotenv

from openai import (
    OpenAI,
    APITimeoutError,
    APIConnectionError,
    APIStatusError,
)


# =========================================================
# 韩中游戏本地化 Agent
# 统一模型调用网关
#
# 所有AI模块最终都应该通过这里调用模型接口：
#
# ai_translator.py
# strict_reviewer.py
# future terminology_suggester.py
# ...
#
#
# 统一管理：
#
# 1. API Key
# 2. Base URL
# 3. 模型名
# 4. Timeout
# 5. Retry
# 6. 日志
# 7. 错误分类
# 8. ConnectTimeout / ReadTimeout区分
# 9. HTTP状态码
# 10. 请求耗时
#
#
# 安全原则：
#
# - 不记录API Key
# - 不记录完整Prompt
# - 不记录完整游戏文本
# - 不输出Base URL
# - 不自动重试
# =========================================================


# =========================================================
# 环境变量
# =========================================================

load_dotenv()


MODEL_API_KEY = os.getenv(
    "MODEL_API_KEY",
    "",
).strip()


MODEL_BASE_URL = os.getenv(
    "MODEL_BASE_URL",
    "",
).strip()


MODEL_NAME = os.getenv(
    "MODEL_NAME",
    "",
).strip()


# =========================================================
# Timeout配置
#
# 当前保持和已经验证过的旧逻辑一致：
#
# connect = 8秒
# read    = 60秒
# write   = 60秒
# pool    = 60秒
#
# SDK自动重试 = 0
# =========================================================

CONNECT_TIMEOUT_SECONDS = 8.0
READ_TIMEOUT_SECONDS = 60.0
WRITE_TIMEOUT_SECONDS = 60.0
POOL_TIMEOUT_SECONDS = 60.0

MAX_RETRIES = 0


# =========================================================
# 日志目录
#
# 永远放在项目目录/logs/
# 而不是依赖当前PowerShell所在目录
# =========================================================

PROJECT_DIR = (
    Path(__file__)
    .resolve()
    .parent
)


LOG_DIR = (
    PROJECT_DIR
    /
    "logs"
)


LOG_DIR.mkdir(
    exist_ok=True
)


LOG_FILE = (
    LOG_DIR
    /
    "model_api.log"
)


# =========================================================
# Logger
# =========================================================

logger = logging.getLogger(
    "model_api"
)


logger.setLevel(
    logging.DEBUG
)


logger.propagate = False


# 防止Streamlit重复加载时重复添加Handler
if not logger.handlers:

    formatter = logging.Formatter(
        (
            "%(asctime)s | "
            "%(levelname)s | "
            "%(name)s | "
            "%(message)s"
        ),
        datefmt=
            "%Y-%m-%d %H:%M:%S",
    )


    # =====================================================
    # Terminal
    # =====================================================

    console_handler = (
        logging.StreamHandler()
    )


    console_handler.setLevel(
        logging.INFO
    )


    console_handler.setFormatter(
        formatter
    )


    logger.addHandler(
        console_handler
    )


    # =====================================================
    # 文件日志
    #
    # 最大5MB
    # 最多保留3个旧日志
    # =====================================================

    file_handler = (
        RotatingFileHandler(
            LOG_FILE,
            maxBytes=
                5 * 1024 * 1024,
            backupCount=
                3,
            encoding=
                "utf-8",
        )
    )


    file_handler.setLevel(
        logging.DEBUG
    )


    file_handler.setFormatter(
        formatter
    )


    logger.addHandler(
        file_handler
    )


# =========================================================
# 配置检查
# =========================================================

def validate_configuration():

    missing = []


    if not MODEL_API_KEY:

        missing.append(
            "MODEL_API_KEY"
        )


    if not MODEL_BASE_URL:

        missing.append(
            "MODEL_BASE_URL"
        )


    if not MODEL_NAME:

        missing.append(
            "MODEL_NAME"
        )


    if missing:

        raise ValueError(
            (
                "缺少Model API配置："
                +
                ", ".join(
                    missing
                )
            )
        )


    return True


# =========================================================
# 创建OpenAI兼容客户端
# =========================================================

def create_client():

    validate_configuration()


    timeout = httpx.Timeout(
        timeout=
            READ_TIMEOUT_SECONDS,

        connect=
            CONNECT_TIMEOUT_SECONDS,

        read=
            READ_TIMEOUT_SECONDS,

        write=
            WRITE_TIMEOUT_SECONDS,

        pool=
            POOL_TIMEOUT_SECONDS,
    )


    logger.debug(
        (
            "创建统一Model API客户端 | "
            "model=%s | "
            "connect_timeout=%.1fs | "
            "read_timeout=%.1fs | "
            "write_timeout=%.1fs | "
            "pool_timeout=%.1fs | "
            "max_retries=%d"
        ),
        MODEL_NAME,
        CONNECT_TIMEOUT_SECONDS,
        READ_TIMEOUT_SECONDS,
        WRITE_TIMEOUT_SECONDS,
        POOL_TIMEOUT_SECONDS,
        MAX_RETRIES,
    )


    client = OpenAI(
        api_key=
            MODEL_API_KEY,

        base_url=
            MODEL_BASE_URL,

        timeout=
            timeout,

        max_retries=
            MAX_RETRIES,
    )


    return client


# =========================================================
# 遍历异常链
#
# OpenAI SDK可能把：
#
# httpcore.ConnectTimeout
# ↓
# httpx.ConnectTimeout
# ↓
# openai.APITimeoutError
#
# 包装多层。
#
# 我们需要找到真正底层原因。
# =========================================================

def get_exception_chain(
    error
):

    result = []

    current = error

    visited = set()


    while current is not None:

        current_id = id(
            current
        )


        if current_id in visited:

            break


        visited.add(
            current_id
        )


        result.append(
            {
                "类型":
                    type(
                        current
                    ).__name__,

                "模块":
                    type(
                        current
                    ).__module__,

                "信息":
                    str(
                        current
                    ),
            }
        )


        next_error = getattr(
            current,
            "__cause__",
            None,
        )


        if next_error is None:

            next_error = getattr(
                current,
                "__context__",
                None,
            )


        current = next_error


    return result


# =========================================================
# 将异常链类型组合成文本
# =========================================================

def exception_type_names(
    error
):

    return [
        item[
            "类型"
        ]

        for item
        in get_exception_chain(
            error
        )
    ]


# =========================================================
# 精确判断超时阶段
# =========================================================

def classify_timeout(
    error
):

    type_names = (
        exception_type_names(
            error
        )
    )


    # =====================================================
    # 建立TCP / TLS连接阶段
    # =====================================================

    if "ConnectTimeout" in type_names:

        return {
            "错误类型":
                "ConnectTimeout",

            "错误阶段":
                "连接阶段",

            "用户说明":
                (
                    "无法在规定时间内与模型接口建立网络连接。"
                    "请求尚未进入模型推理阶段。"
                ),
        }


    # =====================================================
    # 已连接，但等服务器响应超时
    # =====================================================

    if "ReadTimeout" in type_names:

        return {
            "错误类型":
                "ReadTimeout",

            "错误阶段":
                "响应阶段",

            "用户说明":
                (
                    "已经建立网络连接，"
                    "但模型接口没有在规定时间内返回响应。"
                ),
        }


    # =====================================================
    # 写请求时超时
    # =====================================================

    if "WriteTimeout" in type_names:

        return {
            "错误类型":
                "WriteTimeout",

            "错误阶段":
                "请求发送阶段",

            "用户说明":
                (
                    "向模型接口发送请求数据时发生超时。"
                ),
        }


    # =====================================================
    # HTTP连接池
    # =====================================================

    if "PoolTimeout" in type_names:

        return {
            "错误类型":
                "PoolTimeout",

            "错误阶段":
                "连接池阶段",

            "用户说明":
                (
                    "等待可用HTTP连接时发生超时。"
                ),
        }


    # =====================================================
    # 无法继续细分
    # =====================================================

    return {
        "错误类型":
            "APITimeoutError",

        "错误阶段":
            "未知超时阶段",

        "用户说明":
            (
                "模型接口请求发生超时，"
                "当前无法进一步确定超时阶段。"
            ),
    }


# =========================================================
# HTTP状态错误分类
#
# 只在内存中检查错误对象，用于识别模型接口返回的
# “额度不足”等明确业务错误。
#
# 安全原则：
# - 不把HTTP响应正文写入日志
# - 不记录API Key
# - 不记录Prompt
# - 不向用户暴露网关内部request id或额度明细
# =========================================================

def classify_api_status_error(
    error
):

    status_code = getattr(
        error,
        "status_code",
        None,
    )


    # =====================================================
    # 仅用于分类的内存文本
    #
    # OpenAI兼容网关可能把错误信息放在：
    #
    # - str(error)
    # - error.body
    # - error.code
    #
    # 这些内容只用于关键词判断，不写日志。
    # =====================================================

    classification_parts = []


    try:

        classification_parts.append(
            str(
                error
            )
        )

    except Exception:

        pass


    body = getattr(
        error,
        "body",
        None,
    )


    if body is not None:

        try:

            classification_parts.append(
                str(
                    body
                )
            )

        except Exception:

            pass


    code = getattr(
        error,
        "code",
        None,
    )


    if code is not None:

        try:

            classification_parts.append(
                str(
                    code
                )
            )

        except Exception:

            pass


    classification_text = (
        " ".join(
            classification_parts
        )
        .lower()
    )


    # =====================================================
    # Model API额度不足
    #
    # 已实际遇到的公司网关错误示例关键词：
    #
    # token quota is not enough
            #
    # 此类错误虽然HTTP状态码可能是403，
    # 但真正原因不是API Key错误，而是可用额度不足。
    # =====================================================

    quota_markers = (
                "token quota is not enough",
                "insufficient_quota",
        "insufficient quota",
    )


    if any(
        marker in classification_text
        for marker in quota_markers
    ):

        return {
            "错误类型":
                "QuotaInsufficient",

            "错误阶段":
                "模型额度检查",

            "用户说明":
                (
                    "模型接口调用额度不足，"
                    "本次请求未能进入正常模型处理。"
                    "请联系管理员补充或调整调用额度后重试。"
                ),

            "状态码":
                status_code,
        }


    # =====================================================
    # 其他HTTP状态错误
    #
    # 保持原有行为，避免扩大本次修改范围。
    # =====================================================

    return {
        "错误类型":
            "APIStatusError",

        "错误阶段":
            "HTTP响应阶段",

        "用户说明":
            (
                "模型接口返回HTTP错误。"
                f"状态码：{status_code}"
            ),

        "状态码":
            status_code,
    }


# =========================================================
# 统计messages字符数
#
# 只统计长度
# 不记录实际Prompt
# =========================================================

def count_message_characters(
    messages
):

    total = 0


    for message in messages or []:

        content = message.get(
            "content",
            "",
        )


        if isinstance(
            content,
            str,
        ):

            total += len(
                content
            )


        elif content is not None:

            total += len(
                str(
                    content
                )
            )


    return total


# =========================================================
# 统一成功结果
# =========================================================

def success_result(
    content,
    elapsed,
    mode,
    request_label,
):

    return {
        "成功":
            True,

        "内容":
            content,

        "错误类型":
            "",

        "错误阶段":
            "",

        "错误信息":
            "",

        "状态码":
            None,

        "致命错误":
            False,

        "耗时秒":
            round(
                elapsed,
                3,
            ),

        "模式":
            mode,

        "请求标签":
            request_label,

        "模型":
            MODEL_NAME,
    }


# =========================================================
# 统一失败结果
# =========================================================

def error_result(
    error_type,
    error_stage,
    error_message,
    elapsed,
    mode,
    request_label,
    fatal=True,
    status_code=None,
):

    return {
        "成功":
            False,

        "内容":
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

        "状态码":
            status_code,

        "致命错误":
            bool(
                fatal
            ),

        "耗时秒":
            round(
                elapsed,
                3,
            ),

        "模式":
            mode,

        "请求标签":
            request_label,

        "模型":
            MODEL_NAME,
    }


# =========================================================
# 统一模型调用入口
#
# 以后所有业务模块都应该调用：
#
# call_chat_completion(...)
#
# 而不是自己：
#
# OpenAI(...)
# client.chat.completions.create(...)
#
# =========================================================

def call_chat_completion(
    messages,
    mode="快速翻译",
    request_label="",
    model=None,
    temperature=None,
):

    mode = str(
        mode
        or
        "未指定模式"
    ).strip()


    request_label = str(
        request_label
        or
        ""
    ).strip()


    selected_model = str(
        model
        or
        MODEL_NAME
        or
        ""
    ).strip()


    started = (
        time.perf_counter()
    )


    message_count = len(
        messages
        or
        []
    )


    prompt_chars = (
        count_message_characters(
            messages
        )
    )


    logger.info(
        (
            "Model API请求开始 | "
            "mode=%s | "
            "label=%s | "
            "model=%s | "
            "messages=%d | "
            "prompt_chars=%d"
        ),
        mode,
        request_label
        or
        "-",
        selected_model
        or
        "-",
        message_count,
        prompt_chars,
    )


    try:

        client = create_client()


        request_args = {
            "model":
                selected_model,

            "messages":
                messages,
        }


        # =================================================
        # 默认不主动传temperature
        #
        # 防止某些公司模型不支持该参数。
        # =================================================

        if temperature is not None:

            request_args[
                "temperature"
            ] = temperature


        logger.info(
            (
                "正在调用 "
                "client.chat.completions.create | "
                "mode=%s | "
                "label=%s"
            ),
            mode,
            request_label
            or
            "-",
        )


        response = (
            client
            .chat
            .completions
            .create(
                **request_args
            )
        )


        elapsed = (
            time.perf_counter()
            -
            started
        )


        # =================================================
        # 响应结构保护
        # =================================================

        if (
            response is None
            or
            not getattr(
                response,
                "choices",
                None,
            )
        ):

            logger.error(
                (
                    "Model API返回异常：无choices | "
                    "mode=%s | "
                    "label=%s | "
                    "elapsed=%.2fs"
                ),
                mode,
                request_label
                or
                "-",
                elapsed,
            )


            return error_result(
                error_type=
                    "EmptyResponse",

                error_stage=
                    "响应解析阶段",

                error_message=
                    (
                        "模型接口返回结果中没有choices。"
                    ),

                elapsed=
                    elapsed,

                mode=
                    mode,

                request_label=
                    request_label,

                fatal=
                    False,
            )


        first_choice = (
            response.choices[
                0
            ]
        )


        message = getattr(
            first_choice,
            "message",
            None,
        )


        if message is None:

            logger.error(
                (
                    "Model API返回异常：choice没有message | "
                    "mode=%s | "
                    "label=%s | "
                    "elapsed=%.2fs"
                ),
                mode,
                request_label
                or
                "-",
                elapsed,
            )


            return error_result(
                error_type=
                    "EmptyResponse",

                error_stage=
                    "响应解析阶段",

                error_message=
                    (
                        "模型接口返回结果中没有message。"
                    ),

                elapsed=
                    elapsed,

                mode=
                    mode,

                request_label=
                    request_label,

                fatal=
                    False,
            )


        content = getattr(
            message,
            "content",
            "",
        )


        content = str(
            content
            or
            ""
        ).strip()


        if not content:

            logger.error(
                (
                    "Model API返回空内容 | "
                    "mode=%s | "
                    "label=%s | "
                    "elapsed=%.2fs"
                ),
                mode,
                request_label
                or
                "-",
                elapsed,
            )


            return error_result(
                error_type=
                    "EmptyResponse",

                error_stage=
                    "响应解析阶段",

                error_message=
                    "模型接口返回了空内容。",

                elapsed=
                    elapsed,

                mode=
                    mode,

                request_label=
                    request_label,

                fatal=
                    False,
            )


        # =================================================
        # 成功
        # =================================================

        logger.info(
            (
                "Model API请求成功 | "
                "mode=%s | "
                "label=%s | "
                "elapsed=%.2fs | "
                "output_chars=%d"
            ),
            mode,
            request_label
            or
            "-",
            elapsed,
            len(
                content
            ),
        )


        return success_result(
            content=
                content,

            elapsed=
                elapsed,

            mode=
                mode,

            request_label=
                request_label,
        )


    # =====================================================
    # Timeout
    # =====================================================

    except APITimeoutError as error:

        elapsed = (
            time.perf_counter()
            -
            started
        )


        classification = (
            classify_timeout(
                error
            )
        )


        logger.error(
            (
                "Model API超时 | "
                "type=%s | "
                "stage=%s | "
                "mode=%s | "
                "label=%s | "
                "elapsed=%.2fs"
            ),
            classification[
                "错误类型"
            ],
            classification[
                "错误阶段"
            ],
            mode,
            request_label
            or
            "-",
            elapsed,
        )


        logger.debug(
            "异常链：%s",
            get_exception_chain(
                error
            ),
        )


        logger.exception(
            "Model API timeout traceback"
        )


        return error_result(
            error_type=
                classification[
                    "错误类型"
                ],

            error_stage=
                classification[
                    "错误阶段"
                ],

            error_message=
                classification[
                    "用户说明"
                ],

            elapsed=
                elapsed,

            mode=
                mode,

            request_label=
                request_label,

            fatal=
                True,
        )


    # =====================================================
    # Connection Error
    # =====================================================

    except APIConnectionError as error:

        elapsed = (
            time.perf_counter()
            -
            started
        )


        type_names = (
            exception_type_names(
                error
            )
        )


        error_type = (
            "APIConnectionError"
        )


        if "ConnectError" in type_names:

            error_type = (
                "ConnectError"
            )


        logger.error(
            (
                "Model API连接错误 | "
                "type=%s | "
                "mode=%s | "
                "label=%s | "
                "elapsed=%.2fs"
            ),
            error_type,
            mode,
            request_label
            or
            "-",
            elapsed,
        )


        logger.debug(
            "异常链：%s",
            get_exception_chain(
                error
            ),
        )


        logger.exception(
            "Model API connection traceback"
        )


        return error_result(
            error_type=
                error_type,

            error_stage=
                "连接阶段",

            error_message=
                (
                    "无法与模型接口建立正常网络连接。"
                    "请检查公司网络、VPN、代理或网关状态。"
                ),

            elapsed=
                elapsed,

            mode=
                mode,

            request_label=
                request_label,

            fatal=
                True,
        )


    # =====================================================
    # HTTP Status Error
    # =====================================================

    except APIStatusError as error:

        elapsed = (
            time.perf_counter()
            -
            started
        )


        classification = (
            classify_api_status_error(
                error
            )
        )


        status_code = (
            classification.get(
                "状态码"
            )
        )


        logger.error(
            (
                "Model API HTTP错误 | "
                "type=%s | "
                "stage=%s | "
                "status=%s | "
                "mode=%s | "
                "label=%s | "
                "elapsed=%.2fs"
            ),
            classification.get(
                "错误类型",
                "APIStatusError",
            ),
            classification.get(
                "错误阶段",
                "HTTP响应阶段",
            ),
            status_code,
            mode,
            request_label
            or
            "-",
            elapsed,
        )


        # =================================================
        # 安全日志策略
        #
        # QuotaInsufficient的原始异常正文通常包含：
        # - 剩余额度
        # - 预计额度
        # - request id
        #
        # 为避免把这些网关内部信息写入日志，
        # 已识别的额度不足错误不打印原始traceback正文。
        #
        # 其他HTTP错误继续保留原有traceback，
        # 便于排查未知问题。
        # =================================================

        if (
            classification.get(
                "错误类型"
            )
            ==
            "QuotaInsufficient"
        ):

            logger.info(
                (
                    "Model API额度不足已安全分类 | "
                    "status=%s | "
                    "mode=%s | "
                    "label=%s"
                ),
                status_code,
                mode,
                request_label
                or
                "-",
            )


        else:

            logger.exception(
                "Model API status traceback"
            )


        # =================================================
        # 不把HTTP响应正文直接写日志或返回给业务层。
        # =================================================

        return error_result(
            error_type=
                classification.get(
                    "错误类型",
                    "APIStatusError",
                ),

            error_stage=
                classification.get(
                    "错误阶段",
                    "HTTP响应阶段",
                ),

            error_message=
                classification.get(
                    "用户说明",
                    "模型接口返回HTTP错误。",
                ),

            elapsed=
                elapsed,

            mode=
                mode,

            request_label=
                request_label,

            fatal=
                True,

            status_code=
                status_code,
        )


    # =====================================================
    # 配置错误
    # =====================================================

    except ValueError as error:

        elapsed = (
            time.perf_counter()
            -
            started
        )


        logger.error(
            (
                "Model API配置错误 | "
                "mode=%s | "
                "label=%s | "
                "elapsed=%.2fs | "
                "error=%s"
            ),
            mode,
            request_label
            or
            "-",
            elapsed,
            str(
                error
            ),
        )


        logger.exception(
            "Model API configuration traceback"
        )


        return error_result(
            error_type=
                "ConfigurationError",

            error_stage=
                "配置阶段",

            error_message=
                str(
                    error
                ),

            elapsed=
                elapsed,

            mode=
                mode,

            request_label=
                request_label,

            fatal=
                True,
        )


    # =====================================================
    # 其他异常
    # =====================================================

    except Exception as error:

        elapsed = (
            time.perf_counter()
            -
            started
        )


        error_type = (
            type(
                error
            ).__name__
        )


        logger.error(
            (
                "Model API未知异常 | "
                "type=%s | "
                "mode=%s | "
                "label=%s | "
                "elapsed=%.2fs"
            ),
            error_type,
            mode,
            request_label
            or
            "-",
            elapsed,
        )


        logger.exception(
            "Unexpected Model API traceback"
        )


        return error_result(
            error_type=
                error_type,

            error_stage=
                "未知阶段",

            error_message=
                str(
                    error
                ),

            elapsed=
                elapsed,

            mode=
                mode,

            request_label=
                request_label,

            fatal=
                True,
        )


# =========================================================
# 给界面使用的非敏感配置状态
# =========================================================

def get_gateway_status():

    return {
        "API Key":
            (
                "已配置"
                if MODEL_API_KEY
                else
                "未配置"
            ),

        "Base URL":
            (
                "已配置"
                if MODEL_BASE_URL
                else
                "未配置"
            ),

        "模型":
            MODEL_NAME
            or
            "未配置",

        "连接超时":
            f"{CONNECT_TIMEOUT_SECONDS:.0f}秒",

        "响应超时":
            f"{READ_TIMEOUT_SECONDS:.0f}秒",

        "自动重试":
            MAX_RETRIES,

        "日志文件":
            str(
                LOG_FILE
            ),
    }


# =========================================================
# 独立测试
#
# 注意：
# 这里只测试模块能否加载和配置是否存在。
# 不实际请求Model API，避免浪费模型调用。
# =========================================================

if __name__ == "__main__":

    print(
        "统一模型调用网关加载成功。"
    )


    print(
        f"模型："
        f"{MODEL_NAME or '未配置'}"
    )


    print(
        "API Key："
        +
        (
            "已配置"
            if MODEL_API_KEY
            else
            "未配置"
        )
    )


    print(
        "Base URL："
        +
        (
            "已配置"
            if MODEL_BASE_URL
            else
            "未配置"
        )
    )


    print(
        f"连接超时："
        f"{CONNECT_TIMEOUT_SECONDS:.0f}秒"
    )


    print(
        f"响应超时："
        f"{READ_TIMEOUT_SECONDS:.0f}秒"
    )


    print(
        f"日志：{LOG_FILE}"
    )