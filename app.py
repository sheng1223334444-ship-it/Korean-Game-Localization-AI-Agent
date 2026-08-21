import os
from io import BytesIO
from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="韩中游戏本地化 AI Agent | Portfolio Demo",
    page_icon="🎮",
    layout="wide",
)

# ---------------------------------------------------------
# Streamlit Cloud secrets -> environment variables
# No secret is ever printed in the UI.
# ---------------------------------------------------------
for _key in ("MODEL_API_KEY", "MODEL_BASE_URL", "MODEL_NAME"):
    try:
        if _key in st.secrets:
            os.environ.setdefault(_key, str(st.secrets[_key]))
    except Exception:
        pass

from knowledge import KnowledgeBase
from agent_rules import get_priority_summary
from document_processor import translate_text_unit, BatchAIUnavailableError
from model_gateway import get_gateway_status
from xlsx_translator import (
    inspect_xlsx_for_translation,
    preflight_xlsx_translation,
    process_xlsx_translation,
)

MAX_UPLOAD_MB = 2
MAX_TRANSLATE_ROWS = 30
MAX_SINGLE_CHARS = 500

@st.cache_resource
def load_kb():
    return KnowledgeBase()

kb = load_kb()
kb_stats = kb.stats()
gateway = get_gateway_status()

st.title("🎮 韩中游戏本地化 AI Agent")
st.caption("Portfolio Demo Version · Python · Streamlit · Knowledge Base · LLM · Excel QA")

st.info(
    "这是求职作品集公开脱敏版。所有知识库与示例文本均为虚构 Demo 数据，"
    "不包含任何真实公司知识资产、内部接口地址、API Key、日志或未公开游戏文本。"
)

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Demo角色名", f"{kb_stats.get('角色名数量', 0):,}")
with c2:
    st.metric("Demo UWO记录", f"{kb_stats.get('UWO正式译文记录数', 0):,}")
with c3:
    st.metric("Demo Quest记录", f"{kb_stats.get('Quest正式译文记录数', 0):,}")
with c4:
    model_ready = (
        gateway.get("API Key") == "已配置"
        and gateway.get("Base URL") == "已配置"
        and gateway.get("模型") not in ("", "未配置", None)
    )
    st.metric("在线模型", "已配置" if model_ready else "Demo模式")

with st.expander("🔐 公开版说明与能力边界"):
    st.markdown(
        """
- 公开版只使用虚构示例知识库，展示工作流与产品能力。
- 正式业务知识库、真实公司接口、生产日志与未公开文本不包含在此仓库。
- 未配置模型接口时，**正式知识库精确命中仍可直接复用**；非命中内容不会伪装成 AI 结果。
- 在线 Demo 对上传文件大小、单条文本长度和批量处理行数做了限制，避免滥用与额外费用。
        """
    )

tabs = st.tabs(["🏠 产品概览", "🔎 单条翻译", "📄 Excel Demo", "🛡️ QA机制"])

with tabs[0]:
    st.header("从“直接调用 LLM”到“受约束的本地化工作流”")
    st.markdown(
        f"""
**固定知识优先级**

`{get_priority_summary()}`

**核心流程**

`韩文输入 → 正式知识库检索 → 完整句复用 / AI补译 → 格式保护 → 最终QA → 人工确认 → 新文件输出`

公开版重点展示四个能力：

1. **知识复用**：正式完整句优先，避免重复调用模型。
2. **受约束翻译**：角色名、术语、上下文证据先进入翻译决策。
3. **程序级QA**：数字、占位符、标签、换行、韩文残留等自动检查。
4. **Excel安全写回**：已有中文默认保留，原文件不覆盖，异常结果进入审计报告。
        """
    )

    st.success("✅ 生产版与公开版分离：公开作品展示机制，不公开业务资产。")

with tabs[1]:
    st.header("🔎 单条智能翻译")

    presets = [
        "장비를 강화하시겠습니까?",
        "성공 확률은 50%입니다.",
        "{0}골드를 획득했습니다.",
        "<color=#FF0000>위험</color>",
        "루미아에게 보고하세요.",
        "카엘, 북쪽 성문으로 가자.",
    ]

    preset = st.selectbox(
        "快速体验示例",
        ["自定义输入"] + presets,
    )

    default_text = "" if preset == "自定义输入" else preset
    korean_text = st.text_area(
        "韩文原文",
        value=default_text,
        height=120,
        max_chars=MAX_SINGLE_CHARS,
    )

    if st.button("🚀 开始翻译", type="primary", use_container_width=True):
        text = korean_text.strip()

        if not text:
            st.warning("请输入韩文。")
        else:
            with st.spinner("正在执行知识库检索、翻译与最终QA..."):
                result = translate_text_unit(
                    text=text,
                    knowledge_base=kb,
                    extra_context="公开作品集 Demo",
                )

            if result.get("处理方式") == "AI调用失败，保留韩文":
                st.warning(
                    "当前公开 Demo 未配置在线模型，且该句未命中可直接复用的 Demo 正式译文。"
                    "请选择上方示例句体验完整知识库链路，或部署时在 Streamlit Secrets 中配置模型接口。"
                )
                if result.get("错误"):
                    with st.expander("技术状态"):
                        st.write(result.get("错误类型", ""))
                        st.write(result.get("错误阶段", ""))
            else:
                st.success("翻译完成")
                st.code(result.get("译文", ""), language=None)

                a, b, c, d = st.columns(4)
                with a:
                    st.metric("处理方式", result.get("处理方式", ""))
                with b:
                    st.metric("来源", result.get("来源", "") or "AI")
                with c:
                    st.metric(
                        "最终QA",
                        "通过" if result.get("最终QA通过") is True else "需检查",
                    )
                with d:
                    st.metric(
                        "人工确认",
                        "需要" if result.get("需人工确认") else "不需要",
                    )

                if result.get("确认原因"):
                    st.warning(result.get("确认原因"))

                with st.expander("查看审计信息"):
                    safe_keys = [
                        "处理方式", "来源", "格式检查", "最终QA通过",
                        "标点自动规范化", "需人工确认", "确认原因", "AI成功", "模型",
                    ]
                    st.json({k: result.get(k) for k in safe_keys if k in result})

with tabs[2]:
    st.header("📄 Excel 批量翻译 Demo")
    st.caption(
        f"公开版限制：单文件 ≤ {MAX_UPLOAD_MB}MB；单次最多处理 {MAX_TRANSLATE_ROWS} 条待翻译韩文。"
    )

    demo_path = Path(__file__).resolve().parent / "sample_data" / "demo_input.xlsx"
    if demo_path.exists():
        st.download_button(
            "⬇️ 下载示例 Excel",
            data=demo_path.read_bytes(),
            file_name="韩中本地化Agent_Demo输入.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    upload = st.file_uploader("上传 XLSX", type=["xlsx"])

    if upload is not None:
        file_bytes = upload.getvalue()

        if len(file_bytes) > MAX_UPLOAD_MB * 1024 * 1024:
            st.error(f"公开 Demo 仅允许上传不超过 {MAX_UPLOAD_MB}MB 的文件。")
        else:
            try:
                structure = inspect_xlsx_for_translation(file_bytes)
            except Exception as e:
                st.error(f"Excel结构读取失败：{e}")
                structure = None

            if structure:
                sheet_name = st.selectbox("工作表", list(structure.keys()))
                info = structure[sheet_name]

                headers = info.get("列名", [])
                header_map = {x["列号"]: x for x in headers}
                all_cols = [x["列号"] for x in headers]

                def col_label(c):
                    item = header_map.get(c, {})
                    return f"{item.get('列字母', '')} / {item.get('列名', '')}"

                k_default = info.get("推荐韩文列")
                k_index = all_cols.index(k_default) if k_default in all_cols else 0
                korean_col = st.selectbox(
                    "韩文原文列",
                    all_cols,
                    index=k_index,
                    format_func=col_label,
                )

                chinese_options = [None] + [c for c in all_cols if c != korean_col]
                c_default = info.get("推荐中文列")
                c_index = chinese_options.index(c_default) if c_default in chinese_options else 0
                chinese_col = st.selectbox(
                    "中文目标列",
                    chinese_options,
                    index=c_index,
                    format_func=lambda c: "自动识别 / 无则新建 AI中文译文" if c is None else col_label(c),
                )

                preflight_key = (
                    upload.name, len(file_bytes), sheet_name, korean_col, chinese_col
                )

                if st.button("🔍 翻译前预检（不调用模型）", use_container_width=True):
                    bar = st.progress(0.0)
                    status = st.empty()

                    def preflight_progress(event):
                        bar.progress(float(event.get("进度", 0.0)))
                        status.caption(event.get("消息", ""))

                    stats, config = preflight_xlsx_translation(
                        file_bytes=file_bytes,
                        knowledge_base=kb,
                        sheet_name=sheet_name,
                        korean_column=korean_col,
                        chinese_column=chinese_col,
                        progress_callback=preflight_progress,
                    )

                    st.session_state["portfolio_preflight_key"] = preflight_key
                    st.session_state["portfolio_preflight_stats"] = stats
                    st.session_state["portfolio_preflight_config"] = config

                stats = st.session_state.get("portfolio_preflight_stats")
                if (
                    stats
                    and st.session_state.get("portfolio_preflight_key") == preflight_key
                ):
                    st.subheader("预检结果")
                    p1, p2, p3, p4, p5 = st.columns(5)
                    with p1:
                        st.metric("韩文文本", stats.get("韩文文本数", 0))
                    with p2:
                        st.metric("已有中文", stats.get("已有中文保留", 0))
                    with p3:
                        st.metric("待翻译", stats.get("待翻译", 0))
                    with p4:
                        st.metric("正式库预计复用", stats.get("正式完整句预计复用", 0))
                    with p5:
                        st.metric("预计模型请求", stats.get("预计AI请求", 0))

                    pending = int(stats.get("待翻译", 0) or 0)

                    if pending > MAX_TRANSLATE_ROWS:
                        st.error(
                            f"公开 Demo 单次最多处理 {MAX_TRANSLATE_ROWS} 条待翻译韩文；"
                            f"当前为 {pending} 条。请使用更小的演示文件。"
                        )
                    else:
                        if st.button(
                            "🚀 开始 Excel 智能翻译",
                            type="primary",
                            use_container_width=True,
                        ):
                            progress_bar = st.progress(0.0)
                            progress_text = st.empty()
                            live = st.empty()

                            def translate_progress(event):
                                progress_bar.progress(float(event.get("进度", 0.0)))
                                progress_text.caption(event.get("消息", ""))
                                s = event.get("统计", {}) or {}
                                live.info(
                                    "正式库复用："
                                    f"{s.get('正式完整句直接复用', 0)} ｜ "
                                    "AI成功："
                                    f"{s.get('AI翻译成功', 0)} ｜ "
                                    "人工确认："
                                    f"{s.get('需人工确认', 0)} ｜ "
                                    "QA异常："
                                    f"{s.get('格式异常', 0)}"
                                )

                            try:
                                output, result_stats, details, config = process_xlsx_translation(
                                    file_bytes=file_bytes,
                                    knowledge_base=kb,
                                    sheet_name=sheet_name,
                                    korean_column=korean_col,
                                    chinese_column=chinese_col,
                                    progress_callback=translate_progress,
                                )
                            except BatchAIUnavailableError as e:
                                st.error("模型接口异常，批量任务已安全停止，未生成半成品文件。")
                                with st.expander("安全停止信息"):
                                    st.write("停止位置：", e.location or "未知")
                                    st.write("错误类型：", e.error_type or "模型连接错误")
                            except Exception as e:
                                st.error(f"处理失败：{e}")
                            else:
                                st.success("✅ 翻译、最终QA与Excel报告生成完成")

                                r1, r2, r3, r4 = st.columns(4)
                                with r1:
                                    st.metric("新增译文", result_stats.get("新增译文", 0))
                                with r2:
                                    st.metric("正式库复用", result_stats.get("正式完整句直接复用", 0))
                                with r3:
                                    st.metric("人工确认", result_stats.get("需人工确认", 0))
                                with r4:
                                    st.metric("QA异常", result_stats.get("格式异常", 0))

                                st.download_button(
                                    "⬇️ 下载处理结果",
                                    data=output,
                                    file_name=f"{Path(upload.name).stem}_PortfolioDemo_中文翻译版.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                )

                                with st.expander("查看前20条处理明细"):
                                    st.dataframe(details[:20], use_container_width=True, hide_index=True)

with tabs[3]:
    st.header("🛡️ 为什么本地化 Agent 不能只看“翻译得像不像”")

    qa_rows = [
        ("数字 / 百分比", "100、50%、1,000 等必须保持事实一致"),
        ("占位符", "{0}、%s、%d、%1$s 不得丢失或改写"),
        ("标签", "<color>、HTML/XML 等结构必须保留"),
        ("换行", "真实换行与转义符需保持数量和位置约束"),
        ("韩文残留", "最终中文中不应残留未解释的韩文"),
        ("否定 / 条件 / 因果", "避免模型润色改变逻辑关系"),
        ("知识冲突", "历史多译法或上下文不足时进入人工确认"),
        ("Excel安全", "不覆盖原文件；已有中文默认保留；致命错误不输出半成品"),
    ]

    for name, desc in qa_rows:
        st.markdown(f"**{name}** — {desc}")

    st.divider()
    st.markdown(
        """
**Portfolio 设计原则**

公开 Demo 重点证明的是：我能够把真实业务问题抽象成 **知识检索规则、LLM Workflow、程序级 QA、异常处理和可操作界面**，
而不是把业务资产或生产接口公开出来。
        """
    )

st.divider()
st.caption("Portfolio Demo · 作者：朴谋圣 · 公开版仅用于求职作品展示")
