# 韩中游戏本地化 AI Agent — Portfolio Demo

这是一个用于求职作品展示的公开脱敏版本。

## 展示能力

- 韩语 → 简体中文游戏本地化工作流
- 角色名 / UWO / Quest 三层 Demo 知识库
- 正式完整句优先复用
- LLM 缺口补译（可选模型接口）
- 数字、占位符、标签、换行、韩文残留等程序级 QA
- Excel 翻译前预检
- Excel 实时进度
- 已有中文默认保留
- 新文件输出与审计报告
- 致命模型接口错误安全停止

## 隐私与保密

本仓库仅包含虚构 Demo 数据，不包含：

- 真实公司知识库
- 真实公司 API Key / Base URL
- 生产日志
- 未公开游戏文本
- 内部反馈报告
- 生产环境配置

## 本地运行

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

不配置模型接口也可以体验知识库精确命中与 Excel 正式译文复用。

## 可选：配置在线模型

复制 `.streamlit/secrets.toml.example` 为 `.streamlit/secrets.toml`，填写你自己的 OpenAI-compatible 模型接口：

```toml
MODEL_API_KEY = "..."
MODEL_BASE_URL = "https://api.example.com/v1"
MODEL_NAME = "..."
```

**不要把 `secrets.toml` 提交到 GitHub。**

## Streamlit Community Cloud

1. 将本项目上传到 GitHub。
2. 在 Streamlit Community Cloud 创建 App。
3. Main file path 选择 `app.py`。
4. 如需在线模型，在 App 的 Secrets 设置中添加 `MODEL_API_KEY / MODEL_BASE_URL / MODEL_NAME`。
5. 部署完成后，把公开 URL 放进简历和作品集。

## 作者

朴谋圣 — 求职作品集公开版
