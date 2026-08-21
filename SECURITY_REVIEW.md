# Public Release Security Review

此公开版已按以下原则处理：

- 未包含 `.env`
- 未包含原 `knowledge/` 业务资产
- 未包含 `logs/`
- 未包含 `.venv/`
- 未包含生产模型接口测试脚本
- 未发现硬编码 API Key、JWT、邮箱、IPv4 地址或真实 Base URL
- 模型配置统一改为公开通用变量：`MODEL_API_KEY / MODEL_BASE_URL / MODEL_NAME`
- 公开知识库为本次生成的虚构 Demo 数据
- 示例 Excel 为虚构 Demo 数据
- UI 明确标注“Portfolio Demo Version”

部署前仍建议在 GitHub 页面再次检查提交文件，并确认 `.streamlit/secrets.toml` 未被提交。
