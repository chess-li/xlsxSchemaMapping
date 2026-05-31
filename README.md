# Excel Matcher MVP

本项目是一个本地运行的 Excel 智能业务字段匹配系统 MVP。它用于读取业务 Excel 文件，自动识别表头、生成字段画像，并将 Excel 中的字段匹配到一组标准业务字段上，支持人工确认和模板保存，便于后续复用同结构文件的映射结果。

## 项目背景

在订单、CRM、库存、财务等业务系统之间流转数据时，Excel 模板经常存在字段命名不一致的问题。例如同一个客户字段可能被写成“客户名称”、“购买方”、“企业名称”或“公司名称”。如果每次都人工识别和映射字段，效率低且容易出错。

本项目希望提供一个可以在本地 MacBook Pro 运行的轻量化方案：

- 自动解析 `.xlsx` 文件和多 Sheet 工作簿。
- 自动识别标题行、表头行和数据开始行。
- 根据列数据生成字段画像。
- 通过规则、字典、模糊匹配和可选语义匹配生成标准字段映射。
- 通过 Streamlit 提供人工确认界面。
- 将确认后的映射保存为模板，支持后续复用。

## 核心能力

- **Excel 解析**：基于 `openpyxl` 读取单元格、合并单元格、空行、隐藏列和多 Sheet 信息。
- **表头识别**：识别标题行、表头行、数据起始行，并输出置信度和判断依据。
- **字段画像**：识别字段名、列号、数据类型、空值率、唯一值率和样例值。
- **标准字段字典**：内置订单、CRM、库存、财务等领域的标准字段和别名。
- **字段匹配流水线**：按 `Exact -> Dictionary -> RapidFuzz -> Embedding/FAISS -> LLM` 的顺序执行匹配。
- **人工确认**：通过 Streamlit 查看候选结果、调整映射并保存模板。
- **API 服务**：通过 FastAPI 暴露健康检查、字段字典、Excel 分析、字段匹配和模板保存接口。

## 技术栈

- Python 3.11+
- FastAPI
- Streamlit
- openpyxl
- pandas
- Pydantic v2
- RapidFuzz
- SQLite
- pytest
- 可选语义依赖：sentence-transformers、FAISS

## 快速开始

创建虚拟环境并安装基础依赖：

```bash
python3.11 -m venv .venv
.venv/bin/python -m pip install -e ".[test]"
```

如果需要启用 Embedding/FAISS 语义匹配，安装额外依赖：

```bash
.venv/bin/python -m pip install -e ".[test,semantic]"
```

## 运行 API 服务

启动 FastAPI：

```bash
.venv/bin/uvicorn excel_matcher.api.app:app --reload
```

默认访问地址：

```text
http://127.0.0.1:8000
```

接口文档：

```text
http://127.0.0.1:8000/docs
```

## 运行 Streamlit 界面

启动本地 Web 界面：

```bash
.venv/bin/streamlit run excel_matcher/ui/streamlit_app.py
```

使用流程：

1. 打开 Streamlit 页面。
2. 上传 `.xlsx` 文件。
3. 查看工作簿、Sheet、表头识别和字段画像结果。
4. 查看每个 Excel 字段的匹配候选。
5. 手工选择或修正目标标准字段。
6. 点击保存模板，将确认后的映射写入本地 SQLite。

## API 接口概览

### `GET /health`

检查系统组件状态，包括 SQLite、Embedding 和 LLM 阶段的可用性说明。

### `GET /fields`

返回内置标准字段字典。

### `POST /workbooks/analyze`

上传 Excel 文件并返回解析结果、表头识别结果和字段画像。

### `POST /fields/match`

根据字段画像返回标准字段匹配结果。

请求体中的关键字段：

```json
{
  "profiles": [],
  "enable_embedding": false,
  "enable_llm": false
}
```

### `POST /templates`

保存人工确认后的模板映射。

请求体示例：

```json
{
  "signature": "sheet-structure-signature",
  "sheet_name": "销售订单",
  "mappings": {
    "购买方": "customer_name",
    "订单金额": "amount"
  }
}
```

## 语义匹配配置

基础安装默认不启用语义匹配。匹配流水线中 Embedding 和 LLM 阶段默认会被跳过，并在结果中记录跳过原因。

安装语义依赖后，系统使用：

```text
BAAI/bge-small-zh-v1.5
```

本地配置默认值位于 `excel_matcher/config.py`：

```text
SQLite: data/excel_matcher.db
FAISS index: data/faiss/index.faiss
FAISS metadata: data/faiss/metadata.json
Embedding model: BAAI/bge-small-zh-v1.5
```

FAISS 索引是可重建的本地缓存，SQLite 是标准字段、别名和模板数据的主要存储。

## 项目结构

```text
excel_matcher/
  api/          FastAPI 服务入口
  excel/        Excel 解析、表头识别、字段画像
  matcher/      精确匹配、字典匹配、模糊匹配、语义匹配和融合
  services/     工作簿分析、字段匹配、模板签名等服务编排
  storage/      SQLite schema、默认字段和存储实现
  ui/           Streamlit 界面
  vector/       FAISS 向量索引封装
tests/          单元测试和样例生成
docs/           设计文档和执行计划
```

## 测试

运行全部测试：

```bash
.venv/bin/python -m pytest
```

测试覆盖范围包括：

- Excel 解析
- 表头识别
- 字段画像
- 存储层
- 字段匹配器
- 服务层
- API
- 样例准确率
- 语义匹配阶段行为

## 数据与本地文件

运行过程中会在本地生成或使用以下数据：

- `data/excel_matcher.db`：SQLite 数据库。
- `data/faiss/index.faiss`：FAISS 向量索引。
- `data/faiss/metadata.json`：向量索引元数据。

这些文件属于本地运行数据，不是核心源码。语义模型首次使用时可能需要下载依赖模型，后续可使用本地缓存。

## 当前状态

当前项目处于 MVP 阶段，已经包含 Excel 分析、字段匹配、API、Streamlit 界面、SQLite 模板保存和可选语义匹配相关模块。后续可以继续增强标准字段维护、模板复用策略、语义匹配稳定性和更多真实业务样例验证。
