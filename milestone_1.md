# Goal: 本地 Excel 智能业务匹配系统 MVP

## Objective

构建一个能够在本地 MacBook Pro 运行的 Python 系统。

系统能够读取任意 Excel 文件，自动识别表头、分析字段结构、匹配业务字段，并输出标准字段映射结果。

系统必须支持后续模板学习与人工确认。

---

# Success Criteria

系统完成以下能力：

## 1. Excel解析

支持：

* xlsx
* 多Sheet

能够读取：

* 单元格内容
* 合并单元格
* 空行
* 隐藏列

输入：

test.xlsx

输出：

```python
WorkbookData
```

结构化对象。

---

## 2. 表头识别

自动识别：

* 标题行
* 表头行
* 数据开始行

例如：

输入：

销售订单导入模板

订单号 客户名称 金额

SO001 腾讯 1000

输出：

```json
{
  "header_row": 2,
  "data_start_row": 3
}
```

识别准确率：

≥ 80%

测试样例：

至少20个不同Excel模板。

---

## 3. 字段画像

生成：

```json
{
  "column_name":"客户名称",
  "data_type":"TEXT",
  "null_rate":0.01,
  "unique_rate":0.95,
  "samples":[
      "腾讯",
      "阿里"
  ]
}
```

字段画像对象。

支持：

* TEXT
* NUMBER
* DATE
* BOOLEAN

识别。

---

## 4. 字段字典系统

支持维护：

```json
{
  "customer_name":[
      "客户名称",
      "购买方",
      "企业名称"
  ]
}
```

字段别名字典。

存储于 SQLite。

---

## 5. 模糊匹配

支持：

RapidFuzz 匹配。

例如：

购买方

匹配：

customer_name

返回：

```json
{
  "field":"customer_name",
  "score":92
}
```

---

## 6. Embedding匹配

使用：

BAAI/bge-small-zh-v1.5

实现字段语义匹配。

例如：

购买方

匹配：

customer_name

返回：

```json
{
  "field":"customer_name",
  "similarity":0.93
}
```

---

## 7. 向量索引

使用：

FAISS

存储：

* 标准字段
* 字段别名

支持：

TopK查询。

---

## 8. 映射融合引擎

融合：

* 精确匹配
* 字典匹配
* RapidFuzz匹配
* Embedding匹配

输出：

```json
{
  "excel_field":"购买方",
  "target_field":"customer_name",
  "confidence":0.94
}
```

---

## 9. 人工确认界面

使用：

Streamlit

支持：

* 上传Excel
* 查看表头识别结果
* 查看字段映射结果
* 手工修改映射
* 保存模板

---

## 10. 模板学习

用户确认：

购买方 → customer_name

保存模板。

后续同结构Excel：

自动复用。

---

# Technical Stack

Backend

* FastAPI

Excel

* openpyxl
* pandas

Matching

* RapidFuzz

Embedding

* sentence-transformers
* BAAI/bge-small-zh-v1.5

Vector Search

* FAISS

Database

* SQLite

UI

* Streamlit

---

# Project Structure

excel_matcher/

api/

excel/

* parser.py
* header_detector.py
* profiler.py

matcher/

* dictionary_matcher.py
* fuzzy_matcher.py
* embedding_matcher.py
* fusion_matcher.py

vector/

* faiss_store.py

storage/

* sqlite_store.py

ui/

* streamlit_app.py

models/

tests/

---

# Constraints

必须本地运行。

禁止使用：

* OpenAI API
* Claude API
* Gemini API
* 云向量数据库
* 云Embedding服务

禁止依赖：

* Milvus
* Qdrant
* Weaviate

必须支持：

MacBook Pro

最低配置：

* Apple Silicon
* 16GB RAM

---

# Verification

完成后必须满足：

## Unit Test

执行：

```bash
pytest
```

通过率：

100%

---

## Functional Test

测试以下Excel：

* 订单模板
* CRM模板
* 库存模板
* 财务模板

至少20份样例。

---

## Accuracy

表头识别：

≥ 80%

字段匹配：

≥ 85%

---

# Stop Conditions

以下条件全部满足后视为完成：

* Excel解析完成
* 表头识别完成
* 字段画像完成
* RapidFuzz匹配完成
* Embedding匹配完成
* FAISS检索完成
* Streamlit界面完成
* SQLite模板学习完成
* pytest全部通过
* 测试样例全部通过

否则持续迭代修复。

如果出现无法解决的问题：

输出：

* 问题描述
* 根因分析
* 当前阻塞点
* 建议方案

然后停止。
