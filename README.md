# Excel 智能文本分析助手

这是一个基于 AI 和 LangChain 的 Excel 文本分析工具，可以帮助用户对 Excel 文件中的文本数据进行智能分析。

## 功能特点

- 情感分析：自动识别文本中的正面、负面或中性情感
- 关键词提取：从文本中提取最重要的关键词
- 标签分类：根据预定义标签库对文本进行分类
- 可视化展示：通过图表直观展示分析结果
- 人工修正：支持对分析结果进行人工校正
- 结果导出：将分析结果导出为 Excel 文件

## 技术栈

- Python
- Streamlit（Web界面）
- LangChain（AI集成）
- DeepSeek API（大语言模型）
- Pandas（数据处理）
- Plotly/Matplotlib（数据可视化）

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行项目

```bash
streamlit run app.py
```

运行后，项目将在本地 8501 端口提供服务，您可以在浏览器中访问 http://localhost:8501 查看应用。

## 使用说明

1. 上传 Excel 文件（.xlsx 或 .xls 格式）
2. 选择需要分析的列
3. 选择分析类型（情感分析、关键词提取或标签提取）
4. 点击"开始分析"按钮
5. 查看分析结果和可视化图表
6. 可进行人工修正
7. 下载分析结果文件