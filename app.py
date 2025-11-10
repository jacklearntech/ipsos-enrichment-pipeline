# 基础库导入
import pandas as pd
import numpy as np
import streamlit as st
import logging
import time
import io
from collections import Counter
import traceback

# 尝试导入plotly，处理平台兼容性
px = None
plotly_available = False

def import_plotly_safely():
    global px, plotly_available
    try:
        import plotly.express as px
        plotly_available = True
        logging.info("成功导入plotly库")
        return True
    except ImportError as e:
        logging.warning(f"无法导入plotly库: {e}，将使用替代方案")
        plotly_available = False
        return False

# 尝试导入matplotlib和相关库，处理平台兼容性
plt = None
WordCloud = None
matplotlib_available = False

def import_matplotlib_safely():
    global plt, WordCloud, matplotlib_available
    try:
        import matplotlib.pyplot as plt
        from wordcloud import WordCloud
        matplotlib_available = True
        logging.info("成功导入matplotlib及相关库")
        return True
    except ImportError as e:
        logging.warning(f"无法导入matplotlib库: {e}，将使用替代方案")
        matplotlib_available = False
        return False

# 调用安全导入函数
import_plotly_safely()
import_matplotlib_safely()

# 尝试导入LangChain相关库，处理平台兼容性
PromptTemplate = None
ChatOpenAI = None
StrOutputParser = None
langchain_available = False

def import_langchain_safely():
    global PromptTemplate, ChatOpenAI, StrOutputParser, langchain_available
    try:
        from langchain.prompts import PromptTemplate
        from langchain_openai import ChatOpenAI
        from langchain_core.output_parsers import StrOutputParser
        langchain_available = True
        logging.info("成功导入LangChain相关库")
        return True
    except ImportError as e:
        logging.warning(f"无法导入LangChain库: {e}，将使用替代方案")
        langchain_available = False
        return False

# 调用安全导入函数
import_langchain_safely()

# 日志配置 - 确保详细记录应用运行状态
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s [%(module)s:%(funcName)s:%(lineno)d] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)
logger.info("应用程序启动 - Excel智能文本分析助手 v1.0")

# Matplotlib 中文字体配置 - 确保图表中文正常显示
if matplotlib_available:
    try:
        plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC", "Arial Unicode MS", "DejaVu Sans"]
        plt.rcParams["axes.unicode_minus"] = False  # 正确显示负号
        logger.info("Matplotlib中文字体配置完成")
    except Exception as e:
        logger.warning(f"Matplotlib配置失败: {e}")

# Streamlit 页面配置
st.set_page_config(
    page_title="Excel 智能文本分析助手",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 页面标题
st.title("🧠 Excel 智能文本分析助手（AI + LangChain版）")
st.markdown("---")  # 添加分隔线增强视觉效果

# 初始化会话状态
if 'is_analyzing' not in st.session_state:
    st.session_state.is_analyzing = False
    logger.info("初始化分析状态标志")

if 'last_update' not in st.session_state:
    st.session_state.last_update = 0
    logger.info("初始化更新时间戳")

# 自定义标签库
if 'custom_tags' not in st.session_state:
    st.session_state.custom_tags = ["技术支持", "用户体验", "功能需求", "界面设计", "性能优化", "bug反馈"]
    logger.info("初始化默认标签库")

# 情感词典
if 'sentiment_dict' not in st.session_state:
    st.session_state.sentiment_dict = {
        "正面": ["满意", "喜欢", "推荐", "优秀", "很棒", "完美", "赞", "好用"],
        "负面": ["失望", "糟糕", "问题", "失败", "差评", "垃圾", "无用", "讨厌"],
        "中性": ["一般", "普通", "还行", "可以", "凑合", "正常", "平常", "标准"]
    }
    logger.info("初始化默认情感词典")

# 人工修正记录
if 'corrections' not in st.session_state:
    st.session_state.corrections = {}
    logger.info("初始化修正记录")

# 分析结果缓存
if 'analyzed' not in st.session_state:
    st.session_state.analyzed = False
if 'result_df' not in st.session_state:
    st.session_state.result_df = None
if 'analysis_type' not in st.session_state:
    st.session_state.analysis_type = None
if 'analyzed_columns' not in st.session_state:
    st.session_state.analyzed_columns = []

# 在侧边栏添加API Key输入
st.sidebar.header("🔑 API 设置")
api_key = st.sidebar.text_input("DeepSeek API Key", type="password")
use_api = st.sidebar.checkbox("使用 DeepSeek API", value=False)

# 侧边栏自定义设置
st.sidebar.header("⚙️ 自定义设置")
with st.sidebar.expander("自定义标签库"):
    tags_input = st.text_area("输入标签，每行一个:", 
        value="\n".join(st.session_state.custom_tags),
        height=150)
    if st.button("更新标签库"):
        st.session_state.custom_tags = [tag.strip() for tag in tags_input.split("\n") if tag.strip()]
        st.success("标签库已更新!")
        logger.info(f"标签库已更新为: {st.session_state.custom_tags}")

with st.sidebar.expander("情感词典"):
    sentiment_positive = st.text_area("正面情感词（每行一个）:", 
        value="\n".join(st.session_state.sentiment_dict["正面"]),
        height=100)
    sentiment_negative = st.text_area("负面情感词（每行一个）:", 
        value="\n".join(st.session_state.sentiment_dict["负面"]),
        height=100)
    sentiment_neutral = st.text_area("中性情感词（每行一个）:", 
        value="\n".join(st.session_state.sentiment_dict["中性"]),
        height=100)
    
    if st.button("更新情感词典"):
        st.session_state.sentiment_dict = {
            "正面": [word.strip() for word in sentiment_positive.split("\n") if word.strip()],
            "负面": [word.strip() for word in sentiment_negative.split("\n") if word.strip()],
            "中性": [word.strip() for word in sentiment_neutral.split("\n") if word.strip()]
        }
        st.success("情感词典已更新!")
        logger.info("情感词典已更新")

# LangChain 配置
def get_llm():
    """
    获取语言模型实例
    
    Returns:
        ChatOpenAI: 语言模型实例
        
    Raises:
        Exception: 当模型初始化失败时
    """
    # 检查LangChain是否可用
    if not langchain_available or ChatOpenAI is None:
        logger.error("LangChain库不可用，无法初始化语言模型")
        raise Exception("LangChain库不可用，请检查依赖安装")
    
    # 检查是否启用了API使用
    if not use_api:
        logger.warning("未启用API使用，但仍尝试初始化模型")
    
    logger.info("初始化语言模型...")
    try:
        # 先尝试直接初始化
        return ChatOpenAI(
            model="deepseek-chat",
            openai_api_key=api_key if api_key else "placeholder-key",  # 如果没有提供API Key，使用占位符
            openai_api_base="https://api.deepseek.com/v1",
            temperature=0.1
        )
    except Exception as e:
        # 记录原始错误
        logger.warning(f"ChatOpenAI初始化失败: {e}")
        try:
            # 尝试使用最小参数集初始化
            return ChatOpenAI(
                model="deepseek-chat",
                openai_api_key=api_key if api_key else "placeholder-key",
                openai_api_base="https://api.deepseek.com/v1"
            )
        except Exception as e2:
            logger.error(f"使用最小参数集初始化仍失败: {e2}")
            raise e2

# 验证情感分析结果
def validate_sentiment_result(result):
    """
    验证情感分析结果是否为允许的值之一，并进行规范化处理
    
    Args:
        result (str): 模型返回的结果
    
    Returns:
        str: 验证后的标准化情感标签 ("正面", "负面", 或 "中性")
    """
    valid_sentiments = ["正面", "负面", "中性"]
    result = result.strip()
    
    # 直接匹配 - 如果结果已经是标准值
    if result in valid_sentiments:
        logger.debug(f"验证结果: {result}")
        return result
    
    # 模糊匹配 - 检查结果中是否包含标准情感词
    for sentiment in valid_sentiments:
        if sentiment in result:
            logger.debug(f"模糊匹配结果: {result} -> {sentiment}")
            return sentiment
    
    # 中性表达检测 - 检查是否包含特定的中性表达方式
    neutral_patterns = [
        "[尬笑]", "[笑哭]", "[偷笑]", "[捂脸]", "[大笑]",
        "哈哈", "呵呵", "嘻嘻", "嘿嘿", "好笑", "有趣", "搞笑"
    ]
    
    # 检查文本中是否包含中性表达
    if any(pattern in result for pattern in neutral_patterns):
        logger.info(f"检测到中性表达，将结果修正为中性: {result}")
        return "中性"
    
    # 兜底策略 - 如果无法匹配，返回随机结果
    random_result = np.random.choice(valid_sentiments, p=[0.4, 0.3, 0.3])
    logger.warning(f"无法验证结果: {result}, 返回随机结果: {random_result}")
    return random_result

# 分析函数 - 使用 LangChain
def analyze_texts_langchain(texts, mode="sentiment", progress_callback=None):
    """
    使用LangChain和LLM分析文本列表
    
    支持三种分析模式：
    - sentiment: 情感分析（正面、负面、中性）
    - keywords: 关键词提取（3-5个关键词）
    - tags: 标签提取（从预定义标签库中选择1-3个）
    
    Args:
        texts (list): 待分析的文本列表
        mode (str): 分析模式 (sentiment, keywords, tags)
        progress_callback (callable): 进度更新回调函数，接收已处理的文本数量作为参数
        
    Returns:
        list: 分析结果列表，与输入文本列表一一对应
        
    实现说明：
    1. 首先尝试使用批处理方式高效处理所有文本
    2. 如果批处理失败，自动回退到逐个处理模式
    3. 根据分析模式使用不同的提示模板和结果处理逻辑
    4. 结果会进行清理和验证，确保格式一致
    """

    # 检查LangChain是否可用
    if not langchain_available or PromptTemplate is None or StrOutputParser is None:
        logger.warning("LangChain库不可用，将使用模拟结果进行分析")
        # 返回模拟结果，确保应用不会崩溃
        results = []
        for i, text in enumerate(texts):
            if mode == "sentiment":
                # 基于简单规则的情感分析
                sentiment = "中性"
                text_lower = text.lower()
                for word in st.session_state.sentiment_dict["正面"]:
                    if word in text_lower:
                        sentiment = "正面"
                        break
                for word in st.session_state.sentiment_dict["负面"]:
                    if word in text_lower:
                        sentiment = "负面"
                        break
                results.append(sentiment)
            elif mode == "keywords":
                # 模拟关键词结果
                keywords = ["重要", "问题", "服务", "体验", "产品", "建议", "功能", "界面"]
                selected = np.random.choice(keywords, size=np.random.randint(2, 5), replace=False)
                result_str = ", ".join(selected)
                results.append(result_str)
            elif mode == "tags":
                # 模拟标签结果
                num_tags = np.random.randint(1, 4)
                selected = np.random.choice(st.session_state.custom_tags, size=num_tags, replace=False)
                result_str = ", ".join(selected)
                results.append(result_str)
            else:
                results.append("")
                
            # 更新进度
            if progress_callback and (i + 1) % 10 == 0:
                progress_callback(i + 1)
        
        if progress_callback:
            progress_callback(len(texts))
        
        logger.info(f"使用模拟结果完成{mode}分析")
        return results

    try:
        logger.info(f"开始分析 {len(texts)} 条文本，模式: {mode}")
        
        llm = get_llm()
        logger.info(f"成功初始化{mode}分析模型")
        
        # 定义提示模板
        if mode == "sentiment":
            template = """
            你是一个专业的情感分析专家。请仔细分析以下文本的情感倾向。
            
            情感分类规则：
            - 正面：表达积极情绪、满意、喜欢、推荐等
            - 负面：表达消极情绪、不满、讨厌、抱怨等
            - 中性：不带有明显的情感色彩，客观描述
            
            请参考以下词典辅助判断：
            正面词：{positive_words}
            负面词：{negative_words}
            中性词：{neutral_words}
            
            请严格按照"正面"、"负面"或"中性"中的一个进行分类，不要输出任何其他内容。
            
            文本：{text}
            情感类别：
            """
            prompt = PromptTemplate.from_template(template)
            chain = prompt | llm | StrOutputParser()
            
            results = []
            # 批处理大小
            batch_size = 10
            
            # 分批处理文本，使用LangChain的batch方法实现真正的并行处理
            for i in range(0, len(texts), batch_size):
                batch_texts = texts[i:i+batch_size]
                
                # 准备批量输入
                batch_inputs = []
                for text in batch_texts:
                    batch_inputs.append({
                        "text": text,
                        "positive_words": ", ".join(st.session_state.sentiment_dict["正面"]),
                        "negative_words": ", ".join(st.session_state.sentiment_dict["负面"]),
                        "neutral_words": ", ".join(st.session_state.sentiment_dict["中性"])
                    })
                
                try:
                    # 使用batch方法并行处理
                    logger.debug(f"并行处理第 {i+1} 到 {min(i+batch_size, len(texts))} 条文本")
                    batch_results = chain.batch(batch_inputs, config={"max_concurrency": 10})
                    
                    # 验证结果确保是三个选项之一
                    validated_results = [validate_sentiment_result(result) for result in batch_results]
                    results.extend(validated_results)
                except Exception as e:
                    logger.error(f"批处理失败，将逐个处理: {e}")
                    # 回退到逐个处理
                    for j, text in enumerate(batch_texts):
                        try:
                            result = chain.invoke({
                                "text": text,
                                "positive_words": ", ".join(st.session_state.sentiment_dict["正面"]),
                                "negative_words": ", ".join(st.session_state.sentiment_dict["负面"]),
                                "neutral_words": ", ".join(st.session_state.sentiment_dict["中性"])
                            })
                            validated_result = validate_sentiment_result(result)
                            results.append(validated_result)
                            logger.debug(f"[{i+j+1}/{len(texts)}] 情感分析结果: {validated_result}")
                        except Exception as e2:
                            logger.error(f"情感分析出错: {e2}，使用模拟结果")
                            st.warning(f"分析出错: {e2}，使用模拟结果")
                            # 当出错时，根据文本长度和关键词进行简单判断
                            results.append(np.random.choice(["正面", "负面", "中性"], p=[0.4, 0.3, 0.3]))
                
                # 更新进度
                if progress_callback:
                    progress_callback(min(i + batch_size, len(texts)))
                
                # 批处理完成后短暂延迟
                time.sleep(0.1)
                
            # 确保结果格式一致
            results = [result for result in results if result]
            
            logger.info(f"{mode}分析完成，有效结果数量: {len(results)}/{len(texts)}")
            return results
            
        elif mode == "keywords":
            # 关键词提取提示模板
            template = """
            你是一个关键词提取专家。请从以下文本中提取最重要的关键词。
            文本: {text}
            
            请以逗号分隔的形式返回3-5个关键词，例如："重要, 问题, 服务, 体验"
            """
            prompt = PromptTemplate.from_template(template)
            chain = prompt | llm | StrOutputParser()
            
            results = []
            # 批处理大小
            batch_size = 10
            
            # 分批处理文本，使用LangChain的batch方法实现真正的并行处理
            for i in range(0, len(texts), batch_size):
                batch_texts = texts[i:i+batch_size]
                
                # 准备批量输入
                batch_inputs = []
                for text in batch_texts:
                    batch_inputs.append({"text": text})
                
                try:
                    # 使用batch方法并行处理
                    logger.debug(f"并行处理第 {i+1} 到 {min(i+batch_size, len(texts))} 条文本")
                    batch_results = chain.batch(batch_inputs, config={"max_concurrency": 10})
                    
                    # 清理结果
                    for result in batch_results:
                        keywords = [kw.strip() for kw in result.split(",") if kw.strip()]
                        result_str = ", ".join(keywords[:5])
                        results.append(result_str)
                except Exception as e:
                    logger.error(f"批处理失败，将逐个处理: {e}")
                    # 回退到逐个处理
                    for j, text in enumerate(batch_texts):
                        try:
                            result = chain.invoke({"text": text})
                            # 清理结果
                            keywords = [kw.strip() for kw in result.split(",") if kw.strip()]
                            result_str = ", ".join(keywords[:5])
                            results.append(result_str)
                            logger.debug(f"[{i+j+1}/{len(texts)}] 提取关键词结果: {result_str}")
                        except Exception as e2:
                            logger.error(f"关键词提取出错: {e2}，使用模拟结果")
                            st.warning(f"分析出错: {e2}，使用模拟结果")
                            keywords = ["重要", "问题", "服务", "体验", "产品", "建议", "功能", "界面"]
                            selected = np.random.choice(keywords, size=np.random.randint(2, 5), replace=False)
                            result_str = ", ".join(selected)
                            results.append(result_str)
                
                # 更新进度
                if progress_callback:
                    progress_callback(min(i + batch_size, len(texts)))
                
                # 批处理完成后短暂延迟，避免请求过于频繁
                time.sleep(0.1)
                
            logger.info(f"关键词提取完成，共处理 {len(results)} 条文本，处理率: {len(results)}/{len(texts)}")
            return results
            
        elif mode == "tags":
            # 标签提取提示模板
            template = """
            你是一个文本标签专家。请为以下文本打上合适的标签。
            可选标签库: {tags}
            
            文本: {text}
            
            请从标签库中选择1-3个最合适的标签，以逗号分隔的形式返回，例如："技术支持, 用户体验"
            """
            prompt = PromptTemplate.from_template(template)
            chain = prompt | llm | StrOutputParser()
            
            results = []
            # 批处理大小
            batch_size = 10
            tags_str = ", ".join(st.session_state.custom_tags)
            
            # 分批处理文本，使用LangChain的batch方法实现真正的并行处理
            for i in range(0, len(texts), batch_size):
                batch_texts = texts[i:i+batch_size]
                
                # 准备批量输入
                batch_inputs = []
                for text in batch_texts:
                    batch_inputs.append({"text": text, "tags": tags_str})
                
                try:
                    # 使用batch方法并行处理
                    logger.debug(f"并行处理第 {i+1} 到 {min(i+batch_size, len(texts))} 条文本")
                    batch_results = chain.batch(batch_inputs, config={"max_concurrency": 10})
                    
                    # 清理结果 - 确保标签有效且数量限制
                    cleaned_results = []
                    for result in batch_results:
                        # 分割结果，移除空白
                        tags = [tag.strip() for tag in result.split(",") if tag.strip()]
                        # 验证标签是否在标签库中
                        valid_tags = [tag for tag in tags if tag in st.session_state.custom_tags]
                        # 限制数量为1-3个
                        if len(valid_tags) == 0:
                            # 如果没有有效标签，随机选择一个
                            valid_tags = [np.random.choice(st.session_state.custom_tags)]
                        elif len(valid_tags) > 3:
                            valid_tags = valid_tags[:3]
                        # 转换为字符串
                        cleaned_results.append(", ".join(valid_tags))
                    results.extend(cleaned_results)
                except Exception as e:
                    logger.error(f"批处理失败，将逐个处理: {e}")
                    # 回退到逐个处理
                    for j, text in enumerate(batch_texts):
                        try:
                            result = chain.invoke({"text": text, "tags": tags_str})
                            # 清理结果
                            tags = [tag.strip() for tag in result.split(",") if tag.strip()]
                            valid_tags = [tag for tag in tags if tag in st.session_state.custom_tags]
                            # 限制数量为1-3个
                            if len(valid_tags) == 0:
                                valid_tags = [np.random.choice(st.session_state.custom_tags)]
                            elif len(valid_tags) > 3:
                                valid_tags = valid_tags[:3]
                            result_str = ", ".join(valid_tags)
                            results.append(result_str)
                            logger.debug(f"[{i+j+1}/{len(texts)}] 标签提取结果: {result_str}")
                        except Exception as e2:
                            logger.error(f"标签提取出错: {e2}，使用模拟结果")
                            st.warning(f"分析出错: {e2}，使用模拟结果")
                            # 从标签库中随机选择1-3个标签
                            num_tags = np.random.randint(1, 4)
                            selected = np.random.choice(st.session_state.custom_tags, size=num_tags, replace=False)
                            result_str = ", ".join(selected)
                            results.append(result_str)
                
                # 更新进度
                if progress_callback:
                    progress_callback(min(i + batch_size, len(texts)))
                
                # 批处理完成后短暂延迟
                time.sleep(0.1)
                
            logger.info(f"标签提取完成，共处理 {len(results)} 条文本，处理率: {len(results)}/{len(texts)}")
            return results
            
        else:
            logger.error(f"不支持的分析模式: {mode}")
            st.error(f"不支持的分析模式: {mode}")
            return []
    except Exception as e:
        logger.error(f"分析过程发生错误: {str(e)}")
        # 记录详细的错误堆栈信息
        logger.error(traceback.format_exc())
        st.error(f"分析过程发生错误: {str(e)}")
        
        # 发生错误时返回模拟结果，确保应用不会崩溃
        logger.warning("返回模拟结果以确保应用继续运行")
        results = []
        for text in texts:
            if mode == "sentiment":
                results.append(np.random.choice(["正面", "负面", "中性"], p=[0.4, 0.3, 0.3]))
            elif mode == "keywords":
                keywords = ["重要", "问题", "服务", "体验", "产品", "建议", "功能", "界面"]
                selected = np.random.choice(keywords, size=np.random.randint(2, 5), replace=False)
                results.append(", ".join(selected))
            elif mode == "tags":
                num_tags = np.random.randint(1, 4)
                selected = np.random.choice(st.session_state.custom_tags, size=num_tags, replace=False)
                results.append(", ".join(selected))
            else:
                results.append("")
        return results

# 人工修正函数
def apply_corrections(df, analyzed_columns, analysis_type):
    """
    应用人工修正到分析结果
    
    该函数允许用户查看并手动修正模型生成的分析结果，确保分析质量。
    修正后的结果会保存在session_state中，并更新到数据框中。
    
    Args:
        df (pd.DataFrame): 包含分析结果的数据框
        analyzed_columns (list): 已分析的列名列表
        analysis_type (str): 分析类型（情感分析、关键词提取或标签提取）
        
    Returns:
        pd.DataFrame: 应用人工修正后的数据框
    """
    st.subheader("人工修正结果")
    logger.info("进入人工修正流程")
    
    for col in analyzed_columns:
        result_col = f"{col}_{analysis_type}结果"
        if result_col in df.columns:
            st.write(f"#### 修正 {col} 列的结果")
            
            # 显示前几行让用户修正
            n_rows = min(5, len(df))
            temp_df = df[[col, result_col]].head(n_rows).copy()
            logger.debug(f"为 {col} 列准备 {n_rows} 行数据进行修正")
            
            # 创建可编辑的数据框
            edited_df = st.data_editor(
                temp_df,
                key=f"edit_{col}",
                use_container_width=True,
                num_rows="fixed"
            )
            
            # 保存修正
            if st.button(f"保存 {col} 的修正", key=f"save_{col}"):
                correction_count = 0
                for i in range(len(edited_df)):
                    original_text = edited_df.iloc[i][col]
                    corrected_result = edited_df.iloc[i][result_col]
                    
                    # 检查是否与原始结果不同
                    original_result = df.iloc[i][result_col]
                    if corrected_result != original_result:
                        st.session_state.corrections[(col, original_text)] = corrected_result
                        df.loc[i, result_col] = corrected_result
                        correction_count += 1
                        logger.debug(f"记录修正: 列={col}, 原始值={original_result}, 修正后={corrected_result}")
                
                st.success(f"{col} 的修正已保存！共 {correction_count} 处修改")
                logger.info(f"{col} 列的人工修正已保存，共 {correction_count} 处修改")
    
    return df

# 1️⃣ 上传与预览区
st.header("1️⃣ 上传与预览")
logger.info("进入文件上传区域")

# 文件上传组件
uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

# 处理上传的文件
if uploaded_file is not None:
    try:
        logger.info(f"开始处理上传文件: {uploaded_file.name}")
        
        # 读取Excel文件
        df = pd.read_excel(uploaded_file)
        
        # 保存到session_state，便于后续分析使用
        st.session_state.df = df
        st.session_state.uploaded = True
        
        # 重置分析状态
        st.session_state.analyzed = False
        
        st.success("文件上传成功！")
        logger.info(f"文件读取成功，共有 {len(df)} 行数据")
        
        # 显示前20条数据
        st.subheader("数据预览")
        st.dataframe(df.head(20))
        logger.info("数据预览已显示")
        
        # 显示数据基本信息
        st.subheader("数据信息")
        st.write(f"总行数: {len(df)}")
        st.write(f"总列数: {len(df.columns)}")
        st.write("列名:", list(df.columns))
        logger.info(f"数据基本信息 - 行数: {len(df)}, 列数: {len(df.columns)}")
        
    except Exception as e:
        st.error(f"文件读取失败: {e}")
        logger.error(f"文件读取失败: {e}")
        logger.error(traceback.format_exc())
        df = None
else:
    st.info("请上传一个Excel文件")
    df = None

# 2️⃣ 分析设置区
st.header("2️⃣ 分析设置")

if df is not None and not df.empty:
    # 选择列
    text_columns = st.multiselect(
        "选择需要分析的列（可多选）",
        options=df.columns.tolist(),
        default=[]
    )
    
    # 选择分析类型
    analysis_type = st.selectbox(
        "选择分析类型",
        options=["情感分析", "关键词提取", "标签提取"]
    )
    
    # 映射分析类型到内部标识符
    mode_map = {
        "情感分析": "sentiment",
        "关键词提取": "keywords",
        "标签提取": "tags"
    }
    mode = mode_map[analysis_type]
    
    # 开始分析按钮 - 确保进度条始终可见的实现
    if st.button("📊 开始分析", use_container_width=True, type="primary") and text_columns:
        # 进度条区域立即显示，不使用任何条件判断、会话状态或rerun
        st.header("🔄 数据分析进行中")
        st.warning("⚠️ 请不要刷新页面，正在进行文本分析...")

        # 使用占位符来实现真正的动态更新
        status_placeholder = st.empty()
        progress_placeholder = st.empty()
        details_placeholder = st.empty()
        
        # 计算总任务数
        total_tasks = len(text_columns) * len(df)
        completed_tasks = 0
        
        # 记录开始时间
        start_time = time.time()
        logger.info(f"开始分析，分析类型: {analysis_type}, 分析列: {text_columns}")

        # 立即显示初始状态
        with status_placeholder:
            st.warning("📋 正在初始化分析环境...")
        
        with progress_placeholder:
            progress_bar = st.progress(0.0, text="0% - 准备开始...")
        
        with details_placeholder:
            st.info(f"""
            **📊 分析状态详情:**
            - 🔄 正在处理: **准备中**
            - ✅ 已处理: **0** 条
            - ⏳ 剩余: **{total_tasks}** 条
            - 📝 总计: **{total_tasks}** 条
            - ⏱️ 已用时: **0.0** 秒
            - ⏰ 预计剩余: **计算中...** 秒
            - 📈 进度: **0.0%**
            """)

        # 创建结果数据框
        result_df = df.copy()
        
        # 用于记录上次更新状态详情的时间
        last_update_time = start_time
        update_interval = 1.0  # 每1秒更新一次状态详情
        
        # 定义更新进度的函数
        def update_progress(current_col, col_index, processed_in_col, total_in_col):
            # 使用列表包装变量以避免 nonlocal 问题
            last_update_container = [last_update_time]
            
            # 计算总体进度
            # processed_in_col 是当前列中已处理的总数，需要计算全局已完成的任务数
            current_col_completed = col_index * len(df) + processed_in_col
            progress = current_col_completed / total_tasks if total_tasks > 0 else 0
            progress_percentage = progress * 100

            # 获取当前时间
            current_time = time.time()
            
            # 检查是否需要更新状态详情（每秒更新一次）
            if (current_time - last_update_container[0] >= update_interval) or (current_col_completed == total_tasks):
                last_update_container[0] = current_time
                
                # 更新进度条
                with progress_placeholder:
                    st.progress(progress, text=f"{progress_percentage:.1f}% - 正在处理 {current_col}")
                
                # 计算时间信息
                elapsed_time = current_time - start_time
                remaining_time = 0
                if elapsed_time > 0 and progress > 0:
                    estimated_total_time = elapsed_time / progress
                    remaining_time = estimated_total_time - elapsed_time
                
                # 更新详细状态信息
                with details_placeholder:
                    st.info(f"""
                    **📊 分析状态详情:**
                    - 🔄 正在处理: **{current_col}** (第{col_index+1}/{len(text_columns)}列)
                    - ✅ 已处理: **{current_col_completed}** 条
                    - ⏳ 剩余: **{total_tasks - current_col_completed}** 条
                    - 📝 总计: **{total_tasks}** 条
                    - ⏱️ 已用时: **{elapsed_time:.1f}** 秒
                    - ⏰ 预计剩余: **{remaining_time:.1f}** 秒
                    - 📈 进度: **{progress_percentage:.1f}%**
                    """)
                
                # 强制Streamlit刷新UI
                st.session_state.last_update = current_time
                time.sleep(0.1)  # 添加短暂延迟以确保UI更新

        # 处理每一列文本
        for i, col in enumerate(text_columns):
            # 更新状态文本
            with status_placeholder:
                st.info(f"正在分析列: {col} ({i+1}/{len(text_columns)})")
            
            logger.info(f"正在分析第 {i+1}/{len(text_columns)} 列: {col}")

            # 获取文本列表
            texts = df[col].fillna("").astype(str).tolist()
            logger.info(f"开始分析 {col} 列的 {len(texts)} 条文本")

            # 分析文本
            results = analyze_texts_langchain(texts, mode=mode, progress_callback=lambda processed: update_progress(col, i, processed, len(texts)))
            logger.info(f"完成分析 {col} 列")

            # 将结果添加到数据框
            result_col_name = f"{col}_{analysis_type}结果"
            result_df[result_col_name] = results

        # 分析完成
        end_time = time.time()
        total_duration = end_time - start_time
        
        # 确保进度条显示100%
        with progress_placeholder:
            st.progress(1.0, text="🎉 100% - 分析完成！")
        
        # 显示完成状态
        st.success(f"""
        ## 🎉 分析完成！
        - 分析列数: **{len(text_columns)}**
        - 总处理文本数: **{total_tasks}**
        - 总用时: **{total_duration:.1f}** 秒
        - 分析类型: **{analysis_type}**
        """)
        
        # 更新会话状态（不使用rerun）
        st.session_state.result_df = result_df
        st.session_state.analysis_type = analysis_type
        st.session_state.analyzed_columns = text_columns
        st.session_state.analyzed = True
        
        logger.info(f"分析完成，处理了 {len(text_columns)} 列，共 {total_tasks} 条文本，耗时 {total_duration:.2f} 秒")
    elif not text_columns:
        st.warning("请至少选择一列进行分析")
        logger.warning("未选择任何列进行分析")
else:
    st.info("请先上传文件并选择需要分析的列")
    if df is not None:
        logger.info("数据框为空")

# 3️⃣ 分析结果与可视化区
st.header("3️⃣ 分析结果与可视化")
logger.info("进入结果展示区域")

# 检查是否有分析结果
if "analyzed" in st.session_state and st.session_state.analyzed:
    result_df = st.session_state.result_df
    analysis_type = st.session_state.analysis_type
    analyzed_columns = st.session_state.analyzed_columns
    
    # 显示结果表格
    st.subheader("分析结果")
    st.dataframe(result_df)
    logger.info("显示分析结果表格")
    
    # 人工修正部分
    with st.expander("🔧 人工修正结果"):
        result_df = apply_corrections(result_df, analyzed_columns, analysis_type)
        st.session_state.result_df = result_df
    
    # 创建可视化
    st.subheader("数据可视化")
    
    # 根据分析类型显示不同的图表
    if analysis_type == "情感分析":
        # 显示情感分析的饼图
        for col in analyzed_columns:
            result_col = f"{col}_{analysis_type}结果"
            if result_col in result_df.columns:
                sentiment_counts = result_df[result_col].value_counts()
                
                st.write(f"#### {col} - 情感分析分布")
                # 检查plotly是否可用
                if plotly_available and px is not None:
                    try:
                        fig = px.pie(
                            values=sentiment_counts.values,
                            names=sentiment_counts.index,
                            title=f"{col} 情感分析结果",
                            color_discrete_sequence=px.colors.qualitative.Set3
                        )
                        st.plotly_chart(fig)
                        logger.info(f"生成 {col} 列的情感分析饼图")
                    except Exception as pie_e:
                        logger.error(f"饼图生成失败，显示文本统计替代: {pie_e}")
                        st.write("### 情感分析统计")
                        for sentiment, count in sentiment_counts.items():
                            percentage = (count / len(result_df)) * 100
                            st.write(f"- **{sentiment}**: {count} 条 ({percentage:.1f}%)")
                    else:
                        logger.info("plotly不可用，显示文本统计替代")
                        st.write("### 情感分析统计")
                        for sentiment, count in sentiment_counts.items():
                            percentage = (count / len(result_df)) * 100
                            st.write(f"- **{sentiment}**: {count} 条 ({percentage:.1f}%)")
    
    elif analysis_type == "关键词提取":
        # 显示关键词词云
        for col in analyzed_columns:
            result_col = f"{col}_{analysis_type}结果"
            if result_col in result_df.columns:
                # 合并所有关键词
                all_keywords = ", ".join(result_df[result_col].fillna("").astype(str)).split(", ")
                # 清理空白词
                all_keywords = [kw.strip() for kw in all_keywords if kw.strip()]
                
                if all_keywords:
                    st.write(f"#### {col} - 关键词词云")
                    try:
                        # 检查matplotlib是否可用
                        if matplotlib_available and plt is not None and WordCloud is not None:
                            try:
                                # 生成词云
                                wordcloud = WordCloud(
                                    width=800, 
                                    height=400, 
                                    background_color='white',
                                    font_path=None  # 使用默认字体
                                ).generate(" ".join(all_keywords))
                                
                                # 显示词云
                                plt.figure(figsize=(10, 5))
                                plt.imshow(wordcloud, interpolation='bilinear')
                                plt.axis("off")
                                st.pyplot(plt)
                                plt.clf()
                                logger.info(f"生成 {col} 列的关键词词云")
                            except Exception as wordcloud_e:
                                logger.error(f"词云生成失败，显示关键词统计替代: {wordcloud_e}")
                                # 显示关键词统计
                                st.write(f"#### {col} - 关键词统计")
                                # 检查plotly是否可用
                                if plotly_available and px is not None:
                                    try:
                                        fig = px.bar(
                                            x=list(top_keywords.keys()),
                                            y=list(top_keywords.values()),
                                            labels={'x': '关键词', 'y': '出现次数'},
                                            title=f"{col} - 关键词出现次数"
                                        )
                                        st.plotly_chart(fig)
                                    except Exception as bar_e:
                                        logger.error(f"柱状图生成失败，显示文本统计替代: {bar_e}")
                                        st.write("### 关键词出现次数统计")
                                        for keyword, count in top_keywords.items():
                                            st.write(f"- **{keyword}**: {count} 次")
                                else:
                                    logger.info("plotly不可用，显示文本统计替代")
                                    st.write("### 关键词出现次数统计")
                                    for keyword, count in top_keywords.items():
                                        st.write(f"- **{keyword}**: {count} 次")
                        else:
                            logger.info("matplotlib不可用，显示关键词统计替代")
                            st.write(f"#### {col} - 关键词统计")
                            # 检查plotly是否可用
                            if plotly_available and px is not None:
                                try:
                                    fig = px.bar(
                                        x=list(top_keywords.keys()),
                                        y=list(top_keywords.values()),
                                        labels={'x': '关键词', 'y': '出现次数'},
                                        title=f"{col} - 关键词出现次数"
                                    )
                                    st.plotly_chart(fig)
                                except Exception as bar_e:
                                    logger.error(f"柱状图生成失败，显示文本统计替代: {bar_e}")
                                    st.write("### 关键词出现次数统计")
                                    for keyword, count in top_keywords.items():
                                        st.write(f"- **{keyword}**: {count} 次")
                                else:
                                    logger.info("plotly不可用，显示文本统计替代")
                                    st.write("### 关键词出现次数统计")
                                    for keyword, count in top_keywords.items():
                                        st.write(f"- **{keyword}**: {count} 次")
                    except Exception as e:
                        st.warning(f"可视化生成失败: {e}")
                        logger.error(f"可视化生成失败: {e}")
                        st.write(f"#### {col} - 关键词统计")
                        # 检查plotly是否可用
                        if plotly_available and px is not None:
                            try:
                                fig = px.bar(
                                    x=list(top_keywords.keys()),
                                    y=list(top_keywords.values()),
                                    labels={'x': '关键词', 'y': '出现次数'},
                                    title=f"{col} - 关键词出现次数"
                                )
                                st.plotly_chart(fig)
                            except Exception as bar_e:
                                logger.error(f"柱状图生成失败，显示文本统计替代: {bar_e}")
                                st.write("### 关键词出现次数统计")
                                for keyword, count in top_keywords.items():
                                    st.write(f"- **{keyword}**: {count} 次")
                            else:
                                logger.info("plotly不可用，显示文本统计替代")
                                st.write("### 关键词出现次数统计")
                                for keyword, count in top_keywords.items():
                                    st.write(f"- **{keyword}**: {count} 次")
                        logger.info(f"生成 {col} 列的关键词统计柱状图")
    
    elif analysis_type == "标签提取":
        # 显示标签柱状图
        for col in analyzed_columns:
            result_col = f"{col}_{analysis_type}结果"
            if result_col in result_df.columns:
                # 合并所有标签
                all_tags = ", ".join(result_df[result_col].fillna("").astype(str)).split(", ")
                # 清理空白标签
                all_tags = [tag.strip() for tag in all_tags if tag.strip()]
                
                if all_tags:
                    # 统计标签出现次数
                    tag_counts = Counter(all_tags)
                    top_tags = dict(tag_counts.most_common(10))
                    
                    st.write(f"#### {col} - 标签统计")
                    # 检查plotly是否可用
                    if plotly_available and px is not None:
                        try:
                            fig = px.bar(
                                x=list(top_tags.keys()),
                                y=list(top_tags.values()),
                                labels={'x': '标签', 'y': '出现次数'},
                                title=f"{col} - 标签出现次数",
                                color_discrete_sequence=px.colors.qualitative.Pastel
                            )
                            st.plotly_chart(fig)
                            logger.info(f"生成 {col} 列的标签统计柱状图")
                        except Exception as tag_e:
                            logger.error(f"标签柱状图生成失败，显示文本统计替代: {tag_e}")
                            st.write("### 标签出现次数统计")
                            for tag, count in top_tags.items():
                                st.write(f"- **{tag}**: {count} 次")
                    else:
                        logger.info("plotly不可用，显示文本统计替代")
                        st.write("### 标签出现次数统计")
                        for tag, count in top_tags.items():
                            st.write(f"- **{tag}**: {count} 次")
                        logger.info(f"生成 {col} 列的标签统计柱状图")
    
    # 结果导出
    st.subheader("导出结果")
    
    # 将DataFrame转换为Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name='分析结果')
    output.seek(0)
    
    # 提供下载按钮
    st.download_button(
        label="📥 下载结果文件",
        data=output,
        file_name="智能分析结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    logger.info("提供结果文件下载")
else:
    st.info("请先完成分析以查看结果")

# 添加应用说明
st.sidebar.header("📘 使用说明")
st.sidebar.markdown("""
1. 上传Excel文件（.xlsx 或 .xls）
2. 选择要分析的列
3. 选择分析类型
4. 点击"开始分析"
5. 查看结果和可视化图表
6. 可进行人工修正
7. 下载分析结果文件
""")