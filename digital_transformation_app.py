import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import matplotlib.font_manager as fm
import re
import folium
from streamlit_folium import st_folium
import numpy as np
import seaborn as sns

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# 设置白色主题
st.set_page_config(
    page_title='企业数字化转型指数查询系统',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='expanded'
)

# 加载Excel数据
def load_data():
    try:
        # 检查文件是否存在
        file_path = '两版合并后的年报数据_完整版.xlsx'
        if not os.path.exists(file_path):
            st.error(f"文件不存在: {file_path}")
            st.write("当前工作目录:", os.getcwd())
            st.write("当前目录下的文件:", os.listdir('.'))
            return None
        
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        return df
    except Exception as e:
        st.error(f"加载数据失败: {e}")
        st.write("当前工作目录:", os.getcwd())
        st.write("当前目录下的文件:", os.listdir('.'))
        return None

# 加载数据
df = load_data()

if df is not None:
    # 设置词汇分类体系
    VOCABULARY_CLASSIFICATION = {
        '人工智能': [
            '人工智能', 'AI', '机器学习', '深度学习', '神经网络', '自然语言处理', 
            '计算机视觉', '图像理解', '语音识别', '智能决策', '算法模型', 
            '知识图谱', '人机交互', '智能客服', '自动化决策'
        ],
        '大数据': [
            '大数据', '数据挖掘', '数据分析', '数据处理', '数据治理', 
            '数据仓库', '数据湖', '数据中台', '数据可视化', '预测分析', 
            '实时数据', '数据集成', '数据资产', '数据安全'
        ],
        '云计算': [
            '云计算', '云服务', '云平台', '云计算平台', 'IaaS', 'PaaS', 
            'SaaS', '云存储', '云原生', '容器化', '微服务', '弹性计算', 
            '分布式计算', '混合云', '边缘计算'
        ],
        '区块链': [
            '区块链', '分布式账本', '智能合约', '加密货币', '去中心化', 
            '共识机制', '哈希算法', '不可篡改', '数字资产', '区块链技术'
        ],
        '数字技术应用': [
            '数字化转型', '数字经济', '数字金融', '数字营销', '数字制造', 
            '工业互联网', '智能制造', '物联网', 'IoT', '数字孪生', 
            '投资决策系统', '供应链金融', '智慧物流', '智能工厂', 
            '工业4.0', '数字化生产', '智能供应链', '数字管理', '智能运营'
        ]
    }
    
    # 词频统计函数
    def count_word_frequency(text, classification):
        if pd.isna(text):
            return {category: 0 for category in classification.keys()}
        
        text = str(text).lower()
        frequency = {category: 0 for category in classification.keys()}
        
        for category, keywords in classification.items():
            for keyword in keywords:
                # 使用正则表达式进行精确匹配，避免子字符串匹配
                frequency[category] += len(re.findall(r'\b' + re.escape(keyword.lower()) + r'\b', text))
        
        return frequency
    
    # 检查必要列是否存在
    required_columns = ['股票代码', '年份', '企业名称']
    index_columns = [col for col in df.columns if '数字化' in col or '转型' in col or '指数' in col]
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        st.error(f"缺少必要的列: {', '.join(missing_columns)}")
        st.stop()
    
    if not index_columns:
        st.warning("未找到包含'数字化'、'转型'或'指数'的列，请检查数据")
        st.stop()
    
    # 获取唯一的股票代码和年份
    stock_codes = df['股票代码'].unique().tolist()
    years = df['年份'].unique().tolist()
    
    # 排序
    years.sort()
    
    # 侧边栏查询
    with st.sidebar:
        st.title('查询面板')
        st.write('请选择以下参数进行查询')
        
        selected_stock = st.selectbox('股票代码', stock_codes)
        selected_year = st.selectbox('年份', years)
        
        # 查询按钮
        search_button = st.button('查询', key='search_button', help='点击查询数据')
    
    # 主页面内容
    st.title('企业数字化转型指数查询系统')
    
    # 计算统计指标
    index_col = index_columns[0]
    avg_index = df[index_col].mean() if index_col in df.columns else 0
    max_index = df[index_col].max() if index_col in df.columns else 0
    min_index = df[index_col].min() if index_col in df.columns else 0
    median_index = df[index_col].median() if index_col in df.columns else 0
    std_index = df[index_col].std() if index_col in df.columns else 0
    
    # 显示统计概览
    st.subheader('统计概览')
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        st.metric(label="总记录数", value=len(df))
    with col2:
        st.metric(label="企业数量", value=len(df['股票代码'].unique()))
    with col3:
        st.metric(label="年份范围", value=f"{min(years)}-{max(years)}")
    with col4:
        st.metric(label="平均指数", value=f"{avg_index:.2f}")
    with col5:
        st.metric(label="最高指数", value=f"{max_index:.2f}")
    with col6:
        st.metric(label="最低指数", value=f"{min_index:.2f}")
    
    # 显示更多统计信息
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="中位数指数", value=f"{median_index:.2f}")
    with col2:
        st.metric(label="指数标准差", value=f"{std_index:.2f}")
    with col3:
        st.metric(label="数据年份数", value=len(years))
    
    # 数据概览
    st.subheader('数据概览')
    col1, col2 = st.columns([2, 1])
    with col1:
        st.dataframe(df.sample(10))
    with col2:
        st.write("**数据结构**")
        st.write(f"行数: {df.shape[0]}")
        st.write(f"列数: {df.shape[1]}")
        st.write(f"\n**主要列名**")
        st.write("\n".join(df.columns[:10]))
        if len(df.columns) > 10:
            st.write(f"... 等 {len(df.columns)} 列")
    
    # 维度相关性热力图
    st.subheader('维度相关性热力图')
    # 获取数值型列
    numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
    
    if len(numeric_columns) > 1:
        # 计算相关系数矩阵
        corr_matrix = df[numeric_columns].corr()
        
        # 创建热力图
        plt.figure(figsize=(12, 8))
        sns.heatmap(
            corr_matrix, 
            annot=True, 
            cmap='coolwarm', 
            fmt='.2f', 
            linewidths=0.5,
            cbar_kws={'shrink': 0.8}
        )
        plt.title('维度相关性热力图')
        plt.tight_layout()
        st.pyplot(plt)
    else:
        st.info("数据中数值型列不足，无法生成相关性热力图")
    
    # 数字化转型指数分布
    st.subheader('数字化转型指数分布')
    index_col = index_columns[0]
    
    if index_col in df.columns:
        # 直方图和密度图
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
        
        # 直方图
        ax1.hist(df[index_col], bins=20, alpha=0.7, color='#1f77b4')
        ax1.set_title(f'{index_col}分布直方图')
        ax1.set_xlabel(index_col)
        ax1.set_ylabel('企业数量')
        ax1.grid(True, alpha=0.3)
        
        # 密度图
        sns.kdeplot(df[index_col], ax=ax2, fill=True, color='#ff7f0e', alpha=0.7)
        ax2.set_title(f'{index_col}分布密度图')
        ax2.set_xlabel(index_col)
        ax2.set_ylabel('密度')
        ax2.grid(True, alpha=0.3)
        
        plt.tight_layout()
        st.pyplot(fig)
        
        # 数字化转型指数详细统计
        st.subheader('数字化转型指数详细统计')
        index_stats = {
            '平均值': df[index_col].mean(),
            '中位数': df[index_col].median(),
            '标准差': df[index_col].std(),
            '最小值': df[index_col].min(),
            '最大值': df[index_col].max(),
            '25%分位数': df[index_col].quantile(0.25),
            '75%分位数': df[index_col].quantile(0.75)
        }
        
        col1, col2, col3 = st.columns(3)
        for i, (stat_name, value) in enumerate(index_stats.items()):
            with [col1, col2, col3][i % 3]:
                st.info(f"**{stat_name}**\n{value:.4f}")
    else:
        st.info(f"未找到{index_col}列，无法生成指数分布")
    
    # 地理分布地图（优化）
    st.subheader('企业地理分布')
    # 检查是否有地区相关列
    region_columns = [col for col in df.columns if any(keyword in col for keyword in ['地区', '省份', '城市', '地域'])]
    
    if region_columns:
        region_col = region_columns[0]
        
        # 统计各地区企业数量和平均指数
        region_stats = df.groupby(region_col).agg({
            '股票代码': 'nunique',
            index_col: ['mean', 'min', 'max', 'count']
        }).reset_index()
        
        # 重命名列
        region_stats.columns = [region_col, '企业数量', '平均指数', '最低指数', '最高指数', '数据条数']
        
        # 显示地区分布统计
        st.write(f"基于 {region_col} 列的企业分布和指数统计")
        st.dataframe(region_stats)
        
        # 创建地图
        st.subheader('企业分布和指数地图')
        # 这里使用folium创建中国地图
        map_china = folium.Map(location=[35.8617, 104.1954], zoom_start=4, tiles='CartoDB positron')
        
        # 添加企业标记（这里需要实际的经纬度数据，暂时使用示例位置）
        # 为了演示，我们使用随机经纬度
        for _, row in region_stats.iterrows():
            # 这里应该使用实际的经纬度数据
            # 由于数据中可能没有经纬度，我们使用随机位置作为示例
            lat = 20 + np.random.rand() * 30  # 20-50°N
            lon = 70 + np.random.rand() * 60  # 70-130°E
            
            # 创建详细的弹出信息
            popup_content = f"""
            <div style='width: 200px;'>
                <h4>{row[region_col]}</h4>
                <p>企业数量: <strong>{row['企业数量']}</strong></p>
                <p>平均指数: <strong>{row['平均指数']:.2f}</strong></p>
                <p>最低指数: <strong>{row['最低指数']:.2f}</strong></p>
                <p>最高指数: <strong>{row['最高指数']:.2f}</strong></p>
                <p>数据条数: <strong>{row['数据条数']}</strong></p>
            </div>
            """
            
            # 使用不同颜色表示指数高低
            if row['平均指数'] > avg_index + std_index:
                color = 'green'
            elif row['平均指数'] > avg_index:
                color = 'blue'
            elif row['平均指数'] > avg_index - std_index:
                color = 'orange'
            else:
                color = 'red'
            
            # 添加圆形标记，大小表示企业数量
            folium.CircleMarker(
                location=[lat, lon],
                radius=max(5, row['企业数量'] * 0.5),  # 企业数量越多，标记越大
                popup=folium.Popup(popup_content, max_width=300),
                tooltip=f"{row[region_col]}: {row['企业数量']}家企业",
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=0.6
            ).add_to(map_china)
        
        # 添加图层控制
        folium.LayerControl().add_to(map_china)
        
        # 在Streamlit中显示地图
        st_folium(map_china, width=1000, height=600)
        
        # 数字化转型指数热力分布
        st.subheader('数字化转型指数热力分布')
        
        # 创建热力图数据
        heatmap_data = []
        
        # 为每个地区生成多个点，密度与企业数量相关
        for _, row in region_stats.iterrows():
            # 生成经纬度（示例数据）
            base_lat = 20 + np.random.rand() * 30  # 20-50°N
            base_lon = 70 + np.random.rand() * 60  # 70-130°E
            
            # 根据企业数量生成多个点
            num_points = min(row['企业数量'], 10)  # 限制最大点数为10
            
            for _ in range(num_points):
                # 添加一些随机偏移
                lat = base_lat + (np.random.rand() - 0.5) * 2
                lon = base_lon + (np.random.rand() - 0.5) * 2
                
                # 数字越大，热力越强
                heatmap_data.append([lat, lon, row['平均指数']])
        
        # 创建热力图
        heatmap_map = folium.Map(location=[35.8617, 104.1954], zoom_start=4, tiles='CartoDB positron')
        
        # 添加热力图层
        from folium.plugins import HeatMap
        
        HeatMap(
            heatmap_data,
            min_opacity=0.3,
            max_zoom=10,
            radius=15,
            blur=10,
            max_val=max(df[index_col]) if index_col in df.columns else 100,
            gradient={0.4: 'blue', 0.65: 'lime', 0.8: 'yellow', 1: 'red'},
            overlay=True,
            control=True,
            name='数字化转型指数热力图'
        ).add_to(heatmap_map)
        
        # 添加图层控制
        folium.LayerControl().add_to(heatmap_map)
        
        # 在Streamlit中显示热力图
        st.write("热力图说明：颜色越红表示数字化转型指数越高，颜色越蓝表示指数越低")
        st_folium(heatmap_map, width=1000, height=600)
    else:
        st.info("数据中未找到地区相关列，无法生成地理分布地图和热力分布")
        st.write("建议在数据中添加'地区'、'省份'或'城市'列以启用此功能")
    
    # 查询结果
    if search_button or True:  # 默认显示所有数据
        st.markdown('---')
        st.header('查询结果')
        
        # 按股票代码过滤
        stock_data = df[df['股票代码'] == selected_stock]
        
        # 显示该股票的基本信息
        if not stock_data.empty:
            # 获取企业名称
            company_name = stock_data['企业名称'].iloc[0]
            
            # 公司信息卡片
            with st.container():
                st.subheader('公司基本信息')
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.info(f"**企业名称**\n{company_name}")
                with col2:
                    st.info(f"**股票代码**\n{selected_stock}")
                with col3:
                    st.info(f"**数据年份**\n{', '.join(map(str, sorted(stock_data['年份'].unique())))}")
            
            # 获取指定年份的数据
            year_data = stock_data[stock_data['年份'] == selected_year]
            if not year_data.empty:
                # 使用动态检测到的索引列
                index_col = index_columns[0]
                if index_col in year_data.columns:
                    index_value = year_data[index_col].iloc[0]
                    
                    # 指数展示卡片
                    with st.container():
                        st.subheader(f'{selected_year}年数字化转型指数')
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric(
                                label=f"{selected_year}年{index_col}", 
                                value=f"{index_value:.2f}" if isinstance(index_value, (int, float)) else index_value,
                                delta=None
                            )
                else:
                    st.warning(f"未找到{index_col}列")
            else:
                st.warning(f"未找到{selected_stock}在{selected_year}年的数据")
        else:
            st.warning(f"未找到股票代码{selected_stock}的数据")
        
        # 可视化部分
        st.markdown('---')
        st.header('数据可视化')
        
        # 使用动态检测到的索引列
        index_col = index_columns[0]
        
        if index_col in stock_data.columns:
            # 按年份排序
            stock_data_sorted = stock_data.sort_values('年份')
            
            # 折线图和柱状图并排显示
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader('历年指数折线图')
                plt.figure(figsize=(10, 6))
                plt.plot(stock_data_sorted['年份'], stock_data_sorted[index_col], marker='o', linestyle='-', color='#1f77b4')
                plt.title(f'{company_name}({selected_stock})历年{index_col}')
                plt.xlabel('年份')
                plt.ylabel(index_col)
                plt.grid(True, alpha=0.3)
                plt.tight_layout()
                st.pyplot(plt)
            
            with col2:
                st.subheader('历年指数柱状图')
                plt.figure(figsize=(10, 6))
                bars = plt.bar(stock_data_sorted['年份'], stock_data_sorted[index_col], color='#ff7f0e', alpha=0.8)
                plt.title(f'{company_name}({selected_stock})历年{index_col}')
                plt.xlabel('年份')
                plt.ylabel(index_col)
                plt.grid(True, alpha=0.3, axis='y')
                
                # 在柱状图上显示数值
                for bar in bars:
                    height = bar.get_height()
                    plt.text(bar.get_x() + bar.get_width()/2., height + 0.01 * max(stock_data_sorted[index_col]),
                            f'{height:.2f}', ha='center', va='bottom')
                
                plt.tight_layout()
                st.pyplot(plt)
        else:
            st.warning(f"未找到{index_col}列，无法生成趋势图")
        
        # 词频统计
        st.markdown('---')
        st.header('数字技术词频分析')
        
        # 检查是否有文本内容列
        text_columns = [col for col in df.columns if any(keyword in col for keyword in ['内容', '年报', '描述', '文本'])]
        
        if text_columns:
            text_col = text_columns[0]
            st.write(f"基于 {text_col} 列的词频统计")
            
            # 检查该股票是否有文本数据
            stock_text_data = stock_data[stock_data[text_col].notna()]
            
            if not stock_text_data.empty:
                # 计算该股票的总词频
                total_frequency = {category: 0 for category in VOCABULARY_CLASSIFICATION.keys()}
                
                for _, row in stock_text_data.iterrows():
                    text = row[text_col]
                    frequency = count_word_frequency(text, VOCABULARY_CLASSIFICATION)
                    for category, count in frequency.items():
                        total_frequency[category] += count
                
                # 词频柱状图
                st.subheader('数字技术词汇分布')
                plt.figure(figsize=(12, 6))
                categories = list(total_frequency.keys())
                counts = list(total_frequency.values())
                bars = plt.bar(categories, counts, color=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd'])
                plt.title(f'{company_name}数字技术词汇使用分布')
                plt.xlabel('技术类别')
                plt.ylabel('词汇出现次数')
                plt.grid(True, alpha=0.3, axis='y')
                
                # 在柱状图上显示数值
                for bar in bars:
                    height = bar.get_height()
                    plt.text(bar.get_x() + bar.get_width()/2., height + 0.5, f'{height}', ha='center', va='bottom')
                
                plt.tight_layout()
                st.pyplot(plt)
                
                # 词频表格
                st.subheader('词频统计详情')
                frequency_df = pd.DataFrame(list(total_frequency.items()), columns=['技术类别', '词频数'])
                st.dataframe(frequency_df)
            else:
                st.warning(f"未找到{company_name}的文本数据")
        else:
            st.warning("未找到包含文本内容的列，请检查数据")
        
        # 数据表格
        st.markdown('---')
        st.header('详细数据')
        st.dataframe(stock_data)
        
        # 提供下载功能
        st.markdown('---')
        st.header('数据下载')
        csv = stock_data.to_csv(index=False)
        st.download_button(
            label="下载当前股票数据 (CSV)",
            data=csv,
            file_name=f"{company_name}_{selected_stock}_数字化转型数据.csv",
            mime="text/csv"
        )

else:
    st.error("数据加载失败，请检查Excel文件")