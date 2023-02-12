import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image


# 网页名称、标题、子标题
def head_set():
    st.set_page_config(page_title='企业项目成本挣值分析可视化展示')
    st.header('企业项目成本挣值分析可视化展示')
    st.subheader('Designed by Garvey.Wu')


# 下载模板文件
def template_download():
    with open("moban.xlsx", "rb") as file:
        btn = st.download_button(label="下载模板文件",
                                 data=file, file_name="template.xlsx")


# 上传分析文件
def excel_upload():
    global uploaded_file  # 全局变量
    uploaded_file = st.file_uploader("上传分析文件",
                                     type="xlsx")


# 读取上传文件
def upload_read():
    global df  # 全局变量
    # 读取xlsx文件
    df = pd.read_excel(uploaded_file)
    # 展示导入的表格
    st.dataframe(df)


# 在饼图中，显示各活动PV
def pv_pieShow():
    pv_df = pd.read_excel(uploaded_file, usecols="A:B")
    pv_piechart = px.pie(pv_df, title="各活动PV(计划价值)比例", values="PV（计划价值）（万元）", names="活动名称")
    st.plotly_chart(pv_piechart)


# 在进度图中，显示各活动实际进度
def progress_barShow():
    progress_df = pd.read_excel(uploaded_file, usecols=[0, 5])
    progress_bar = px.bar(progress_df, title="实际进度图(0到1)*100%",
                          x='活动名称', y='当前进度')
    st.plotly_chart(progress_bar)


# 挣值分析指标图
def figure1_barShow():
    # 挣值分析指标图
    figure1_df = pd.read_excel(uploaded_file, usecols="I:J")
    figure1_df = figure1_df.head(9)
    st.markdown("挣值分析指标表格")
    st.table(figure1_df)

    # streamlit的多重选择（选项数据）
    figure1 = figure1_df['挣值分析指标'].unique().tolist()
    figure1_selection = st.multiselect("挣值分析指标", figure1, default=figure1)  # 多选，默认全选
    mask = figure1_df['挣值分析指标'].isin(figure1_selection)  # 根据某属性选取指定条件的行,完成过滤
    figure1_bar = px.bar(figure1_df[mask], title="挣值分析指标可视化", x='指标值（万元）', y='挣值分析指标', text="指标值（万元）",
                         orientation="h", color_discrete_sequence=['#F63366'] * len(figure1_df),
                         template='plotly_white')
    st.plotly_chart(figure1_bar)


# 绩效指数展示
def performShow():
    global cpi_data, spi_data, tcpi_data  # 全局变量，便于performJudge()调用
    cpi_data = df.iloc[[9], [9]].values[0][0]  # 选取在excel中cpi值
    spi_data = df.iloc[[10], [9]].values[0][0]  # 选取在excel中spi值
    tcpi_data = df.iloc[[11], [9]].values[0][0]  # 选取在excel中tcpi值
    col1, col2, col3 = st.columns(3)
    col1.metric("CPI", cpi_data, "成本绩效指数")
    col2.metric("SPI", spi_data, "进度绩效指数")
    col3.metric("TCPI", tcpi_data, "完工尚需绩效指数")


# 判断绩效执行状态
def performJudge():
    st.markdown("绩效执行状态:")
    # 成本绩效判断
    if cpi_data > 1:
        st.write("成本节约")
    elif cpi_data < 1:
        st.write("成本超支")
    else:
        st.write("成本平衡")

    # 进度绩效判断
    if spi_data > 1:
        st.write("进度提前")
    elif spi_data < 1:
        st.write("进度落后")
    else:
        st.write("进度平衡")


# 图片展示
def imageShow():
    image = Image.open("survey.jpg")  # 在未导入excel时，展示图片
    st.image(image, clamp=False,
             channels="RGB", output_format="auto")


# 调用
head_set()
template_download()
excel_upload()

# 如果文件已上传
if uploaded_file is not None:
    upload_read()
    pv_pieShow()
    progress_barShow()
    figure1_barShow()
    performShow()
    performJudge()

# 否则展示图片
else:
    imageShow()

