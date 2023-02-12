import streamlit as st
import pandas as pd
import plotly.express as px
from PIL import Image
import xlwt

# 网页名称、标题、子标题
def head_set():
    st.set_page_config(page_title='质差用户数据分析系统')
    st.header('质差用户数据分析分析可视化展示')

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
    #df = pd.read_excel(uploaded_file, header=0, usecols=[1,3,46,38,41,14,61])
    # 展示导入的表格
    #st.dataframe(df)
def shuju():
    df = pd.read_excel(uploaded_file, header=0, usecols=[1, 3, 38, 41, 46, 61])
    def fun(df):
        dic = {}
        for i in df.columns:
            dic[i] = df[i].value_counts()
        return dic

    dd = fun(df)

    f = xlwt.Workbook()  # 创建工作薄
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5
    style = xlwt.XFStyle()
    style.pattern = pattern
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style.alignment = al
    # 获取字典的键
    list_ = [k for k in dd]
    k = 0
    l = 0
    for s in range(len(dd)):
        l = k + 1
        # 写入第一行
        # sheet1.write_merge(0, 1, k, l, list_[s], style)
        # sheet1.write_merge(0, 1, k, l, list_[s], style)
        sheet1.write(0, k, list_[s])
        sheet1.write(0, l, list_[s] + "重复次数")
        # 写入内容
        j = 1
        for v, h in zip(dd[list_[s]], dd[list_[s]].index):
            sheet1.write(j, k, h)  # 循环写入 竖着写
            sheet1.write(j, l, v)  # 循环写入 竖着写
            j = j + 1
        k = k + 2
    f.save('统计1.xlsx')
    excel_da = '统计1.xlsx'
    da = pd.read_excel("统计1.xlsx")
    sheet_name1 = 'sheet1'
    # 筛选满足要求的数据
    exectop5 = da[(da['机顶盒业务账号重复次数'] >= 3) & (da['olt设备ip重复次数'] >= 20)]
    cess = da[da['olt设备ip重复次数'] >= 20]
    ba = exectop5.to_excel('筛选数据后的文件1.xlsx', sheet_name='Sheet1', index=False, header=True)
    # 再次读取数据
    sx_file = '筛选数据后的文件1.xlsx'
    sx_name = 'Sheet1'

    # 业务
    df_iptv = pd.read_excel(uploaded_file,
                       #sheet_name=sheet_name,
                       usecols='D,BJ,B',
                       header=0)
    # olt
    df_olt = pd.read_excel(uploaded_file,
                        #   sheet_name=sheet_name,
                           usecols='D,BJ,AU',
                           header=0)
    # 此处为各部门参加问卷调查人数
    df_participants = pd.read_excel(sx_file,
                                    sheet_name=sx_name,
                                    usecols='A:P',
                                    header=0)

    # header指定那一行作为列名0，hearder=1：选择第二行为表头，第一行数据就不要了。其他以此类推
    # df_participants.dropna(inplace=True)

    # with st.spinner('请等待...'):
    #    time.sleep(5)

    # st.balloons()
    # 根据选择分组数据

    # 多页面程序
    class MultiApp:
        def __init__(self):
            self.apps = []
            self.app_dict = {}

        def add_app(self, title, func):
            if title not in self.apps:
                self.apps.append(title)
                self.app_dict[title] = func

        def run(self):
            title = st.sidebar.radio(
                '菜单栏',
                self.apps,
                format_func=lambda title: str(title))
            self.app_dict[title]()

    def foo1():
        # 设置网页标题
        st.header('质差用户数据详情')
        # 设置网页子标题
        # 展示导入的表格
        st.dataframe(df)
        # st.subheader('IPTV质差用户数据可视化')
        st.balloons()
        # streamlit的多重选择(选项数据)
        department = df_iptv['时间'].unique().tolist()
        # streamlit的滑动条(年龄数据)
        ages = df_iptv['卡顿次数'].unique().tolist()
        # 滑动条, 最大值、最小值、区间值
        age_selection = st.slider('卡顿次数:',
                                  min_value=min(ages),
                                  max_value=max(ages),
                                  value=(min(ages), max(ages)))
        # 多重选择, 默认全选
        department_selection = st.multiselect('时间:',
                                              department,
                                              default=department)
        # 根据选择过滤数据
        mask = (df_iptv['卡顿次数'].between(*age_selection)) & (df_iptv['时间'].isin(department_selection))
        number_of_result = df_iptv[mask].shape[0]

        # 根据筛选条件, 得到有效数据
        st.markdown(f'*有效数据: {number_of_result}*')
        # 根据选择分组数据
        df_grouped = df_iptv[mask].groupby(by=['机顶盒业务账号']).count()[['卡顿次数']]
        df_grouped = df_grouped.rename(columns={'卡顿次数': '次数'})
        df_grouped = df_grouped.reset_index()

        bar_chart = px.bar(df_grouped.head(15),
                           y='机顶盒业务账号',
                           x='次数',
                           text='次数',
                           color_discrete_sequence=['#F63366'] * len(df_grouped),
                           template='plotly_white',
                           orientation='h',
                           width=950,
                           height=700,
                           )
        st.plotly_chart(bar_chart)

        # 添加图片和交互式表格
        col1, col2 = st.columns(2)
        image = Image.open('wanfeng.jpg')
        col1.image(image,
                   caption='业务帐号卡顿详情->',
                   width=500,
                   use_column_width=True)
        col2.dataframe(df_iptv[mask], width=500)

        pie_chart = px.pie(df_participants,
                           title='时间分布图',
                           values='时间重复次数',
                           names='时间',
                           height=700, )
        st.plotly_chart(pie_chart)

    def bar2():
        # streamlit的多重选择(选项数据)
        departmenta = df_olt['时间'].unique().tolist()
        # streamlit的滑动条(年龄数据)
        ages = df_olt['卡顿次数'].unique().tolist()
        # 滑动条, 最大值、最小值、区间值
        age_selection = st.slider('卡顿次数:',
                                  min_value=min(ages),
                                  max_value=max(ages),
                                  value=(min(ages), max(ages)))
        # 多重选择, 默认全选
        department_selection = st.multiselect('时间:',
                                              departmenta,
                                              default=departmenta)
        # 根据选择过滤数据
        mask1 = (df_olt['卡顿次数'].between(*age_selection)) & (df_olt['时间'].isin(department_selection))
        number_of_result = df_olt[mask1].shape[0]

        # 根据筛选条件, 得到有效数据
        st.markdown(f'*有效数据: {number_of_result}*')
        # 根据选择分组数据
        df_grouped = df_olt[mask1].groupby(by=['olt设备ip']).count()[['卡顿次数']]
        df_grouped = df_grouped.rename(columns={'卡顿次数': '次数'})
        df_grouped = df_grouped.reset_index()

        bar_chart = px.bar(df_grouped.head(15),
                           y='olt设备ip',
                           x='次数',
                           text='次数',
                           color_discrete_sequence=['#F63366'] * len(df_grouped),
                           template='plotly_white',
                           orientation='h',
                           width=950,
                           height=700,
                           )
        st.plotly_chart(bar_chart)

        # 添加图片和交互式表格
        col1, col2 = st.columns(2)
        image = Image.open('mao1.jpg')
        col1.image(image,
                   caption='Mao',
                   use_column_width=True)
        col2.dataframe(df_olt[mask1], width=700)
        # 扇形图
        pie_chart = px.pie(df_participants,
                           title='bras设备ip分布图',
                           values='bras设备ip重复次数',
                           names='bras设备ip',
                           height=700, )
        st.plotly_chart(pie_chart)

    def sur3():
        # 扇形图
        pie_chart = px.pie(df_participants,
                           title='硬件型号分布图',
                           values='硬件型号重复次数',
                           names='硬件型号',
                           height=700, )
        st.plotly_chart(pie_chart)
        # 扇形图
        pie_chart = px.pie(df_participants,
                           title='机顶盒软件版本号分布图',
                           values='机顶盒软件版本号重复次数',
                           names='机顶盒软件版本号',
                           height=700, )
        st.plotly_chart(pie_chart)

    app = MultiApp()
    app.add_app("业务帐号时间分布情况", foo1)
    app.add_app("OLT、MSE分布趋势", bar2)
    app.add_app('机顶盒型号详情信息', sur3)
    app.run()


# 图片展示
def imageShow():
    image = Image.open("lu.jpg")  # 在未导入excel时，展示图片
    st.image(image, clamp=False,
             channels="RGB", output_format="auto")


# 调用
head_set()
template_download()
excel_upload()

# 如果文件已上传
if uploaded_file is not None:
    upload_read()
    shuju()

# 否则展示图片
else:
    imageShow()