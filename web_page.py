import streamlit as st
import re
import pythoncom

pythoncom.CoInitialize()



# 设置网页标题，以及使用宽屏模式
st.set_page_config(
    page_title="工程设计设备数量计数工具",
    layout="centered", initial_sidebar_state="auto", page_icon=":shark:",
    menu_items={
        'About': "Developer: XAB and CX, CASCO!",
        'Get Help': None,
        'Report a bug': None
    }
)

# 隐藏右边的菜单以及页脚
# hide_streamlit_style = """
#                         <style>
#                         #MainMenu {visibility: hidden;}
#                         footer {visibility: hidden;}
#                         </style>
#                         """
# st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# 隐藏页脚
hide_streamlit_style = """
                        <style>
                        footer {visibility: hidden;}
                        </style>
                        """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# 左侧的导航栏
st.sidebar.write("导航栏")
sidebar = st.sidebar.radio(
    "系统功能",
    ("组合继电器", "防雷分线柜", "分线综合柜", "IBP盘面按钮", "铁路信号设备", "图纸数量"),
)
if sidebar == "组合继电器":
    st.title("组合继电器计数")
    jdq_file_name = st.text_input('请输入需要计数的文件位置', key="jdq_equ")
    if not jdq_file_name:
        st.warning('请输入计数文件的文件夹位置')
    uploaded_file = st.file_uploader("请选择需要计数的文件", key="jdq_equ")
    if uploaded_file is not None:
        # st.stop()
        jdq_type = st.multiselect(
            "请选择设备类型",
            ["全部", "JWXC-1700", 'JPXC-1000', 'JWJXC-H125/80', 'JYJXC-160/260', 'JWJXC-480', 'JSBXC-850',
             'JWXC-H340', 'JZXC-H18', "JWJXC-H125/0.44", 'BD2-7', 'QDX1-S13/30', 'SP-905-04']
        )
        jdq_path = uploaded_file.name
        jdq_location = jdq_file_name + '\\' + jdq_path
        button_select = st.button("开始计数", key=0)
        jdq_add = st.text_input('请输入设备类型')
        button_add = st.button("开始计数", key=1)
#         if button_select:
#             if '全部' in jdq_type:
#                 jdq_new_type = ["JWXC-1700", 'JPXC-1000', 'JWJXC-H125/80', 'JYJXC-160/260', 'JWJXC-480',
#                                 'JSBXC-850',
#                                 'JWXC-H340', 'JZXC-H18', "JWJXC-H125/0.44", 'BD2-7', 'QDX1-S13/30', 'SP-905-04']
#                 count_system.resultOutputTime()
#                 count_system.countFunction(jdq_location, jdq_new_type).ZHJDQ()
#             else:
#                 count_system.resultOutputTime()
#                 count_system.countFunction(jdq_location, jdq_type).ZHJDQ()
#         if button_add:
#             jdq_add = re.split('[,，]', jdq_add)
#             jdq_add = list(filter(None, jdq_add))
#             count_system.resultOutputTime()
#             count_system.countFunction(jdq_location, jdq_add).ZHJDQ()

if sidebar == "防雷分线柜":
    st.title("防雷分线柜计数")
    fl_file_name = st.text_input('请输入需要计数的文件位置', key="fl_equ")
    if not fl_file_name:
        st.warning('请输入计数文件的文件夹位置')
    uploaded_file = st.file_uploader("请选择需要计数的文件", key="fl_equ")
    if uploaded_file is not None:
        # st.stop()
        fl_type = st.multiselect(
            "请选择设备类型",
            ["全部", "SFLM-220", "SFLM-120", "SFLM-60", "SFLM-C"]
        )
        fl_path = uploaded_file.name
        fl_location = fl_file_name + '\\' + fl_path
        button_select = st.button("开始计数", key=0)
        fl_add = st.text_input('请输入设备类型')
        button_add = st.button("开始计数", key=1)
#         if button_select:
#             if '全部' in fl_type:
#                 fl_new_type = ["SFLM-220", "SFLM-120", "SFLM-60", "SFLM-C"]
#                 count_system.resultOutputTime()
#                 count_system.countFunction(fl_location, fl_new_type).FLFXG()
#             else:
#                 count_system.resultOutputTime()
#                 count_system.countFunction(fl_location, fl_type).FLFXG()
#         if button_add:
#             fl_add = re.split('[,，]', fl_add)
#             fl_add = list(filter(None, fl_add))
#             count_system.resultOutputTime()
#             count_system.countFunction(fl_location, fl_add).FLFXG()

if sidebar == "分线综合柜":
    st.title("分线综合柜计数")
    fx_file_name = st.text_input('请输入需要计数的文件位置', key="fx_equ")
    if not fx_file_name:
        st.warning('请输入计数文件的文件夹位置')
    uploaded_file = st.file_uploader("请选择需要计数的文件", key="fx_equ")
    if uploaded_file is not None:
        # st.stop()
        fx_type = st.multiselect(
            "请选择设备类型",
            ["全部", "SFLM-220", "SFLM-60"]
        )
        fx_path = uploaded_file.name
        fx_location = fx_file_name + '\\' + fx_path
        button_select = st.button("开始计数", key=0)
        fx_add = st.text_input('请输入设备类型')
        button_add = st.button("开始计数", key=1)
#         if button_select:
#             if '全部' in fx_type:
#                 fx_new_type = ["SFLM-220", "SFLM-60"]
#                 count_system.resultOutputTime()
#                 count_system.countFunction(fx_location, fx_new_type).FXZHG()
#             else:
#                 count_system.resultOutputTime()
#                 count_system.countFunction(fx_location, fx_type).FXZHG()
#         if button_add:
#             fx_add = re.split('[,，]', fx_add)
#             fx_add = list(filter(None, fx_add))
#             count_system.resultOutputTime()
#             count_system.countFunction(fx_location, fx_add).FXZHG()

if sidebar == "IBP盘面按钮":
    st.title("IBP盘面按钮计数")
    IBP_file_name = st.text_input('请输入需要计数的文件位置', key="an_equ")
    if not IBP_file_name:
        st.warning('请输入计数文件的文件夹位置')
    uploaded_file = st.file_uploader("请选择需要计数的文件", key="an_equ")
    if uploaded_file is not None:
        # st.stop()
        an_type = st.multiselect(
            "请选择设备类型",
            ["全部", "计轴复位按钮", "紧急停车按钮", '紧停取消按钮', '扣车按钮', '终止扣车按钮',
             '信号试灯按钮', '报警切除按钮', '故障旁路按钮', '故障旁路恢复按钮', '紧停报警蜂鸣器']
        )
        IBP_path = uploaded_file.name
        an_location = IBP_file_name + '\\' + IBP_path
        button_select = st.button("开始计数", key=0)
        an_add = st.text_input('请输入设备类型')
        button_add = st.button("开始计数", key=1)
#         if button_select:
#             if '全部' in an_type:
#                 an_new_type = ["计轴复位按钮", "紧急停车按钮", '紧停取消按钮', '扣车按钮', '终止扣车按钮',
#                                '信号试灯按钮', '报警切除按钮', '故障旁路按钮', '故障旁路恢复按钮', '紧停报警蜂鸣器']
#                 count_system.resultOutputTime()
#                 count_system.countFunction(an_location, an_new_type).IBP()
#             else:
#                 count_system.resultOutputTime()
#                 count_system.countFunction(an_location, an_type).IBP()
#         if button_add:
#             an_add = re.split('[,，]', an_add)
#             an_add = list(filter(None, an_add))
#             count_system.resultOutputTime()
#             count_system.countFunction(an_location, an_add).IBP()

if sidebar == "铁路信号设备":
    st.title("铁路信号设备计数")
    rail_equ_name = st.text_input('请输入需要计数的文件位置', key="railway_equ")
    if not rail_equ_name:
        st.warning('请输入计数文件的文件夹位置')
    uploaded_file = st.file_uploader("请选择需要计数的文件", key="railway_equ")
    if uploaded_file is not None:
        # st.stop()
        rail_equ_path = uploaded_file.name
        rail_equ_location = rail_equ_name + '\\' + rail_equ_path
        rail_equ_type = st.multiselect(
            "请选择设备类型",
            ['AX', '无源', '有源', "预告信号机", '调车信号机', '进站信号机',
             '总出发信号机', '总折返信号机', '矮柱出站信号机', '高柱出站信号机'])  # list类型
        button_select = st.button("开始计数", key=0)
        rail_equ_add = st.text_input('请输入设备类型')
        button_add = st.button("开始计数", key=1)
#         if button_select:
#             count_system.resultOutputTime()
#             count_system.rail_equ_count(rail_equ_location, rail_equ_type, None)
#         if button_add:
#             rail_equ_add = re.split('[,，]', rail_equ_add)
#             rail_equ_add = list(filter(None, rail_equ_add))
#             count_system.resultOutputTime()
#             count_system.rail_equ_count(rail_equ_location, [], rail_equ_add)

if sidebar == "图纸数量":
    st.title("图纸数量计数")
    page_file_name = st.text_input('请输入需要计数的文件位置', key="page_count")
    button_start = st.button("开始计数", key="page_count")
    if button_start:
        page_path_list = count_system.getfiles(page_file_name)
        for file in page_path_list:
            pycad = count_system.open_close_CAD(file)
            count_system.resultOutputTime()
            count, text_count = count_system.pageCount(pycad)
            st.write(pycad.doc.Name + '内的图纸数量为' + str(count))
            pycad.ActiveDocument.Close()
