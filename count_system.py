import pyautocad
import datetime
import os
import streamlit as st
import re
import time

Autocad = pyautocad.Autocad


def resultOutputTime():
    st.write('***************************************** ' +
             datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')
             + ' ******************************************')


def openCAD(fpath1):
    pycad = Autocad()
    try:
        (filepath, temp_filename) = os.path.split(fpath1)  # 返回当前路径以及文件名（没有路径，只是文件的名字）
        if pycad.doc.Name == temp_filename:
            return pycad
        else:
            pycad = Autocad(create_if_not_exists=True)
            pycad.app.Application.Documents.Open(fpath1)
            return pycad
    except:
        pycad = Autocad(create_if_not_exists=True)
        pycad.app.Application.Documents.Open(fpath1)
        return pycad


def open_close_CAD(path):  # 针对对图纸数量的计数不需要显示所有cad
    pycad = Autocad(create_if_not_exists=True)
    pycad.app.Visible = 0
    time.sleep(0.1)
    pycad.app.Application.Documents.Open(path)
    # time.sleep(0.1)
    return pycad


def getfiles(file):  # file为文件夹
    path_list = []  # 文件的绝对路径
    for filepath, dirnames, filenames in os.walk(file):
        # os.walk()方法：可以递归的找出目表路径下的所有文件
        for filename in filenames:
            if filename.endswith('.dwg'):
                path_list.append(os.path.join(filepath, filename))
    # filenames=os.listdir(file)
    # print(filenames)
    return path_list


def pageCount(pycad):  # 图纸数量计数
    st.write('正在识别' + pycad.doc.Name)
    count = 0
    text_count = []
    for text in pycad.iter_objects('Text'):
        if re.fullmatch('第[0-9]+页[,，]共[0-9]+页', text.TextString):
            count += 1
    return count, text_count


def display(equ, equipment_count):
    st.write(equ + '的数量为：' + str(len(equipment_count)))
    for equ_com in equipment_count:
        st.write(equ_com)


class countFunction:
    def __init__(self, filepath, equ_type):
        self.filepath = filepath
        self.equ_type = equ_type

    def ZHJDQ(self):  # 对组合继电器类型表的图内设备进行计数
        pycad = openCAD(self.filepath)
        st.write('正在识别' + pycad.doc.Name)
        equ_type_count = len(self.equ_type)  # 选择要识别的继电器类型的数量
        cad_text_zb_int = []
        YL_count = []  # 记录预留设备数量
        equ_count = [0 for index in range(equ_type_count)]  # 生成与继电器类型的数量相等的0作为其初始化
        equ_type_count_dic = dict(zip(self.equ_type, equ_count))
        for text in pycad.iter_objects('Text'):
            x = text.InsertionPoint[:2]
            x = list(x)
            x = (int(text.InsertionPoint[0]), int(text.InsertionPoint[1]))
            x = tuple(x)
            if x in cad_text_zb_int:
                continue
            else:
                cad_text_zb_int.append(x)
                if text.TextString in self.equ_type:
                    equ_type_count_dic[text.TextString] += 1
                    text.Highlight(HighlightFlag=True)
                if re.search('[(（]预留[)）]', text.TextString):
                    YL_count.append(text.TextString)
        for equ in equ_type_count_dic.keys():
            st.write(equ + "的数量是" + str(equ_type_count_dic[equ]))
        if YL_count:
            display('预留设备', YL_count)

    def FXZHG(self):  # 对分线综合柜的图内设备进行计数
        pycad = openCAD(self.filepath)
        st.write('正在识别' + pycad.doc.Name)
        equ_type_count = len(self.equ_type)  # 选择要识别的设备类型的数量
        cad_text_zb_int = []
        YL_count = []  # 记录预留设备数量
        equ_count = [0 for index in range(equ_type_count)]  # 生成与设备类型的数量相等的0作为其初始化
        equ_type_count_dic = dict(zip(self.equ_type, equ_count))
        for text in pycad.iter_objects('Text'):
            x = text.InsertionPoint[:2]
            x = list(x)
            x = (int(text.InsertionPoint[0]), int(text.InsertionPoint[1]))
            x = tuple(x)
            if x in cad_text_zb_int:
                continue
            else:
                cad_text_zb_int.append(x)
                if text.TextString in self.equ_type:
                    equ_type_count_dic[text.TextString] += 1
                    text.Highlight(HighlightFlag=True)
                if re.search('[(（]预留[)）]', text.TextString):
                    YL_count.append(text.TextString)
        for equ in equ_type_count_dic.keys():
            st.write(equ + "的数量是" + str(equ_type_count_dic[equ]))
        if YL_count:
            st.write('预留设备', YL_count)

    def FLFXG(self):  # 对防雷分线柜的图内设备进行计数
        pycad = openCAD(self.filepath)
        st.write('正在识别' + pycad.doc.Name)
        equ_type_count = len(self.equ_type)  # 选择要识别的设备类型的数量
        cad_text_zb_int = []
        YL_count = []  # 记录预留设备数量
        equ_count = [0 for index in range(equ_type_count)]  # 生成与设备类型的数量相等的0作为其初始化
        equ_type_count_dic = dict(zip(self.equ_type, equ_count))
        for text in pycad.iter_objects('Text'):
            x = text.InsertionPoint[:2]
            x = list(x)
            x = (int(text.InsertionPoint[0]), int(text.InsertionPoint[1]))
            x = tuple(x)
            if x in cad_text_zb_int:
                continue
            else:
                cad_text_zb_int.append(x)
                if text.TextString in self.equ_type:
                    equ_type_count_dic[text.TextString] += 1
                    text.Highlight(HighlightFlag=True)
                if re.search('[(（]预留[)）]', text.TextString):
                    YL_count.append(text.TextString)
        for equ in equ_type_count_dic.keys():
            st.write(equ + "的数量是" + str(equ_type_count_dic[equ]))
        if YL_count:
            display('预留设备', YL_count)

    def IBP(self):  # 对IBP盘盘面布置图的图内设备进行计数
        p_JZFW = re.compile("(ST[0-9]+)|(T[0-9]+)|(ST[0-9]+[()（）]预留[()（）])|(T[0-9]+[()（）]预留[()（）])")  # 计轴复位按钮匹配模板
        p_JJTC = re.compile("紧急停车[0-9]*")  # 紧急停车按钮匹配模板
        p_JTQX = re.compile('取消紧停[0-9]*')  # 紧停取消按钮匹配模板
        p_KC = re.compile('扣车[0-9]*')  # 扣车按钮匹配模板
        p_ZZKC = re.compile('终止扣车[0-9]*')  # 终止扣车按钮匹配模板
        p_XHSD = re.compile('信号试灯[0-9]*')  # 信号试灯按钮匹配模板
        p_BJQC = re.compile("报警切除[0-9]*")  # 报警切除按钮匹配模板
        p_GZPL = re.compile('W' + '[0-9]+' + '故障旁路[0-9]*')  # 故障旁路按钮匹配模板
        p_GZPLHF = re.compile('W' + '[0-9]+' + '故障旁路恢复')  # 故障旁路恢复按钮匹配模板
        p_JTFMQ = re.compile('紧停报警[0-9]*')  # 紧停报警蜂鸣器匹配模板

        equipment_BUTTON_dict = {"计轴复位按钮": p_JZFW, "紧急停车按钮": p_JJTC, '紧停取消按钮': p_JTQX,
                                 '扣车按钮': p_KC, '终止扣车按钮': p_ZZKC,
                                 '信号试灯按钮': p_XHSD, '报警切除按钮': p_BJQC, '故障旁路按钮': p_GZPL,
                                 '故障旁路恢复按钮': p_GZPLHF,
                                 '紧停报警蜂鸣器': p_JTFMQ
                                 }
        pycad = openCAD(self.filepath)
        st.write('正在识别' + pycad.doc.Name)
        equ_type_count = len(self.equ_type)  # 选择要识别的设备类型的数量
        cad_text_zb_int = []
        YL_count = []  # 记录预留设备数量
        equ_count = [0 for index in range(equ_type_count)]  # 生成与设备类型的数量相等的0作为其初始化
        equ_type_count_dic = dict(zip(self.equ_type, equ_count))
        # print(equ_type_count_dic)
        for text in pycad.iter_objects('Text'):
            x = text.InsertionPoint[:2]
            x = list(x)
            x = (int(text.InsertionPoint[0]), int(text.InsertionPoint[1]))
            x = tuple(x)
            if x in cad_text_zb_int:
                continue
            else:
                cad_text_zb_int.append(x)
                # equipment_count = list(filter(lambda equ: re.fullmatch(equipment_BUTTON_dict[equ], text.TextString) != None, equ_type_count_dic.keys()))
                count_list = list(
                    filter(lambda equ: re.fullmatch(equipment_BUTTON_dict[equ], text.TextString), self.equ_type))
                # 循环对每个text.TextString进行循环正则匹配，匹配有两层对应，第一层为输入的self.equ_type对应equipment_BUTTON_dict的键进行循环，并获取其字典的值作为
                # 匹配模板，匹配得到的结果返回为equipment_BUTTON_dict某个键的列表形式，第二层为self.equ_type新生成的字典，其为设备：设备数量的字典，
                # 所以第二层对应为匹配返回的列表元素与新生成的字典的键对应，匹配返回的列表长度与新生成的字典的值对应
                # print(count_list)
                if count_list:
                    equ_type_count_dic[count_list[0]] += len(count_list)
                if re.search('[(（]预留[)）]', text.TextString):
                    YL_count.append(text.TextString)
        for equ in equ_type_count_dic.keys():
            if equ == '紧停报警蜂鸣器':
                st.write(equ + "的数量是" + str(equ_type_count_dic[equ] - 1))
            else:
                st.write(equ + "的数量是" + str(equ_type_count_dic[equ]))
        if YL_count:
            display('预留设备', YL_count)


def rail_equ_count(filepath, equ_type, input_equ):
    pycad = openCAD(filepath)
    if len(equ_type) != 0:
        st.write('正在识别' + pycad.doc.Name)
        equ_type_count = len(equ_type)  # 选择要识别的设备类型的数量
        cad_ax_int = []  # AX坐标列表，实现去重
        cad_gz_int = []  # 高柱出站信号机坐标列表，实现去重
        cad_az_int = []  # 矮柱出站信号机坐标列表，实现去重
        cad_dc_int = []  # 调车信号机坐标列表，实现去重
        cad_jz_int = []  # 进站信号机坐标列表，实现去重
        cad_wy_int = []  # 无源信标坐标列表，实现去重
        cad_yy_int = []  # 有源信标坐标列表，实现去重
        cad_zcf_int = []  # 总出发信号机坐标列表，实现去重
        cad_yg_int = []  # 预告信号机坐标列表，实现去重
        cad_zzf_int = []  # 总折返信号机坐标列表，实现去重
        equ_list = [0 for index in range(equ_type_count)]  # 生成与设备类型的数量相等个数的0，进行初始化
        equ_type_count_dic = dict(zip(equ_type, equ_list))
        for blockObj in pycad.iter_objects('BlockReference'):
            if "AX" in equ_type:
                if blockObj.Name == "AX":
                    # x = blockObj.InsertionPoint[:2]
                    # x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_ax_int:
                        continue
                    else:
                        cad_ax_int.append(x)
                        equ_type_count_dic["AX"] += 1
                        # blockObj.Highlight(HighlightFlag=True)
            if "矮柱出站信号机" in equ_type:
                if blockObj.Name in ["出站信号机-矮", "矮柱出站", "出发信号机-矮"]:
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_az_int:
                        continue
                    else:
                        cad_az_int.append(x)
                        equ_type_count_dic["矮柱出站信号机"] += 1
            if "高柱出站信号机" in equ_type:
                if blockObj.Name in ["出站信号机-高", "高柱出站", "出发信号机-高"]:
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_gz_int:
                        continue
                    else:
                        cad_gz_int.append(x)
                        equ_type_count_dic["高柱出站信号机"] += 1
            if "调车信号机" in equ_type:
                if blockObj.Name in ["调车信号机", "调车"]:
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_dc_int:
                        continue
                    else:
                        cad_dc_int.append(x)
                        equ_type_count_dic["调车信号机"] += 1
            if "进站信号机" in equ_type:
                if blockObj.Name in ["进站信号机", "进站"]:
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_jz_int:
                        continue
                    else:
                        cad_jz_int.append(x)
                        equ_type_count_dic["进站信号机"] += 1
            if "无源" in equ_type:
                if blockObj.Name == "无源":
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_wy_int:
                        continue
                    else:
                        cad_wy_int.append(x)
                        equ_type_count_dic["无源"] += 1
            if "有源" in equ_type:
                if blockObj.Name == "有源":
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_yy_int:
                        continue
                    else:
                        cad_yy_int.append(x)
                        equ_type_count_dic["有源"] += 1
            if "总出发信号机" in equ_type:
                if blockObj.Name in ["总出发信号机", "总出发"]:
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_zcf_int:
                        continue
                    else:
                        cad_zcf_int.append(x)
                        equ_type_count_dic["总出发信号机"] += 1
            if "预告信号机" in equ_type:
                if blockObj.Name in ["预告信号机", "预告"]:
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_yg_int:
                        continue
                    else:
                        cad_yg_int.append(x)
                        equ_type_count_dic["预告信号机"] += 1
            if "总折返信号机" in equ_type:
                if blockObj.Name in ["总折返信号机", "总折返"]:
                    x = blockObj.InsertionPoint[:2]
                    x = list(x)
                    x = (int(blockObj.InsertionPoint[0]), int(blockObj.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_zzf_int:
                        continue
                    else:
                        cad_zzf_int.append(x)
                        equ_type_count_dic["总折返信号机"] += 1
        for equ in equ_type_count_dic.keys():
            st.write(equ + "的数量: " + str(equ_type_count_dic[equ]))
    else:
        st.write('正在识别' + pycad.doc.Name)
        cad_equ_int = []  # 输入设备的坐标列表，实现去重
        equ_type_count = len(input_equ)  # 选择要识别的继电器类型的数量
        equ_count = [0 for index in range(equ_type_count)]  # 生成与继电器类型的数量相等的0作为其初始化
        equ_type_count_dic = dict(zip(input_equ, equ_count))
        for blockObject in pycad.iter_objects('BlockReference'):
            for input_name in input_equ:
                if blockObject.Name.lower() == input_name.lower() or blockObject.Name == input_name:
                    x = (int(blockObject.InsertionPoint[0]), int(blockObject.InsertionPoint[1]))
                    x = tuple(x)
                    if x in cad_equ_int:
                        continue
                    else:
                        cad_equ_int.append(x)
                        blockObject.Highlight(HighlightFlag=True)
                        equ_type_count_dic[str(input_name)] += 1
        for equ in equ_type_count_dic.keys():
            st.write(equ + "的数量是" + str(equ_type_count_dic[equ]))
