import pandas as pd
from openpyxl import load_workbook
import xlwt
import os


def split_excel(file_in, sheetName, sheet_rows, file_out):
    """
    将原excel中某个sheet的某些字段数据拆分成 n条/sheet 保存成新的excel中的新的sheet
    :param file_in: 目标读取excel或其路径
    :param sheetName: excel中对应的sheet名
    :param sheet_rows: 每个sheet几条数据
    :param file_out: 保存文件的路径
    :return: None
    """
    df = pd.read_excel(file_in, sheet_name=sheetName)
    df = df['支付宝商户号'].astype(str)  # 修改字段类型，未转为字符型的话会导致商户号变为科学计数法

    sheet_nums = df.count()//sheet_rows if df.count() % sheet_rows == 0 else df.count()//sheet_rows + 1

    wb = xlwt.Workbook()  # 创建xls文件对象
    sheet_list = []
    for i in range(sheet_nums):
        # 循环创建需要的sheet
        sheet = wb.add_sheet(str(i+1))
        sheet_list.append(sheet)

    wb.save(file_out)

    # DataFrame进行保存时为了避免被不断地覆盖，这里engine使用了openpyxl
    writer = pd.ExcelWriter(file_out, engine='openpyxl')

    df_list = []
    for i in range(sheet_nums):
        df_i = df.iloc[5000*i:5000*(i+1)]  # 选取每张sheet对应的行
        df_list.append(df_i)
        df_list[i].to_excel(excel_writer=writer, sheet_name=str(i+1), index=None, encoding="utf-8")
        writer.save()

    writer.close()


def getFileName(file_in):
    """
    使用os模块walk函数，搜索出某目录下的全部excel文件
    :param file_in: 目标读取exceld的文件夹
    :return: excel路径列表
    """
    # 获取同一个文件夹下的所有excel文件名
    file_list = []
    for root, dirs, files in os.walk(file_in):
        for filespath in files:
            print(os.path.join(root, filespath))
            file_list.append(os.path.join(root, filespath))

    return file_list


def MergeExcel(file_in, file_out):
    """
    将某一文件夹里的excel合并为一份（excel的表格字段应该一致）
    :param file_in: 目标读取exceld的文件夹路径
    :param file_out: 保存文件的路径
    :return: None
    """
    file_list = getFileName(file_in)
    result = pd.DataFrame()
    # 合并多个excel文件
    for each in file_list:
        # 读取xlsx格式文件，涉及到卡号/编号等数字前带0的要添加参数converters，使其保持文本格式，避免丢失0
        data1 = pd.read_excel(each, converters={u'终端号': str,u'商户号': str,u'发卡行': str,u'收单行': str,u'卡号': str})
        result = result.append(data1)

    # 写出数据xlsx格式
    writer = pd.ExcelWriter(file_in + '/' + file_out, engine='xlsxwriter', options={'strings_to_urls': False})
    result.to_excel(writer, index=False)
    writer.close()


if __name__ == "__main__":
    file_in = 'D:/项目文档/【支付宝商户】2020.xlsx'
    sheetName = 1
    sheet_rows = 5000
    file_out = 'D:/项目文档/【支付宝商户】2020_拆分.xls'
    split_excel(file_in, sheetName, sheet_rows, file_out)

    # file_in = 'D:/项目文档/test'
    # file_out = 'result.xlsx'
    # MergeExcel(file_in, file_out)
