import pandas as pd
from openpyxl import load_workbook
import xlwt


def join(file_inl, file_inr, file_out, right_index, left_key):
    """
    将需求发过来的excel与查询结果excel进行左连接
    :param file_inl: 需求发过来的excel路径(左表)
    :param file_inr: 查询结果excel路径（右表）
    :param file_out: 保存文件的路径
    :param right_index: 右表设置的索引（用于连接）
    :param left_key: 左表用于连接的字段
    :return: None
    """
    df1 = pd.read_excel(file_inl, converters={u'商户号': str}, sheet_name=0)
    df2 = pd.read_excel(file_inr,
                        converters={u'左端商户号': str, u'法人手机号': str, u'持卡人身份证号码': str, u'营业执照号码': str},
                        sheet_name=0)
    writer = pd.ExcelWriter(file_out, engine='xlsxwriter', options={'strings_to_urls': False})

    # 方法一：df左连接默认按双方的索引连接，可修改索引双方索引进行连接，左表的商户号会提到第一列显示
    # df = df1.set_index('商户号').join(df2.set_index('左端商户号'), how='left')
    # df.to_excel(writer)

    # 方法二：仅修改右表索引，用on选择左表用于连接的字段（此时左表的索引仍为默认），左表完全保持原样
    df = df1.join(df2.set_index(right_index), on=left_key)
    df.to_excel(writer, index=False)

    writer.close()


if __name__ == "__main__":
    # file_inl = 'D:/项目文档/副本导联系方式不加密安徽.xlsx'
    file_inl = 'C:/Users/huangzesen/Desktop/工作簿2.xlsx'
    file_inr = 'C:/Users/huangzesen/Desktop/0804.xlsx'
    file_out = 'C:/Users/huangzesen/Desktop/result_0804.xlsx'
    right_index = '左端商户号'
    left_key = '商户号'
    join(file_inl, file_inr, file_out, right_index, left_key)
