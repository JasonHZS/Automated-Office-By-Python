import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import matplotlib.ticker as mtick

"""
       Axes这个不是你画图的xy坐标抽！可以把axes理解为你要放到画布上的各个物体!
       如果你的figure只有一张图，那么你只有一个axes
       如果你的figure有subplot，那么每一个subplot就是一个axes
       如果同一画布想要画柱形图与折线图，那么就添加多个axes，如下
"""

# 设置生成的图表支持中文显示
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['font.family'] = 'sans-serif'
# 将df中的科学计数法数值格式化为保留5位小数
pd.set_option('display.float_format', lambda x: '%.5f' % x)

# pd.set_option('display.max_columns', None)  #  全部显示科学计数法
# plt.style.use('ggplot')


def sign_data(ax, type, x, y, format):
    """

    :param ax:  图表对象
    :param type: 'num'或'rate'
    :param x: x轴数据list
    :param y: y轴数据list
    :param format: 数据格式（例如保留一位小数的百分数：'%.0f%%'，或者保留一位小数点：'%.1f'）
    :return:
    """
    for a, b in zip(x, y):
        if type == 'rate' and b < 0:
            ax.text(a, b, format % b, ha='center', va='bottom', c='white', fontsize=14)
        if type == 'rate' and b >= 0:
            ax.text(a, b, format % b, ha='center', va='bottom', fontsize=14)
        if type == 'num':
            ax.text(a, b, format % b, ha='center', va='bottom', fontsize=14)


class Sheet4(object):
    def __init__(self):
        self.path = "D:/项目文档/分析报告/3月大POS分析-副本.xlsx"
        self.sheet = 3

    def c_barline_trend(self):
        """
        sheet3对应数据生成柱形图与折线图：大POS新增入网商户数趋势
        :return:
        """
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=22, skip_footer=38)
        df1 = df.iloc[3:15, [0,1]]  # 柱形图
        df2 = df.iloc[3:15, [0,2]]  # 折线图
        df2['增长率'] = df2['增长率'] * 100
        month = df1['月份'].astype('str').values
        # df1['新增商户数'] = df1['新增商户数'].apply(lambda x: "{:.1f}".format(float(x)))
        num = df1['新增商户数'].values
        rate = df2['增长率'].values

        fig, ax = plt.subplots(figsize=(14, 7))
        plt.grid(axis="y")  # 打开网格
        ax.set_axisbelow(True)  # 设置柱形图在网格上层
        width = 0.5
        ax.bar(month, num, width, label='新增商户数')
        sign_data(ax, 'num', month, num, '%.1f')
        ax2 = ax.twinx()  # 设置共用x轴
        ax2.plot(month, rate, label='增长率', marker='.', color='orange')
        sign_data(ax2, 'rate', month, rate, '%.0f%%')

        ax.set_title('大POS新增入网商户数趋势', fontsize=17)
        fig.legend(loc='upper right')
        plt.text(0.1, 480, '单位：万', fontsize=12)
        plt.savefig("D:/项目文档/分析报告/pic/大POS新增入网商户数趋势.jpg", dpi=500)
        # plt.show()

    def c_barline_activity(self):
        """
        sheet3对应数据生成柱形图与折线图：大POS商户活跃度
        :return:
        """
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=51, skip_footer=68)
        df1 = df.iloc[3:15, [0,1,2,3,4]]

        month = df1['月份'].astype('str').values
        num_old = df1['老商户活跃数'].values
        num_new = df1['新商户活跃数'].values

        df1['老商户活跃率'] = df1['老商户活跃率'] * 100
        df1['新商户活跃率'] = df1['新商户活跃率'] * 100
        rate_old = df1['老商户活跃率'].values
        rate_new = df1['新商户活跃率'].values

        fig, ax = plt.subplots(figsize=(14, 7))
        plt.grid(axis="y")  # 打开网格
        ax.set_axisbelow(True)  # 设置柱形图在网格上层
        # 柱形图
        ax.bar(month, num_old, width=0.5, bottom=num_new, label='老商户活跃数')
        sign_data(ax, 'num', month, num_old, '%.1f')
        ax.bar(month, num_new, width=0.5, label='新商户活跃数')
        sign_data(ax, 'num', month, num_new, '%.1f')
        # 折线图
        ax2 = ax.twinx()  # 设置共用x轴
        ax2.plot(month, rate_old, label='老商户活跃率', marker='.', color='orange')
        sign_data(ax2, 'rate', month, rate_old, '%.0f%%')
        ax2.plot(month, rate_new, label='新商户活跃率', marker='.', color='red')
        sign_data(ax2, 'rate', month, rate_new, '%.0f%%')

        yticks = mtick.FormatStrFormatter(fmt='%.2f%%')  # 设置百分比形式的坐标轴
        ax2.yaxis.set_major_formatter(yticks)
        # 设置标注
        box = ax.get_position()
        ax.set_position([box.x0, box.y0, box.width, box.height * 0.9])
        ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.15), ncol=3, frameon=False, fontsize=13)
        ax2.legend(loc='lower center', bbox_to_anchor=(0.5, 1.1), ncol=3, frameon=False, fontsize=13)

        ax.set_title('大POS商户活跃度', fontsize=19)
        plt.text(0.1, 75, '单位：万', fontsize=12)
        plt.savefig("D:/项目文档/分析报告/pic/大POS商户活跃度.jpg", dpi=500)
        # plt.show()

    def c_pie_merchant(self):
        """
        sheet3对应数据生成饼图
        :return:
        """
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=40, skip_footer=45)
        df1 = df.iloc[0:2, [0, 2]].set_index('新增类别')
        df2 = df.iloc[0:3, [3, 5]].set_index('存量类型')
        df1.plot(kind='pie', autopct='%.0f%%', fontsize=15, subplots=True, explode=(0.01, 0.01), startangle=90,
                 legend=False, figsize=(7, 7))
        plt.title('3月大POS新增入网商户分布', size=25)
        plt.savefig("D:/项目文档/分析报告/pic/3月大POS新增入网商户分布.jpg", dpi=500)
        df2.plot(kind='pie', autopct='%.0f%%', fontsize=15, subplots=True, explode=(0.01, 0.01, 0.01), startangle=90,
                 legend=False, figsize=(7, 7))
        plt.title('大POS存量商户分布', size=25)
        plt.savefig("D:/项目文档/分析报告/pic/大POS存量商户分布.jpg", dpi=500)
        # plt.show()


class Sheet3(object):
    def __init__(self):
        self.path = "D:/项目文档/分析报告/3月大POS分析-副本.xlsx"
        self.sheet = 2

    def c_barline_amount(self):
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=1)
        df1 = df.iloc[3:15, [0,1,2,3]]

        month = df1['月份'].astype('str').values
        old_mer_amo = df1['老商户交易金额'].values
        df1['增长率'] = df1['增长率'] * 100
        new_mer_amo = df1['新商户交易金额'].values
        rate = df1['增长率'].values

        fig, ax = plt.subplots(figsize=(14, 7))
        plt.grid(axis="y")  # 打开网格
        ax.set_axisbelow(True)  # 设置柱形图在网格上层
        # 柱形图
        ax.bar(month, old_mer_amo, width=0.5, bottom=new_mer_amo, label='老商户交易金额')
        sign_data(ax, 'num', month, old_mer_amo, '%.1f')
        ax.bar(month, new_mer_amo, width=0.5, label='新商户交易金额')
        sign_data(ax, 'num', month, new_mer_amo, '%.1f')
        # 折线图
        ax2 = ax.twinx()  # 设置共用x轴
        ax2.plot(month, rate, label='增长率', marker='.', color='orange')
        sign_data(ax2, 'rate', month, rate, '%.0f%%')
        yticks = mtick.FormatStrFormatter(fmt='%.1f%%')  # 设置百分比形式的坐标轴
        ax2.yaxis.set_major_formatter(yticks)
        ax.set_title('大POS交易金额趋势图', fontsize=19)
        box = ax.get_position()
        ax.set_position([box.x0, box.y0, box.width, box.height * 0.9])
        ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.15), ncol=3, frameon=False, fontsize=13)
        ax2.legend(loc='lower center', bbox_to_anchor=(0.5, 1.1), ncol=3, frameon=False, fontsize=13)
        plt.text(0.01, 180, '单位：亿元', fontsize=12)
        plt.savefig("D:/项目文档/分析报告/pic/大POS交易金额趋势图.jpg", dpi=500)
        # plt.show()

    def c_ring_pnums(self):
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=55)
        df1 = df.iloc[0:7, [0, 2]].set_index(keys='交易类型')
        df2 = df.iloc[8:10, [0, 2]].set_index('交易类型')
        explode = (0.04, 0.01, 0.01, 0.02, 0.05, 0.1, 0.01)
        """ 参数说明：
            labeldistance，文本的位置离远点有多远，1.1指1.1倍半径的位置
            autopct，圆里面的文本格式，%3.1f%%表示小数有三位，整数有一位的浮点数
            shadow，饼是否有阴影
            startangle，起始角度，0，表示从0开始逆时针转，为第一块。一般选择从90度开始比较好看
            pctdistance，百分比的text离圆心的距离
        """
        ax2 = df2.plot.pie(autopct='%.2f%%', fontsize=14, subplots=True, explode=(0.01, 0.01), startangle=220,
                           radius=0.65, wedgeprops=dict(width=0.4, edgecolor='w'), legend=False, labeldistance=0.7,
                           pctdistance=0.5, figsize=(8, 7))
        df1.plot.pie(ax=ax2, autopct='%.2f%%', fontsize=15, subplots=True, explode=explode, startangle=100,
                     radius=1, wedgeprops=dict(width=0.3, edgecolor='w'), legend=False, labeldistance=1.05,
                     pctdistance=0.85, figsize=(8, 7))
        plt.axis('equal')  # 设置x，y轴刻度一致，这样饼图才能是圆的
        plt.title('3月各交易类型交易笔数', size=22)
        plt.savefig("D:/项目文档/分析报告/pic/3月各交易类型交易笔数.jpg", dpi=500)
        # plt.show()
        df1.sort_values(by="笔数占比", ascending=False, inplace=True)
        return df1[0:3].index.values, df1[0:3].values

    def c_ring_pamount(self):
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=72)
        df1 = df.iloc[0:7, [0, 1]].set_index('交易类型')
        # df1['金额占比'] = df1['金额占比'].apply(lambda x: "{:.2%}".format(float(x)))
        df2 = df.iloc[7:9, [0, 1]].set_index('交易类型')
        # df2['金额占比'] = df2['金额占比'].apply(lambda x: "{:.2%}".format(float(x)))

        explode = (0.1, 0.01, 0.01, 0.02, 0.05, 0.05, 0.05)
        ax2 = df2.plot.pie(autopct='%.2f%%', fontsize=14, subplots=True, explode=(0.01, 0.01), startangle=220,
                           radius=0.65, wedgeprops=dict(width=0.4, edgecolor='w'), legend=False, labeldistance=0.7,
                           pctdistance=0.5, figsize=(8, 7))
        df1.plot.pie(ax=ax2, autopct='%.2f%%', fontsize=15, subplots=True, explode=explode, startangle=180,
                     radius=1, wedgeprops=dict(width=0.3, edgecolor='w'), legend=False, labeldistance=1.05,
                     pctdistance=0.85, figsize=(8, 7))
        plt.axis('equal')  # 设置x，y轴刻度一致，这样饼图才能是圆的
        plt.title('3月各交易类型交易金额', size=22)
        plt.savefig("D:/项目文档/分析报告/pic/3月各交易类型交易金额.jpg", dpi=500)
        # plt.show()
        df1.sort_values(by="金额占比", ascending=False, inplace=True)
        return df1[0:3].index.values, df1[0:3].values


class Sheet1(object):
    def __init__(self):
        self.path = "D:/项目文档/分析报告/3月大POS分析-副本.xlsx"
        self.sheet = 0

    def t_profit_ranking(self):
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=19, skip_footer=31)
        df1 = df.iloc[0:11, [0, 1, 2, 3, 4]]
        # 将该列格式化（加上‘+’或‘-’）输出,':+.0f'的对象必须是浮点数
        df1['Unnamed: 1'][1:11] = df1['Unnamed: 1'][1:11].apply(lambda x: "{:+.0f}".format(float(x)) if x != '-' else x)
        return df1, df1.shape[0], df1.shape[1]

    def t_bpos_profit(self):
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=34)
        df1 = df.iloc[0:14, [0, 1, 2, 3, 4, 5, 6, 7, 8]]
        df1['Unnamed: 1'][1:14] = df1['Unnamed: 1'][1:14].apply(lambda x: "{:.0f}".format(float(x)))
        df1['Unnamed: 2'][1:14] = df1['Unnamed: 2'][1:14].apply(lambda x: "{:.0f}".format(float(x)))
        df1['Unnamed: 3'][1:14] = df1['Unnamed: 3'][1:14].apply(lambda x: "{:.0f}".format(float(x)))
        df1['Unnamed: 4'][1:14] = df1['Unnamed: 4'][1:14].apply(lambda x: "{:.0f}".format(float(x)))
        df1['Unnamed: 5'][1:13] = df1['Unnamed: 5'][1:13].apply(lambda x: "{:.0%}".format(float(x)))
        df1['Unnamed: 6'][1:13] = df1['Unnamed: 6'][1:13].apply(lambda x: "{:.0%}".format(float(x)))
        df1['Unnamed: 7'][1:13] = df1['Unnamed: 7'][1:13].apply(lambda x: "{:.0%}".format(float(x)))
        df1['Unnamed: 8'][1:13] = df1['Unnamed: 8'][1:13].apply(lambda x: "{:.0%}".format(float(x)))
        return df1, df1.shape[0], df1.shape[1]

    def c_pie_profit(self):
        """
        sheet1对应数据生成3月收益分布饼图
        :return:
        """
        df = pd.read_excel(self.path, sheet_name=self.sheet)
        df1 = df.iloc[0:3, [12, 13]].set_index(keys='当月收益分布')
        explode = (0.01, 0.01, 0.01)
        df1.plot(kind='pie', autopct='%.0f%%', fontsize=15, subplots=True, explode=explode, startangle=90, legend=False, figsize=(7, 7))
        df2 = df.iloc[[0,15], [4,5,6,7,8]]
        plt.savefig("D:/项目文档/分析报告/pic/3月收益分布.jpg", dpi=500)
        return df1, df2

    def c_barline_profit(self):
        df = pd.read_excel(self.path, sheet_name=self.sheet, skiprows=1)
        df1 = df.iloc[3:15, [0, 1, 2, 3, 4, 5, 6]]  # 柱形图——收益

        month = df1['月份'].astype('str').values
        trans_profit = df1['交易收益'].values
        withdraw_profit = df1['提现收益'].values
        server_profit = df1['服务费收益'].values
        df1['交易收益增长率'] = df1['交易收益增长率'] * 100
        df1['提现收益增长率'] = df1['提现收益增长率'] * 100
        df1['服务费收益增长率'] = df1['服务费收益增长率'] * 100
        trans_rate = df1['交易收益增长率'].values
        withdraw_rate = df1['提现收益增长率'].values
        server_rate = df1['服务费收益增长率'].values

        fig, ax = plt.subplots(figsize=(10, 5))
        plt.grid(axis="y")  # 打开网格
        ax.set_axisbelow(True)  # 设置柱形图在网格上层
        ax.set_title('大POS收益趋势图', fontsize=19)
        plt.text(0.1, 330, '单位：万', fontsize=14)
        # 柱形图
        ax.bar(month, trans_profit, width=0.5, bottom=withdraw_profit, label='交易收益')
        sign_data(ax, 'num', month, trans_profit, '%.1f')
        ax.bar(month, withdraw_profit, width=0.5, bottom=server_profit, label='提现收益')
        sign_data(ax, 'num', month, withdraw_profit, '%.1f')
        ax.bar(month, server_profit, width=0.5, label='服务费收益')
        sign_data(ax, 'num', month, server_profit, '%.1f')
        # 折线图
        ax2 = ax.twinx()  # 设置共用x轴
        ax2.plot(month, trans_rate, label='交易收益增长率', marker='*')
        ax2.plot(month, withdraw_rate, label='提现收益增长率', marker='*')
        ax2.plot(month, server_rate, label='服务费收益增长率', marker='*')

        yticks = mtick.FormatStrFormatter(fmt='%.1f%%')  # 设置百分比形式的坐标轴
        ax2.yaxis.set_major_formatter(yticks)
        plt.tick_params(labelsize=14)
        box = ax.get_position()
        ax.set_position([box.x0, box.y0, box.width, box.height * 0.9])
        ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.15), ncol=3, frameon=False, fontsize=13)
        ax2.legend(loc='lower center', bbox_to_anchor=(0.5, 1.1), ncol=3, frameon=False, fontsize=13)
        plt.savefig("D:/项目文档/分析报告/pic/大POS收益趋势图.jpg", dpi=500)
        # plt.show()
