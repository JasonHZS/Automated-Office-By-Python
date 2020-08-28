from PPTExcelUtils import pptMaker, TableChartMaker

# # plt.grid(True)  # 打开网格
# # 调整x轴标签，从垂直变成水平
# plt.xticks(rotation=360)
# # 设置坐标轴名称
# plt.xlabel("资产类别")
# plt.ylabel("数额")


def run():
    # ppt第2页操作
    pptMaker.PPT().p2_operate()

    # ppt第3页操作
    pptMaker.PPT().p3_operate()

    # ppt第5页操作
    pptMaker.PPT().p5_operate()

    # ppt第9页操作
    pptMaker.PPT().p9_operate()


if __name__ == "__main__":
    run()
    # TableChartMaker.Sheet4().c_barline_trend()
    # TableChartMaker.Sheet4().c_barline_activity()
    # TableChartMaker.Sheet4().c_pie_merchant()
    # TableChartMaker.Sheet3().c_barline_amount()
    # TableChartMaker.Sheet1().c_barline_profit()
    # TableChartMaker.Sheet1().c_pie_profit()
    # TableChartMaker.Sheet1().t_bpos_profit()
    # TableChartMaker.Sheet1().t_profit_ranking()
    # TableChartMaker.Sheet3().c_ring_pamount()
    # TableChartMaker.Sheet3().c_ring_pnums()
