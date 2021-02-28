# coding=utf-8
import io

import las
import matplotlib
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlrd
import xlwt
from matplotlib.widgets import MultiCursor, RadioButtons, SpanSelector
from PyQt5.QtWidgets import (QApplication, QFileDialog, QMainWindow,
                             QTableWidgetItem)


class MATPLOTLIB_MFC40:
    def __init__(self, fileDir):
        self.damage_Tag = ''
        self.lines = []  # 生成的成果list
        self.log = las.LASReader(fileDir, null_subs=np.nan)
        self.fig1 = plt.figure('MIT油套管快速评价系统', figsize=(12, 8))
        xls = xlrd.open_workbook(".\\casing_data.xls")
        table = xls.sheet_by_name('Sheet1')
        # 注意下面几个的类型为excel单元格对象
        self.outer_diameter = table.cell(0, 2)
        self.inner_diameter = table.cell(0, 3)
        self.thickness = table.cell(0, 4)

        self.scale_left = float(self.inner_diameter.value) / 2 - 20
        self.scale_right = float(self.inner_diameter.value) / 2 + 120

        self.scale_left_min = float(self.inner_diameter.value) - 30
        self.scale_right_max = float(self.inner_diameter.value) + 30

        # 定义RadioButtons
        axcolor = 'lightgoldenrodyellow'
        rax = plt.axes([0.75, 0.05, 0.12, 0.07], facecolor=axcolor)
        radio = RadioButtons(rax, (u'Penetration', u'Projection', u'Transformation'), active=-1, activecolor='purple')
        plt.subplots_adjust(bottom=0.15, top=0.95, right=0.9, left=0.10, wspace=0.60)
        radio.on_clicked(self.actionfunc)
        #####################################################################################
        # 坐标轴1
        self.ax1 = plt.subplot(141)
        # 下面赋值加逗号是为了使得type(self.line1)为matplotlib.lines.Line2D对象，而不是list
        self.line1, = self.ax1.plot(self.log.data['D01'], self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D02'] + 2.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D03'] + 5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D04'] + 7.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D05'] + 10, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['D06'] + 12.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D07'] + 15, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D08'] + 17.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D09'] + 20, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D10'] + 22.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['D11'] + 25, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D12'] + 27.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D13'] + 30, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D14'] + 32.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D15'] + 35, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['D16'] + 37.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D17'] + 40, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D18'] + 42.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D19'] + 45, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D20'] + 47.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['D21'] + 50, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D22'] + 52.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D23'] + 55, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D24'] + 57.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D25'] + 60, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D26'] + 62.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['D27'] + 65, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D28'] + 67.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D29'] + 70, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D30'] + 72.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D31'] + 75, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['D32'] + 77.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D33'] + 80, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D34'] + 82.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D35'] + 85, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D36'] + 87.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['D37'] + 90, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D38'] + 92.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D39'] + 95, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['D40'] + 97.5, self.log.data['DEPT'], 'g-', lw=0.3)

        self.ax1.set_xlim(self.scale_left, self.scale_right)
        self.ax1.set_ylim(self.log.start, self.log.stop)
        self.ax1.invert_yaxis()

        span1 = SpanSelector(self.ax1, self.onselect1, 'vertical', useblit=False,
                             rectprops=dict(alpha=0.5, facecolor='yellow'), span_stays=True)
        # plt.ylabel(self.log.curves.DEPT.descr + " (%s)" % self.log.curves.DEPT.units)
        # plt.xlabel(self.log.curves.D01.descr + " (%s)" % self.log.curves.D01.units)
        # plt.title(self.log.well.WELL.data)
        plt.ylabel('Measured Depth(m)')
        plt.title('Original')

        plt.gca().spines['bottom'].set_position(('data', 0))
        plt.gca().spines['top'].set_position(('data', 0))
        plt.grid()
        #####################################################################################
        # 坐标轴2
        self.ax2 = plt.subplot(142)
        self.line2, = self.ax2.plot(self.log.data['D01'], self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D02'] + 2.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D03'] + 5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D04'] + 7.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D05'] + 10, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['D06'] + 12.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D07'] + 15, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D08'] + 17.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D09'] + 20, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D10'] + 22.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['D11'] + 25, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D12'] + 27.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D13'] + 30, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D14'] + 32.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D15'] + 35, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['D16'] + 37.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D17'] + 40, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D18'] + 42.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D19'] + 45, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D20'] + 47.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['D21'] + 50, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D22'] + 52.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D23'] + 55, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D24'] + 57.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D25'] + 60, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D26'] + 62.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['D27'] + 65, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D28'] + 67.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D29'] + 70, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D30'] + 72.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D31'] + 75, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['D32'] + 77.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D33'] + 80, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D34'] + 82.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D35'] + 85, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D36'] + 87.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['D37'] + 90, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D38'] + 92.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D39'] + 95, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['D40'] + 97.5, self.log.data['DEPT'], 'g-', lw=0.3)

        self.ax2.set_xlim(self.scale_left, self.scale_right)
        self.ax2.set_ylim(self.log.start, self.log.stop)
        self.ax2.invert_yaxis()

        span2 = SpanSelector(self.ax2, self.onselect2, 'vertical', useblit=False,
                             rectprops=dict(alpha=0.5, facecolor='yellow'), span_stays=True)
        plt.title('Middle')
        plt.gca().spines['bottom'].set_position(('data', 0))
        plt.gca().spines['top'].set_position(('data', 0))
        self.ax2.grid()
        #####################################################################################
        # 坐标轴3
        self.ax3 = plt.subplot(143)
        self.line3, = self.ax3.plot(self.log.data['D01'], self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D02'] + 2.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D03'] + 5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D04'] + 7.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D05'] + 10, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['D06'] + 12.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D07'] + 15, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D08'] + 17.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D09'] + 20, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D10'] + 22.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['D11'] + 25, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D12'] + 27.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D13'] + 30, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D14'] + 32.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D15'] + 35, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['D16'] + 37.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D17'] + 40, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D18'] + 42.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D19'] + 45, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D20'] + 47.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['D21'] + 50, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D22'] + 52.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D23'] + 55, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D24'] + 57.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D25'] + 60, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D26'] + 62.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['D27'] + 65, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D28'] + 67.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D29'] + 70, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D30'] + 72.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D31'] + 75, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['D32'] + 77.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D33'] + 80, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D34'] + 82.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D35'] + 85, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D36'] + 87.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['D37'] + 90, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D38'] + 92.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D39'] + 95, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['D40'] + 97.5, self.log.data['DEPT'], 'g-', lw=0.3)

        self.ax3.set_xlim(self.scale_left, self.scale_right)
        self.ax3.set_ylim(self.log.start, self.log.stop)
        self.ax3.invert_yaxis()

        # self.span3_cyan = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
        #                           rectprops=dict(alpha=0.5, facecolor='cyan'), span_stays=True)

        plt.title('Large')
        plt.gca().spines['bottom'].set_position(('data', 0))
        plt.gca().spines['top'].set_position(('data', 0))
        self.ax3.grid()
        #####################################################################################
        # 坐标轴4
        self.ax4 = plt.subplot(144)
        self.ax4.plot(self.log.data['IDMX'], self.log.data['DEPT'], 'r--', lw=0.3)
        self.ax4.plot(self.log.data['IDMN'], self.log.data['DEPT'], 'b-', lw=0.3)
        self.ax4.plot(self.log.data['IDAV'], self.log.data['DEPT'], 'k--', lw=0.3)

        self.ax4.set_xlim(self.scale_left_min, self.scale_right_max)
        self.ax4.set_ylim(self.log.start, self.log.stop)
        self.ax4.invert_yaxis()

        plt.title('Min-Max')
        plt.gca().spines['bottom'].set_position(('data', 0))
        plt.gca().spines['top'].set_position(('data', 0))
        self.ax4.grid()

        multi = MultiCursor(plt.gcf().canvas, (self.ax1, self.ax2, self.ax3, self.ax4), color='r', lw=1,
                            horizOn=True, vertOn=False)
        #########################

        plt.show()

    def onselect1(self, ymin, ymax):
        indmin, indmax = np.searchsorted(self.log.data['DEPT'], (ymin, ymax))
        indmax = min(len(self.log.data['DEPT']) - 1, indmax)
        thisx = self.log.data['D01'][indmin:indmax]
        thisy = self.log.data['DEPT'][indmin:indmax]
        self.line2.set_data(thisx, thisy)
        self.ax2.set_ylim(thisy[-1], thisy[0])
        self.line3.set_data(thisx, thisy)
        self.ax3.set_ylim(thisy[-1], thisy[0])
        self.ax4.set_ylim(thisy[-1], thisy[0])
        plt.gcf().canvas.draw()

    def onselect2(self, ymin, ymax):
        indmin, indmax = np.searchsorted(self.log.data['DEPT'], (ymin, ymax))
        indmax = min(len(self.log.data['DEPT']) - 1, indmax)
        thisx = self.log.data['D01'][indmin:indmax]
        thisy = self.log.data['DEPT'][indmin:indmax]
        self.line3.set_data(thisx, thisy)
        self.ax3.set_ylim(thisy[-1], thisy[0])
        self.ax4.set_ylim(thisy[-1], thisy[0])
        plt.gcf().canvas.draw()

    def onselect3(self, ymin, ymax):
        ymin = round(ymin, 3)
        ymax = round(ymax, 3)
        if ymin != ymax:
            self.lines = []  # 清空一下
            print('井段为：', ymin, '-', ymax)
            self.lines.append(ymin)
            self.lines.append(ymax)
        elif ymin == ymax:
            print('极值深度为：', ymax)
            self.lines.append(ymax)
            index = np.where(abs(self.log.data['DEPT'] - ymax) <= float(abs(self.log.step) - 0.001))
            index = index[0][0]
            D01_value = self.log.data['D01'][index]
            D02_value = self.log.data['D02'][index]
            D03_value = self.log.data['D03'][index]
            D04_value = self.log.data['D04'][index]
            D05_value = self.log.data['D05'][index]
            D06_value = self.log.data['D06'][index]
            D07_value = self.log.data['D07'][index]
            D08_value = self.log.data['D08'][index]
            D09_value = self.log.data['D09'][index]
            D10_value = self.log.data['D10'][index]
            D11_value = self.log.data['D11'][index]
            D12_value = self.log.data['D12'][index]
            D13_value = self.log.data['D13'][index]
            D14_value = self.log.data['D14'][index]
            D15_value = self.log.data['D15'][index]
            D16_value = self.log.data['D16'][index]
            D17_value = self.log.data['D17'][index]
            D18_value = self.log.data['D18'][index]
            D19_value = self.log.data['D19'][index]
            D20_value = self.log.data['D20'][index]
            D21_value = self.log.data['D21'][index]
            D22_value = self.log.data['D22'][index]
            D23_value = self.log.data['D23'][index]
            D24_value = self.log.data['D24'][index]
            D25_value = self.log.data['D25'][index]
            D26_value = self.log.data['D26'][index]
            D27_value = self.log.data['D27'][index]
            D28_value = self.log.data['D28'][index]
            D29_value = self.log.data['D29'][index]
            D30_value = self.log.data['D30'][index]
            D31_value = self.log.data['D31'][index]
            D32_value = self.log.data['D32'][index]
            D33_value = self.log.data['D33'][index]
            D34_value = self.log.data['D34'][index]
            D35_value = self.log.data['D35'][index]
            D36_value = self.log.data['D36'][index]
            D37_value = self.log.data['D37'][index]
            D38_value = self.log.data['D38'][index]
            D39_value = self.log.data['D39'][index]
            D40_value = self.log.data['D40'][index]
            self.Min_value = self.log.data['IDMN'][index]
            self.Ave_value = self.log.data['IDAV'][index]
            self.Max_value = self.log.data['IDMX'][index]
            print(str(self.Min_value), ' ', str(self.Ave_value), ' ', str(self.Max_value))
            if self.damage_Tag == 'Transformation':
                self.lines.append(round(self.Min_value, 3))
                self.lines.append(round(self.Ave_value, 3))
                self.lines.append(round(self.Max_value, 3))
                print(self.lines)
                # 保存list到文件
                self.lines_temp = pd.DataFrame(self.lines, columns=['value'])
                writer = pd.ExcelWriter('Transformation.xlsx')
                self.lines_temp.to_excel(writer, 'Sheet1')
                writer.save()
            elif self.damage_Tag == 'Penetration':
                All_Ders = [D01_value, D02_value, D03_value, D04_value, D05_value, D06_value,
                               D07_value, D08_value, D09_value, D10_value, D11_value, D12_value,
                               D13_value, D14_value, D15_value, D16_value, D17_value, D18_value,
                               D19_value, D20_value, D21_value, D22_value, D23_value, D24_value,
                               D25_value, D26_value, D27_value, D28_value, D29_value, D30_value,
                               D31_value, D32_value, D33_value, D34_value, D35_value, D36_value,
                               D37_value, D38_value, D39_value, D40_value]
                All_Ders_Dict = {'D01_value': D01_value,
                                    'D02_value': D02_value,
                                    'D03_value': D03_value,
                                    'D04_value': D04_value,
                                    'D05_value': D05_value,
                                    'D06_value': D06_value,
                                    'D07_value': D07_value,
                                    'D08_value': D08_value,
                                    'D09_value': D09_value,
                                    'D10_value': D10_value,
                                    'D11_value': D11_value,
                                    'D12_value': D12_value,
                                    'D13_value': D13_value,
                                    'D14_value': D14_value,
                                    'D15_value': D15_value,
                                    'D16_value': D16_value,
                                    'D17_value': D17_value,
                                    'D18_value': D18_value,
                                    'D19_value': D19_value,
                                    'D20_value': D20_value,
                                    'D21_value': D21_value,
                                    'D22_value': D22_value,
                                    'D23_value': D23_value,
                                    'D24_value': D24_value,
                                    'D25_value': D25_value,
                                    'D26_value': D26_value,
                                    'D27_value': D27_value,
                                    'D28_value': D28_value,
                                    'D29_value': D29_value,
                                    'D30_value': D30_value,
                                    'D31_value': D31_value,
                                    'D32_value': D32_value,
                                    'D33_value': D33_value,
                                    'D34_value': D34_value,
                                    'D35_value': D35_value,
                                    'D36_value': D36_value,
                                    'D37_value': D37_value,
                                    'D38_value': D38_value,
                                    'D39_value': D39_value,
                                    'D40_value': D40_value,
                                    }
                FING_Max_value = max(All_Ders_Dict.values())
                FING_String = ''
                for key, value in All_Ders_Dict.items():
                    if value == FING_Max_value:
                        FING_String = key
                        self.lines.append(int(FING_String[1:3]))
                        value_3_digits = round(value, 3) # 保留三位有效数字
                        self.lines.append(value_3_digits)
                        break

                # 求个平均值
                normal_Depth1 = float(self.lines[0]) - 0.2  # 在井段开始深度上方0.2m处取一个正常的点进行读值
                normal_Depth2 = float(self.lines[0]) - 0.4  # 在井段开始深度上方0.4m处取一个正常的点进行读值
                normal_Depth3 = float(self.lines[0]) - 0.6  # 在井段开始深度上方0.6m处取一个正常的点进行读值
                normal_Depth4 = float(self.lines[0]) - 0.8  # 在井段开始深度上方0.8m处取一个正常的点进行读值
                normal_Depth5 = float(self.lines[0]) - 1.0  # 在井段开始深度上方1.0m处取一个正常的点进行读值

                index1 = np.where(abs(self.log.data['DEPT'] - normal_Depth1) <= float(abs(self.log.step) - 0.001))
                index1 = index1[0][0]
                FING_value1 = self.log.data[FING_String[0:3]][index1]

                index2 = np.where(abs(self.log.data['DEPT'] - normal_Depth2) <= float(abs(self.log.step) - 0.001))
                index2 = index2[0][0]
                FING_value2 = self.log.data[FING_String[0:3]][index2]

                index3 = np.where(abs(self.log.data['DEPT'] - normal_Depth3) <= float(abs(self.log.step) - 0.001))
                index3 = index3[0][0]
                FING_value3 = self.log.data[FING_String[0:3]][index3]

                index4 = np.where(abs(self.log.data['DEPT'] - normal_Depth4) <= float(abs(self.log.step) - 0.001))
                index4 = index4[0][0]
                FING_value4 = self.log.data[FING_String[0:3]][index4]

                index5 = np.where(abs(self.log.data['DEPT'] - normal_Depth5) <= float(abs(self.log.step) - 0.001))
                index5 = index5[0][0]
                FING_value5 = self.log.data[FING_String[0:3]][index5]

                FING_value = round(sum([FING_value1, FING_value2, FING_value3, FING_value4, FING_value5]) / 5.0, 3)

                self.lines.append(FING_value)
                self.lines.append(round(self.Min_value, 3))
                self.lines.append(round(self.Ave_value, 3))
                self.lines.append(round(self.Max_value, 3))
                print(self.lines)
                # 保存list到文件
                self.lines_temp = pd.DataFrame(self.lines, columns=['value'])
                writer = pd.ExcelWriter('Penetration.xlsx')
                self.lines_temp.to_excel(writer, 'Sheet1')
                writer.save()
            elif self.damage_Tag == 'Projection':
                All_Ders = [D01_value, D02_value, D03_value, D04_value, D05_value, D06_value,
                               D07_value, D08_value, D09_value, D10_value, D11_value, D12_value,
                               D13_value, D14_value, D15_value, D16_value, D17_value, D18_value,
                               D19_value, D20_value, D21_value, D22_value, D23_value, D24_value,
                               D25_value, D26_value, D27_value, D28_value, D29_value, D30_value,
                               D31_value, D32_value, D33_value, D34_value, D35_value, D36_value,
                               D37_value, D38_value, D39_value, D40_value]
                All_Ders_Dict = {'D01_value': D01_value,
                                    'D02_value': D02_value,
                                    'D03_value': D03_value,
                                    'D04_value': D04_value,
                                    'D05_value': D05_value,
                                    'D06_value': D06_value,
                                    'D07_value': D07_value,
                                    'D08_value': D08_value,
                                    'D09_value': D09_value,
                                    'D10_value': D10_value,
                                    'D11_value': D11_value,
                                    'D12_value': D12_value,
                                    'D13_value': D13_value,
                                    'D14_value': D14_value,
                                    'D15_value': D15_value,
                                    'D16_value': D16_value,
                                    'D17_value': D17_value,
                                    'D18_value': D18_value,
                                    'D19_value': D19_value,
                                    'D20_value': D20_value,
                                    'D21_value': D21_value,
                                    'D22_value': D22_value,
                                    'D23_value': D23_value,
                                    'D24_value': D24_value,
                                    'D25_value': D25_value,
                                    'D26_value': D26_value,
                                    'D27_value': D27_value,
                                    'D28_value': D28_value,
                                    'D29_value': D29_value,
                                    'D30_value': D30_value,
                                    'D31_value': D31_value,
                                    'D32_value': D32_value,
                                    'D33_value': D33_value,
                                    'D34_value': D34_value,
                                    'D35_value': D35_value,
                                    'D36_value': D36_value,
                                    'D37_value': D37_value,
                                    'D38_value': D38_value,
                                    'D39_value': D39_value,
                                    'D40_value': D40_value,
                                    }
                FING_Min_value = min(All_Ders_Dict.values())
                FING_String = ''
                for key, value in All_Ders_Dict.items():
                    if value == FING_Min_value:
                        FING_String = key
                        self.lines.append(int(FING_String[1:3]))
                        value_3_digits = round(value, 3)  # 保留三位有效数字
                        self.lines.append(value_3_digits)
                        break

                # 求个平均值
                normal_Depth1 = float(self.lines[0]) - 0.2  # 在井段开始深度上方0.2m处取一个正常的点进行读值
                normal_Depth2 = float(self.lines[0]) - 0.4  # 在井段开始深度上方0.4m处取一个正常的点进行读值
                normal_Depth3 = float(self.lines[0]) - 0.6  # 在井段开始深度上方0.6m处取一个正常的点进行读值
                normal_Depth4 = float(self.lines[0]) - 0.8  # 在井段开始深度上方0.8m处取一个正常的点进行读值
                normal_Depth5 = float(self.lines[0]) - 1.0  # 在井段开始深度上方1.0m处取一个正常的点进行读值

                index1 = np.where(abs(self.log.data['DEPT'] - normal_Depth1) <= float(abs(self.log.step) - 0.001))
                index1 = index1[0][0]
                FING_value1 = self.log.data[FING_String[0:3]][index1]

                index2 = np.where(abs(self.log.data['DEPT'] - normal_Depth2) <= float(abs(self.log.step) - 0.001))
                index2 = index2[0][0]
                FING_value2 = self.log.data[FING_String[0:3]][index2]

                index3 = np.where(abs(self.log.data['DEPT'] - normal_Depth3) <= float(abs(self.log.step) - 0.001))
                index3 = index3[0][0]
                FING_value3 = self.log.data[FING_String[0:3]][index3]

                index4 = np.where(abs(self.log.data['DEPT'] - normal_Depth4) <= float(abs(self.log.step) - 0.001))
                index4 = index4[0][0]
                FING_value4 = self.log.data[FING_String[0:3]][index4]

                index5 = np.where(abs(self.log.data['DEPT'] - normal_Depth5) <= float(abs(self.log.step) - 0.001))
                index5 = index5[0][0]
                FING_value5 = self.log.data[FING_String[0:3]][index5]

                FING_value = round(sum([FING_value1, FING_value2, FING_value3, FING_value4, FING_value5]) / 5.0, 3)

                self.lines.append(FING_value)
                self.lines.append(round(self.Min_value, 3))
                self.lines.append(round(self.Ave_value, 3))
                self.lines.append(round(self.Max_value, 3))
                print(self.lines)
                # 保存list到文件
                self.lines_temp = pd.DataFrame(self.lines, columns=['value'])
                writer = pd.ExcelWriter('Projection.xlsx')
                self.lines_temp.to_excel(writer, 'Sheet1')
                writer.save()

            self.fig2 = plt.figure('截面图')
            # 设置下面所需要的参数
            barSlices1 = 40
            barSlices2 = 100

            # theta指每个标记所在射线与极径的夹角，下面表示均分角度
            theta1 = np.linspace(0.0, 2 * np.pi, barSlices1, endpoint=False)
            theta2 = np.linspace(0.0, 2 * np.pi, barSlices2, endpoint=False)

            # # r表示点距离圆心的距离，np.random.rand(barSlices)表示返回返回服从“0-1”均匀分布的随机样本值
            # r = 2 * np.random.rand(barSlices) + 50
            r = [D01_value, D02_value, D03_value, D04_value, D05_value, D06_value,
                   D07_value, D08_value, D09_value, D10_value, D11_value, D12_value,
                   D13_value, D14_value, D15_value, D16_value, D17_value, D18_value,
                   D19_value, D20_value, D21_value, D22_value, D23_value, D24_value,
                   D25_value, D26_value, D27_value, D28_value, D29_value, D30_value,
                   D31_value, D32_value, D33_value, D34_value, D35_value, D36_value,
                   D37_value, D38_value, D39_value, D40_value]

            # 网上搜的方法，不知道怎么就可以闭合了(黑人问号)
            r = np.concatenate((r, [r[0]]))  # 闭合
            theta1 = np.concatenate((theta1, [theta1[0]]))  # 闭合
            theta2 = np.concatenate((theta2, [theta2[0]]))  # 闭合

            inside_radius = [float(self.inner_diameter.value) / 2] * barSlices2
            outside_radius = [float(self.outer_diameter.value) / 2] * barSlices2

            inside_radius = np.concatenate((inside_radius, [inside_radius[0]]))  # 闭合
            outside_radius = np.concatenate((outside_radius, [outside_radius[0]]))  # 闭合

            # 绘图之前先清理一下
            plt.clf()
            # polar表示绘制极坐标图，颜色，线宽，标志点样式
            plt.polar(theta1, r, color="blue", linewidth=1, marker="", mfc="b", ms=10)
            plt.gca().set_theta_zero_location('N')
            # plt.gca().set_rlim(0, 57.15)  # 设置显示的极径范围
            plt.gca().set_rlim(0, float(self.outer_diameter.value) / 2)
            plt.gca().fill(theta1, r, facecolor='w', alpha=0.2)  # 填充颜色
            plt.gca().fill_between(theta2, outside_radius, inside_radius, facecolor='cyan', alpha=0.8)  # 填充之间颜色
            plt.gca().patch.set_facecolor('1')
            plt.gca().set_rgrids(np.arange(0, 100, 100))

            # label = np.array([j for j in range(1, 25)])  # 定义标签
            # plt.gca().set_thetagrids(np.arange(0, 360, 15), label)
            plt.gca().set_thetagrids(np.arange(0, 360, 360))

            text_show = ''.join(['Depth:', str(ymax), 'm\nOD: ', str(self.outer_diameter.value), 'mm\nID: ', \
                                 str(self.inner_diameter.value),'mm\nWT: ', str(self.thickness.value), 'mm'])
            plt.text(7 * np.pi / 4, 90, text_show)
            # 绘图展示
            plt.show()

    # RadioButtons行为定义
    def actionfunc(self, damage_type):

        # 损伤SpanSelector对象
        self.span3_red = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
                                      rectprops=dict(alpha=0.5, facecolor='red'), span_stays=True)
        # 结垢SpanSelector对象
        self.span3_green = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
                                        rectprops=dict(alpha=0.5, facecolor='green'), span_stays=True)
        # 变形SpanSelector对象
        self.span3_yellow = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
                                         rectprops=dict(alpha=0.5, facecolor='yellow'), span_stays=True)
        if damage_type == 'Penetration':
            print(damage_type)
            self.damage_Tag = 'Penetration'
            # del self.span3_red
            del self.span3_green
            del self.span3_yellow
        elif damage_type == 'Projection':
            print(damage_type)
            self.damage_Tag = 'Projection'
            del self.span3_red
            # del self.span3_green
            del self.span3_yellow
        elif damage_type == 'Transformation':
            print(damage_type)
            self.damage_Tag = 'Transformation'
            del self.span3_red
            del self.span3_green
            # del self.span3_yellow
