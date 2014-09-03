#! /usr/bin/env python
# -*- coding: utf-8 -*-

import fnmatch
import logging
import numpy
import os
import xlrd
from xlwt import Utils
import matplotlib.mlab as mlab
import matplotlib.pyplot as plt

log = logging.getLogger(__name__)


class Histogram(object):

    DATA_CELL_RANGE = 'A1:A10000'

    def __init__(self, data=None, csv_name=None,
        xls_name=None, sheet_name=None, data_start_cell='A1', data_end_cell='A10000', skip_empty=True):
        self.data = data
        self.csv_name = csv_name
        self.xls_name = xls_name
        self.sheet_name = sheet_name

        if self.data is None:
            if self.xls_name is None:
                files = fnmatch.filter(os.listdir('.'), '*.xls')
                #files = [f for f in os.listdir('.') if fnmatch.fnmatch(f, '*.xls|*.xlsx')]
                if len(files) > 0:
                    print('@@@ files=%s' % files)
                    self.xls_name = files[0]
                else:
                    raise Exception('aaaa')
            print('@@@ self.xls_name=%s' % self.xls_name)
            book = xlrd.open_workbook(self.xls_name)

            if self.sheet_name is None:
                sheet = book.sheet_by_index(0)
            else:
                sheet = book.sheet_by_name('Statistics (total score)')
            print '@@@ sheet.nrows=%s' % sheet.nrows
            print '@@@ sheet.ncols=%s' % sheet.ncols

            start_row, start_col, end_row, end_col = Utils.cellrange_to_rowcol_pair(self.DATA_CELL_RANGE)
            print '@@@ start_row=%s' % start_row
            print '@@@ start_col=%s' % start_col
            print '@@@ end_row=%s' % end_row
            print '@@@ end_col=%s' % end_col
            data = []
            for row in range(start_row, min(end_row + 1, sheet.nrows)):
                for col in range(start_col, min(end_col + 1, sheet.ncols)):
                    value = sheet.cell(row, col).value
                    if value != '':
                        data.append(value)
            print '@@@ data=%s' % data

#            start_cell = Histogram.cell_to_rowcol(self.DATA_START_CELL)
#            print '@@@ sheet.nrows=%s' % sheet.nrows
#            #print '@@@ col_slice=%s' % sheet.col_slice(0, 1, sheet.nrows) # emptyが入っちゃう。。
#            col_slices = sheet.col_slice(start_cell[0], start_cell[1], sheet.nrows)
#            print '@@@ col_slices=%s' % col_slices
#            print '@@@ len(col_slices)=%s' % len(col_slices)
#            col_slices = [cell.value for cell in col_slices if cell.value != '']
#            print '@@@ col_slices=%s' % col_slices
#            print '@@@ len(col_slices)=%s' % len(col_slices)

        self.data = numpy.array(data)
        print 'data=%s' % self.data
        print 'size of data=%d' % len(self.data)
        self.mu = numpy.mean(self.data)
        print 'mean value=%.1f' % self.mu
        print 'median value=%.1f' % numpy.median(self.data)
        #self.sigma = numpy.std(self.data)
        #print 'standard deviation=%.2f' % self.sigma

    @property
    def sigma(self):
        return numpy.std(self.data)

    def draw(self):
        # 共通初期設定
        plt.rc('font', **{'family': 'serif'})
        # キャンバス
        fig = plt.figure()
        # プロット領域（1x1分割の1番目に領域を配置せよという意味）
        ax = fig.add_subplot(111)
        # ヒストグラム（normedをTrueで指定すると確率表示になる）
        ax.hist(self.data, bins=25, range=(0, 50), normed=False, facecolor='#0066cc', alpha=0.8) #2079b4
        # X, Y方向の表示範囲
        ax.set_xlim(0, 50)
        ax.set_ylim(0, 20)
        # 平均値
        ave = ax.axvline(x=self.mu, linewidth=1, color='#ff9900', linestyle='--', zorder=0)
        # タイトル
        ax.set_title('Histogram of Scores', size=16)
        ax.set_xlabel('Score', size=14)
        ax.set_ylabel('Frequency', size=14)
        # 凡例
        ax.legend((ave,), ('Ave.',), scatterpoints=1, loc='upper right', fontsize=12)
        # 任意のテキスト
        ax.annotate('(N=%d)' % len(self.data),
                    xy=(0.98, 1.01), xycoords='axes fraction',
                    horizontalalignment='right', verticalalignment='bottom', fontsize=12)
        # グリッド表示
        ax.grid(True)
        #file_name = os.path.splitext(os.path.basename(__file__))[0]
        file_name = 'images/%s.png' % os.path.splitext(os.path.basename(__file__))[0]
        print('@@@ file_name=%s' % file_name)
        print('@@@ dirname=%s' % os.path.dirname(file_name))
        if not os.path.exists(os.path.dirname(file_name)):
            os.makedirs(os.path.dirname(file_name))
        plt.savefig(file_name)
        plt.show()

    @staticmethod
    def cell_to_rowcol(cell):
        return Utils.cell_to_rowcol2('A1')


if __name__ == '__main__':
    histogram = Histogram()
    histogram.draw()
