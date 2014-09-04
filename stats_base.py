#! /usr/bin/env python
# -*- coding: utf-8 -*-

import fnmatch
import inspect
import logging
import numpy
import os
import xlrd
from xlwt import Utils
import matplotlib.mlab as mlab
import matplotlib.pyplot as plt

log = logging.getLogger(__name__)


class StatsColor(object):
    pass


class StatsBase(object):
    pass


class StatsData(numpy.ndarray):
    '''
    property
        size
        sum
        mu
        median
        sigma
    '''
    def __new__(cls, input_array, legend=None):
        if isinstance(input_array, numpy.ndarray):
            return input_array
        return numpy.asarray(input_array).view(cls)

    @property
    def size(self):
        return super(StatsData, self).size

    @property
    def sum(self):
        return super(StatsData, self).sum()

    @property
    def mean(self):
        return super(StatsData, self).mean()

    #TODO
    #@property
    #def median(self):
    #    return numpy.median(self)

    @property
    def std(self):
        return super(StatsData, self).std()


class StatsBase(object):

#    TITLE = 'Title'
#    X_LABEL = 'Label X'
#    Y_LABEL = 'Label Y'

    def __init__(
        self,
        data=None,
        csv_name=None,
        xls_name=None, sheet_name=None, data_start_cell='A1', data_end_cell='A10000', skip_empty=True,
        title=None, xlabel=None, ylabel=None,
    ):
        self.title = title
        self.xlabel = xlabel
        self.ylabel = ylabel

        if data is None and csv_name is None and xls_name is None:
            files = fnmatch.filter(os.listdir('.'), '*.xls')
            #files = [f for f in os.listdir('.') if fnmatch.fnmatch(f, '*.xls|*.xlsx')]
            if len(files) > 0:
                print('@@@ files=%s' % files)
                xls_name = files[0]
            else:
                raise Exception('aaaa')
            print('@@@ xls_name=%s' % xls_name)

        if data is not None:
            self.data = data
            self.data2 = data2
            self.data3 = data3
        elif csv_name is not None:
            # TODO
            pass
        elif xls_name is not None:
            book = xlrd.open_workbook(xls_name)

            if sheet_name is None:
                sheet = book.sheet_by_index(0)
            else:
                sheet = book.sheet_by_name('Statistics (total score)')
            print '@@@ sheet.nrows=%s' % sheet.nrows
            print '@@@ sheet.ncols=%s' % sheet.ncols

            start_row, start_col = self.to_rowcol(data_start_cell)
            end_row, end_col = self.to_rowcol(data_end_cell)
            #start_row, start_col, end_row, end_col = Utils.cellrange_to_rowcol_pair(self.DATA_CELL_RANGE)
            print '@@@ start_row=%s' % start_row
            print '@@@ start_col=%s' % start_col
            print '@@@ end_row=%s' % end_row
            print '@@@ end_col=%s' % end_col
            datas = []
            for col in range(start_col, min(end_col + 1, sheet.ncols)):
                data = []
                for row in range(start_row, min(end_row + 1, sheet.nrows)):
                    value = sheet.cell(row, col).value
                    if value != '':
                        data.append(value)
                datas.append(data)
#                print '@@@ data=%s' % data
            print '@@@ datas=%s' % datas

            self.data = StatsData(datas[0]) if len(datas) > 0 else None
            self.data2 = StatsData(datas[1]) if len(datas) > 1 else None
            self.data3 = StatsData(datas[2]) if len(datas) > 2 else None

#            start_cell = Histogram.cell_to_rowcol(self.DATA_START_CELL)
#            print '@@@ sheet.nrows=%s' % sheet.nrows
#            #print '@@@ col_slice=%s' % sheet.col_slice(0, 1, sheet.nrows) # emptyが入っちゃう。。
#            col_slices = sheet.col_slice(start_cell[0], start_cell[1], sheet.nrows)
#            print '@@@ col_slices=%s' % col_slices
#            print '@@@ len(col_slices)=%s' % len(col_slices)
#            col_slices = [cell.value for cell in col_slices if cell.value != '']
#            print '@@@ col_slices=%s' % col_slices
#            print '@@@ len(col_slices)=%s' % len(col_slices)
        else:
            raise Exception()

        #self.data = StatsData(data)
        print 'data=%s' % self.data
        print 'dir(data)=%s' % dir(self.data)
        print 'size of data=%d' % self.data.size
        print 'sum total=%.1f' % self.data.sum
        print 'mean value=%.1f' % self.data.mean
        #print 'median value=%.1f' % self.data.median
        print 'standard deviation(sigma)=%.2f' % self.data.std

        # 共通初期設定
        plt.rc('font', **{'family': 'serif'})
        # キャンバス
        fig = plt.figure()
        # プロット領域（1x1分割の1番目に領域を配置せよという意味）
        self.ax = fig.add_subplot(111)

    def set_title(self, title=None):
        self.ax.set_title(title or self.title or '', size=16)

    def set_labels(self, xlabel=None, ylabel=None):
        self.ax.set_xlabel(xlabel or self.xlabel or '', size=14)
        self.ax.set_ylabel(ylabel or self.ylabel or '', size=14)

    def save_image(self, filename=None):
        if not filename:
            co_filename = inspect.currentframe().f_back.f_code.co_filename
            print('@@@ co_filename=%s' % co_filename)
            name, ext = os.path.splitext(os.path.basename(co_filename))
            filename = 'images/%s.png' % name
        print('@@@ filename=%s' % filename)
        if not os.path.exists(os.path.dirname(filename)):
            os.makedirs(os.path.dirname(filename))
        plt.savefig(filename)

    def to_rowcol(self, cell):
        return Utils.cell_to_rowcol2(cell)


if __name__ == '__main__':
    pass
