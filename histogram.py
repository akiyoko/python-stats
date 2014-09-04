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

from stats_base import StatsBase

log = logging.getLogger(__name__)


class StatsHistogram(StatsBase):

    X_LIMIT = 50
    Y_LIMIT = 50
    BINS = 10

    def draw(self, show_title=True, show_labels=True, show_numbers=True, show_ave=True, show_legend=True, show_grid=True):
        self.ax.hist(self.data, bins=self.BINS, range=(0, self.X_LIMIT), facecolor='#0066cc', alpha=0.8)
        # X, Y方向の表示範囲
        self.ax.set_xlim(0, self.X_LIMIT)
        self.ax.set_ylim(0, self.Y_LIMIT)

        # タイトル
        if show_title:
            self.set_title()
        # ラベル
        if show_labels:
            self.set_labels()
        # N数
        if show_numbers:
            self.ax.annotate(
                '(N=%d)' % self.data.size,
                xy=(0.98, 1.01),
                xycoords='axes fraction',
                horizontalalignment='right',
                verticalalignment='bottom',
                fontsize=12
            )
        # 平均値
        if show_ave:
            ave = self.ax.axvline(x=self.data.mean, linewidth=1, color='#ff9900', linestyle='--', zorder=0)
            self.ax.legend((ave,), ('Ave.',), scatterpoints=1, loc='upper right', fontsize=12)
        # 凡例
        if show_legend:
            #TODO
            pass
        # グリッド表示
        if show_grid:
            self.ax.grid(True)
        # 画像を保存
        self.save_image()
        plt.show()


if __name__ == '__main__':
    histogram = StatsHistogram(
        title='Histogram of Scores', xlabel='Score', ylabel='Frequency')
#    histogram.Y_LIMIT = 50
#    histogram.BINS = 10
    histogram.draw()
