# -*- coding: utf-8 -*-
"""
Created on Sun Jun 10 21:25:44 2018

@author: zelord
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from fbprophet import Prophet
from pandas import DataFrame
"""
g_dates = [20180608, 20180607, 20180605, 20180604, 20180601, 20180531, 20180530, 20180529, 20180528, 20180525, 20180524, 20180523, 20180521, 20180518, 20180517, 20180516, 20180515, 20180514, 20180511, 20180510, 20180509, 20180508, 20180504, 20180503, 20180502, 20180430, 20180427, 20180426, 20180425, 20180424, 20180423, 20180420, 20180419, 20180418, 20180417, 20180416, 20180413, 20180412, 20180411, 20180410, 20180409, 20180406, 20180405, 20180404, 20180403, 20180402, 20180330, 20180329, 20180328, 20180327, 20180326, 20180323, 20180322, 20180321, 20180320, 20180319, 20180316, 20180315, 20180314, 20180313, 20180312, 20180309, 20180308, 20180307, 20180306, 20180305, 20180302, 20180228, 20180227, 20180226, 20180223, 20180222, 20180221, 20180220, 20180219, 20180214, 20180213, 20180212, 20180209, 20180208, 20180207, 20180206, 20180205, 20180202, 20180201, 20180131, 20180130, 20180129, 20180126, 20180125, 20180124, 20180123, 20180122, 20180119, 20180118, 20180117, 20180116, 20180115, 20180112, 20180111, 20180110, 20180109, 20180108, 20180105, 20180104, 20180103, 20180102, 20171228, 20171227, 20171226, 20171222, 20171221, 20171220, 20171219, 20171218, 20171215, 20171214, 20171213, 20171212, 20171211, 20171208, 20171207, 20171206, 20171205, 20171204, 20171201, 20171130, 20171129, 20171128, 20171127, 20171124, 20171123, 20171122, 20171121, 20171120, 20171117, 20171116, 20171115, 20171114, 20171113, 20171110, 20171109, 20171108, 20171107, 20171106, 20171103, 20171102, 20171101, 20171031, 20171030, 20171027, 20171026, 20171025, 20171024, 20171023, 20171020, 20171019, 20171018, 20171017, 20171016, 20171013, 20171012, 20171011, 20171010, 20170929, 20170928, 20170927, 20170926, 20170925, 20170922, 20170921, 20170920, 20170919, 20170918, 20170915, 20170914, 20170913, 20170912, 20170911, 20170908, 20170907, 20170906, 20170905, 20170904, 20170901, 20170831, 20170830, 20170829, 20170828, 20170825, 20170824, 20170823, 20170822, 20170821, 20170818, 20170817, 20170816, 20170814, 20170811, 20170810, 20170809, 20170808, 20170807, 20170804, 20170803, 20170802, 20170801, 20170731, 20170728, 20170727, 20170726, 20170725, 20170724, 20170721, 20170720, 20170719, 20170718, 20170717, 20170714, 20170713, 20170712, 20170711, 20170710, 20170707, 20170706, 20170705, 20170704, 20170703, 20170630, 20170629, 20170628, 20170627, 20170626, 20170623, 20170622, 20170621, 20170620, 20170619, 20170616, 20170615, 20170614, 20170613, 20170612, 20170609, 20170608, 20170607, 20170605, 20170602, 20170601, 20170531, 20170530, 20170529]
g_closes = [49650, 50600, 51300, 51100, 51300, 50700, 49500, 51300, 52300, 52700, 51400, 51800, 50000, 49500, 49400, 49850, 49200, 50100, 51300, 51600, 50900, 52600, 51900, 53000, 53000, 53000, 53000, 52140, 50400, 50460, 51900, 51620, 52780, 51360, 49980, 50340, 49800, 49000, 48860, 48880, 49200, 48400, 48740, 46920, 48120, 48540, 49220, 49040, 48700, 49980, 50280, 49720, 51780, 51060, 51200, 50740, 51140, 51540, 51760, 51660, 49740, 49740, 49200, 48620, 47020, 45200, 46020, 47060, 47380, 47380, 47220, 46760, 47280, 47400, 48380, 49000, 47540, 45720, 44700, 46000, 45800, 47420, 47920, 47700, 49820, 49900, 49800, 51220, 50780, 50260, 49340, 49160, 48240, 49320, 49900, 49620, 50000, 48540, 48200, 48240, 48840, 50400, 52020, 52120, 51080, 51620, 51020, 50960, 49360, 48200, 49700, 49140, 50880, 51560, 51200, 50620, 51060, 51320, 52100, 51780, 52000, 50740, 50020, 51260, 51340, 50840, 50800, 52600, 53280, 52640, 55460, 55300, 55960, 55280, 55200, 55820, 55780, 55340, 55920, 56380, 56400, 56340, 56760, 56100, 56380, 56380, 57060, 57220, 55080, 54040, 53080, 52400, 53900, 54040, 54300, 53840, 52980, 54760, 54800, 53920, 54000, 54800, 54640, 52800, 51280, 51260, 51680, 51660, 53620, 53000, 52800, 52220, 52120, 52480, 50400, 50300, 49620, 49600, 49800, 49080, 48120, 47000, 46760, 46040, 46480, 46320, 46200, 46080, 46100, 47020, 47520, 47480, 47000, 46840, 46900, 47040, 46200, 45000, 44620, 45900, 46280, 47720, 47580, 47700, 47780, 49000, 48600, 48200, 47760, 49800, 49840, 50000, 50860, 51080, 51200, 50740, 50840, 50640, 50480, 50560, 49880, 49000, 48660, 47860, 48060, 47580, 47000, 47220, 47540, 47940, 47700, 48300, 48280, 47620, 47960, 47480, 48140, 46560, 45580, 45680, 45360, 45400, 45380, 46100, 45160, 45300, 45940, 45960, 44680, 44700, 44640, 45620]
new_dates = []


for i in range(len(g_dates)):
    yyyy = int(g_dates[i] / 10000)
    mm = int(g_dates[i] - (yyyy * 10000))
    dd = mm % 100
    mm = mm / 100
    val = '%04d-%02d-%02d' %(yyyy, mm, dd)
    new_dates.append(val)
    
print(new_dates)
"""

market_df = pd.read_csv('SP500.csv', index_col='DATE', parse_dates=True)
#market_df.plot()