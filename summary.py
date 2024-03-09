import sortexcel_szse
import sortexcel_bse
import sortexcel_sse
import sse_zrz
import szse_zrz
import sse_bgcz
import szse_bgcz
import pandas as pd
import os
import shutil

newpath = r'C:/整合'
if os.path.exists(newpath):
    shutil.rmtree(newpath)
if not os.path.exists(newpath):
    os.makedirs(newpath)

data1 = {'IPO项目': ['排队', '终止', '注册生效', '合计'],
         '上海主板': [sortexcel_sse.cat1, sortexcel_sse.cat2, sortexcel_sse.cat3,
                  sortexcel_sse.cat1 + sortexcel_sse.cat2 + sortexcel_sse.cat3],
         '科创板': [sortexcel_sse.cat4, sortexcel_sse.cat5, sortexcel_sse.cat6,
                 sortexcel_sse.cat6 + sortexcel_sse.cat5 + sortexcel_sse.cat4],
         '深圳主板': [sortexcel_szse.cat1, sortexcel_szse.cat2, sortexcel_szse.cat3,
                  sortexcel_szse.cat1 + sortexcel_szse.cat2 + sortexcel_szse.cat3],
         '创业板': [sortexcel_szse.cat4, sortexcel_szse.cat5, sortexcel_szse.cat6,
                 sortexcel_szse.cat6 + sortexcel_szse.cat5 + sortexcel_szse.cat4],
         '北交所': [sortexcel_bse.cat1, sortexcel_bse.cat2, sortexcel_bse.cat3,
                 sortexcel_bse.cat1 + sortexcel_bse.cat2 + sortexcel_bse.cat3]}

data2 = {'再融资项目': ['排队','终止','注册生效','合计'],
        '上海主板': [sse_zrz.cat1,sse_zrz.cat2,sse_zrz.cat3,sse_zrz.cat1+sse_zrz.cat2+sse_zrz.cat3],
        '科创板':[sse_zrz.cat4,sse_zrz.cat5,sse_zrz.cat6,sse_zrz.cat6+sse_zrz.cat5+sse_zrz.cat4],
        '深圳主板':[szse_zrz.cat1,szse_zrz.cat2,szse_zrz.cat3,szse_zrz.cat1+szse_zrz.cat2+szse_zrz.cat3],
        '创业板':[szse_zrz.cat4,szse_zrz.cat5,szse_zrz.cat6,szse_zrz.cat6+szse_zrz.cat5+szse_zrz.cat4],
        '北交所':['','','','']}

data3 = {'并购重组项目': ['排队','终止','注册生效','合计'],
        '上海主板': [sse_bgcz.cat1,sse_bgcz.cat2,sse_bgcz.cat3,sse_bgcz.cat1+sse_bgcz.cat2+sse_bgcz.cat3],
        '科创板':[sse_bgcz.cat4,sse_bgcz.cat5,sse_bgcz.cat6,sse_bgcz.cat6+sse_bgcz.cat5+sse_bgcz.cat4],
        '深圳主板':[szse_bgcz.cat1,szse_bgcz.cat2,szse_bgcz.cat3,szse_bgcz.cat1+szse_bgcz.cat2+szse_bgcz.cat3],
        '创业板':[szse_bgcz.cat4,szse_bgcz.cat5,szse_bgcz.cat6,szse_bgcz.cat6+szse_bgcz.cat5+szse_bgcz.cat4],
        '北交所':['','','','']}

df1 = pd.DataFrame(data1)
df2 = pd.DataFrame(data2)
df3 = pd.DataFrame(data3)

df1.to_csv('C:/整合/IPO整合.csv',sep=",",index=False,encoding='utf-16')
df2.to_csv('C:/整合/再融资整合.csv',sep=",",index=False,encoding='utf-16')
df3.to_csv('C:/整合/并购重组整合.csv',sep=",",index=False,encoding='utf-16')

