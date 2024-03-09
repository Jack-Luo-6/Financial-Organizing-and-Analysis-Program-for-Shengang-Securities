import pandas as pd
import sortexcel_sse
import sortexcel_szse
import sortexcel_bse
import sse_zrz
import sse_bgcz
import szse_zrz
import szse_bgcz
import summary
import os

def excel_create(rootpath):
    filename = os.path.join(rootpath, "download/综合数据.xlsx")
    with pd.ExcelWriter(filename) as writer:
        summary.df1.to_excel(writer, sheet_name='IPO项目整理')
        summary.df2.to_excel(writer, sheet_name='再融资项目整理')
        summary.df3.to_excel(writer, sheet_name='并购重组项目整理')
        sortexcel_bse.df.to_excel(writer, sheet_name='北交所IPO')
        sortexcel_szse.df.to_excel(writer, sheet_name='深交所IPO')
        szse_zrz.df.to_excel(writer, sheet_name='深交所再融资')
        szse_bgcz.df.to_excel(writer, sheet_name='深交所并购重组')
        sortexcel_sse.df.to_excel(writer, sheet_name='上交所IPO')
        sse_zrz.df.to_excel(writer, sheet_name='上交所再融资')
        sse_bgcz.df.to_excel(writer, sheet_name='上交所并购重组')
