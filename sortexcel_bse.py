import pandas as pd
import requests
import os
import shutil

url = 'https://www.bse.cn/projectNewsController/excelExport.do?companyCode&statetypes&sortfield=updateDate&sorttype=desc&startDate&endDate&keyword'

r = requests.get(url)
if r.status_code==200:
    newpath = r'C:/北交所ipo'
    if os.path.exists(newpath):
        shutil.rmtree(newpath)
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    filename = 'C:/北交所ipo/原初数据.xls'
    f = open(filename, 'wb')
    f.write(r.content)
    f.close()

    df = pd.read_excel(filename)

    reorderlist = ["上市委会议通过","上市委会议暂缓","提交注册","已受理","已问询","中止","终止","上市委会议未通过","终止注册","不予注册","注册",'核准']

    def sorter(column):
        reorder = reorderlist
        cat = pd.Categorical(column, categories=reorder, ordered=True)
        return pd.Series(cat)

    df=df.sort_values(['审核状态'], ascending=True, key=sorter)
    df['序号'] = df.reset_index().index + 1
    col1=df.pop('序号')
    df.insert(0,'序号',col1)
    df['审核状态'] = pd.Categorical(df['审核状态'], categories=reorderlist)
    cat1=int(df['审核状态'].value_counts()['上市委会议通过']+df['审核状态'].value_counts()['上市委会议暂缓']+df['审核状态'].value_counts()['提交注册']+df['审核状态'].value_counts()['已受理']+df['审核状态'].value_counts()['已问询']+df['审核状态'].value_counts()['中止'])
    cat11 = int(df['审核状态'].value_counts()['上市委会议通过'])
    cat12 = int(df['审核状态'].value_counts()['上市委会议暂缓'])
    cat13 = int(df['审核状态'].value_counts()['提交注册'])
    cat14 = int(df['审核状态'].value_counts()['已受理'])
    cat15 = int(df['审核状态'].value_counts()['已问询'])
    cat16 = int(df['审核状态'].value_counts()['中止'])
    cat2=int(df['审核状态'].value_counts()['终止']+df['审核状态'].value_counts()['终止注册']+df['审核状态'].value_counts()['不予注册']+df['审核状态'].value_counts()["上市委会议未通过"])
    cat21 = int(df['审核状态'].value_counts()['终止'])
    cat22 = int(df['审核状态'].value_counts()['终止注册'])
    cat23 = int(df['审核状态'].value_counts()['不予注册'])
    cat24 = int(df['审核状态'].value_counts()['上市委会议未通过'])
    cat3=int(df['审核状态'].value_counts()['注册']+df['审核状态'].value_counts()['核准'])
    cat31 = int(df['审核状态'].value_counts()['注册'])
    cat32 = int(df['审核状态'].value_counts()['核准'])

    new_col=['','','','','','','','','','']
    if cat1+cat2+cat3==0:
        new_col1 = ['', '', '主板：' + str(cat1 + cat2 + cat3) , '',
                      '', '', '', '', '', '', '']
    else:
        new_col1=['','','主板：'+str(cat1+cat2+cat3)+'（第'+str(1)+'号到第'+str(cat1+cat2+cat3)+'号）','','','','','','','']
    if cat1==0:
        new_col2= ['', '', '排队审核与待发行：' + str(cat1) , '', '', '', '', '', '', '']
    else:
        new_col2=['','','排队审核与待发行：'+str(cat1)+'（第'+str(1)+'号到第'+str(cat1)+'号）','','','','','','','']
    if cat2==0:
        new_col3 = ['', '', '终止上市：' + str(cat2) , '', '', '', '',
                     '', '', '']
    else:
        new_col3=['','','终止上市：'+str(cat2)+'（第'+str(cat1+1)+'号到第'+str(cat1+cat2)+'号）','','','','','','','']
    if cat3==0:
        new_col4 = ['', '', '注册生效：' + str(cat3) , '',
                     '', '', '', '', '', '']
    else:
        new_col4=['','','注册生效：'+str(cat3)+'（第'+str(cat1+cat2+1)+'号到第'+str(cat1+cat2+cat3)+'号）','','','','','','','']

    df.loc[len(df.index)] = new_col
    df.loc[len(df.index)] = new_col1
    df.loc[len(df.index)] = new_col
    df.loc[len(df.index)] = new_col2
    df.loc[len(df.index)] = ['', '', '上市委会议通过：' + str(cat11), '', '', '', '', '', '', '']
    df.loc[len(df.index)] = ['', '', '上市委会议暂缓：' + str(cat12), '', '', '', '', '', '', '']
    df.loc[len(df.index)] = ['', '', '提交注册：' + str(cat13), '', '', '', '', '', '', '']
    df.loc[len(df.index)] = ['', '', '已受理：' + str(cat14), '', '', '', '', '', '', '']
    df.loc[len(df.index)] = ['', '', '已问询：' + str(cat15), '', '', '', '', '', '', '']
    df.loc[len(df.index)] = ['', '', '中止：' + str(cat16), '', '', '', '', '', '', '']
    df.loc[len(df.index)] = ['', '', '', '', '', '', '', '', '', '']
    df.loc[len(df.index)] = new_col3
    df.loc[len(df.index)] = ['', '', '终止：' + str(cat21), '', '', '', '', '', '', '']
    df.loc[len(df.index)] = ['', '', '', '', '', '', '', '', '', '']
    df.loc[len(df.index)] = new_col4
    df.loc[len(df.index)] = ['', '', '注册：' + str(cat31), '', '', '', '', '', '', '']
    df.loc[len(df.index)] = ['', '', '核准：' + str(cat32), '', '', '', '', '', '', '']

    df.to_csv('C:/北交所ipo/北交所整理数据.csv',sep=",",index=False,encoding='utf-16')
else:
    print("bse download failed!")
