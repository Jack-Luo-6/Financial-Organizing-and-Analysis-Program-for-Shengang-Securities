import pandas as pd
import requests
import os
import shutil

url = 'http://listing.szse.cn/api/ras/projectrends/export?bizType=3&random=0.41041080966625076'
r = requests.get(url, allow_redirects=True)

newpath = r'C:/深交所并购重组'
if os.path.exists(newpath):
    shutil.rmtree(newpath)
if not os.path.exists(newpath):
    os.makedirs(newpath)

filename = 'C:/深交所并购重组/深交所原初数据.xls'
open(filename, 'wb').write(r.content)

df = pd.read_excel(filename)

df=df.sort_values(['板块'], ascending=True)
c=int(df['板块'].value_counts()['主板'])
df1 = df.iloc[:c,:]
df2 = df.iloc[c:,:]

reorderlist = ["并购重组委会议通过","审核通过","提交注册","新受理","已受理","已问询",'已回复',"暂缓审议","中止","终止(撤回)","终止(审核不通过)","终止(未在规定时限内回复)","终止注册","不予注册","注册生效"]

def sorter(column):
    reorder = reorderlist
    cat = pd.Categorical(column, categories=reorder, ordered=True)
    return pd.Series(cat)

df1=df1.sort_values(['审核状态'], ascending=True, key=sorter)
df1['序号'] = df1.reset_index().index + 1
col1=df1.pop('序号')
df1.insert(0,'序号',col1)

df1['审核状态'] = pd.Categorical(df1['审核状态'], categories=reorderlist,ordered=True)
cat1=int(df1['审核状态'].value_counts()['并购重组委会议通过']+df1['审核状态'].value_counts()['审核通过']+df1['审核状态'].value_counts()['提交注册']+df1['审核状态'].value_counts()['已受理']+df1['审核状态'].value_counts()['已问询']+df1['审核状态'].value_counts()['已回复']+df1['审核状态'].value_counts()['中止']+df1['审核状态'].value_counts()['新受理'])
cat11=int(df1['审核状态'].value_counts()['并购重组委会议通过'])
cat12=int(df1['审核状态'].value_counts()['审核通过'])
cat13=int(df1['审核状态'].value_counts()['提交注册'])
cat14=int(df1['审核状态'].value_counts()['已受理'])
cat15=int(df1['审核状态'].value_counts()['已问询'])
cat16=int(df1['审核状态'].value_counts()['已回复'])
cat17=int(df1['审核状态'].value_counts()['中止'])
cat18=int(df1['审核状态'].value_counts()['新受理'])
cat2=int(df1['审核状态'].value_counts()['终止(撤回)']+df1['审核状态'].value_counts()['终止注册']+df1['审核状态'].value_counts()['终止(未在规定时限内回复)']+df1['审核状态'].value_counts()['不予注册']+df1['审核状态'].value_counts()['终止(审核不通过)'])
cat21=int(df1['审核状态'].value_counts()['终止(撤回)'])
cat22=int(df1['审核状态'].value_counts()['终止注册'])
cat23=int(df1['审核状态'].value_counts()['终止(未在规定时限内回复)'])
cat24=int(df1['审核状态'].value_counts()['不予注册'])
cat25=int(df1['审核状态'].value_counts()['终止(审核不通过)'])
cat3=int(df1['审核状态'].value_counts()['注册生效'])
cat31=int(df1['审核状态'].value_counts()['注册生效'])

new_col10=['','','','','','','','','','','']
if cat1+cat2+cat3==0:
    new_col100 = ['', '', '主板：' + str(cat1 + cat2 + cat3) , '',
                  '', '', '', '', '', '', '']
else:
    new_col100=['','','主板：'+str(cat1+cat2+cat3)+'（第'+str(1)+'号到第'+str(cat1+cat2+cat3)+'号）','','','','','','','','']
if cat1==0:
    new_col11 = ['', '', '排队审核与待发行：' + str(cat1) , '', '', '', '', '', '', '',
                 '']
else:
    new_col11=['','','排队审核与待发行：'+str(cat1)+'（第'+str(1)+'号到第'+str(cat1)+'号）','','','','','','','','']
if cat2==0:
    new_col12 = ['', '', '终止上市：' + str(cat2) , '', '', '', '',
                 '', '', '', '']
else:
    new_col12=['','','终止上市：'+str(cat2)+'（第'+str(cat1+1)+'号到第'+str(cat1+cat2)+'号）','','','','','','','','']
if cat3==0:
    new_col13 = ['', '', '注册生效：' + str(cat3) , '',
                 '', '', '', '', '', '', '']
else:
    new_col13=['','','注册生效：'+str(cat3)+'（第'+str(cat1+cat2+1)+'号到第'+str(cat1+cat2+cat3)+'号）','','','','','','','','']

df2=df2.sort_values(['审核状态'], ascending=True, key=sorter)
df2['序号'] = df2.reset_index().index + 1
col1=df2.pop('序号')
df2.insert(0,'序号',col1)

df2['审核状态'] = pd.Categorical(df2['审核状态'], categories=reorderlist,ordered=True)
cat4=int(df2['审核状态'].value_counts()['提交注册']+df2['审核状态'].value_counts()['已问询']+df2['审核状态'].value_counts()['已回复']+df2['审核状态'].value_counts()['中止'])
cat41=int(df2['审核状态'].value_counts()['提交注册'])
cat42=int(df2['审核状态'].value_counts()['已问询'])
cat43=int(df2['审核状态'].value_counts()['已回复'])
cat44=int(df2['审核状态'].value_counts()['中止'])
cat5=int(df2['审核状态'].value_counts()['终止(撤回)']+df2['审核状态'].value_counts()['终止(审核不通过)'])
cat51=int(df2['审核状态'].value_counts()['终止(撤回)'])
cat52=int(df2['审核状态'].value_counts()['终止(审核不通过)'])
cat6=int(df2['审核状态'].value_counts()['注册生效'])
cat61=int(df2['审核状态'].value_counts()['注册生效'])

new_col20=['','','','','','','','','','','']
if cat4+cat5+cat6==0:
    new_col200 = ['', '', '创业板：' + str(cat4 + cat5 + cat6) , '',
                  '', '', '', '', '', '', '']
else:
    new_col200=['','','创业板：'+str(cat4+cat5+cat6)+'（第'+str(1)+'号到第'+str(cat4+cat5+cat6)+'号）','','','','','','','','']
if cat4==0:
    new_col21 = ['', '', '排队审核与待发行：' + str(cat4), '', '', '', '', '', '', '','']
else:
    new_col21=['','','排队审核与待发行：'+str(cat4)+'（第'+str(1)+'号到第'+str(cat4)+'号）','','','','','','','','']
if cat5==0:
    new_col22 = ['', '', '终止上市：' + str(cat5) , '', '', '', '',
                 '', '', '', '']
else:
    new_col22=['','','终止上市：'+str(cat5)+'（第'+str(cat4+1)+'号到第'+str(cat4+cat5)+'号）','','','','','','','','']
if cat6==0:
    new_col23 = ['', '', '注册生效：' + str(cat6) , '',
                 '', '', '', '', '', '', '']
else:
    new_col23=['','','注册生效：'+str(cat6)+'（第'+str(cat4+cat5+1)+'号到第'+str(cat4+cat5+cat6)+'号）','','','','','','','','']

list=[df1,df2]
df=pd.concat(list)

df.loc[len(df.index)] = new_col10
df.loc[len(df.index)] = new_col100
df.loc[len(df.index)] = new_col10
df.loc[len(df.index)] = new_col11
df.loc[len(df.index)] = ['', '', '提交注册：' + str(cat13) , '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = ['', '', '已问询：' + str(cat15) , '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = ['', '', '中止：' + str(cat17) , '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = ['', '', '', '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = new_col12
df.loc[len(df.index)] = ['', '', '终止（撤回）：' + str(cat21) , '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = ['', '', '', '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = new_col13
df.loc[len(df.index)] = ['', '', '注册生效：' + str(cat31) , '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = new_col20
df.loc[len(df.index)] = new_col20
df.loc[len(df.index)] = new_col20
df.loc[len(df.index)] = new_col200
df.loc[len(df.index)] = new_col20
df.loc[len(df.index)] = new_col21
df.loc[len(df.index)] = ['', '', '中止：' + str(cat44) , '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = ['', '', '', '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = new_col22
df.loc[len(df.index)] = ['', '', '终止（撤回）：' + str(cat51) , '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = ['', '', '终止（审核不通过）：' + str(cat52) , '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = ['', '', '', '', '', '', '', '', '', '', '']
df.loc[len(df.index)] = new_col23
df.loc[len(df.index)] = ['', '', '注册生效：' + str(cat61) , '', '', '', '', '', '', '', '']


df.to_csv('C:/深交所并购重组/深交所整理数据.csv',sep=",",index=False,encoding='utf-16')

