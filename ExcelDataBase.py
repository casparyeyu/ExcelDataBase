#-*- coding:utf-8 -*-
import xlwings as xw
from multiprocessing import Pool
import pandas as pd
import re
import requests
import warnings
import time
from datetime import date, datetime, timedelta
from DataAndViewBase import *
from sqlBase import *
import itertools
from loguru import logger
#import duckdb

logger.add('log.log', rotation="100 MB", retention="10 days", enqueue=True, format="{time} {level} {message}")
warnings.filterwarnings('ignore')
sqlean.extensions.enable_all()
false = False
true = True
null = none = NULL = NONE = Null = ''

@xw.func
def instr(src_str: str, search_str: str, split_flag='#*&'):
    srcs = src_str.split(split_flag)
    #searchs = search_str.split(split_flag)
    x = 0
    for s in srcs:
        if s in search_str:
            x += 1
    return x

def fetchDatatoExcel(tableName):
    t1 = time.time()
    wb = xw.Book.caller()
    if tableName not in [sht.name for sht in wb.sheets]:
        sht = wb.sheets.add(name = tableName, after=wb.sheets.active)
    else:
        sht = wb.sheets[tableName]
        sht.select()
        sht.clear_contents()
        if sht.range("A1").api.Comment:
            sht.range("A1").api.Comment.Delete()
    sht.api.Tab.Color = 255

    df = getTableData(tableName)
    #print(df)
    df_hz = pd.DataFrame(columns=df.columns)
    if len(df)>0:
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df_hz.loc['合计',col]=df[col].sum()
            else:
                df_hz.loc['合计',col]='-'
                
    for i in range(0, len(df.columns)):
        #print(df[df.columns[i]])
        if pd.api.types.is_string_dtype(df[df.columns[i]].dtype):
            sht.range((1,i+2),(len(df)+2,i+2)).number_format='@'
        if pd.api.types.is_float_dtype(df[df.columns[i]].dtype):#df[df.columns[i]]
            sht.range((1,i+2),(len(df)+2,i+2)).number_format='#,##0.00'
    sht.range("A1").options(formatter=tabledata_format).value = df
    if len(df)>0:
        sht.range((len(df)+2, 1)).options(header=False).value = df_hz
    
    sht.range((1,2),(len(df)+1,len(df.columns)+1)).name = "DATA_" + tableName
    sht.range("B2").select()          # "A2"冻结首行
    wb.app.api.ActiveWindow.FreezePanes = True
    #sht.names.add("DATA_" + tableName, '='+sht.range((1,2),(len(df)+1,len(df.columns)+1)).address)
    sht.range("DATA_" + tableName).autofit()
    
    sht.api.Tab.Color = 5296274
    t2 = time.time()
    sht.range("A1").api.AddComment('行数：{}行\n耗时：{:.2f}秒'.format(len(df),t2 - t1))
    sht.activate()

def fetchViewDataWithHeadtoExcel(viewName):
    '''
    查询视图结果，适用于无参数或有参数但均有默认值的脚本。有非默认值的参数的脚本，近展示参数等待输入。
    '''
    t1 = time.time()
    wb = xw.Book.caller()
    if viewName!="SQL":
        if viewName not in [sht.name for sht in wb.sheets]:
            sht = wb.sheets.add(name = viewName, after=wb.sheets.active)
        else:
            sht = wb.sheets[viewName]
            sht.select()
            sht.clear_contents()
            if sht.range("A1").api.Comment:
                sht.range("A1").api.Comment.Delete()
    else:
        sht = wb.sheets.active
        sht.range((2,1),(sht.used_range.last_cell.row + 1, sht.used_range.last_cell.column)).clear_contents()
        if sht.range("A1").api.Comment:
            sht.range("A1").api.Comment.Delete()

    sht.api.Tab.Color = 255
    sht.range("B2").select()   # "A2"冻结首行
    wb.app.api.ActiveWindow.FreezePanes = False
    viewBase = loadView()
    param_keys = []
    if viewName in viewBase:
        #根据viewName去找对应的文件，然后读取sql脚本
        filename = viewBase[viewName]['file']
        with open(filename,'r', encoding='utf-8') as f:
            sqlstr = f.read()
            param_keys = list(set(re.findall('(?!regexp).*(\[[\|,，\-A-Za-z0-9\u4E00-\u9FA5]*?\])', sqlstr)))
    else:
        if viewName=="SQL":
            sqlstr = sht.range(1,2).value
        else:
            sqlstr = ''
    
    params = {}
    if param_keys:
        it = True   #是否全部有默认值
        for k in param_keys:
            if '默认' in k:
                params[k+":"] = k.split('|')[1].split(']')[0]
                if params[k+":"] == 'null' or not params[k+":"]:
                    params[k+":"] = "value          "
            else:
                params[k+":"]="value          "
                it = False
        
    if len(params) > 0:
        sht.range(1,1).value = "请输入参数"
        params = dict(sorted(params.items()))
        sht.range(1,2).options(numbers=lambda x:str(x)).value = params
        sht.range((1,2),(len(params),3)).name =  "PARAMS_" + viewName
        #sht.names.add("PARAMS_" + viewName, "="+sht.range((1,2),(len(params),3)).address)
        sht.range((1,1),(len(params),3)).autofit()
        sht.range((1,1),(len(params),3)).number_format='@'
        #sht.range((1,3),(len(params),3)).api.FONT.ITALIC = True

        if it:
            updateViewDataWithHeadtoExcelWithParameter(viewName)
    else:
        df = xsql(sqlstr) if sqlstr else pd.DataFrame()
        df_hz = pd.DataFrame(columns=df.columns)
        if len(df)>0:
            for col in df.columns:
                if pd.api.types.is_numeric_dtype(df[col]):
                    df_hz.loc['合计',col]=df[col].sum()
                else:
                    df_hz.loc['合计',col]='-'
        
        sqlline = 1 if viewName=='SQL' else 0

        for i in range(0, len(df.columns)):
            if pd.api.types.is_string_dtype(df[df.columns[i]].dtype):
                sht.range((1+sqlline,i+2),(len(df)+2+sqlline,i+2)).number_format='@'
            if pd.api.types.is_float_dtype(df[df.columns[i]].dtype):
                sht.range((1+sqlline,i+2),(len(df)+2+sqlline,i+2)).number_format='#,##0.00'

        sht.range(1+sqlline,1).value = df
        sht.range(2+sqlline,2).select()          # "A2"冻结首行
        wb.app.api.ActiveWindow.FreezePanes = True
        if len(df)>0:
            sht.range((len(df)+2+sqlline, 1)).options(header=False).value = df_hz
        
        sht.range((1+sqlline,2),(len(df)+1+sqlline,len(df.columns)+1)).name = "VIEW_" + viewName + (sht.name if viewName=="SQL" else '')
        #sht.names.add("VIEW_" + viewName + (sht.name if viewName=="SQL" else ''), "="+sht.range((1+sqlline,2),(len(df)+1+sqlline,len(df.columns)+1)).address)
        sht.range("VIEW_" + viewName + (sht.name if viewName=="SQL" else '')).autofit()
        
        if viewName!="SQL":
            #根据配置表进行默认的pivot
            if 'index' in viewBase[viewName] and 'columns' in viewBase[viewName]:
                sht.range(1,len(df.columns)+4).value = df.pivot_table(index=viewBase[viewName]['index'], columns=viewBase[viewName]['columns'], sort=False)
                sht.range(1,len(df.columns)+4).current_region.autofit()
        
        t2 = time.time()
        sht.range("A1").api.AddComment('行数：{}行\n耗时：{:.2f}秒'.format(len(df),t2 - t1))

    sht.api.Tab.Color = 5296274
    sht.activate()

def updateViewDataWithHeadtoExcelWithParameter(viewName):
    '''
    对于有参数的脚本，根据输入的参数，更新结果
    '''
    t1 = time.time()
    wb = xw.Book.caller()
    sht = wb.sheets[viewName]
    sht.api.Tab.Color = 255
    
    params = sht.range("PARAMS_" + viewName).options(dict, dtype=str, numbers=lambda x:str(int(x))).value
    params2 = {}
    for k in params.keys():
        if params[k] is None or params[k].strip()=='value' or params[k].strip()=='null':
            params[k] = ''
        params2[k.rsplit(':',1)[0]] = params[k]

    sht.range((len(params)+1,1),(sht.used_range.last_cell.row+1, sht.used_range.last_cell.column)).clear_contents()
    if sht.range("A1").api.Comment:
        sht.range("A1").api.Comment.Delete()

    viewBase = loadView()
    if viewName in viewBase:
        filename = viewBase[viewName]['file']
        with open(filename,'r', encoding='utf-8') as f:
            sqlstr = f.read()
            for k in params2.keys():
                sqlstr = sqlstr.replace(k, params2[k])
    else:
        sqlstr = ''
    df = xsql(sqlstr) if sqlstr else pd.DataFrame()
    df_hz = pd.DataFrame(columns=df.columns)
    if len(df)>0:
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df_hz.loc['合计',col]=df[col].sum()
            else:
                df_hz.loc['合计',col]='-'
                    
    for i in range(0, len(df.columns)):
        if pd.api.types.is_string_dtype(df[df.columns[i]].dtype):
            sht.range((len(params)+1,i+2),(len(params)+len(df)+2,i+2)).number_format='@'
        if pd.api.types.is_float_dtype(df[df.columns[i]].dtype):
            sht.range((len(params)+1,i+2),(len(params)+len(df)+2,i+2)).number_format='#,##0.00'

    sht.range(len(params)+1,1).value = df
    if len(df)>0:
        sht.range((len(params)+len(df)+2, 1)).options(header=False).value = df_hz
        
    sht.range((len(params)+1,2),(len(params)+len(df)+1,len(df.columns)+1)).name = "VIEW_" + viewName 
    sht.range(len(params)+2,1).select()          # "A2"冻结首行
    wb.app.api.ActiveWindow.FreezePanes = True
    #sht.names.add("VIEW_" + viewName, '='+sht.range((len(params)+1,2),(len(params)+len(df)+1,len(df.columns)+1)).address)
    sht.range("VIEW_" + viewName).autofit()
        
    if 'index' in viewBase[viewName] and 'columns' in viewBase[viewName]:
        sht.range(len(params)+1,len(df.columns)+4).value = df.pivot_table(index=viewBase[viewName]['index'], columns=viewBase[viewName]['columns'], sort=False)
        sht.range(len(params)+1,len(df.columns)+4).current_region.autofit()
    
    sht.api.Tab.Color = 5296274
    t2 = time.time()
    sht.range("A1").api.AddComment('行数：{}行\n耗时：{:.2f}秒'.format(len(df),t2 - t1))
    sht.activate()

def fetchViewDataOnlytoExcel(area_index, value_head, viewName):
    t1 = time.time()
    wb = xw.Book.caller()
    sht = wb.sheets.active
    #获取行列的组合
    #标题行的行数
    head_row_cnt = sht.range(value_head).rows.count
    #列标题，可能有多行，标准化成同一格式[[A,B],[C,D]];[[A,B]],[[A]
    if head_row_cnt == 1:
        target_columns = [[sht.range(value_head).value]] if sht.range(value_head).count==1 else [sht.range(value_head).value] #兼容一列及多列
    else:
        target_columns = [[x] for x in sht.range(value_head).value] if sht.range(value_head).columns.count==1 else  sht.range(value_head).value
    if ',' in area_index or '，' in area_index:
        #target_index = sht.range(area_index.split(',',1)[1]).options(pd.DataFrame, index=False, header=False).value.ffill()
        #行、列的个数相关，索引有多少行、列有多少列、标题行有多少行，都用影响数据结构
        index_head = sht.range(re.split('[,，]',area_index,1)[0]).value
        if sht.range(re.split('[,，]',area_index,1)[0]).rows.count>1:
            #索引标题有多行，取第一行
            index_head = sht.range(re.split('[,，]',area_index,1)[0]).value[0]
        if sht.range(re.split('[,，]',area_index,1)[0]).columns.count==1:
            index_head = [index_head]
        
        if sht.range(re.split('[,，]',area_index,1)[1]).rows.count > 1:
            #多于一行，取全部
            index_value = sht.range(re.split('[,，]',area_index,1)[1]).value
        else:
            index_value = [sht.range(re.split('[,，]',area_index,1)[1]).value]
        
        target_index = pd.DataFrame(index_value, columns=index_head).ffill()
    else:
        target_index = sht.range(area_index).options(pd.DataFrame,index=False).value.ffill().loc[head_row_cnt-1:]
    
    viewBase = loadView()
    if viewName in viewBase:
        filename = viewBase[viewName]['file']
    else:
        for key in viewBase.keys():
            #print(viewBase[key])
            if 'alias' in viewBase[key] and viewName == viewBase[key]['alias']:
                viewName = key
                filename = viewBase[key]['file']
                break
        if viewName not in viewBase:
            return None
    
    with open(filename,'r', encoding='utf-8') as f:
        sqlstr = f.read()
    #print(sqlstr)
    df = xsql(sqlstr)
    #xw.view(df)
    head_item=[]
    for rowno in range(len(target_columns)):
        for item in df.columns.tolist():
            if item not in target_index.columns.tolist() and (df[item].dtype==str or df[item].dtype==object):
                #print(item, target_columns[rowno][0], df[item].values)
                if target_columns[rowno][0] in df[item].values:
                    head_item.append(item)
                    break
    
    value_item=''
    for item in df.columns.tolist():
        if item not in target_index.columns.tolist() and item not in head_item:
            value_item=item
            break
    df = pd.pivot_table(df, index=target_index.columns.tolist(), columns=head_item, values=value_item, sort=False) 
    #df = df.pivot_table(index=target_index.columns.tolist(), columns=head_item, values=value_item, sort=False)
    #xw.view(df)
    if head_row_cnt>1:
        #改变表头的组成方式
        new_target_columns = []
        for col in sht.range(value_head).columns:
            new_target_columns.append(tuple(col.value))
    else:
        new_target_columns = target_columns[0]
    target_df = pd.merge(left=target_index, right=df, how='left', left_on=target_index.columns.tolist(), right_on=target_index.columns.tolist()) #python 2.0版本会报错
    if ',' in area_index or '，' in area_index:
        sht.range(sht.range(re.split('[,，]',area_index,1)[1]).row,sht.range(value_head).column).value = target_df[new_target_columns].fillna(0).values
    else:
        sht.range(value_head).offset(head_row_cnt).value = target_df[new_target_columns].fillna(0).values
    
    t2 = time.time()
    xw.apps.active.alert('填入完成，耗时：{:.2f}秒'.format(t2 - t1))

def getDataMenu():
    '''
    在Excel中加载数据库表的列表，列明parent（分类）和表名，展示在当前Sheet中。
    '''
    DataBase = loadData()
    menu = [[a[1]['parent'],a[0]] for a in sorted(DataBase.items(), key =lambda x:x[1]['parent']+x[0])]
    wb = xw.Book.caller()
    sheetName = '数据列表'
    if sheetName not in [sht.name for sht in wb.sheets]:
        sht = wb.sheets.add(name = sheetName, after=wb.sheets.active)
    else:
        sht = wb.sheets[sheetName]
        sht.activate()
        sht.used_range.delete()
    sht['A1'].value=['分类','表名']
    sht['A2'].value = menu
    autoMergeCells()
    table_format(sht['A1'].current_region)
    
def getViewMenu():
    '''
    在Excel中加载视图的列表，列明parent（分类）和表名，展示在当前Sheet中。
    '''
    ViewBase = loadView()
    menu = [[a[1]['parent'],a[0]] for a in sorted(ViewBase.items(), key =lambda x:x[1]['parent']+x[0])]
    wb = xw.Book.caller()
    sheetName = '报表列表'
    if sheetName not in [sht.name for sht in wb.sheets]:
        sht = wb.sheets.add(name = sheetName, after=wb.sheets.active)
    else:
        sht = wb.sheets[sheetName]
        sht.activate()
        sht.used_range.delete()
    sht['A1'].value=['分类','表名']
    sht['A2'].value = menu
    autoMergeCells()
    table_format(sht['A1'].current_region)

def fetch(sheetName):
    #根据入参SheetName判断是否为表名或者视图名，或者根据备注去判断是表还是视图，然后刷新数据
    #支持视图有参数需要调整时，重新进入查询
    DataBase = loadData()
    viewBase = loadView()
    if sheetName in viewBase:
        #判断是否有参数区域
        wb = xw.Book.caller()
        sht = wb.sheets[sheetName]
        hasparams = list(filter(lambda x:sheetName in x and 'PARAMS' in x,[n.name for n in wb.names]))
        if hasparams:
            updateViewDataWithHeadtoExcelWithParameter(sheetName)
        else:
            fetchViewDataWithHeadtoExcel(sheetName)
    elif sheetName in DataBase:
        wb = xw.Book.caller()
        sht = wb.sheets[sheetName]
        fetchDatatoExcel(sheetName)
    else:
        wb = xw.Book.caller()
        sht = wb.sheets[sheetName]
        if str.upper(sht.range(1,1).value) == 'SQL':
            fetchViewDataWithHeadtoExcel('SQL')
        else:
            pass
