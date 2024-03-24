import re
import sqlean
from sql_metadata import Parser
from DataAndViewBase import *

def getTableData(tableName):
    DataBase = loadData()
    if tableName in DataBase:
        df = pd.DataFrame()
        for file in DataBase[tableName]['files'].values():
            #需要进一步获取sheetName,作为下一步的入参
            sheetName = 0 if DataBase[tableName]['sheetName']=='' or DataBase[tableName]['sheetName']=='nan' else DataBase[tableName]['sheetName']
            #需要根据文件名称，重新提取会计期间等信息
            fileInfo, configInfo = getTableInfoFromFileName(file, sheetName)[0]
            header = 0 if configInfo['是否有标题行']=='是' else None
            usecols = configInfo['指定列']  if ':' in configInfo['指定列'] else None
            skiprows = int(configInfo['起始行'])-1
            skipfooter = 1 if configInfo['是否有汇总行']=='是' else  0
            if 'xls' in file:
                try:
                    newdf = pd.read_excel(file, sheet_name=sheetName, header=header, usecols=usecols, skiprows=skiprows, skipfooter=skipfooter, dtype={'科目编码':str,'产品编码':str,'总账科目编码':str,'总账辅助核算编码':str,'账号':str,'银行账号':str,'银行账户编码':str,'对方账号':str,'附件业务数据ID':str,'年':str,'月':str,'日':str,'Proc_Num':str,'日期':str,'付款帐号':str,'收款帐号':str,'科目':str,'核算科目':str,'起存日':str,'到期日':str})
                except Exception as e:
                    #xw.apps.active.alert('错误：'+ str(e))
                    io_excel = read_excel_with_password(file)
                    newdf = pd.read_excel(io_excel, sheet_name=sheetName, engine='openpyxl', header=header, usecols=usecols, skiprows=skiprows, skipfooter=skipfooter, dtype={'科目编码':str,'产品编码':str,'总账科目编码':str,'总账辅助核算编码':str,'账号':str,'银行账号':str,'银行账户编码':str,'对方账号':str,'附件业务数据ID':str,'年':str,'月':str,'日':str,'Proc_Num':str,'日期':str,'付款帐号':str,'收款帐号':str,'科目':str,'核算科目':str,'起存日':str,'到期日':str})
            if 'csv' in file or 'txt' in file:
                try:
                    newdf = pd.read_csv(file, sep='[,\t\|]', quotechar= "'",header=header, usecols=usecols, skiprows=skiprows, skipfooter=skipfooter, dtype={'科目编码':str,'产品编码':str,'总账科目编码':str,'总账辅助核算编码':str,'账号':str,'银行账号':str,'银行账户编码':str,'对方账号':str,'附件业务数据ID':str,'日期':str,'付款帐号':str,'收款帐号':str,'科目':str,'核算科目':str,'起存日':str,'到期日':str}, engine='python', on_bad_lines="warn",encoding='gbk')
                except:
                    try:
                        newdf = pd.read_csv(file, sep=',', quotechar= "'",header=header, usecols=usecols, skiprows=skiprows, skipfooter=skipfooter, dtype={'科目编码':str,'产品编码':str,'总账科目编码':str,'总账辅助核算编码':str,'账号':str,'银行账号':str,'银行账户编码':str,'对方账号':str,'附件业务数据ID':str,'日期':str,'付款帐号':str,'收款帐号':str,'科目':str,'核算科目':str,'起存日':str,'到期日':str}, engine='python', on_bad_lines="warn",encoding='gbk')
                    except:
                        try:
                            newdf = pd.read_csv(file, sep='\t', quotechar= "'",header=header, usecols=usecols, skiprows=skiprows, skipfooter=skipfooter, dtype={'科目编码':str,'产品编码':str,'总账科目编码':str,'总账辅助核算编码':str,'账号':str,'银行账号':str,'银行账户编码':str,'对方账号':str,'附件业务数据ID':str,'日期':str,'付款帐号':str,'收款帐号':str,'科目':str,'核算科目':str,'起存日':str,'到期日':str}, engine='python', on_bad_lines="warn",encoding='gbk')
                        except:
                            newdf = pd.read_csv(file, sep='\|', quotechar= "'",header=header, usecols=usecols, skiprows=skiprows, skipfooter=skipfooter, dtype={'科目编码':str,'产品编码':str,'总账科目编码':str,'总账辅助核算编码':str,'账号':str,'银行账号':str,'银行账户编码':str,'对方账号':str,'附件业务数据ID':str,'日期':str,'付款帐号':str,'收款帐号':str,'科目':str,'核算科目':str,'起存日':str,'到期日':str}, engine='python', on_bad_lines="warn",encoding='gbk')
            if len(newdf)==0:
                xw.apps.active.alert('当前文件未获取到数据：' + file)
            newdf.columns = newdf.columns.str.strip()
            for key in list(fileInfo.keys())[2:]:
                #不展示前两个key信息，即表名、sheetName
                newdf[key] = fileInfo[key]
            try:
                df = pd.concat([df,newdf], sort=False, ignore_index=True)
            except Exception as e:
                xw.apps.active.alert(file + '\nconcat合并失败：' + str(e))
        df.index = df.index + 1
        return df
    else:
        return pd.DataFrame()

def getSQLParams(sqlfiles):
    '''
    sqlfiles:sql文件的字典列表
    '''
    params = []
    for s in sqlfiles.keys():
        if '.sql' in sqlfiles[s]:
            params.extend(re.findall('(?!regexp).*(\[[\|,，A-Za-z0-9\u4E00-\u9FA5]*?\])', open(sqlfiles[s], 'r').read()))
    return list(set(params))

def clean(sql_str):
    q = re.sub(r"/\*[^*]*\*+(?:[^*/][^*]*\*+)*/", "", sql_str)
    lines = [line for line in q.splitlines() if not re.match("^\s*(--|#)", line)]
    q = " ".join([re.split("--|#", line)[0] for line in lines])
    q = ' '.join(q.split())
    return q

def xsql(sqlstr):
    """
    自动识别其中的表名,将表名替换为对应的df,然后执行获取结果
    """
    conn =sqlean.connect(':memory:')
    #conn.create_function("REGEXP", 2, regexp)
    tablenames = Parser(sqlstr).tables
    for tb in tablenames:
        df = getTableData(tb)
        if not df.empty:
            df.to_sql(tb, conn)
    df = pd.read_sql(sqlstr, conn)
    df.index = df.index + 1
    return df
