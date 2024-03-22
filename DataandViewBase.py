import natsort
import os
import io
import pandas as pd
import json
import re
import msoffcrypto
import xlwings as xw

'''
首先，你需要定义你的不同的Excel。定义的方式为：根据“文件名+sheet名”映射到某一个表名，如将满足“文件名匹配'.*工作记录表-(.*?)-(\d{8})'且sheetName='工作记录'”的excel，定义为表“员工每日工作记录表”。
后续将可以通过“select * from 员工每日工作记录表”获取全部类似Excel的数据。
文件名支持正则，sheetName暂不支持正则。
'''


def judgeIfcomeinNew(path, tablefile):
    """
    用于判断path目录下的各个文件，和配置文件tablefile谁更新。如果path目录下的文件更新，则需要更新配置文件。
    Inputs:
        path:数据文件所在的路径
        tablefile:每次更新配置，生成的tables.json文件。该文件记录每个“数据库表名”对应的Excel的文件列表。
    Returns:
        True/False:如果为True，提示有文件发生更新，可能有新增文件，要更新配置文件。
    """
    newest_time = os.path.getmtime(tablefile)
    for root, dirs, files in os.walk(path):
        for file in files:
            file_time = os.path.getmtime(os.path.join(root, file))
            if file_time > newest_time:
                return True
    return False

def getTableInfoFromFileName(filename, sheetName=0):
    """
    用于建立“文件名+sheet名”到“数据库表名”的映射关系。
    filename为文件的绝对路径; home_path用来获取根目录，用于支持子文件夹的情况; name为文件名或二级子目录+文件名；
    sheetName=0,表示不要求填sheetName,直接取默认的第1个Sheet,用于获取配置信息、或读文件时sheetName为空的情况。
    返回：
    target_info:{'表名':最终的登记表名,'会计期间':会计期间,'其他信息':其他信息,}, 用于根据文件名对应到表名，生成配置文件
    config_info:{'起始行':, '指定列':'', '是否有标题行':, '是否有汇总行':}, 用于实际读取数据时
    """
    config = ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__),'config\\config.ini'), encoding='utf-8')
    if config.get("public_config","using_file")=='True':
        df_config = pd.read_excel(os.path.join(config.get("public_config","config_path"), config.get("public_config","data_decode_file")),keep_default_na=False) 
    else:
        df_config = pd.DataFrame(get_opsData(config["public_config_ops"]["appkey"],config["public_config_ops"]["sign"],config["public_config_ops"]["data_decode"].split(',')[0],config["public_config_ops"]["data_decode"].split(',')[1]))
        df_config.drop(columns=['创建时间','创建人','拥有者','uaid','最近修改时间','autoid','wfftime','allowdelete','controlpermissions'],inplace=True)
    #data_dir的配置为：项目 = 目录,密码   或   ps.项目=ops的地址,sign,key
    all_paths = config.items("data_dir")
    home_path = ''
    #home_path为选中的data_dir对应的目录，即文件所属的根目录，当文件属于该文件夹或该文件夹的子文件夹时，循环终止
    #print(filename)
    for path in all_paths:
        #print(path[1])
        if path[1].split(',')[0] in filename:
            home_path = path[1].split(',')[0]
            break
    #name取值：home_path下只有一级目录时，取文件名；有二级目录时，取二级目录+文件名，以此类推；有二级目录的，二级目录同样参与文件名格式的匹配
    name = filename.split(home_path+'\\')[-1].split('\\',1)[-1].rsplit('.',1)[0]   
    #sheetName在进入配置时，不再读取文件，而直接以配置为准
    target_config = df_config[df_config['文件名格式'].apply(lambda x: True if re.findall(x, name) else False) & df_config['sheetName'].apply(lambda x:True if (x==sheetName and sheetName!=0 or sheetName==0) else False)].reset_index(drop=True)
    target_info = {}
    config_info = {}

    #print(home_path, name, target_config)
    
    if len(target_config)==0:
        #表示没有进行配置的表，按通用简单表格文件处理
        target_info = {'表名':name,'sheetName':'','文件名':name, }
        config_info = {'起始行':1, '指定列':'', '是否有标题行':'是', '是否有汇总行':'否'}
        return [(target_info,config_info)]
    target_info_list = []
    for index, row in target_config.iterrows():
        #nameinfo:根据文件名格式，从文件名中获取关键信息，用于组成表名、会计期间、其他信息中的关键值
        nameinfo = list(re.findall(row['文件名格式'], name)[0]) if isinstance(re.findall(row['文件名格式'], name)[0], tuple) else re.findall(row['文件名格式'], name)
        #根据表名信息，对其他信息进行替换
        for i in range(len(nameinfo)):
            row['表名'] = row['表名'].replace('\\'+str(i+1),nameinfo[i])
            row['会计期间'] = row['会计期间'].replace('\\'+str(i+1),nameinfo[i])
            row['其他信息'] = row['其他信息'].replace('\\'+str(i+1),nameinfo[i])
        #print(name, nameinfo, row['表名'], row['会计期间'], row['其他信息'])
        target_info['表名'] = row['表名']
        target_info['sheetName'] = row['sheetName'] 
        if row['会计期间']:
            target_info['会计期间'] = row['会计期间']
        if row['其他信息']:
            target_info.update(eval(row['其他信息']))
        
        default = {'起始行':1,'指定列':None,'是否有标题行':'是','是否有汇总行':'否'}
        for i, v in row.items():
            if i not in ['表名','文件名格式','会计期间','其他信息','sheetName']:
                config_info[i] = default[i] if pd.isna(v) else v
        target_info_list.append((target_info,config_info))
    return target_info_list

def listDataDir(path, res, home_path):
    '''
    Returns:
    database={table1:{'parent':'\子目录','files':{key1:file1,key2:file2}}
    DataBase的结构:
    1)有哪些数据表
    2)每个数据表对应的文件列表，不同文件的区分方式
    与原Excel组织方式的不同：不再关心数据的层次关系，即不再关系目录
    '''
    if path[:4]!='http':
        dirlist = os.listdir(path.split(',')[0])
        dirlist = list(filter(lambda x:'.json' not in x and 'bak' not in x and 'Bak' not in x and 'BAK' not in x and '~' not in x and '表格解析配置' not in x, dirlist))
        for f in natsort.natsorted(dirlist):
            temp_path = os.path.join(path.split(',')[0], f)
            if os.path.isdir(temp_path):
                listDataDir(temp_path, res, home_path)
            else:
                #print(temp_path)
                for info in getTableInfoFromFileName(temp_path):
                    fileinfo = info[0]
                    #print(fileinfo)
                    if fileinfo['表名'] not in res:
                        #print(temp_path, home_path, temp_path.split(home_path.split(',')[0]+'\\')[-1].split('\\',1)[0])
                        res[fileinfo['表名']] = {'parent':temp_path.split(home_path.split(',')[0]+'\\')[-1].split('\\',1)[0],'files':{},'passwds':'','sheetName':fileinfo['sheetName']}
                    res[fileinfo['表名']]['files'].update({'_'.join(list(fileinfo.values())[1:]):temp_path})
                    if ',' in path:
                        res[fileinfo['表名']]['passwds']=path.split(',')[1]

def loadData(refresh = False):
    '''
    加载数据库表：如果有新的数据文件，则重新读取所有数据文件结构；否则，直接读取表格配置文件
    Returns:
      DataBase:数据表:{'表名':{'parent':'文件夹','files':[]}}
    '''
    config = ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__),'config\\config.ini'), encoding='utf-8')
    DataBase = {}
    if not refresh:
        with open(os.path.join(os.path.dirname(__file__),'config\\tables.json'), 'r', encoding='utf-8') as f:
            DataBase = json.load(f)
    
    update = False
    for key in config.options("data_dir"):
        currentDataBase= {}
        if 'ops.'!=key[:4] and (refresh or judgeIfcomeinNew(config["data_dir"][key],os.path.join(os.path.dirname(__file__),'config\\tables.json'))):
            listDataDir(config["data_dir"][key], currentDataBase, config["data_dir"][key])
            DataBase.update(currentDataBase)
            update = True
        elif 'ops.'==key[:4] and refresh:
            currentDataBase[key.split('.')[2]]={'parent':key[4:].split('.')[1],'files':{},'passwds':''}
            currentDataBase[key.split('.')[2]]['files'].update({'ops':config["data_dir"][key].split(',')[0]})
            currentDataBase[key.split('.')[2]]['passwds']=config["data_dir"][key].split(',',1)[1]
            DataBase.update(currentDataBase)
            update = True
        else:
            pass

    if update:
        with open(os.path.join(os.path.dirname(__file__),'config\\tables.json'), 'w', encoding='utf-8') as f:
            json.dump(DataBase, f, ensure_ascii=False) 
            
    return DataBase

def listViewDir(path, res):
    '''
    Returns:
    database={table1:{'parent':'\子目录','files':{key1:file1,key2:file2}}
    DataBase的结构:
    1)有哪些数据表
    2)每个数据表对应的文件列表，不同文件的区分方式
    与原Excel组织方式的不同：不再关心数据的层次关系，即不再关系目录
    '''
    dirlist = os.listdir(path)
    dirlist = list(filter(lambda x:'.json' not in x and 'bak' not in x and 'Bak' not in x and 'BAK' not in x and '~' not in x and '表格解析配置' not in x and 'xls' not in x, dirlist))
    for f in natsort.natsorted(dirlist):
        temp_path = os.path.join(path, f)
        if os.path.isdir(temp_path):
            listViewDir(temp_path, res)
        else:
            filaname = f.rsplit('.',1)[0]
            res[filaname] = {'parent':path.split('\\')[-1],'file':temp_path}

def loadView(refresh = False):
    '''
    加载快报视图表：如果有新的视图文件，则重新读取所有数据文件结构；否则，直接读取表格配置文件
    Returns:
      ViewBase:数据表:{'表名':{'parent':'文件夹','files':[]}}
    '''
    config = ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__),'config\\config.ini'), encoding='utf-8')
    ViewBase = {}
    if not refresh:
        with open(os.path.join(os.path.dirname(__file__),'config\\views.json'), 'r', encoding='utf-8') as f:
            ViewBase = json.load(f)

    update =False
    for key in config.options("view_dir"):
        currentViewBase = {}
        if refresh or judgeIfcomeinNew(config["view_dir"][key],os.path.join(os.path.dirname(__file__),'config\\views.json')):
            listViewDir(config["view_dir"][key], currentViewBase)
            if config.get("public_config","using_file")=='True':
                df = pd.read_excel(os.path.join(config["public_config"]["config_path"], config["public_config"]["view_config_file"]),index_col=0)
            else:
                df = pd.DataFrame(get_opsData(config["public_config_ops"]["appkey"], config["public_config_ops"]["sign"], config["public_config_ops"]["view_config"].split(',')[0], config["public_config_ops"]["view_config"].split(',')[1]))
                df.set_index('视图名', inplace=True)
            for view in currentViewBase.keys():
                if view in df.index.tolist():
                    currentViewBase[view]['index'] = df.loc[view]['行字段'].split(',')
                    currentViewBase[view]['columns'] = df.loc[view]['列字段'].split(',')
                    currentViewBase[view]['values'] = df.loc[view]['值字段']
                    currentViewBase[view]['alias'] = df.loc[view]['其他名字']
            ViewBase.update(currentViewBase)
            update = True
        else:
            pass
    
    if update:
        with open(os.path.join(os.path.dirname(__file__),'config\\views.json'), 'w', encoding='utf-8') as f:
            json.dump(ViewBase, f, ensure_ascii=False) 
    return ViewBase

def updateConfig(viewID, userName):
    '''
    根据低码端的配置，为不同的使用者，进行不同的配置。
    Inputs:
        viewID:
        userName:
    '''
    appKey = '6b014274c34294d5'
    sign = 'ZTc3NmI0OWQ2NDkxYjkzNGY2MmMzODQ3YWQ1NTUzOTE3MGQ1Y2U1NmQ4MWFkZTBmYjAzMjYxMmE0ODFmMWMyZQ=='
    rs = get_opsData(appKey, sign, '64cb6dc6b4448501e994f735', viewID)

    config = ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__),'config\\config.ini'), encoding='utf-8')

    for d in rs:
        if 'ops.' not in d['数据分类']:
            config.set('data_dir', d['数据分类'] + '.' +  d['文件夹'].rsplit('\\',1)[1], d['文件夹'].rsplit('\\',1)[0]+ ',' + d['key'])
            if d['key']:
                config.set('password', d['数据分类'] + d['文件夹'].rsplit('\\',1)[1], d['key'])
        else:
            config.set('data_dir', d['数据分类'] + '.' +  d['文件名'], d['文件夹'] + ',' + d['key'])
        

    #再更新View配置，避免展示过多无用的配置
    rs = get_opsData(appKey, sign, '64ccbfeeb4448501e994f86d', '64ccbfeeb4448501e994f877')
    for d in rs:
        if userName in ','.join(u['fullname'] for u in json.loads(d['报表使用者'])):
            config.set('view_dir', d['视图分类'], d['文件夹'].rsplit('\\',1)[0])

    with open(os.path.join(os.path.dirname(__file__),'config\\config.ini'), 'w', encoding='utf-8') as f:
        config.write(f)    

def load(refresh = False, viewID = '', userName = ''):
    if refresh and viewID and userName:
        try:
            updateConfig(viewID, userName)
        except Exception as e:
            xw.apps.active.alert(str(e))
            xw.apps.active.alert('连接低码获取目录及密钥失败，请手动更新config.ini或联系张家炜')

    loadData(refresh)
    loadView(refresh)
    if refresh:
        xw.apps.active.alert("更新完成，开始使用吧！")
    


def read_excel_with_password(file):
    config = ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__),'config\\config.ini'), encoding='utf-8')
    pwds = config.items("password")
    temp = io.BytesIO()
    with open(file, 'rb') as f:
        excel = msoffcrypto.OfficeFile(f)
        for pwd in pwds:
            try:
                excel.load_key(pwd[1])
                excel.decrypt(temp)
                break
            except:
                print('wrong:',pwd[0])
                temp = io.BytesIO()
    if len(temp.getvalue())==0:
        xw.apps.active.alert(file+'已加密，未匹配到正确的密码')

    return temp

def read_excel_by_xlwings(file, sheet_index=0, sheet_name='', header=0, usecols='', skiprows=0, skipfooter=0, tdtype=str):
    """
    可以用于解密，但速度很慢，且对于多列列明相同时，会导致后续无法支持
    """
    config = ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__),'config\\config.ini'), encoding='utf-8')
    pwds = config.items("password")
    book = None
    
    with xw.App() as app:
        app.visible = False
        app.screen_updating = False
        book = app.books.open(file, read_only=True)
        try:
            book = app.books.open(file, read_only=True)
        except:
            for pwd in pwds:
                try:
                    book = app.books.open(file, read_only=True, password = pwd[1])
                    break
                except:
                    pass
        print(book)
        if book:
            sht = book.sheets[sheet_name] if sheet_name else book.sheets[sheet_index]
            startrow = skiprows+1
            endrow = sht.used_range.rows.count-skipfooter
            if usecols and ':' in usecols:
                rng = sht.range(usecols.split(':')[0] + str(startrow) + ':' + usecols.split(':')[1] + str(endrow))
            else:
                rng = sht.range((startrow, 1),(endrow, sht.used_range.columns.count))
            df = rng.options(pd.DataFrame, index=False, header=1, dtype=str, empty='').value
            book.close()
        else:
            df = pd.DataFrame()
        
    return df
