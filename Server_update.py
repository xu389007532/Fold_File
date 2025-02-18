import os
import datetime
import sqlite3
from sqlite3 import connect
from Share import Honour_Share
import pymssql


def mssql_insert(sql,data):
    # 连接数据库
    userID, userPWD, serverName, databaseName = Honour_Share.Py_Decrypto(os.path.abspath('D:/xu/Python/Fold_File/Source/userCommon.xml'), os.path.abspath('D:/xu/Python/Fold_File/Source/Honour.dll'))
    connection = pymssql.connect(server=serverName, user=userID, password=userPWD, database=databaseName)
    try:
        with connection.cursor() as cursor:
            # 创建插入语句（使用%s作为占位符）
            # sql = "INSERT INTO Quotation_Update (FileName, Fold, CheckFold, Type,UpdateTime) VALUES (%s, %s, %s, %s, %s)"
            # 准备数据（每个元组代表一行数据）
            # data = [
            #     ('value1_1', 'value1_2'),
            #     ('value2_1', 'value2_2'),
            #     # 添加更多数据...
            # ]
            # 执行批量插入
            cursor.executemany(sql, data)

            # 提交事务
            connection.commit()
    except Exception as e:
        print('MSSQL出现错误，回滚事务')
    finally:
        # 关闭连接
        connection.close()

def update_Folds(sql,data):
    bsdata = connect("D:/xu/Python/Fold_File/Source/FoldFiles.db")
    cursor = bsdata.cursor()
    cursor.execute(sql)
    data1 = cursor.fetchone()
    if data1 is None:
        cursor.execute('INSERT INTO Folds (Fold, FoldLastModify) VALUES (?, ?)',data)
    else:
        sql_update=f"UPDATE Folds SET FoldLastModify = '{data[1]}' WHERE Fold='{data[0]}'"
        cursor.execute(sql_update)
    bsdata.commit()
    cursor.close()
    bsdata.close()

def commit_SQL(sql):
    bsdata = connect("D:/xu/Python/Fold_File/Source/FoldFiles.db")
    cursor = bsdata.cursor()
    cursor.execute(sql)
    bsdata.commit()
    cursor.close()
    bsdata.close()

def read_data(sql):
    bsdata = connect("D:/xu/Python/Fold_File/Source/FoldFiles.db")
    cursor = bsdata.cursor()
    cursor.execute(sql)
    data1 = cursor.fetchone()
    cursor.close()
    bsdata.close()
    return data1

def read_data_all(sql):
    bsdata = connect("D:/xu/Python/Fold_File/Source/FoldFiles.db")
    cursor = bsdata.cursor()
    cursor.execute(sql)
    data1 = cursor.fetchall()
    cursor.close()
    bsdata.close()
    return data1

def commit(data,sql):


    conn = sqlite3.connect("D:/xu/Python/Fold_File/Source/FoldFiles.db")
    cursor = conn.cursor()

    # # 准备数据
    # data = [
    #     (1, 'item1', 10.0),
    #     (2, 'item2', 20.0),
    #     (3, 'item3', 30.0),
    #     # 添加更多数据...
    # ]

    # 使用事务批量插入数据，然后一次性提交事务以优化性能
    try:
        conn.execute('BEGIN')  # 开始事务（可选，默认情况下SQLite会自动开启）

        cursor.executemany(sql, data)
        conn.commit()  # 提交事务，一次性完成所有插入操作。
    except Exception as e:
        print('出现错误，回滚事务')
        print(sql)
        print(data)
        conn.rollback()  # 如果出现错误，回滚事务。
    finally:
        conn.close()  # 最后关闭连接。无论成功还是失败，都应关闭连接。

def check_Folds(directory,lastupdate):
    lastupdate1 = datetime.datetime.strptime(lastupdate, "%Y-%m-%d %H:%M:%S")
    Folds = []
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        dir_path = item_path.replace('\\', '/')
        if os.path.isdir(dir_path):
            timestamp = os.path.getmtime(dir_path)
            date_time = datetime.datetime.fromtimestamp(timestamp)
            formatted_date_time = date_time.strftime('%Y-%m-%d %H:%M:%S')

            if date_time>lastupdate1:
                Folds.append((dir_path, formatted_date_time))
                print(f"目录: {dir_path} 最后更新日期和时间: {formatted_date_time}")
    return Folds

def list_Folds(directory):
    # lastupdate1 = datetime.datetime.strptime(lastupdate, "%Y-%m-%d %H:%M:%S")
    Folds = []
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        dir_path = item_path.replace('\\', '/')
        if os.path.isdir(dir_path):
            timestamp = os.path.getmtime(dir_path)
            date_time = datetime.datetime.fromtimestamp(timestamp)
            formatted_date_time = date_time.strftime('%Y-%m-%d %H:%M:%S')
            print(f"目录: {dir_path} 最后更新日期和时间: {formatted_date_time}")
            Folds.append((dir_path, formatted_date_time))
    return Folds

def list_Folds_bak(directory):
    Folds=[]
    for root, dirs, files in os.walk(directory):
        # print('root:',root,'dirs:',dirs)
        for name in dirs:
            dir_path1 = os.path.join(root, name)
            dir_path = dir_path1.replace('\\','/')
            timestamp = os.path.getmtime(dir_path)
            date_time = datetime.datetime.fromtimestamp(timestamp)
            formatted_date_time = date_time.strftime('%Y-%m-%d %H:%M:%S')
            print(f"目录: {dir_path} 最后更新日期和时间: {formatted_date_time}")
            Folds.append((dir_path,formatted_date_time))
    return Folds

def list_files(directory):

    AllFiles=[]
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for root, dirs, files in os.walk(directory):
        # print('root:',root,'dirs:',dirs)

        for f in files:
            path=root.replace('\\', '/')
            if path.find('/', len(directory) + 1)==-1:
                checkFold=path
            else:
                checkFold=path[:path.find('/', len(directory) + 1)]
            AllFiles.append((path, f,now,checkFold))
            # print(f"{os.path.abspath(root)}/{f}")
    return AllFiles

def first_update(directory_path):

    updatetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 指定你要搜索的目录路径
    print(datetime.datetime.now())
    mssql_insert('Delete Quotation_Update',[()])
    commit_SQL("DELETE FROM Files")
    commit_SQL("DELETE FROM TempFiles")
    commit_SQL("DELETE FROM Folds")

    # directory_path = '//hmpfs01/DG3 Public Library/ITProgram/PDFtypesetting'
    # directory_path = '//10.5.85.5/hp-pricing/SZO Estimation Folder/Quotation to client in Excel format'
    Folds=list_Folds(directory_path)
    sql_insert1='INSERT INTO Folds (Fold, FoldLastModify) VALUES (?, ?)'
    commit(Folds,sql_insert1)
    print(datetime.datetime.now())


    Files=list_files(directory_path)
    sql_insert2='INSERT INTO Files (Fold, FileName,FileUpdateTime,CheckFold) VALUES (?, ?, ?, ?)'
    commit(Files,sql_insert2)

    sql_update=f"UPDATE config SET LastModify = '{updatetime}'"
    commit_SQL(sql_update)

# first_update()

def update_Fold_File(directory_path,local, lastupdate):
    # directory_path = '//hmpfs01/DG3 Public Library/ITProgram/PDFtypesetting'
    updatetime=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print('更新時間:',updatetime)
    # lastupdate="2024-12-01 15:30:45"
    Check=check_Folds(directory_path,lastupdate)
    commit_SQL("DELETE FROM TempFiles")
    for c in Check:
        sql_update_Folds="SELECT * FROM Folds WHERE Fold='"+c[0]+"'"
        update_Folds(sql_update_Folds,c)
        Files = list_files(c[0])
        sql_insert2 = 'INSERT INTO TempFiles (Fold, FileName,FileUpdateTime,CheckFold) VALUES (?, ?, ?, ?)'
        commit(Files, sql_insert2)

    SQL_DELETE='SELECT A.FileName, A.Fold, A.CheckFold, B.FileName as FileName1 FROM  View_CheckFold as A LEFT OUTER JOIN TempFiles As B on A.FileName=B.FileName and  A.Fold=B.Fold where FileName1 is NULL'
    delete_data=read_data_all(SQL_DELETE)
    delete_data_list=[]
    delete_data_list2 = []
    for d in delete_data:
        delete_data1=(d[0],d[1],d[2],"Delete",updatetime)
        delete_data2 = (d[0], d[1])
        delete_data_list.append(delete_data1)
        delete_data_list2.append(delete_data2)

    SQL_ADD='SELECT A.FileName, A.Fold, A.CheckFold, B.FileName as FileName1 FROM  TempFiles as A LEFT OUTER JOIN View_CheckFold As B on A.FileName=B.FileName and  A.Fold=B.Fold where FileName1 is NULL'
    add_data=read_data_all(SQL_ADD)
    add_data_list=[]
    add_data_list2 = []
    for a in add_data:
        add_data1=(a[0],a[1],a[2],"Add",updatetime)
        add_data2 = (a[0], a[1], a[2], updatetime)
        add_data_list.append(add_data1)
        add_data_list2.append(add_data2)

    print('增加文件：',add_data_list)
    print('刪除文件：', delete_data_list)


    # mssql 記錄增加和刪除的文件
    insert_sql = "INSERT INTO Quotation_Update (FileName, Fold, CheckFold, Type,UpdateTime) VALUES (%s, %s, %s, %s, %s)"
    mssql_insert(insert_sql, add_data_list)
    mssql_insert(insert_sql, delete_data_list)

    #SQLite 更新
    sql_insert2 = 'INSERT INTO Files (FileName, Fold, CheckFold, FileUpdateTime) VALUES (?, ?, ?, ?)'
    commit(add_data_list2,sql_insert2)
    sql_delete2 = 'DELETE FROM Files WHERE FileName=? and Fold=?'
    commit(delete_data_list2,sql_delete2)

    sql_update=f"UPDATE config SET LastModify = '{updatetime}'"
    commit_SQL(sql_update)


sql_LastModify = 'SELECT * FROM config LIMIT 1'
directory_path,local, lastupdate, *other_config = read_data(sql_LastModify)

# first_update(directory_path)
update_Fold_File(directory_path,local, lastupdate)