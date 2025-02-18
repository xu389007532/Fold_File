#################实现自动先把ui 文件转为python 文件#################开始.
import qt_ui_to_py
qt_ui_to_py.runMain()
#################实现自动先把ui 文件转为python 文件#################结束...
import ui_main_Search
from PyQt5 import QtCore
from PyQt5.QtGui import QStandardItemModel,QStandardItem
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QFileDialog, QAbstractItemView, QMessageBox, QInputDialog, QLineEdit
from PyQt5.QtCore import QModelIndex,Qt,QThread,pyqtSignal
import re
import sys
import os
import fnmatch
import shutil
from sqlite3 import connect
import sqlite3
import datetime
from Share import Honour_Share


def commit(data,sql):


    conn = sqlite3.connect("./Source/Client_FoldFiles.db")
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
        conn.rollback()  # 如果出现错误，回滚事务。
    finally:
        conn.close()  # 最后关闭连接。无论成功还是失败，都应关闭连接。

def list_Folds(directory):
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
            AllFiles.append((f,root.replace('\\', '/')))
            # print(f"{os.path.abspath(root)}/{f}")
    return AllFiles



def read_data(sql):
    bsdata = connect("./Source/Client_FoldFiles.db")
    cursor = bsdata.cursor()
    cursor.execute(sql)
    data1 = cursor.fetchone()
    cursor.close()
    bsdata.close()
    return data1

def select_data(sql):
    bsdata = connect("./Source/Client_FoldFiles.db")
    cursor = bsdata.cursor()
    cursor.execute(sql)
    data1 = cursor.fetchall()
    cursor.close()
    bsdata.close()
    return data1

class mainwindow(QMainWindow):
    """
    主程序窗口
    """
    def __init__(self, Window):
        super().__init__()
        self.ui = ui_main_Search.Ui_MainWindow()
        self.ui.setupUi(self)
        self.load_data()
        self.show()
        self.ui_event()


    def ui_event(self):
        """
        主窗口事件
        :return:
        """
        self.ui.pushButton_UploadToNewwork.clicked.connect(self.fun_UploadToNewwork)  #
        self.ui.lineEdit_Network.returnPressed.connect(self.fun_Find_network)
        self.ui.lineEdit_Local.returnPressed.connect(self.fun_Find_local)
        selectionModel = self.ui.tableView.selectionModel()
        selectionModel.selectionChanged.connect(self.on_selectionChanged)

        selectionModel2 = self.ui.tableView_2.selectionModel()
        selectionModel2.selectionChanged.connect(self.on_selectionChanged2)

        self.ui.pushButton_CopyToLocal.clicked.connect(self.fun_CopyToLocal)
        self.ui.pushButton_Dele_network.clicked.connect(self.fun_Dele_network)
        self.ui.actionupdate.triggered.connect(self.fun_updateConfig)
        self.ui.actionDelete_File.triggered.connect(self.fun_Delete_File_config)


    def load_data(self):
        userID, userPWD, serverName, databaseName = Honour_Share.Py_Decrypto(os.path.abspath('./Source/userCommon.xml'), os.path.abspath('./Source/Honour.dll'))
        sql='SELECT * FROM config LIMIT 1'
        self.network,self.local,self.LastModify, *other_config=read_data(sql)
        self.other_config=other_config
        if list(self.other_config).__contains__('Delete_File_No'):
            self.ui.pushButton_Dele_network.setEnabled(False)
        else:
            self.ui.pushButton_Dele_network.setEnabled(True)
        if not os.path.exists(self.local):
            os.makedirs(self.local)
        # 檢查MSSQL Quotation_Update 更新的數據， 如有， 再更新返本地 start
        sql_str=f"SELECT FileName,Fold,CheckFold,Type,UpdateTime FROM Quotation_Update where UpdateTime>'{self.LastModify}' order by id"
        update_data=Honour_Share.read_sql_fetchall(sql_str,userID, userPWD,serverName,databaseName)
        for ud in update_data:
            if ud[3]=='Delete':
                delete_data_list2=[(ud[0],ud[1])]
                sql_delete2 = 'DELETE FROM Files WHERE FileName=? and Fold=?'
                commit(delete_data_list2, sql_delete2)
                print('Delete file: ',ud[0])

            elif ud[3]=='Add':
                add_data_list2=[(ud[0],ud[1],ud[2],ud[4])]
                sql_insert2 = 'INSERT INTO Files (FileName, Fold, CheckFold, FileUpdateTime) VALUES (?, ?, ?, ?)'
                commit(add_data_list2, sql_insert2)
                print('Add file: ', ud[0])

        updatetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sql_update = f"UPDATE config SET LastModify = '{updatetime}'"
        commit([()],sql_update)
        # 檢查MSSQL Quotation_Update 更新的數據， 如有， 再更新返本地 end

        self.ui.lineEdit_Network_path.setText(self.network)
        self.ui.lineEdit_Local_path.setText(self.local)

        self.model_view = QStandardItemModel()
        self.model_view.setHorizontalHeaderLabels(['文件名','更新時間','路徑','','檢查路徑'])
        self.ui.tableView.setModel(self.model_view)

        local_file=list_files(self.local)
        self.model_view2 = QStandardItemModel()
        self.model_view2.setHorizontalHeaderLabels(['文件名','路徑'])
        self.ui.tableView_2.setModel(self.model_view2)
        for row, list_value in enumerate(local_file):
            for col, v in enumerate(list_value):
                self.model_view2.setItem(row, col, QStandardItem(v))

        # self.ui.tableView.setModel(self.model_view)
        # self.ui.process_tableView2.resizeColumnToContents(0)
        self.ui.tableView_2.resizeColumnsToContents()
        self.ui.tableView_2.resizeRowsToContents()

    def fun_Delete_File_config(self):
        Dele_File_config,okPressed=QInputDialog.getText(None,"是否啟用刪除文件功能!", "是否啟用刪除文件功能. 請輸入 Yes 或 No",QLineEdit.Normal,"")
        if Dele_File_config and okPressed:
            if Dele_File_config=='Yes':
                data = [('Delete_File_Yes',)]
            else:
                data = [('Delete_File_No',)]
            update_config_sql = 'UPDATE config SET Delete_File = ?'
            commit(data, update_config_sql)
        sql='SELECT * FROM config LIMIT 1'
        self.network,self.local,self.LastModify, *other_config=read_data(sql)
        self.other_config=other_config
        if list(self.other_config).__contains__('Delete_File_No'):
            self.ui.pushButton_Dele_network.setEnabled(False)
        else:
            self.ui.pushButton_Dele_network.setEnabled(True)

    def fun_updateConfig(self):
        self.network=self.ui.lineEdit_Network_path.text()
        self.local=self.ui.lineEdit_Local_path.text()
        data=[(self.network,self.local)]
        update_config_sql='UPDATE config SET Network_Flod = ?, Local_Flod = ?'
        commit(data,update_config_sql)
        if not os.path.exists(self.local):
            os.makedirs(self.local)

    def on_selectionChanged2(self):  # 本地數據
        indexes=self.ui.tableView_2.selectionModel().selectedIndexes()
        rows = self.ui.tableView_2.selectionModel().selectedRows()
        self.select_item2=[]
        for r in rows:
            index0 = self.model_view2.index(r.row(), 0)
            index1 = self.model_view2.index(r.row(),1)
            self.select_item2.append((self.model_view2.data(index0),self.model_view2.data(index1)))
            # print(self.model_view.data(index0),self.model_view.data(index2))
            # print(r.data())
        print(self.select_item2)

    def on_selectionChanged(self, selected, deselected):  # 使用選擇改變信號來獲取數據
        indexes=self.ui.tableView.selectionModel().selectedIndexes()
        rows = self.ui.tableView.selectionModel().selectedRows()
        self.select_item=[]
        for r in rows:
            index0 = self.model_view.index(r.row(), 0)
            index2 = self.model_view.index(r.row(),2)
            self.select_item.append((self.model_view.data(index2),self.model_view.data(index0)))
            # print(self.model_view.data(index0),self.model_view.data(index2))
            # print(r.data())
        print(self.select_item)

        # print(f"Selected item: {self.model_view.itemFromIndex(indexes[0]).text()}{self.model_view.itemFromIndex(indexes[2]).text()}")
        # for index in indexes:
        #     print(f"Selected item: {self.model_view.itemFromIndex(index).text()}")
        # print("selected: ")
        # for ix in selected.indexes():
        #     print(ix.data())

        # print("deselected: ")
        # for ix in deselected.indexes():
        #     print(ix.data())

    def fun_Find_local(self):
        key='*'+self.ui.lineEdit_Local.text()+'*'
        AllFiles = []
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for root, dirs, files in os.walk(self.local):
            # print('root:',root,'dirs:',dirs)

            for f in files:
                if fnmatch.fnmatch(f, key):
                    AllFiles.append((f, root.replace('\\', '/')))
                # print(f"{os.path.abspath(root)}/{f}")
        self.model_view2.clear()
        self.model_view2.setHorizontalHeaderLabels(['文件名', '路徑'])
        for row, list_value in enumerate(AllFiles):
            for col, v in enumerate(list_value):
                self.model_view2.setItem(row, col, QStandardItem(v))

        # self.ui.tableView.setModel(self.model_view)
        # self.ui.process_tableView2.resizeColumnToContents(0)
        self.ui.tableView_2.resizeColumnsToContents()
        self.ui.tableView_2.resizeRowsToContents()

    def fun_Find_network(self):
        self.model_view.clear()
        self.model_view.setHorizontalHeaderLabels(['文件名', '更新時間', '路徑', '', '檢查路徑'])
        find=self.ui.lineEdit_Network.text()
        find=find.replace('*','%').replace('?','_')
        if find!='':
            sql="SELECT * FROM Files WHERE FileName like '%"+find+"%'"
            data1=select_data(sql)

            for row, list_value in enumerate(data1):
                for col, v in enumerate(list_value):
                    self.model_view.setItem(row, col, QStandardItem(v))

            # self.ui.tableView.setModel(self.model_view)
            # self.ui.process_tableView2.resizeColumnToContents(0)
            self.ui.tableView.resizeColumnsToContents()
            self.ui.tableView.resizeRowsToContents()

            # print(data1)

    def re_fold(self,data):
        # data = 'Cost-QNM034290-R1 4097773-2 SBA Easter 4 Cards （20241122 LL）.xlsx'
        # data = 'QNM033879-R2  4092633-2 IFG-STML Christmas Cards Acquisiton (240906 SF).xlsx'
        re_QNM = re.compile(r'\S*QNM([0-9]{6})\S*', re.IGNORECASE)
        re_QN = re.compile(r'\S*QN([0-9]{6})\S*', re.IGNORECASE)
        # print('QNM:', re_QNM.findall(data))
        # print('QN:', re_QN.findall(data))
        QNM = re_QNM.findall(data)
        if QNM:
            QNM_number = int(QNM[0])
            fold = 'QNM' + (str(int(QNM_number / 1000)) + '001').rjust(6, '0') + ' to ' + 'QNM' + (str(int(33879 / 1000) + 1) + '000').rjust(6, '0')
            print('Fold:', fold)
            return fold
        else:
            QMessageBox.information(self, "提示!", "文件名找不到QNM編號, 此文件會放在QNM文件夾裏. 請注意." +'\r\n'+ data)
            return 'QNM'
    def fun_UploadToNewwork(self):

        start=datetime.datetime.now()
        for si in self.select_item2:
            fold=self.re_fold(si[0])
            mk_fold=self.network+'/'+fold
            if not os.path.exists(mk_fold):
                os.makedirs(mk_fold)

            src=si[1]+'/'+si[0]
            dst=self.network+'/'+fold+'/'+si[0]
            if os.path.exists(src):
                shutil.copyfile(src,dst)
            else:
                QMessageBox.information(self, "錯誤提示!", "本地文件找不到, 不能上傳此文件, 請檢查: "+src)
        end = datetime.datetime.now()
        time=(end-start).seconds
        self.ui.label_Status.setText("狀態: 上傳文件共需時間 "+ time.__str__()+' 秒.')

    def fun_CopyToLocal(self):


        start=datetime.datetime.now()

        for si in self.select_item:
            src=si[0]+'/'+si[1]
            if self.other_config[0] == 'download_AddFold_Today':
                today = start.date().strftime("%Y-%m-%d")
                down_fold=self.local + '/' + today
                if not os.path.exists(down_fold):
                    os.makedirs(down_fold)
                dst = down_fold + '/' + si[1]
            else:
                dst=self.local+'/'+si[1]
            if os.path.exists(src):
                shutil.copyfile(src,dst)
            else:
                QMessageBox.information(self, "錯誤提示!", "共享盤文件找不到, 不能下載此文件, 請檢查: "+src)
        end = datetime.datetime.now()
        time=(end-start).seconds
        self.ui.label_Status.setText("狀態: 下載文件共需時間 "+ time.__str__()+' 秒.')

    def fun_Dele_network(self):
        start = datetime.datetime.now()
        for si in self.select_item:
            filename=si[0]+'/'+si[1]
            if os.path.exists(filename):
                os.remove(filename)
            else:
                QMessageBox.information(self, "錯誤提示!", "共享盤文件找不到, 不能刪除此文件, 請檢查: "+filename)

        end = datetime.datetime.now()
        time=(end-start).seconds
        self.ui.label_Status.setText("狀態: 刪除文件共需時間 "+ time.__str__()+' 秒.')

if __name__ == "__main__":
    Honour_Share.kill_process('PyApp_')
    app = QApplication([])
    check_status = 0
    Honour_Share.update_ver("./ver_Search.txt")
    ui_mainwindow = mainwindow(QMainWindow())
    sys.exit(app.exec_())

