from PyQt4 import QtCore, QtGui
from PyQt4.QtCore import *
from PyQt4.QAxContainer import *
import sys
import datetime
import time
import sqlite3
import telegram

bot = telegram.Bot(token="")
teleid = 000000000
today = datetime.datetime.today().strftime("%y%m%d")
today = "20" + today

def GetTimenow():
    now = time.localtime()
    timenow = str(now[0]) + "/" + str(now[1]) + "/" + str(now[2]) + " " + str(now[3]) + ":" + str(now[4]) + ":" + str(now[5])
    return timenow

def SendTeleMessage(string):
    bot.send_message(chat_id=teleid, text=today + " " + string)

def LogWriting(string):
    log = open("./log/log.txt", "a")
    log.write(today + "\t")
    log.write(string + "\n")
    log.close()

class MainWindow(QAxWidget):
    def __init__(self):
        super().__init__()
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')
        self.connect(self, SIGNAL("OnEventConnect(int)"), self.OnEventConnect)
        self.connect(self, SIGNAL("OnReceiveMsg(QString, QString, QString, QString)"), self.OnReceiveMsg)
        self.connect(self, SIGNAL(
            "OnReceiveTrData(QString, QString, QString, QString, QString, int, QString, QString, QString)"),
                     self.OnReceiveTrData)
        self.connect(self, SIGNAL("OnReceiveChejanData(QString, int, QString)"), self.OnReceiveChejanData)

    def UIsetup(self, Form):
        Form.setWindowTitle("DBUPDATE V1.1 ")
        Form.setObjectName("Form")
        Form.resize(350, 155)
        self.groupBox1 = QtGui.QGroupBox(Form)
        self.groupBox1.setTitle(" DB Update\t\t\t\t\t" + "실행시간\t\t"+ GetTimenow())
        self.groupBox1.setGeometry(QtCore.QRect(10, 10, 330, 140))
        self.groupBox1.setObjectName("groupBox1")
        self.formLayoutWidget = QtGui.QWidget(self.groupBox1)
        self.formLayoutWidget.setGeometry(QtCore.QRect(9, 20, 200, 200))
        self.formLayoutWidget.setObjectName("formLayoutWidget")
        self.formLayout = QtGui.QFormLayout(self.formLayoutWidget)
        self.formLayout.setObjectName("formLayout")

        self.DBLocLabel = QtGui.QLabel(self.formLayoutWidget)
        self.DBLocLabel.setText("Update Start")
        self.DBLocLabel.setObjectName("DBLocLabel")
        self.formLayout.setWidget(0, QtGui.QFormLayout.LabelRole, self.DBLocLabel)
        
        self.DBTimeLabel = QtGui.QLabel(self.formLayoutWidget)
        self.DBTimeLabel.setText(GetTimenow())
        self.DBTimeLabel.setObjectName("DBTimeLabel")
        self.formLayout.setWidget(1, QtGui.QFormLayout.LabelRole, self.DBTimeLabel)

        self.DBProcLabel = QtGui.QLabel(self.formLayoutWidget)
        self.DBProcLabel.setText("0/2400")
        self.DBProcLabel.setObjectName("DBProcLabel")
        self.formLayout.setWidget(2, QtGui.QFormLayout.LabelRole, self.DBProcLabel)

        self.DBProctimeLabel = QtGui.QLabel(self.formLayoutWidget)
        self.DBProctimeLabel.setText(GetTimenow())
        self.DBProctimeLabel.setObjectName("DBProctimeLabel")
        self.formLayout.setWidget(3, QtGui.QFormLayout.LabelRole, self.DBProctimeLabel)

        self.LoginLabel = QtGui.QLabel(self.formLayoutWidget)
        self.LoginLabel.setText("Login 시도중...")
        self.LoginLabel.setObjectName("LoginLabel")
        self.formLayout.setWidget(5, QtGui.QFormLayout.LabelRole, self.LoginLabel)

        self.login()

    def login(self):
        ret = self.dynamicCall("CommConnect()")

    def quit(self):
        self.dynamicCall("CommTerminate()")
        sys.exit()

    def OnEventConnect(self, nErrCode):
        if nErrCode == 0:
            SendTeleMessage("DB 갱신 시작")
            self.LoginLabel.setText("Login 완료")
            self.DBTimeLabel.setText(GetTimenow())
            codelist = self.GetCodeList()
        else:
            bot.send_message(chat_id=teleid, text=today + " 서버 연결에 실패했습니다.")
            time.sleep(3)
            self.login()
        i = 0
        while i < len(codelist):
            self.IntLat(i)
            self.DBProcLabel.setText(str(i) + "/" + str(len(codelist)))
            self.DBProctimeLabel.setText(GetTimenow())
            self.LoginLabel.setText(codelist[i])
            self.dynamicCall("SetInputValue(QString, QString)", "종목코드", codelist[i])
            self.dynamicCall("CommRQData(QString, QString, int, QString)", "UpdateRQ", "opt10081", 0, codelist[i])
            self.tr_event_loop = QEventLoop()
            self.tr_event_loop.exec_()
            i += 1
        SendTeleMessage("DB 업데이트 완료")
        self.quit()
        exit()

    def OnReceiveTrData(self, sScrNo, sRQName, sTrCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        if sRQName == "UpdateRQ":
            test = self.dynamicCall('GetCommDataEx(QString,QString)', sTrCode, sRQName)
            try:
                test[0][0] = sScrNo
                return self.UpdateDataBase(test)
            except:
                SendTeleMessage("\t"+ str(sScrNo) + "종목 pass 발생.")
                self.tr_event_loop.exit()

    def GetCodeList(self):
        totallist = []
        kospi = self.dynamicCall("GetCodeListByMarket(QString)", ["0"])
        kospi_code_list = kospi.split(';')
        kospi_code_list = kospi_code_list[:len(kospi_code_list)-1]
        kosdaq = self.dynamicCall("GetCodeListByMarket(QString)", ["10"])
        kosdaq_code_list = kosdaq.split(';')
        totallist = totallist + kospi_code_list + kosdaq_code_list
        totallist = totallist[:len(totallist)-1]
        return totallist

    def UpdateDataBase(self, list):
        file = sqlite3.connect('./db/' + list[0][0] + '.db')
        cursor = file.cursor()
        try:
            cursor.execute("SELECT * FROM Dailytable ORDER BY Date DESC")
            lastdata = cursor.fetchone()
            lastdate = lastdata[0]
            latestdate = list[0][4]
            if lastdate == latestdate:
                cursor.execute("UPDATE Dailytable SET Open = ?, Closing = ?, High = ?, Low = ?, Volume = ?, TrValue = ? WHERE Date = ?",
                        (list[0][5], list[0][1], list[0][6], list[0][7], list[0][2], list[0][3], list[0][4]))
                file.commit()
                file.close()
            else:
                i = 0
                while i < len(list):
                    if int(lastdate) >= int(list[i][4]):
                        break
                    i += 1
                j = 0
                while j < i:
                    cursor.execute("INSERT INTO Dailytable VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? )", (
                        list[j][4], list[j][5], list[j][1], list[j][6], list[j][7], list[j][2], list[j][3], 0, 0, 0, 0, 0, 0, 0,))
                    file.commit()
                    j += 1
        except:
            cursor.execute("CREATE TABLE Dailytable(Date text PRIMARY KEY, Open int, Closing int, High int, Low int, Volume int, TrValue int, Cap int, W52High int, W52Low int, MA5 int, MA20 int, MA60 int, MA120 int)")
            j = 0
            while j < len(list):
                cursor.execute("INSERT INTO Dailytable VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? )", (
                    list[j][4], list[j][5], list[j][1], list[j][6], list[j][7], list[j][2], list[j][3], 0, 0, 0, 0,
                    0, 0, 0,))
                file.commit()
                j += 1
        file.close()
        LogWriting(list[0][0])
        self.tr_event_loop.exit()

    def IntLat(self, i):
        interval = i % 5
        latency = 1
        if interval == 0:
            time.sleep(latency)

if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    Form = QtGui.QWidget()
    ui = MainWindow()
    ui.UIsetup(Form)
    Form.show()
    sys.exit(app.exec_())
