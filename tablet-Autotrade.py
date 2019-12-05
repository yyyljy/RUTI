from PyQt4 import QtCore, QtGui
from PyQt4.QtCore import *
from PyQt4.QAxContainer import *
import datetime
import time
import sqlite3
import telegram
import sys
import pythoncom

today = datetime.datetime.today().strftime("%y%m%d")
today = "20" + today
pi_news = "pi-news"
tablet_apidb = "tablet-api&db"
tablet_trader = "tablet-trader"
Acclist = []
trlist = [[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0]]
portvalue = 1000000
RealAcc = "5080811710" ##실제투자
TestAcc = "8084364011" ##모의투자
status = TestAcc ##AutoSendOrder
EarningRate = 1.02

def SendTeleMessage(string, who):
    if who == "tablet-api&db":
        bot = telegram.Bot(token="255382445:AAEjVx8xzeusWLeaUhp2F1SiA9mIa7G1V_s")
    elif who == "pi-news":
        bot = telegram.Bot(token="196997404:AAEWruipuqcHcomrYcWVCwoetOjysxsAQ8s")
    elif who == "tablet-trader":
        bot = telegram.Bot(token="229690897:AAE8f5FIiumGqw1jceOV9uvuZ5BQWluS6yk")
    bot.send_message(chat_id=233869620, text=today + " " + string)

def GetTimenow():
    now = time.localtime()
    timenow = str(now[0]) + "/" + str(now[1]) + "/" + str(now[2]) + " " + str(now[3]) + ":" + str(now[4]) + ":" + str(now[5])
    return timenow

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
        Form.setWindowTitle("AutoTrader V1.4")
        Form.setObjectName("Form")
        Form.resize(705, 505)

        self.groupBox1 = QtGui.QGroupBox(Form)
        self.groupBox1.setTitle(" 오늘 매매 종목 ")
        self.groupBox1.setGeometry(QtCore.QRect(10, 15, 215, 215))
        self.groupBox1.setObjectName("groupBox1")
        self.formLayoutWidget = QtGui.QWidget(self.groupBox1)
        self.formLayoutWidget.setGeometry(QtCore.QRect(9, 20, 200, 200))
        self.formLayoutWidget.setObjectName("formLayoutWidget")
        self.formLayout = QtGui.QFormLayout(self.formLayoutWidget)
        self.formLayout.setObjectName("formLayout")

        self.groupBox2 = QtGui.QGroupBox(Form)
        self.groupBox2.setTitle(" 예 측 등 락 율  ")
        self.groupBox2.setGeometry(QtCore.QRect(10, 240, 215, 215))
        self.groupBox2.setObjectName("groupBox2")
        self.formLayoutWidget2 = QtGui.QWidget(self.groupBox2)
        self.formLayoutWidget2.setGeometry(QtCore.QRect(9, 20, 200, 200))
        self.formLayoutWidget2.setObjectName("formLayoutWidget")
        self.formLayout2 = QtGui.QFormLayout(self.formLayoutWidget2)
        self.formLayout2.setObjectName("formLayout2")

        self.PredictTextEdit = QtGui.QTextEdit(self.formLayoutWidget)
        self.PredictTextEdit.setGeometry(QtCore.QRect(5, 5, 187, 178))
        self.PredictTextEdit.setObjectName("PredictTextEdit")

        self.EarningTextEdit = QtGui.QTextEdit(self.formLayoutWidget2)
        self.EarningTextEdit.setGeometry(QtCore.QRect(5, 5, 187, 178))
        self.EarningTextEdit.setObjectName("EarningTextEdit")
        #self.EarningTextEdit.append("---오늘의 예측---")

        self.SystemTextEdit = QtGui.QTextEdit(Form)
        self.SystemTextEdit.setGeometry(QtCore.QRect(237, 40, 445, 402))
        self.SystemTextEdit.setObjectName("SystemTextEdit")
        self.SystemTextEdit.append(GetTimenow() + "\t프로그램 시작.")

        self.ACCBox = QtGui.QComboBox(Form)
        self.ACCBox.setGeometry(QtCore.QRect(237, 5, 110, 23))
        self.ACCBox.setObjectName("ACCBox")

        self.login()

    def login(self):
        ret = self.dynamicCall("CommConnect()")

    def quit(self):
        self.dynamicCall("CommTerminate()")
        sys.exit()

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        if sRQName == "매수주문":
            self.SystemTextEdit.append(GetTimenow() + "\t" + sScrNo + " : " + sMsg)
            SendTeleMessage(sScrNo + "\n" + sMsg, tablet_trader)
        if sRQName == "매도주문":
            self.SystemTextEdit.append(GetTimenow() + "\t" + sScrNo + " : " + sMsg)
            SendTeleMessage(sScrNo + "\n" + sMsg, tablet_trader)
            if sMsg.find("정상") != -1:
                i = 0
                while i < len(trlist):
                    if str(trlist[i][0]) == str(sScrNo):
                        trlist[i][5] = 0
                        break
                    i += 1
        else:
            self.SystemTextEdit.append(GetTimenow() + "\t" + sMsg)
            SendTeleMessage("OnReceiveMsg : " + sScrNo + "\n" + sMsg + "//" + sRQName, tablet_trader)

    def OnEventConnect(self, nErrCode):
        if nErrCode == 0:
            self.SystemTextEdit.append(GetTimenow() + "\t로그인 성공.")
            ACC_NO = self.dynamicCall('GetLoginInfo("ACCNO")')
            ACClist = ACC_NO.split(';')
            self.ACCBox.addItems(ACClist[:len(ACClist) - 1])
            SendTeleMessage("Autotrade 시작.", tablet_trader)
            self.TomorrowPredict()
            self.OrderBuy()
        else:
            self.SystemTextEdit.append(GetTimenow() + "\t로그인 실패.")

    def OnReceiveTrData(self, sScrNo, sRQName, sTrCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        if sRQName == "수익율요청":
            table = self.dynamicCall('GetCommDataEx(QString,QString)', sTrCode, sRQName)
            i = 0
            while i < len(trlist):
                if str(trlist[i][0]) == str(sScrNo):
                    j = 0
                    while j < len(table):
                        if str(table[j][1]) == str(sScrNo):
                            if int(trlist[i][5]) == int(table[j][6]):
                                trlist[i][2] = int(table[j][4])
                                trlist[i][3] = round((trlist[i][2] * EarningRate), -1)
                                self.AutoSendOrder(str(sScrNo), int(trlist[i][5]), int(trlist[i][3]), 2)
                        j += 1
                i += 1
        elif sRQName == "매수주문":
            ordernum = self.dynamicCall('GetCommData(QString,QString,int,QString)', sTrCode, sRQName, 0, "주문번호")
            test = self.dynamicCall('GetCommDataEx(QString,QString)', sTrCode, sRQName)

    def OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        ordernum = self.GetChjanData(9203)
        name = self.GetChjanData(302)
        orderqty = self.GetChjanData(900)
        price = self.GetChjanData(901)
        getqty = self.GetChjanData(930)
        code = self.GetChjanData(9001)
        code = code[1:]
        if (sGubun == "0") and (ordernum != "0"):
            self.SystemTextEdit.append(GetTimenow() + "\t" + code + " " + orderqty + "주 주문")
        if sGubun == "1":
            self.SystemTextEdit.append(GetTimenow() + "\t" + code + " 체결 " + getqty + "주 보유 중")
            i = 0
            while i < len(trlist):
                if trlist[i][0] == code:
                    trlist[i][5] = int(getqty.strip("'"))
                    if trlist[i][5] == trlist[i][4]:
                        SendTeleMessage("\t" + "Time to sell " + code, tablet_trader)
                        self.SystemTextEdit.append(GetTimenow() + "\t" + code + " 매도 주문")
                        self.GetEarningRate(code)
                i += 1

    def GetChjanData(self, nFid):
        chjang = self.dynamicCall('GetChejanData(QString)', nFid)
        return chjang

    def AutoSendOrder(self, Code, Qty, Price, Type):
        Code = str(Code)
        sScrNo = Code
        ACCNO = status ##RealAcc // TestAcc
        if Type == 1:
            print(Code + "\t" + str(int(Qty)) + "주 매수 주문")
            Order = self.dynamicCall('SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                                     ["매수주문", sScrNo, ACCNO, Type, Code, int(Qty), 0, "03", ""])
        elif Type == 2:
            print(Code + "\t" + str(int(Qty)) + "주 매도 주문")
            SendTeleMessage(Code + "\t" + str(int(Qty)) + "주 매도 주문", tablet_trader)
            Order = self.dynamicCall('SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                                     ["매도주문", sScrNo, ACCNO, Type, Code, int(Qty), Price, "00", ""])
        pythoncom.PumpWaitingMessages()

    def TomorrowPredict(self):
        list = self.GetCodeList()
        codei = 0
        self.EarningTextEdit.append("---오늘의 예측---")
        while codei < len(list):
            Code = list[codei]
            filename = "./db/" + str(Code) + ".db"
            file = sqlite3.connect(filename)
            cursor = file.cursor()
            try:
                cursor.execute("SELECT * FROM Dailytable ORDER BY Date ASC")
            except:
                codei += 1
                continue
            dbtuple = cursor.fetchall()
            file.close()
            day = len(dbtuple) - 1
            if(codei == 1):
                self.SystemTextEdit.append(GetTimenow() + "\t" + "조회일 : " + dbtuple[day][0])
            if len(dbtuple) < 267:
                codei += 1
                continue
            if (dbtuple[day][5] < 300000) or (dbtuple[day][2] >= 1000000) or (dbtuple[day][2] <= 1500):
                codei += 1
                continue
            statematrix = ["0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0\n" * 16]
            statematrix = statematrix[0].split('\n')
            i = 0
            while i < 16:
                statematrix[i] = statematrix[i].split(' ')
                i += 1
            i = 0
            j = 0
            while i < 16:
                while j < 16:
                    statematrix[i][j] = 0
                    j += 1
                j = 0
                i += 1
            i = 0
            statematrix.remove("")

            while i < len(dbtuple) - 2:
                if int(dbtuple[i][0]) < 20150615:
                    i += 1
                    continue
                if dbtuple[i][5] == 0:
                    i += 1
                    continue
                if dbtuple[i][2] <= int(dbtuple[i][1] * 0.76):
                    prevstate = 0
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 0.82):
                    prevstate = 1
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 0.88):
                    prevstate = 2
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 0.94):
                    prevstate = 3
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 0.955):
                    prevstate = 4
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 0.97):
                    prevstate = 5
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 0.985):
                    prevstate = 6
                elif dbtuple[i][2] <= int(dbtuple[i][1]):
                    prevstate = 7
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 1.015):
                    prevstate = 8
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 1.03):
                    prevstate = 9
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 1.045):
                    prevstate = 10
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 1.06):
                    prevstate = 11
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 1.12):
                    prevstate = 12
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 1.18):
                    prevstate = 13
                elif dbtuple[i][2] <= int(dbtuple[i][1] * 1.24):
                    prevstate = 14
                else:
                    prevstate = 15
                if dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 0.76):
                    statematrix[prevstate][0] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 0.82):
                    statematrix[prevstate][1] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 0.88):
                    statematrix[prevstate][2] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 0.94):
                    statematrix[prevstate][3] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 0.955):
                    statematrix[prevstate][4] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 0.97):
                    statematrix[prevstate][5] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 0.985):
                    statematrix[prevstate][6] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 1.00):
                    statematrix[prevstate][7] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 1.015):
                    statematrix[prevstate][8] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 1.03):
                    statematrix[prevstate][9] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 1.045):
                    statematrix[prevstate][10] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 1.06):
                    statematrix[prevstate][11] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 1.12):
                    statematrix[prevstate][12] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 1.18):
                    statematrix[prevstate][13] += 1
                elif dbtuple[i + 1][2] <= int(dbtuple[i + 1][1] * 1.24):
                    statematrix[prevstate][14] += 1
                else:
                    statematrix[prevstate][15] += 1
                i += 1
            codei += 1
            if dbtuple[day][2] <= int(dbtuple[day][1] * 0.76):
                state = 0
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 0.82):
                state = 1
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 0.88):
                state = 2
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 0.94):
                state = 3
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 0.955):
                state = 4
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 0.97):
                state = 5
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 0.985):
                state = 6
            elif dbtuple[day][2] <= int(dbtuple[day][1]):
                state = 7
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 1.015):
                state = 8
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 1.03):
                state = 9
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 1.045):
                state = 10
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 1.06):
                state = 11
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 1.12):
                state = 12
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 1.18):
                state = 13
            elif dbtuple[day][2] <= int(dbtuple[day][1] * 1.24):
                state = 14
            else:
                state = 15
            j = 0
            maxvalue = max(statematrix[state])
            total = sum(statematrix[state])
            while statematrix[state][j] != maxvalue:
                j += 1
            predic1 = j
            statematrix[state][j] = 0
            secvalue = max(statematrix[state])
            statematrix[state][j] = maxvalue
            j = 0
            while statematrix[state][j] != secvalue:
                j += 1
            predic2 = j
            if (4 <= predic1) or (predic1 <= 11):
                if secvalue == 0:
                    weight = 1 + (0.015 * (predic1 - 8))
                else:
                    weight = 1 + (0.015 * (predic1 - 8)) + (predic2 - predic1) * (secvalue / total) * 0.01
            else:
                if secvalue == 0:
                    if predic1 <= 3:
                        weight = 1 - (0.015 * 4) + (0.06 * (predic1 - 4))
                    else:
                        weight = 1 + (0.015 * 4) + (0.06 * (predic1 - 12))
                else:
                    if predic1 <= 3:
                        weight = 1 - (0.015 * 4) + (0.06 * (predic1 - 4)) + (predic2 - predic1) * (secvalue / total) * 0.01
                    else:
                        weight = 1 + (0.015 * 4) + (0.06 * (predic1 - 12)) + (predic2 - predic1) * (secvalue / total) * 0.01
            if weight >= 1.05:
                self.PredictTextEdit.append(str(Code) + "\t\\" + str(dbtuple[day][2]))
                index = 0
                while index < len(trlist):
                    if trlist[index][0] == 0:
                        trlist[index][0] = Code
                        trlist[index][1] = dbtuple[day][2]
                        trlist[index][4] = int(portvalue / dbtuple[day][2])
                        break
                    index += 1
                self.EarningTextEdit.append(str(Code) + "\t" + str(round(weight-1, 3) * 100) + "%")
            pythoncom.PumpWaitingMessages()
        self.SystemTextEdit.append(GetTimenow() + "\t" + "매매 종목 조회 완료.")
        message = "\n"
        i = 0
        while i < len(trlist):
            if trlist[i][0] == 0:
                break
            message += str(trlist[i][0]) + " " + str(trlist[i][1]) + "원 " + str(trlist[i][4]) + "주\n"
            i += 1
        if message == "\n":
            SendTeleMessage("오늘은 매매 종목이 없습니다.", tablet_trader)
        else:
            SendTeleMessage(message, tablet_trader)

    def GetCodeList(self):
        totallist = []
        kospi = self.dynamicCall("GetCodeListByMarket(QString)", ["0"])
        kospi_code_list = kospi.split(';')
        kospi_code_list = kospi_code_list[:len(kospi_code_list) - 1]
        kosdaq = self.dynamicCall("GetCodeListByMarket(QString)", ["10"])
        kosdaq_code_list = kosdaq.split(';')
        totallist = totallist + kospi_code_list + kosdaq_code_list
        totallist = totallist[:len(totallist) - 1]
        return totallist

    def OrderBuy(self):
        i = 0
        while i < len(trlist):
            if trlist[i][0] == 0:
                break
            self.AutoSendOrder(trlist[i][0], trlist[i][4], 0, 1)
            i += 1
            pythoncom.PumpWaitingMessages()

    def OrderSell(self):
        i = 0
        buy = 0
        sell = 0
        while i < len(trlist):
            if trlist[i][2] != 0:
                buy += 1
            if trlist[i][3] != 0:
                sell += 1
            i += 1

    def GetEarningRate(self, code):
        #ret = self.dynamicCall('SetInputValue(QString, QString)', "종목코드", str(Code))
        #ret = self.dynamicCall('SetInputValue(QString, QString)', "조회구분", "0")
        ret = self.dynamicCall('SetInputValue(QString, QString)', "계좌번호", status)
        #ret = self.dynamicCall('SetInputValue(QString, QString)', "비밀번호", "0000")
        #ret = self.dynamicCall('CommRqData(QString, QString, int, QString)', "수익율요청", "OPT10076", "0", str(Code))
        ret = self.dynamicCall('CommRqData(QString, QString, int, QString)', "수익율요청", "OPT10085", "0", str(code))

if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    Form = QtGui.QWidget()
    ui = MainWindow()
    ui.UIsetup(Form)
    Form.show()
    sys.exit(app.exec_())