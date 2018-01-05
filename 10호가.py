import sys
import time
import pythoncom
import win32com.client
import threading
import pandas
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from PyQt5 import QtWidgets
from PyQt5 import QtCore  # QtCore를 명시적으로 보여주기 위해
from pandas import Series, DataFrame
import locale
# cp object
g_instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")  #1
g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr") #2
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")  #3
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil") #4
locale.setlocale(locale.LC_ALL, '')

class TestThread(QThread):
    # 쓰레드의 커스텀 이벤트
    # 데이터 전달 시 형을 명시해야 함
    threadEvent = QtCore.pyqtSignal(int)
    
    def __init__(self, parent=None):
        super().__init__()
        self.n = 0
        self.main = parent
        self.isRun = False
 
    def run(self):
        while self.isRun:
            print('쓰레드 : ' + str(self.n))
 
            # 'threadEvent' 이벤트 발생
            # 파라미터 전달 가능(객체도 가능)
            self.threadEvent.emit(self.n)
            self.n += 1
            self.sleep(1)

class stockPricedData:
    def __init__(self):
        self.dicEx = {ord('0'): "동시호가/장중 아님", ord('1'): "동시호가", ord('2'): "장중"}
        self.code = ""
        self.name = ""
        self.cur = 0        # 현재가
        self.diff = 0       # 대비
        self.diffp = 0      # 대비율
        self.offer = [0 for _ in range(10)]     # 매도호가
        self.bid = [0 for _ in range(10)]       # 매수호가
        self.offervol = [0 for _ in range(10)]     # 매도호가 잔량
        self.bidvol = [0 for _ in range(10)]       # 매수호가 잔량
        self.totOffer = 0       # 총매도잔량
        self.totBid = 0         # 총매수 잔량
        self.vol = 0            # 거래량
        self.tvol = 0           # 순간 체결량
        self.baseprice = 0      # 기준가
        self.high = 0
        self.low = 0
        self.open = 0
        self.volFlag = ord('0')  # 체결매도/체결 매수 여부
        self.time = 0
        self.sum_buyvol = 0
        self.sum_sellvol = 0
        self.vol_str = 0
        
        # 예상체결가 정보
        self.exFlag= ord('2')
        self.expcur = 0         # 예상체결가
        self.expdiff = 0        # 예상 대비
        self.expdiffp = 0       # 예상 대비율
        self.expvol = 0         # 예상 거래량
        self.objCur = CpPBStockCur()
        self.objOfferbid = CpPBStockBid()

    def __del__(self):
        self.objCur.Unsubscribe()
        self.objOfferbid.Unsubscribe()


    # 전일 대비 계산
    def makediffp(self):
        lastday = 0
        if (self.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            if self.baseprice > 0  :
                lastday = self.baseprice
            else:
                lastday = self.expcur - self.expdiff
            if lastday:
                self.expdiffp = (self.expdiff / lastday) * 100
            else:
                self.expdiffp = 0
        else:
            if self.baseprice > 0  :
                lastday = self.baseprice
            else:
                lastday = self.cur - self.diff
            if lastday:
                self.diffp = (self.diff / lastday) * 100
            else:
                self.diffp = 0

    def getCurColor(self):
        diff = self.diff
        if (self.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            diff = self.expdiff
        if (diff > 0):
            return 'color: red'
        elif (diff == 0):
            return  'color: black'
        elif (diff < 0):
            return 'color: blue'
        
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, rpMst, parent):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.parent = parent  # callback 을 위해 보관
        self.rpMst = rpMst


    # PLUS 로 부터 실제로 시세를 수신 받는 이벤트 핸들러
    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량
 
        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)
 
        
        if self.name == "stockcur":
            # 현재가 체결 데이터 실시간 업데이트
            self.rpMst.exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            code = self.client.GetHeaderValue(0)
            diff = self.client.GetHeaderValue(2)
            cur= self.client.GetHeaderValue(13)  # 현재가
            vol = self.client.GetHeaderValue(9)  # 거래량

            # 예제는 장중만 처리 함.
            if (self.rpMst.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
                # 예상체결가 정보
                self.rpMst.expcur = cur
                self.rpMst.expdiff = diff
                self.rpMst.expvol = vol
            else:
                self.rpMst.cur = cur
                self.rpMst.diff = diff
                self.rpMst.makediffp()
                self.rpMst.vol = vol
                self.rpMst.open = self.client.GetHeaderValue(4)
                self.rpMst.high = self.client.GetHeaderValue(5)
                self.rpMst.low = self.client.GetHeaderValue(6)
                self.rpMst.tvol = self.client.GetHeaderValue(17)
                self.rpMst.volFlag = self.client.GetHeaderValue(14)  # '1'  매수 '2' 매도
                self.rpMst.time = self.client.GetHeaderValue(18)
                self.rpMst.sum_buyvol = self.client.GetHeaderValue(16)  #누적매수체결수량 (체결가방식)
                self.rpMst.sum_sellvol = self.client.GetHeaderValue(15) #누적매도체결수량 (체결가방식)
                if (self.rpMst.sum_sellvol) :
                    self.rpMst.volstr = self.rpMst.sum_buyvol / self.rpMst.sum_sellvol * 100
                else :
                    self.rpMst.volstr = 0

            self.rpMst.makediffp()
            # 현재가 업데이트
            self.parent.monitorPriceChange()

            return

        elif self.name == "stockbid":
            # 현재가 10차 호가 데이터 실시간 업데이c
            code = self.client.GetHeaderValue(0)
            dataindex = [3, 7, 11, 15, 19, 27, 31, 35, 39, 43]
            obi = 0
            for i in range(10):
                self.rpMst.offer[i] = self.client.GetHeaderValue(dataindex[i])
                self.rpMst.bid[i] = self.client.GetHeaderValue(dataindex[i] + 1)
                self.rpMst.offervol[i] = self.client.GetHeaderValue(dataindex[i] + 2)
                self.rpMst.bidvol[i] = self.client.GetHeaderValue(dataindex[i] + 3)

            self.rpMst.totOffer = self.client.GetHeaderValue(23)
            self.rpMst.totBid = self.client.GetHeaderValue(24)
            # 10차 호가 변경 call back 함수 호출
            self.parent.monitorOfferbidChange()
            return
    

# CpStockCur: 실시간 현재가 요청 클래스
class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()
 
    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()


# Cp6033 : 주식 잔고 조회
class Cp6033:
    def __init__(self):
        # 통신 OBJECT 기본 세팅
        self.objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = self.objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            return
 
        # 
        acc = self.objTrade.AccountNumber[0]  # 계좌번호
        accFlag = self.objTrade.GoodsList(acc, 1)  # 주식상품 구분
        #print(acc, accFlag[0])
 
        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  #  요청 건수(최대 50)
        
 
    # 실제적인 6033 통신 처리
    def rq6033(self, retcode):
        self.objRq.BlockRequest()
 
        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
 
        cnt = self.objRq.GetHeaderValue(7)
        print(cnt)
 
        print("종목코드 종목명 신용구분 체결잔고수량 체결장부단가 평가금액 평가손익")
        for i in range(cnt):
            code = self.objRq.GetDataValue(12, i)  # 종목코드
            name = self.objRq.GetDataValue(0, i)  # 종목명
            retcode.append(code)
            if len(retcode) >=  200:       # 최대 200 종목만,
                break
            cashFlag = self.objRq.GetDataValue(1, i)  # 신용구분
            date = self.objRq.GetDataValue(2, i)  # 대출일
            amount = self.objRq.GetDataValue(7, i) # 체결잔고수량
            buyPrice = self.objRq.GetDataValue(17, i) # 체결장부단가
            evalValue = self.objRq.GetDataValue(9, i) # 평가금액(천원미만은 절사 됨)
            evalPerc = self.objRq.GetDataValue(11, i) # 평가손익
 
            print(code, name, cashFlag, amount, buyPrice, evalValue, evalPerc)
 
    def Request(self, retCode):
        self.rq6033(retCode)
 
        # 연속 데이터 조회 - 200 개까지만.
        while self.objRq.Continue:
            self.rq6033(retCode)
            print(len(retCode))
            if len(retCode) >= 200:
                break
        # for debug
        size = len(retCode)
        for i in range(size):
            print(retCode[i])
        return True


class CpMarketEye:
    def Request(self, codes, rqField):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
 
        # 관심종목 객체 구하기
        objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        # 요청 필드 세팅 - 종목코드, 종목명, 시간, 대비부호, 대비, 현재가, 거래량
        # rqField = [0,17, 1,2,3,4,10]
        objRq.SetInputValue(0, rqField) # 요청 필드
        objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        objRq.BlockRequest()
 
 
        # 현재가 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()
        rqRet = objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
 
        cnt  = objRq.GetHeaderValue(2)
 
        for i in range(cnt):
            rpCode = objRq.GetDataValue(0, i)  # 코드
            rpName = objRq.GetDataValue(1, i)  # 종목명
            rpTime= objRq.GetDataValue(2, i)  # 시간
            rpDiffFlag = objRq.GetDataValue(3, i)  # 대비부호
            rpDiff = objRq.GetDataValue(4, i)  # 대비
            rpCur = objRq.GetDataValue(5, i)  # 현재가
            rpVol = objRq.GetDataValue(6, i)  # 거래량
            print(rpCode, rpName, rpTime,  rpDiffFlag, rpDiff, rpCur, rpVol)
 
        return True
  
# SB/PB 요청 ROOT 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False
 
    def Subscribe(self, var, rpMst, parent):
        if self.bIsSB:
            self.Unsubscribe()
 
        if (len(var) > 0):
            self.obj.SetInputValue(0, var)
 
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, rpMst, parent)
        self.obj.Subscribe()
        self.bIsSB = True
 
    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False
 
# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__("stockcur", "DsCbo1.StockCur")
 
# CpPBStockBid: 실시간 10차 호가 요청 클래스
class CpPBStockBid(CpPublish):
    def __init__(self):
        super().__init__("stockbid", "Dscbo1.StockJpBid")
 
 
# SB/PB 요청 ROOT 클래스
class CpPBConnection:
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpUtil.CpCybos")
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, "connection", None)
 
 
# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPCurrentPrice:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        return
 
 
    def Request(self, code, rtMst, callbackobj):
        # 현재가 통신
        rtMst.objCur.Unsubscribe()
        rtMst.objOfferbid.Unsubscribe()
 
        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print("통신상태", self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False
 
 
        # 수신 받은 현재가 정보를 rtMst 에 저장
        rtMst.code = code
        rtMst.name = g_objCodeMgr.CodeToName(code)
        rtMst.cur =  self.objStockMst.GetHeaderValue(11)  # 종가
        rtMst.diff =  self.objStockMst.GetHeaderValue(12)  # 전일대비
        rtMst.baseprice  =  self.objStockMst.GetHeaderValue(27)  # 기준가
        rtMst.vol = self.objStockMst.GetHeaderValue(18)  # 거래량
        rtMst.exFlag = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        rtMst.expcur = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        rtMst.expdiff = self.objStockMst.GetHeaderValue(56)  # 예상체결대비
        rtMst.makediffp()
 
        rtMst.totOffer = self.objStockMst.GetHeaderValue(71)  # 총매도잔량
        rtMst.totBid = self.objStockMst.GetHeaderValue(73)  # 총매수잔량
 
 
        # 10차호가
        for i in range(10):
            rtMst.offer[i] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            rtMst.bid[i] = (self.objStockMst.GetDataValue(1, i) ) # 매수호가
            rtMst.offervol[i] = (self.objStockMst.GetDataValue(2, i))  # 매도호가 잔량
            rtMst.bidvol[i] = (self.objStockMst.GetDataValue(3, i) ) # 매수호가 잔량
 
 
        rtMst.objCur.Subscribe(code,rtMst, callbackobj)
        rtMst.objOfferbid.Subscribe(code,rtMst, callbackobj)
 
 


# CpWeekList:  일자별 리스트 구하기
class CpWeekList:
    def __init__(self):
        self.objWeek = win32com.client.Dispatch("Dscbo1.StockWeek")
        return


    def Request(self, code, caller):
        # 현재가 통신
        self.objWeek.SetInputValue(0, code)
        # 데이터들
        dates = []
        opens = []
        highs = []
        lows = []
        closes = []
        diffs = []
        vols = []
        diffps = []
        foreign_vols = []
        foreign_diff = []
        foreign_p = []

        # 누적 개수 - 100 개까지만 하자
        sumCnt = 0
        while True:
            ret = self.objWeek.BlockRequest()
            if self.objWeek.GetDibStatus() != 0:
                print("통신상태", self.objWeek.GetDibStatus(), self.objWeek.GetDibMsg1())
                return False

            cnt = self.objWeek.GetHeaderValue(1)
            sumCnt += cnt
            if cnt == 0:
                break

            for i in range(cnt):
                dates.append(self.objWeek.GetDataValue(0, i))
                opens.append(self.objWeek.GetDataValue(1, i))
                highs.append(self.objWeek.GetDataValue(2, i))
                lows.append(self.objWeek.GetDataValue(3, i))
                closes.append(self.objWeek.GetDataValue(4, i))

                temp = self.objWeek.GetDataValue(5, i)
                diffs.append(temp)
                vols.append(self.objWeek.GetDataValue(6, i))

                temp2 = self.objWeek.GetDataValue(10, i)
                if (temp < 0):
                    temp2 *= -1
                diffps.append(temp2)

                foreign_vols.append(self.objWeek.GetDataValue(7, i)) # 외인보유
                foreign_diff.append(self.objWeek.GetDataValue(8, i)) # 외인보유 전일대비
                foreign_p.append(self.objWeek.GetDataValue(9, i)) # 외인비중

            if (sumCnt > 100):
                break

            if self.objWeek.Continue == False:
                break

        if len(dates) == 0:
            return False

        caller.rpWeek = None
        weekCol = {'close': closes,
                   'diff':  diffs,
                   'diffp': diffps,
                    'vol': vols,
                    'open':opens,
                    'high': highs,
                    'low': lows,
                    'for_v' : foreign_vols,
                    'for_d': foreign_diff,
                    'for_p': foreign_p,
                   }
        caller.rpWeek = DataFrame(weekCol, index=dates)
        return True


# CpStockBid:  시간대별 조회
class CpStockBid:
    def __init__(self):
        self.objSBid = win32com.client.Dispatch("Dscbo1.StockBid")
        return


    def Request(self, code, caller):
        # 현재가 통신
        self.objSBid.SetInputValue(0, code)
        self.objSBid.SetInputValue(2, 80)  # 요청개수 (최대 80)
        self.objSBid.SetInputValue(3, ord('C'))  # C 체결가 비교 방식 H 호가 비교방식

        times = []
        curs = []
        diffs = []
        tvols = []
        offers = []
        bids = []
        vols = []
        offerbidFlags = [] # 체결 상태 '1' 매수 '2' 매도
        volstrs = [] # 체결강도
        marketFlags = [] # 장구분 '1' 동시호가 예상체결' '2' 장중

        # 누적 개수 - 100 개까지만 하자
        sumCnt = 0
        while True:
            ret = self.objSBid.BlockRequest()
            if self.objSBid.GetDibStatus() != 0:
                print("통신상태", self.objSBid.GetDibStatus(), self.objSBid.GetDibMsg1())
                return False

            cnt = self.objSBid.GetHeaderValue(2)
            sumCnt += cnt
            if cnt == 0:
                break

            strcur = ""
            strflag = ""
            strflag2 = ""
            for i in range(cnt):
                cur = self.objSBid.GetDataValue(4, i)
                times.append(self.objSBid.GetDataValue(9, i))
                diffs.append(self.objSBid.GetDataValue(1, i))
                vols.append(self.objSBid.GetDataValue(5, i))
                tvols.append(self.objSBid.GetDataValue(6, i))
                offers.append(self.objSBid.GetDataValue(2, i))
                bids.append(self.objSBid.GetDataValue(3, i))
                flag = self.objSBid.GetDataValue(7, i)
                if (flag == ord('1')):
                    strflag = "체결매수"
                else:
                    strflag = "체결매도"
                offerbidFlags.append(strflag)
                volstrs.append(self.objSBid.GetDataValue(8, i))
                flag = self.objSBid.GetDataValue(10, i)
                if (flag == ord('1')):
                    strflag2 = "예상체결"
                    #strcur = '*' + str(cur)
                else:
                    strflag2 = "장중"
                    #strcur = str(cur)
                marketFlags.append(strflag2)
                curs.append(cur)


            if (sumCnt > 100):
                break

            if self.objSBid.Continue == False:
                break

        if len(times) == 0:
            return False

        caller.rpStockBid = None
        sBidCol = {'time': times,
                   'cur':  curs,
                   'diff': diffs,
                    'vol': vols,
                    'tvol':tvols,
                    'offer': offers,
                    'bid': bids,
                    'flag': offerbidFlags,
                    'market': marketFlags,
                    'volstr': volstrs}
        caller.rpStockBid = DataFrame(sBidCol)
        print(caller.rpStockBid)
        return True


 
class Form(QtWidgets.QDialog):
    def __init__(self):
        #QtWidgets.QDialog.__init__(self, parent)
        super().__init__()
        self.ui = uic.loadUi("hoga.ui", self)
        self.ui.pushButton_2.clicked.connect(self.threadStart)
        self.ui.show()
        self.objMst = CpRPCurrentPrice()
        self.item = stockPricedData()
 
        self.setCode("000660")
 
        self.th = TestThread(self)
        self.th.threadEvent.connect(self.threadEventHandler)
        
    def threadStart(self):  
        if not self.th.isRun:
            print('메인 : 쓰레드 시작')
            self.th.isRun = True
            self.th.start()

    @pyqtSlot()
    def threadStop(self):
        if self.th.isRun:
            print('메인 : 쓰레드 정지')
            self.th.isRun = False
            
    # 쓰레드 이벤트 핸들러
    # 장식자에 파라미터 자료형을 명시
    @pyqtSlot(int)
    def threadEventHandler(self):
        print("handler")
        
    @pyqtSlot()
    def slot_codeupdate(self):
        print("codeupdate")
        code = self.ui.editCode.toPlainText()
        self.setCode(code)
 
    def slot_codechanged(self):
        print("codechange")
        code = self.ui.editCode.toPlainText()
        self.setCode(code)
 
 
    def monitorPriceChange(self):
        self.displyHoga()
 
    def monitorOfferbidChange(self):
        self.displyHoga()
 
    def setCode(self, code):
        if len(code) < 6 :
            return
 
        print(code)
        if not (code[0] == "A"):
            code = "A" + code
 
        name = g_objCodeMgr.CodeToName(code)
        print(name)
        if len(name) == 0:
            print("종목코드 확인")
            return
 
        self.ui.label_name.setText(name)
 
        if (self.objMst.Request(code, self.item, self) == False):
            return
        self.displyHoga()
 
    def displyHoga(self):
        self.ui.label_offer10.setText(format(self.item.offer[9],','))
        self.ui.label_offer9.setText(format(self.item.offer[8],','))
        self.ui.label_offer8.setText(format(self.item.offer[7],','))
        self.ui.label_offer7.setText(format(self.item.offer[6],','))
        self.ui.label_offer6.setText(format(self.item.offer[5],','))
        self.ui.label_offer5.setText(format(self.item.offer[4],','))
        self.ui.label_offer4.setText(format(self.item.offer[3],','))
        self.ui.label_offer3.setText(format(self.item.offer[2],','))
        self.ui.label_offer2.setText(format(self.item.offer[1],','))
        self.ui.label_offer1.setText(format(self.item.offer[0],','))
 
        self.ui.label_offer_v10.setText(format(self.item.offervol[9],','))
        self.ui.label_offer_v9.setText(format(self.item.offervol[8],','))
        self.ui.label_offer_v8.setText(format(self.item.offervol[7],','))
        self.ui.label_offer_v7.setText(format(self.item.offervol[6],','))
        self.ui.label_offer_v6.setText(format(self.item.offervol[5],','))
        self.ui.label_offer_v5.setText(format(self.item.offervol[4],','))
        self.ui.label_offer_v4.setText(format(self.item.offervol[3],','))
        self.ui.label_offer_v3.setText(format(self.item.offervol[2],','))
        self.ui.label_offer_v2.setText(format(self.item.offervol[1],','))
        self.ui.label_offer_v1.setText(format(self.item.offervol[0],','))
 
        self.ui.label_bid10.setText(format(self.item.bid[9],','))
        self.ui.label_bid9.setText(format(self.item.bid[8],','))
        self.ui.label_bid8.setText(format(self.item.bid[7],','))
        self.ui.label_bid7.setText(format(self.item.bid[6],','))
        self.ui.label_bid6.setText(format(self.item.bid[5],','))
        self.ui.label_bid5.setText(format(self.item.bid[4],','))
        self.ui.label_bid4.setText(format(self.item.bid[3],','))
        self.ui.label_bid3.setText(format(self.item.bid[2],','))
        self.ui.label_bid2.setText(format(self.item.bid[1],','))
        self.ui.label_bid1.setText(format(self.item.bid[0],','))
 
        self.ui.label_bid_v10.setText(format(self.item.bidvol[9],','))
        self.ui.label_bid_v9.setText(format(self.item.bidvol[8],','))
        self.ui.label_bid_v8.setText(format(self.item.bidvol[7],','))
        self.ui.label_bid_v7.setText(format(self.item.bidvol[6],','))
        self.ui.label_bid_v6.setText(format(self.item.bidvol[5],','))
        self.ui.label_bid_v5.setText(format(self.item.bidvol[4],','))
        self.ui.label_bid_v4.setText(format(self.item.bidvol[3],','))
        self.ui.label_bid_v3.setText(format(self.item.bidvol[2],','))
        self.ui.label_bid_v2.setText(format(self.item.bidvol[1],','))
        self.ui.label_bid_v1.setText(format(self.item.bidvol[0],','))
 
        cur = self.item.cur
        diff = self.item.diff
        diffp = self.item.diffp
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            cur = self.item.expcur
            diff = self.item.expdiff
            diffp = self.item.expdiffp
 
 
        strcur = format(cur, ',')
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            strcur = "*" + strcur
 
        curcolor = self.item.getCurColor()
        self.ui.label_cur.setStyleSheet(curcolor)
        self.ui.label_cur.setText(strcur)
        strdiff = str(diff) + "  " + format(diffp, '.2f')
        strdiff += "%"
        self.ui.label_diff.setText(strdiff)
        self.ui.label_diff.setStyleSheet(curcolor)
 
        self.ui.label_totoffer.setText(format(self.item.totOffer,','))
        self.ui.label_totbid.setText(format(self.item.totBid,','))
 
 
 
if __name__ == '__main__':
        app = QtWidgets.QApplication(sys.argv)
        w = Form()
        sys.exit(app.exec())
