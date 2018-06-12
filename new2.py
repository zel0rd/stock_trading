# -*- coding: utf-8 -*-
"""
Created on Thu Nov 16 00:45:18 2017

@author: Zelord.Kwoun
"""
import os
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
import ctypes

import pandas as pd
from pandas import DataFrame
import numpy as np
import matplotlib.pyplot as plt

from fbprophet import Prophet
from datetime import datetime




g_instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")  #1
g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr") #2
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")  #3
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil") #4
locale.setlocale(locale.LC_ALL, '')
# 잔고조회
g_cnt = 0
g_code = []
g_name = []
g_amount = []
g_buyPrice =[]
g_evalValue = []
g_evalPerc = []
g_rate = []
g_money = []
g_test = 1

g_code1=[]
g_name1=[]
g_price1=[]

g_code2=[]
g_name2=[]
g_price2=[]

#매수매도
g_buycode=0

#미체결조회

g_dates = []
g_closes = []


a_code  = []
a_name  = []
a_orderDesc = []
a_amount  = []
a_price  = []
a_ContAmount = []
#hooseNum = 0
#buyName = ""

d_code = []
d_name = []


################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
 
 
################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False
 
    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False
 
    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False
 
    return True
 
 
################################################
    

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
            self.n += 10
            self.sleep(10)
     
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

    def set_params2(self,client):
        self.client = client

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
        handler.set_params2(self.objStockCur)
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
        #print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
        global g_cnt
        g_cnt = self.objRq.GetHeaderValue(7)
        #print(cnt)
        #print("AAAAAAAAAAAAAAAAAAA",len(g_code))
        if len(g_code) == 0:
            #print("종목코드 종목명 신용구분 체결잔고수량 체결장부단가 평가금액 평가손익")
            for i in range(g_cnt):
                g_code.append(self.objRq.GetDataValue(12, i))  # 종목코드
                g_name.append(self.objRq.GetDataValue(0, i))  # 종목명
                #g_cashFlag = self.objRq.GetDataValue(1, i)  # 신용구분
                #g_date = self.objRq.GetDataValue(2, i)  # 대출일
                g_amount.append("%d주"%self.objRq.GetDataValue(7, i)) # 체결잔고수량
                g_buyPrice.append("%d원"%int(self.objRq.GetDataValue(17, i))) # 체결장부단가
                g_evalValue.append("%d원"%int(self.objRq.GetDataValue(9, i))) # 평가금액(천원미만은 절사 됨)
                g_evalPerc.append("%d%%"%int(self.objRq.GetDataValue(11, i))) # 평가손익
                #평가금액 - 수량 * 장부가 = money
                g_money.append("%d원"%int(self.objRq.GetDataValue(9, i) - self.objRq.GetDataValue(7, i)*self.objRq.GetDataValue(17, i)))
                
                
                
                #  평 가 금 액 / 잔 고 수 량 / 장 부 단 가 - 1 * 100
         
            print("할당됨할당됨할당됨할당됨할당됨할당됨할당됨할당됨")
                #print("g_code :",g_code, g_name, g_cashFlag, g_amount, g_buyPrice, g_evalValue, g_evalPerc)
                #print("g_code :",g_code, "g_name :",g_name, "g_amount :", g_amount, "장부가 :%d원"%g_buyPrice, "평가금액 : %d원"%g_evalValue,"손익 : %d원"% g_evalPerc)
        
        #print("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")
        #print(g_cnt)
        #print(g_code)
        #print(g_name)
        #print(g_amount)
        #print(g_buyPrice)
        #print(g_evalValue)
        #print(g_evalPerc)
    
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
            time.sleep(1)
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
            #if self.objWeek.GetDibStatus() != 0:
            #    print("통신상태", self.objWeek.GetDibStatus(), self.objWeek.GetDibMsg1())
            #    return False

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
                
            if (sumCnt > 300):
                break

            if self.objWeek.Continue == False:
                break
        """
        sumCnt2 = 0
        dates1 = []
        closes1 = []
        while True:

            cnt = self.objWeek.GetHeaderValue(1)
            sumCnt2 += cnt
            if cnt == 0:
                break

            for i in range(cnt):
                dates1.append(self.objWeek.GetDataValue(0, i))
                closes1.append(self.objWeek.GetDataValue(4, i))

               
            if (sumCnt2 > 1000):
                break

            if self.objWeek.Continue == False:
                break

        global g_dates,g_closes
        g_dates = dates1
        g_closes = closes1
        """
        global g_dates,g_closes
        g_dates = dates
        g_closes = closes
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

class orderData:
    def __init__(self):
        self.code = ""          # 종목코드
        self.name = ""          # 종목명
        self.orderNum = 0       # 주문번호
        self.orderPrev = 0      # 원주문번호
        self.orderDesc = ""     # 주문구분내용
        self.amount = 0     # 주문수량
        self.price = 0      # 주문 단가
        self.ContAmount = 0  # 체결수량
        self.credit = ""     # 신용 구분 "현금" "유통융자" "자기융자" "유통대주" "자기대주"
        self.modAvali = 0  # 정정/취소 가능 수량
        self.buysell = ""  # 매매구분 코드  1 매도 2 매수
        self.creditdate = ""    # 대출일
        self.orderFlag = ""     # 주문호가 구분코드
        self.orderFlagDesc = "" # 주문호가 구분 코드 내용
 
        # 데이터 변환용
        self.concdic = {"1": "체결", "2": "확인", "3": "거부", "4": "접수"}
        self.buyselldic = {"1": "매도", "2": "매수"}
 
    def debugPrint(self):
        print("%s, %s, 주문번호 %d, 원주문 %d, %s, 주문수량 %d, 주문단가 %d, 체결수량 %d, %s, "
              "정정가능수량 %d, 매수매도: %s, 대출일 %s, 주문호가구분 %s %s"
              %(self.code, self.name, self.orderNum, self.orderPrev, self.orderDesc, self.amount, self.price,
                self.ContAmount,self.credit,self.modAvali, self.buyselldic.get(self.buysell),
                self.creditdate,self.orderFlag, self.orderFlagDesc))
 
 

 
class Cp5339:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpTrade.CpTd5339")
        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분
 
 
    def Request5339(self, dicOrderList, orderList):
        self.objRq.SetInputValue(0, self.acc)
        self.objRq.SetInputValue(1, self.accFlag[0])
        self.objRq.SetInputValue(4, "0") # 전체
        self.objRq.SetInputValue(5, "1") # 정렬 기준 - 역순
        self.objRq.SetInputValue(6, "0") # 전체
        self.objRq.SetInputValue(7, 20) # 요청 개수 - 최대 20개
 
        print("[Cp5339] 미체결 데이터 조회 시작")
        # 미체결 연속 조회를 위해 while 문 사용
        while True :
            ret = self.objRq.BlockRequest()
            if self.objRq.GetDibStatus() != 0:
                print("통신상태", self.objRq.GetDibStatus(), self.objRq.GetDibMsg1())
                return False
 
            if (ret == 2 or ret == 3):
                print("통신 오류", ret)
                return False;
 
            # 통신 초과 요청 방지에 의한 요류 인 경우
            while (ret == 4) : # 연속 주문 오류 임. 이 경우는 남은 시간동안 반드시 대기해야 함.
                remainTime = g_objCpStatus.LimitRequestRemainTime
                print("연속 통신 초과에 의해 재 통신처리 : ",remainTime/1000, "초 대기" )
                time.sleep(remainTime / 1000)
                ret = self.objRq.BlockRequest()
 
 
            # 수신 개수
            cnt = self.objRq.GetHeaderValue(5)
            print("[Cp5339] 수신 개수 ", cnt)
            if cnt == 0 :
                break
            
            for i in range(cnt):
                item = orderData()
                item.orderNum = self.objRq.GetDataValue(1, i)
                item.orderPrev  = self.objRq.GetDataValue(2, i)
                item.code  = self.objRq.GetDataValue(3, i)  # 종목코드
                item.name  = self.objRq.GetDataValue(4, i)  # 종목명
                item.orderDesc  = self.objRq.GetDataValue(5, i)  # 주문구분내용
                item.amount  = self.objRq.GetDataValue(6, i)  # 주문수량
                item.price  = self.objRq.GetDataValue(7, i)  # 주문단가
                item.ContAmount = self.objRq.GetDataValue(8, i)  # 체결수량
                item.credit  = self.objRq.GetDataValue(9, i)  # 신용구분
                item.modAvali  = self.objRq.GetDataValue(11, i)  # 정정취소 가능수량
                item.buysell  = self.objRq.GetDataValue(13, i)  # 매매구분코드
                item.creditdate  = self.objRq.GetDataValue(17, i)  # 대출일
                item.orderFlagDesc  = self.objRq.GetDataValue(19, i)  # 주문호가구분코드내용
                item.orderFlag  = self.objRq.GetDataValue(21, i)  # 주문호가구분코드
 
                            
                a_code.append(self.objRq.GetDataValue(3, i))
                a_name.append(self.objRq.GetDataValue(4, i))
                a_orderDesc.append(self.objRq.GetDataValue(5, i))
                a_amount.append(self.objRq.GetDataValue(6, i))
                a_price.append(self.objRq.GetDataValue(7, i))
                a_ContAmount.append(self.objRq.GetDataValue(8, i))
                
                # 사전과 배열에 미체결 item 을 추가
                dicOrderList[item.orderNum] = item
                orderList.append(item)
 
            print(dicOrderList)
            print(orderList[0])
            print("AAAAAAA"+item.code)
            
            print("AAAAAAA"+item.name)
            
            print(orderList)
            
            
            # 연속 처리 체크 - 다음 데이터가 없으면 중지
            if self.objRq.Continue == False :
                print("[Cp5339] 연속 조회 여부: 다음 데이터가 없음")
                break
 
        return True
    
 # Cp8537 : 종목검색 전략 조회
class Cp8537:
    def __init__(self):
        self.objRq = None
        return
 
    def requestList(self, caller):
        #caller.data8537 = {}
        print(caller)
        self.objRq = None
        self.objRq = win32com.client.Dispatch("CpSysDib.CssStgList")
 
        # 예제 전략에서 전략 리스트를 가져옵니다.
        self.objRq.SetInputValue(0, ord('0'))   # '0' : 예제전략, '1': 나의전략
        self.objRq.BlockRequest()
 
        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            return False
 
        cnt = self.objRq.GetHeaderValue(0) # 0 - (long) 전략 목록 수
        flag = self.objRq.GetHeaderValue(1) # 1 - (char) 요청구분
        print('종목검색 전략수:', cnt)
 
 
        for i in range(cnt):
            item = {}
            item['전략명'] = self.objRq.GetDataValue(0, i)
            item['ID'] = self.objRq.GetDataValue(1, i)
            item['전략등록일시'] = self.objRq.GetDataValue(2, i)
            item['작성자필명'] = self.objRq.GetDataValue(3, i)
            item['평균종목수'] = self.objRq.GetDataValue(4, i)
            item['평균승률'] = self.objRq.GetDataValue(5, i)
            item['평균수익'] = self.objRq.GetDataValue(6, i)
            caller.data8537[item['전략명']] = item
            
        return True
 
    def requestStgID(self, id, caller):
        caller.dataStg = []
        self.objRq = None
        self.objRq = win32com.client.Dispatch("CpSysDib.CssStgFind")
        self.objRq.SetInputValue(0, id) # 전략 id 요청
        self.objRq.BlockRequest()
        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            return False
 
        cnt = self.objRq.GetHeaderValue(0)  # 0 - (long) 검색된 결과 종목 수
        totcnt = self.objRq.GetHeaderValue(1)  # 1 - (long) 총 검색 종목 수
        stime = self.objRq.GetHeaderValue(2)  # 2 - (string) 검색시간
        print('검색된 종목수:', cnt, '전체종목수:', totcnt, '검색시간:', stime)
 
        for i in range(cnt):
            item = {}
            item['code'] = self.objRq.GetDataValue(0, i)
            item['종목명'] = g_objCodeMgr.CodeToName(item['code'])
            caller.dataStg.append(item)
 
        return True   
class Form(QtWidgets.QDialog):
    def __init__(self):
        #threading.Thread.__init__(self) 
        super().__init__() 
        self.ui = uic.loadUi("hoga_2.ui", self)
        
        #CONNECTIN CHECK
        if InitPlusCheck() == False:
            exit()
    
        self.ui.btn1.clicked.connect(self.threadStart)
        self.ui.btn2.clicked.connect(self.threadStop)
        #self.ui.pushButton_2.clicked.connect(self.pushButton_2action)
        self.ui.btnStart.clicked.connect(self.btnStart_clicked)
        self.ui.btnStart_2.clicked.connect(self.btnStart_2clicked)
        self.ui.btnStart_3.clicked.connect(self.btnStart_3clicked)
        
        self.ui.pushButton_4.clicked.connect(self.pushButton_4clicked)
        #self.ui.pushButton_5.clicked.connect(self.StopSubscribe)
        self.ui.tableWeek_3.resizeColumnsToContents()
        self.ui.tableWeek_3.resizeRowsToContents()
        
        self.ui.pushButton_2.clicked.connect(self.pushButton_2clicked)
        self.ui.pushButton_3.clicked.connect(self.pushButton_3clicked)
        self.ui.pushButton_5.clicked.connect(self.pushButton_5clicked)
        
        self.comboBox_4.currentIndexChanged.connect(self.comboChanged)
        self.ui.pushButton_6.clicked.connect(self.pushButton_6clicked)
        self.ui.pushButton_7.clicked.connect(self.pushButton_7clicked)
       # self.ui.pushButton_7.clicked.connect(self.pushButton_7clicked)
        
        
        
        self.isSB = False
        
        
        self.ui.show()
        
        self.objMst = CpRPCurrentPrice()
        self.item = stockPricedData()
        
        # 일자별
        self.objWeek = CpWeekList()
        self.rpWeek = DataFrame()   # 일자별 데이터프레임


        # 시간대별
        self.rpStockBid = DataFrame()
        self.objStockBid = CpStockBid()
        self.todayIndex = 0

        #self.setCode("005930")
        self.setCode("002240")
                 
        
        self.th = TestThread(self)
        self.th.threadEvent.connect(self.threadEventHandler)
        
        
        
        #self.ui.pushButton_2 = QpushButton("aaa",self)
        #self.ui.pushButton_2.clicked(self.pushButton_2)
        #self.ui.editCode_2("Aaaa",self)
        
        #self.setupUI()
    #def DisconnectClicked(self):
    #    self.CpUtil.CpCybos::PlusDisconnect
    
    ###엑셀출력
    
    def pushButton_7clicked(self):
        
        new_dates = []
        for i in range(len(g_dates)):
            yyyy = int(g_dates[i] / 10000)
            mm = int(g_dates[i] - (yyyy * 10000))
            dd = mm % 100
            mm = mm / 100
            val = '%04d-%02d-%02d' %(yyyy, mm, dd)
            new_dates.append(val)
        
        global d_code,d_name
        
        
        #print data set
        print("\n\n\n\n\n" + d_code + "(" + d_name + ") : " + "Data Sets")
        t = {'dates' : new_dates, 'closes' : g_closes}
        tf = pd.DataFrame(data=t)
        print(tf)
        #print(g_closes)
        #print(new_dates) 
        
        plt.plot(new_dates,g_closes)
        plt.show()
        
        
        print("2017 Data Set")
        d = {'ds':new_dates,'y':g_closes}
        df = pd.DataFrame(data=d)
        df_temp = df.drop(df.index[0:100])
        m = Prophet()
        m.fit(df_temp)
        future = m.make_future_dataframe(periods=100)
        forecast = m.predict(future)
        m.plot(forecast)
        plt.show()
        
        
        #Time series forecast
        print("\n\n\n\n\nTime series forecast")
        
        m = Prophet()
        m.fit(df)
        future = m.make_future_dataframe(periods=100)
        forecast = m.predict(future)
        m.plot(forecast)
        plt.show()
        m.plot_components(forecast)
        plt.show()
        
        
        #Saturating forecast
        print("\n\n\n\n\nSaturating forecast")
        df['cap'] = g_closes[0]
        m = Prophet(growth='logistic')
        m.fit(df)
        
        future = m.make_future_dataframe(periods=10)
        future['cap'] = g_closes[0]
        fcst = m.predict(future)
        m.plot(fcst)
        plt.show()
        
        #Adjusting trend flexibility
        print("\n\n\n\n\nAdjusting trend flexibility")
        m = Prophet(changepoint_prior_scale=0.5)
        forecast = m.fit(df).predict(future)
        m.plot(forecast)
        plt.show()
        
        #Outliers
        print("\n\n\n\n\nOutliers")
        m = Prophet()
        m.fit(df)
        future = m.make_future_dataframe(periods=100)
        forecast = m.predict(future)
        m.plot(forecast)
        plt.show()
         
        #Sub-daily data
        print("\n\n\n\n\nSub-daily data")
        m = Prophet(changepoint_prior_scale=1).fit(df)
        future = m.make_future_dataframe(periods=300, freq='H')
        fcst = m.predict(future)
        m.plot(fcst)
        plt.show()
        
        
    #전략 정보조회
    def pushButton_6clicked(self):
        self.obj8537 = Cp8537()
        self.data8537 = {}
        self.dataStg = []
        self.obj8537.requestList(self)
 
        for k, v in self.data8537.items():
            self.comboBox_4.addItem(k)
        return
    
    def comboChanged(self):
        cur = self.comboBox_4.currentText()
        print(cur)
        self.requestStgID(cur)
    
    
    def requestStgID(self, stgName):
        item = self.data8537[stgName]
        id = item['ID']
        name = item['전략명']
 
        self.obj8537.requestStgID(id, self)
 
        print('검색전략:', id, '전략명:', name, '검색종목수:', len(self.dataStg))
        for item in self.dataStg:
            print(item)
            
    
        return
#####################
    
    
    def pushButton_5clicked(self):
        self.bTradeInit = False
        # 연결 여부 체크
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
        if (g_objCpTrade.TradeInit(0) != 0):
            print("주문 초기화 실패")
            return False
        self.bTradeInit = True
 
        # 미체결 리스트를 보관한 자료 구조체
        self.diOrderList= dict()  # 미체결 내역 딕셔너리 - key: 주문번호, value - 미체결 레코드
        self.orderList = []       # 미체결 내역 리스트 - 순차 조회 등을 위한 미체결 리스트
 
        # 미체결 통신 object
        self.obj = Cp5339()
        # 주문 취소 통신 object
#        self.objOrder = CpRPOrder()
 
        # 실시간 주문 체결
#        self.contsb = CpConclution()
#        self.contsb.Subscribe("", self)
    
        #def request5339
        if self.bTradeInit == False :
            print("TradeInit 실패")
            return False
 
        self.diOrderList = {}
        self.orderList = []
        self.obj.Request5339(self.diOrderList, self.orderList)
        
        #print(self.obj.item.code)
        for item in self.orderList:
            item.debugPrint()
        print("[Reqeust5339]미체결 개수 ", len(self.orderList))
        
        #print(a_code)
        #print(a_name)
        #print(a_orderDesc)
        #print(a_amount)
        #print(a_price)
        #print(a_ContAmount)
        a_cnt =len(a_code)
        myorder= {'a_code':a_code,'a_name':a_name,'a_orderDesc':a_orderDesc,'a_amount':a_amount,'a_price':a_price,'a_ContAmount':a_ContAmount}
        print(myorder)
        
        column_idx_lookup = {'a_code':0,'a_name':1,'a_orderDesc':2,'a_amount':3,'a_ContAmount':4,'a_price':5}
        column_headers = ['종목코드', '종목명','주문구분','수량','체결','금액']
        self.ui.tableWeek_4.setRowCount(a_cnt)
        self.ui.tableWeek_4.setHorizontalHeaderLabels(column_headers)
       
        for k, v in myorder.items():
            col = column_idx_lookup[k]
            for row, val in enumerate(v):
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWeek_4.setItem(row, col, item)
                
        self.ui.tableWeek_4.resizeColumnsToContents()
        self.ui.tableWeek_4.resizeRowsToContents()


        

        ##매 수 주 문
    def pushButton_2clicked(self):
        
        
        global g_buycode
        price = int(self.ui.lineEdit.text())
        count = self.spinBox.value()
        
        objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            exit()
         
        acc = objTrade.AccountNumber[0] #계좌번호
        accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])
        objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
        objStockOrder.SetInputValue(0, "2")   # 2: 매수
        objStockOrder.SetInputValue(1, acc )   #  계좌번호
        objStockOrder.SetInputValue(2, accFlag[0])   # 상품구분 - 주식 상품 중 첫번째
        objStockOrder.SetInputValue(3, g_buycode)   # 종목코드 - A003540 - 대신증권 종목
        objStockOrder.SetInputValue(4, count)   # 매수수량 10주
        objStockOrder.SetInputValue(5, price)   # 주문단가  - 14,100원
        objStockOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        objStockOrder.SetInputValue(8, "01")   # 주문호가 구분코드 - 01: 보통
         
        # 매수 주문 요청
        objStockOrder.BlockRequest()
         
        rqStatus = objStockOrder.GetDibStatus()
        rqRet = objStockOrder.GetDibMsg1()
        print("통신상태", "rqStatus:",rqStatus,"rqRet :", rqRet)


        ##매 도 주 
    def pushButton_3clicked(self):
        objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = objTrade.TradeInit(0)
        
        if (initCheck != 0):
            print("주문 초기화 실패")
            exit()
         
        
        choose = self.comboBox_3.currentText()
        code = choose[0:7]
        count = self.spinBox_2.value()
        #price = int(self.editCode_2.toPainText())
        price = int(self.ui.lineEdit_2.text())
        # 주식 매도 주문
        acc = objTrade.AccountNumber[0] #계좌번호
        accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])
        objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
        objStockOrder.SetInputValue(0, "1")   #  1: 매도
        objStockOrder.SetInputValue(1, acc )   #  계좌번호
        objStockOrder.SetInputValue(2, accFlag[0])   #  상품구분 - 주식 상품 중 첫번째
        objStockOrder.SetInputValue(3, code)   #  종목코드 - A003540 - 대신증권 종목
        objStockOrder.SetInputValue(4, count)   #  매도수량 10주
        objStockOrder.SetInputValue(5, price)   #  주문단가  - 14,100원
        objStockOrder.SetInputValue(7, "0")   #  주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        objStockOrder.SetInputValue(8, "01")   # 주문호가 구분코드 - 01: 보통
         
        # 매도 주문 요청
        objStockOrder.BlockRequest()
         
        rqStatus = objStockOrder.GetDibStatus()
        rqRet = objStockOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()
            
            
        
        
        
    def pushButton_4clicked(self):
        self.StopSubscribe();
        global g_cnt
        codes = []
        obj6033 = Cp6033()
        if obj6033.Request(codes) == False:
            return
    
        print("BBBBBBBBBBBBBBBBBBBBBBBBBB")
        print(g_cnt)
        print(g_code)
        print(g_name)
        print(g_amount)
        print(g_buyPrice)
        print(g_evalValue)
        print(g_evalPerc)
        print(g_rate)
        print(g_money)
        for i in range(g_cnt):
            #self.ui.comboBox_3.addItem(g_name[i])
            self.ui.comboBox_3.addItem(g_code[i]+"  :  "+g_name[i])
            
            
        mycount = {'g_code':g_code,'g_name':g_name,'g_amount':g_amount,'g_buyPrice':g_buyPrice,'g_evalValue':g_evalValue,'g_evalPerc':g_evalPerc,'g_money':g_money}
        print(mycount)
        
        column_idx_lookup = {'g_code':0,'g_name':1,'g_amount':2,'g_buyPrice':3,'g_evalValue':4,'g_evalPerc':5,'g_money':6}
        
        column_headers = ['종목코드', '종목명','잔고수량','매입가','평가금액','평가손익','손익금액']
        self.ui.tableWeek_3.setRowCount(g_cnt)
        self.ui.tableWeek_3.setHorizontalHeaderLabels(column_headers)
        
        for k, v in mycount.items():
            col = column_idx_lookup[k]
            for row, val in enumerate(v):
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWeek_3.setItem(row, col, item)

        
        self.ui.tableWeek_3.resizeColumnsToContents()
        self.ui.tableWeek_3.resizeRowsToContents()

        
    @pyqtSlot()
    def slot_codeupdate(self):
        code = self.ui.editCode.toPlainText()
        self.setCode(code)
 
    def slot_codechanged(self):
        print("codechange")
        code = self.ui.editCode.toPlainText()
        self.setCode(code)
 
    
    def monitorPriceChange(self):
        self.displyHoga()
        self.updateWeek()
        self.updateStockBid()

    def monitorOfferbidChange(self):
        self.displyHoga()

    def setCode(self, code):
        if len(code) < 6 :
            return

        print(code)
        if not (code[0] == "A"):
            code = "A" + code

        name = g_objCodeMgr.CodeToName(code)
        if len(name) == 0:
            print("종목코드 확인")
            return
            
        global d_code, d_name
        d_code = code
        d_name = name
        
        self.ui.label_name.setText(name)
        
        
        if (self.objMst.Request(code, self.item, self) == False):
            return
        self.displyHoga()


        # 일자별
        self.ui.tableWeek.clearContents()
        if (self.objWeek.Request(code, self) == True):
            print(self.rpWeek)
            self.displyWeek()
                    
        # 시간대별
        self.ui.tableStockBid.clearContents()
        if (self.objStockBid.Request(code, self) == True):
            self.displyStockBid()


    def btnStart_clicked(self):
        codeList1 = g_objCodeMgr.GetStockListByMarket(1)
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)
        #codeList = codeList1 + codeList2
        #combine1=[]
        for i, code in enumerate(codeList1):
            secondCode = g_objCodeMgr.GetStockSectionKind(code)
            s_name = g_objCodeMgr.CodeToName(code)
            stdPrice = g_objCodeMgr.GetStockStdPrice(code)
            
            g_code1.append(codeList1[i])
            g_name1.append(s_name)
            g_price1.append(stdPrice)
            #print(i, code, secondCode, stdPrice, s_name)
            self.ui.comboBox.addItem("[%s] "%codeList1[i]+s_name)
       #combine1.append(name)
            
            
            
        #for i in range(0,len(code)):
        #    print(combine1[i])
        
        for i, code in enumerate(codeList2):
            secondCode = g_objCodeMgr.GetStockSectionKind(code)
            s_name = g_objCodeMgr.CodeToName(code)
            stdPrice = g_objCodeMgr.GetStockStdPrice(code)
            
            g_code2.append(codeList2[i])
            g_name2.append(s_name)
            g_price2.append(stdPrice)
            
            #print(i, code, secondCode, stdPrice, s_name)
            #combine2 = {name[i]: code[i]}
            #print(combine2)
            self.ui.comboBox_2.addItem("[%s] "%codeList2[i]+s_name)
    
    def btnStart_2clicked(self):
        choose = self.comboBox.currentText()
        self.ui.label_3.setText(choose)
        name = choose[10:]
        global g_buycode
        g_buycode= g_code1[g_name1.index(name)]
        price = g_price1[g_name1.index(name)]
        print(price)
        self.ui.label_16.setText(str(price))
        
    def btnStart_3clicked(self):
        choose = self.comboBox_2.currentText()
        self.ui.label_3.setText(choose)
        name = choose[10:]
        global g_buycode
        g_buycode= g_code2[g_name2.index(name)]
        price = g_price2[g_name2.index(name)]
        self.ui.label_16.setText(str(price))


    def StopSubscribe(self):
        if self.isSB:
            cnt = len(self.objCur)
            for i in range(cnt):
                self.objCur[i].Unsubscribe()
            print(cnt, "종목 실시간 해지되었음")
        self.isSB = False
 
        self.objCur = []
 
    #def pushButton_2action(self):



    
    def threadStart(self):  
    #if g_instCpCybos == 1:
        #self.objMst = CpRPCurrentPrice()
        #self.item = stockPricedData()
       
        
        #btnStart_clicked
        self.StopSubscribe();
        codes = []
        obj6033 = Cp6033()
        
        #주식계좌조회
        self.objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
        acc = self.objTrade.AccountNumber[0]
        
        self.isSB = True

        #self.btnStart_clicked(self)
        #self.objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
        #acc = self.objTrade.AccountNumber[0]
        #self.ui.label_2.setText(acc)
        
        
        #elif g_instCpCybos == 0:
        #    print("Connection error")
            
        if not self.th.isRun:
            print('메인 : 쓰레드 시작')
            self.th.isRun = True
            self.th.start()

    @pyqtSlot()
    def threadStop(self):
        self.ui.label_Cp.setText("Stop")
        self.ui.label_Obj.setText("Stop")
        self.ui.account_2.setText("Stop")
        if self.th.isRun:
            print('메인 : 쓰레드 정지')
            self.th.isRun = False
            
    # 쓰레드 이벤트 핸들러
    # 장식자에 파라미터 자료형을 명시
    @pyqtSlot(int)
    def threadEventHandler(self):
        pythoncom.CoInitialize()
        #g_instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        if g_instCpCybos.Isconnect == 1:
            C_Connection = "Success"
            col_C_Connection = "<font color= blue>"+C_Connection
        elif g_instCpCybos.Isconnect == 0:
            C_Connection = "Fail"
            col_C_Connection = "<font color= red>"+C_Connection
        if g_objCpStatus.Isconnect == 1:
            O_Connection = "Success"
            col_O_Connection = "<font color= blue>"+O_Connection
        elif g_objCpStatus.Isconnect == 0:
            O_Connection = "Fail"
            col_O_Connection = "<font color= red>"+O_Connection

            
        
        #self.statusBar.showMessage(self.lineEdit.text()) ;+str(g_objCodeMgr)+str(g_objCpStatus)+str(g_objCpTrade.Isconnect)
        self.objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
        #acc = 0
        #print(self.objTrade.AccountNumber[0])
        acc = self.objTrade.AccountNumber[0]
        self.ui.account_2.setText("<font color=blue>"+acc)
        
        self.ui.label_Cp.setText(col_C_Connection)
        self.ui.label_Obj.setText(col_O_Connection)
        #self.statusBar.showMessage("instCpCyobs : "+str(g_instCpCybos.Isconnect)+"  objCpStatus : "+str(g_objCpStatus.Isconnect))
        #self.label2.setText(Connection)
        #print('메인 : threadEvent(self,' + str(n) + ')')
        
    
            
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
        strdiff = format(diffp, '.2f')
        strdiff += "%"
        strdiff_2 = str(diff)
        self.ui.label_diff_2.setText(strdiff_2)
        self.ui.label_diff_2.setStyleSheet(curcolor)
        self.ui.label_diff.setText(strdiff)
        self.ui.label_diff.setStyleSheet(curcolor)
        self.ui.label_totoffer.setText(format(self.item.totOffer,','))
        self.ui.label_totbid.setText(format(self.item.totBid,','))
        
    # 일자별 리스트 UI 채우기
    def displyWeek(self):
        rowcnt = len(self.rpWeek.index)
        if rowcnt == 0:
            return
        self.ui.tableWeek.setRowCount(rowcnt)

        nRow = 0

        for index, row in self.rpWeek.iterrows():
            datas = [index, row['close'],row['diff'],row['diffp'],row['vol'],row['open'],row['high'],row['low'],
                     row['for_v'], row['for_d'], row['for_p']]
            for col in range(len(datas)) :
                val = ''
                if (col == 0):  # 일자
                    # 20170929 --> 2017/09/29
                    yyyy = int(datas[col] / 10000)
                    mm = int(datas[col] - (yyyy * 10000))
                    dd = mm % 100
                    mm = mm / 100
                    val = '%04d/%02d/%02d' %(yyyy, mm, dd)
                elif (col == 3 or col == 10): # 대비율
                    val = locale.format('%.2f', datas[col], 1)
                    val += "%"

                else:
                    val = locale.format('%d', datas[col], 1)

                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.ui.tableWeek.setItem(nRow, col, item)

            if (nRow == 0) :
                self.todayIndex = index
            nRow += 1

            self.tableWeek.resizeColumnsToContents()
        return

    # 일자별 리스트 UI 채우기 - 오늘 날짜 업데이트
    def updateWeek(self):
        rowcnt = len(self.rpWeek.index)
        if rowcnt == 0:
            return

        # 오늘 날짜 데이터 업데이트
        self.rpWeek.set_value(self.todayIndex, 'close', self.item.cur)
        self.rpWeek.set_value(self.todayIndex, 'open', self.item.open)
        self.rpWeek.set_value(self.todayIndex, 'high', self.item.high)
        self.rpWeek.set_value(self.todayIndex, 'low', self.item.low)
        self.rpWeek.set_value(self.todayIndex, 'vol', self.item.vol)
        self.rpWeek.set_value(self.todayIndex, 'diff', self.item.diff)
        self.rpWeek.set_value(self.todayIndex, 'diffp', self.item.diffp)

        datas = [self.todayIndex, self.item.cur,self.item.diff, self.item.diffp, self.item.vol,
                 self.item.open, self.item.high, self.item.low]
        for col in range(len(datas)) :
            val = ''
            if (col == 0):  # 일자
                # 20170929 --> 2017/09/29
                yyyy = int(datas[col] / 10000)
                mm = int(datas[col] - (yyyy * 10000))
                dd = mm % 100
                mm = mm / 100
                val = '%04d/%02d/%02d' %(yyyy, mm, dd)
            elif (col == 3): # 대비율
                val = locale.format('%.2f', datas[col], 1)
                val += "%"

            else:
                val = locale.format('%d', datas[col], 1)

            item = QTableWidgetItem(val)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.ui.tableWeek.setItem(0, col, item)

        return



    # 시간대별 리스트 UI  채우기
    def displyStockBid(self):
        rowcnt = len(self.rpStockBid.index)
        if rowcnt == 0:
            return
        self.ui.tableStockBid.setRowCount(rowcnt)

        nRow = 0

        for index, row in self.rpStockBid.iterrows():
            # 행 내에 표시할 데이터 - 컬럼 순
            datas = [row['time'], row['cur'], row['diff'], row['offer'], row['bid'], row['vol'], row['tvol'],
                     row['tvol'], row['volstr']]
            market = row['market']
            for col in range(len(datas)):
                val = ''
                if col == 0: # 시각
                    # 155925 --> 15:59:25
                    hh = int(datas[col] / 10000)
                    mm = int(datas[col] - (hh * 10000))
                    ss = mm % 100
                    mm = mm / 100
                    val = '%02d:%02d:%02d' %(hh, mm, ss)
                elif col == 6: # 체결매도
                    market = row['flag']
                    if (market == "체결매도") :
                        val = locale.format('%d', datas[col], 1)
                elif col == 7: # 체결매수
                    market = row['flag']
                    if (market == "체결매수"):
                        val = locale.format('%d', datas[col], 1)
                elif col == 8:  # 체결강도
                    val = locale.format('%.2f', datas[col], 1)
                elif col == 1: # 현재가
                    val = locale.format('%d', datas[col], 1)
                    if (market == "예상체결"):
                        val = '*' + val
                else:          # 기타
                    val = locale.format('%d', datas[col], 1)
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.ui.tableStockBid.setItem(nRow, col, item)
            nRow += 1

        self.tableStockBid.resizeColumnsToContents()
        return


    def updateStockBid(self):
        rowcnt = len(self.rpStockBid.index)
        if rowcnt == 0:
            return
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            return

        buyvol = sellvol = 0
        if self.item.volFlag == ord('1') :
                buyvol = self.item.tvol
        if self.item.volFlag == ord('2') :
                sellvol = self.item.tvol
        line = DataFrame({"time": self.item.time,
                          "cur": self.item.cur,
                          "diff": self.item.diff,
                          "offer": self.item.offer[0],
                          "bid": self.item.bid[0],
                          "vol": self.item.vol,
                          "tvol": buyvol,
                          "tvol": sellvol,
                          "volstr": self.item.volstr},
                         index=[0])

        self.rpStockBid = pandas.concat([line, self.rpStockBid.ix[:]]).reset_index(drop=True)

        # 행 내에 표시할 데이터 - 컬럼 순
        datas = [self.item.time, self.item.cur, self.item.diff, self.item.offer[0], self.item.bid[0],
                 self.item.vol, sellvol, buyvol, self.item.volstr]
        self.ui.tableStockBid.insertRow(0)
        for col in range(len(datas)):
            val = ''
            if col == 0: # 시각
                # 155925 --> 15:59:25
                hh = int(datas[col] / 10000)
                mm = int(datas[col] - (hh * 10000))
                ss = mm % 100
                mm = mm / 100
                val = '%02d:%02d:%02d' %(hh, mm, ss)
            elif col == 6: # 체결매도
                val = locale.format('%d', datas[col], 1)
            elif col == 7: # 체결매수
                val = locale.format('%d', datas[col], 1)
            elif col == 8: # 체결강도
                val = locale.format('%.2f', datas[col], 1)
            else:          # 기타
                val = locale.format('%d', datas[col], 1)

            item = QTableWidgetItem(val)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.ui.tableStockBid.setItem(0, col, item)


        return

# 일자별 리스트 UI 채우기
    def have(self):
        rowcnt = len(self.rpWeek.index)
        if rowcnt == 0:
            return
        self.ui.tableWeek.setRowCount(rowcnt)

        nRow = 0

        for index, row in self.rpWeek.iterrows():
            datas = [index, row['close'],row['diff'],row['diffp'],row['vol'],row['open'],row['high'],row['low'],
                     row['for_v'], row['for_d'], row['for_p']]
            for col in range(len(datas)) :
                val = ''
                if (col == 0):  # 일자
                    # 20170929 --> 2017/09/29
                    yyyy = int(datas[col] / 10000)
                    mm = int(datas[col] - (yyyy * 10000))
                    dd = mm % 100
                    mm = mm / 100
                    val = '%04d/%02d/%02d' %(yyyy, mm, dd)
                elif (col == 3 or col == 10): # 대비율
                    val = locale.format('%.2f', datas[col], 1)
                    val += "%"

                else:
                    val = locale.format('%d', datas[col], 1)

                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.ui.tableWeek.setItem(nRow, col, item)

            if (nRow == 0) :
                self.todayIndex = index
            nRow += 1

            self.tableWeek.resizeColumnsToContents()
        return


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w =Form()
    app.exec_()