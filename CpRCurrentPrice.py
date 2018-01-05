# -*- coding: cp949 -*-
#CpRPCurrentPrice
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
