# -*- coding: cp949 -*-
#CpRPCurrentPrice
class CpRPCurrentPrice:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS�� ���������� ������� ����. ")
            return
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        return


    def Request(self, code, rtMst, callbackobj):
        # ���簡 ���
        rtMst.objCur.Unsubscribe()
        rtMst.objOfferbid.Unsubscribe()

        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print("��Ż���", self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False


        # ���� ���� ���簡 ������ rtMst �� ����
        rtMst.code = code
        rtMst.name = g_objCodeMgr.CodeToName(code)
        rtMst.cur =  self.objStockMst.GetHeaderValue(11)  # ����
        rtMst.diff =  self.objStockMst.GetHeaderValue(12)  # ���ϴ��
        rtMst.baseprice  =  self.objStockMst.GetHeaderValue(27)  # ���ذ�
        rtMst.vol = self.objStockMst.GetHeaderValue(18)  # �ŷ���
        rtMst.exFlag = self.objStockMst.GetHeaderValue(58)  # �����÷���
        rtMst.expcur = self.objStockMst.GetHeaderValue(55)  # ����ü�ᰡ
        rtMst.expdiff = self.objStockMst.GetHeaderValue(56)  # ����ü����
        rtMst.makediffp()

        rtMst.totOffer = self.objStockMst.GetHeaderValue(71)  # �Ѹŵ��ܷ�
        rtMst.totBid = self.objStockMst.GetHeaderValue(73)  # �Ѹż��ܷ�


        # 10��ȣ��
        for i in range(10):
            rtMst.offer[i] = (self.objStockMst.GetDataValue(0, i))  # �ŵ�ȣ��
            rtMst.bid[i] = (self.objStockMst.GetDataValue(1, i) ) # �ż�ȣ��
            rtMst.offervol[i] = (self.objStockMst.GetDataValue(2, i))  # �ŵ�ȣ�� �ܷ�
            rtMst.bidvol[i] = (self.objStockMst.GetDataValue(3, i) ) # �ż�ȣ�� �ܷ�


        rtMst.objCur.Subscribe(code,rtMst, callbackobj)
        rtMst.objOfferbid.Subscribe(code,rtMst, callbackobj)
