import os
import random
import queue
import pickle
import threading

from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import *
from config.errorCode import *
from config.kiwoomFID import *
from pykiwoom.kiwoom import *
from tkinter import *
from tkinter import filedialog
import datetime
from enum import Enum
from UIHelper import *
import pandas as pd
pd.set_option('display.max_columns', 50)
pd.set_option('display.max_rows', 100)
from pandas import DataFrame
from multiprocessing import Queue
import logging

from config.Enum모음 import *

logging.basicConfig(level=logging.debug)

logging.debug("debug")
logging.info("info")
logging.warning("warning")
logging.error("error")
logging.critical("critical")

# { <editor-fold desc="---Worker---">
# 실시간으로 들어오는 데이터를 보고 주문 여부를 판단하는 스레드
class Worker(QThread):
    # argument는 없는 단순 trigger
    # 데이터는 queue를 통해서 전달됨
    trigger = pyqtSignal()

    def __init__(self, data_queue, order_queue):
        super().__init__()
        self.data_queue = data_queue                # 데이터를 받는 용
        self.order_queue = order_queue              # 주문 요청용
        self.timestamp = None
        self.limit_delta = datetime.timedelta(seconds=1)
        self.isBlock = False
        self.isRun = True

    def run(self):
        while self.isRun:
            if not self.data_queue.empty():
                data = self.data_queue.get()
                while self.isRun and self.process_data(data) == False:
                    pass
                if not self.isRun:
                    break
                self.order_queue.put(data)                      # 주문 Queue에 주문을 넣음
                self.timestamp = datetime.datetime.now()        # 이전 주문 시간을 기록함
                self.trigger.emit()


    def process_data(self, data):
        # 시간 제한을 충족하는가?
        time_meet = False
        if self.timestamp is None:
            time_meet = True
        else:
            now = datetime.datetime.now()                           # 현재시간
            delta = now - self.timestamp                            # 현재시간 - 이전 주문 시간
            if delta >= self.limit_delta:
                time_meet = True

        algo_meet = False
        # 알고리즘을 충족하는가?
        if not self.isBlock or not self.isRun:
            algo_meet = True

        # 알고리즘과 주문 가능 시간 조건을 모두 만족하면
        if time_meet and algo_meet:
            return True
        else:
            return False

# } </editor-fold>

class MyWindow(QMainWindow):

# { <editor-fold desc="---__init__---">
    def __init__(self, data_queue, order_queue):
        super().__init__()

        # queue
        self.data_queue = data_queue
        self.order_queue = order_queue

        # thread start
        self.worker = Worker(data_queue, order_queue)
        self.worker.trigger.connect(self.pop_order)
        self.worker.start()
        self.initUI()
        self.Main()

    def push_data(self,data):
        self.data_queue.put(data)


    @pyqtSlot()
    def pop_order(self):
        if not self.order_queue.empty():
            data = self.order_queue.get()
            # print('get:%s' % data)
            if ',' in data:
                datas = str(data).split(',')
                length = len(datas)
                if length == 3:
                    func = getattr(MyWindow, datas[0],datas[1])
                    func(self,datas[2])
                else:
                    print('pop_order error')
            else:
                func = getattr(MyWindow, data)
                func(self)

    @pyqtSlot()
    def workerStart(self):
        self.worker.isBlock = False


    @pyqtSlot()
    def workerPause(self):
        self.worker.isBlock = True


    @pyqtSlot()
    def workerStop(self):
        self.worker.isRun = False

# } </editor-fold>

# { <editor-fold desc="---UI---">

    def initUI(self):
        self.setWindowTitle('[키움증권] 자동매매 프로그램')
        self.setGeometry(0, 0, 500, 1000)
        self.top_qlable_namelist = ['예수금','당일매수', '매매수익','수수료+세금','당일실현손익']
        self.qLineEdit = {}  # QLineEdit 리스트

        self.top_qpblist = {}  # QPushButton 리스트
        self.top_qpushbutton_namelist = ('매수내역상세요청', '매도내역상세요청', '당일매매상위', '잔고내역', '미체결내역','다른이름으로저장','불러오기')

        QGridLayout_내자산 = QGridLayout()

        for value in self.top_qpushbutton_namelist:
            self.top_qpblist.update({value:QPushButton(value, self)})
            self.top_qpblist[value].resize(self.top_qpblist[value].sizeHint())
            self.top_qpblist[value].clicked.connect(lambda state, x=self.top_qpblist[value]: self.btn_button_Clicked(x))

        self.qllist = {}  # QLineEdit 리스트(읽기전용)
        self.optionnamelist = {'int_purchase_amount':'매매금액', # 회당 매수 금액
                               'int_before_store_purchase_amount': '장전매매금액',  # 장 시작전 회당 매수 금액
                               'int_strong_delaytime': '체크타임', # 최대 확인 시간

                               'float_default_strong_limit': '최소체결강도',  # 최소 강도 제한
                               'float_ignore_highpoint': '등락율제한',  #등락율이 클시 구매를 제한한다.
                               'float_fluctuation_detection': '등락대비제한',  # 등락대비가 클시 구매를 제한한다.

                               'float_buy_strong_limit': '강도급상승매수',  # 체결강도 급상승에 따른 무조건 매수
                               'float_condition_fluctuations_strong_highpoint': '강도상승매수',  # 강도상승감지 구매
                               'float_jango_buy_fluctuation': '잔고추매등락율',  # 손익율이 기준보다 하락시 추매 조건을 갖는다.

                               'int_예수금유지금액': '예수금유지금액',  # 예수금유지금액
                               'float_condition_fluctuations_strong_lowpoint': '강도급하락매도',  # 강도하락감지 판매
                               'float_condition_fluctuations_price_lowpoint': '가격급하락매도',  # 가격하락감지 판매

                               'float_jango_reg_fluctuation': '잔고관리등락율',  # 손익율이 기준보다 높을시 관리 조건을 갖는다.
                               'float_jango_condition_fluctuations_strong_lowpoint': '잔고강도급하락',  # 강도하락감지 판매 (잔고 내 손해 종목의)
                               'float_jango_condition_fluctuations_price_lowpoint': '잔고가격급하락',  # 가격하락감지 판매 (잔고 내 손해 종목의)

                               'float_sell_ignore_strong_limit': '판매제한강도1',  # 체결강도 이상시 판매 제한
                               'float_sell_strong_limit': '판매제한강도2',  # 체결강도 조건에 따른 판매 보류
                               'float_최대수익구간': '최대수익구간',  # 최대수익구간 달성시 시장가 매도

                               'int_장마무리예수금': '장마무리예수금',
                               'int_장마무리매수단위': '장마무리매수단위',
                               'float_매수단위최고치': '매수단위최고치',

                               'time_Cancel_Outstanding': '미체결취소',  # 장 시작 후 미체결 종목 취소 시간
                               'time_Long_Standby': '장시간대기',  # 장 시작 후 대기하는 시간
                               'time_Intermediate_Finish_Time': '중간마무리',  # 장 중간 마무리 시간
                               # 'float_strong_sell': '강도미달판매',  # 강도 기준치 보다 하락시 판매
                               # 'float_condition_lowpoint_today':'손익하락매도',  # 판매 제한 강도보다 낮으며 손익율이 기준보다 하락시 판매한다.
                               # 'float_sell_strong_limit_and_price_lowpoint': '가격하락율매도',  # 보류 체결강도보다 낮은 상태의 가격하락감지 판매
                               # 'int_transaction_volume_limit': '최소거래량',  # 최소 거래량 제한
                               # 'int_ignore_buy_and_off_time': '매수매도딜레이',  # 매수 후 바로 매도 금지 시간
                               # 'int_transaction_volume_detection': '거래량증가량',  # 거래량증가량 감지
                               }

        self.qpblist = {}  # QPushButton 리스트

        for value in self.optionnamelist:
            self.qllist.update({value: QLineEdit(self)})
            self.qllist[value].setReadOnly(True)
            self.qllist[value].setFixedSize(80,20)
            # self.qLineEdit.update({value:QLineEdit(self)})
            # self.qLineEdit[value].setFixedSize(80, 20)
            self.qpblist.update({value:QPushButton('APPLY',self)})
            self.qpblist[value].setFixedSize(80, 20)
            self.qpblist[value].clicked.connect(lambda state, x=self.qpblist[value]: self.btn_apply_Clicked(x))
        
        self.qclist = {} #QCheckBox 리스트 - DEBUG OPTION

        hboxlist = {}
        hboxlist.update({'control': QHBoxLayout()})

        hboxlist.update({'top_myinfo': QHBoxLayout()})
        for value in self.top_qlable_namelist:
            hboxlist['top_myinfo'].addWidget(QLabel(value, self))
            self.qLineEdit.update({value: QLineEdit(self)})
            self.qLineEdit[value].setReadOnly(True)
            self.qLineEdit[value].setFixedSize(70, 20)
            # self.qLineEdit[value].resize(self.qLineEdit[value].sizeHint())
            hboxlist['top_myinfo'].addWidget(self.qLineEdit[value])

        QGridLayout_checkbox = QGridLayout()

        debugoption_cnt = 0
        linelen = 8

        for value in DEBUGTYPE:
            key = '%s_%s' % ('DEBUGTYPE', value.name)
            self.qclist.update({key:QCheckBox(value.name, self)})
            self.qclist[key].clicked.connect(lambda state, x=key: self.checkBox_clicked(x))
            QGridLayout_checkbox.addWidget(self.qclist[key], int(debugoption_cnt / linelen), debugoption_cnt % linelen)
            debugoption_cnt += 1

        line = int(debugoption_cnt / linelen) + 1

        self.qclist.update({'자동거래종목': QCheckBox('자동거래종목', self)})
        self.qclist['자동거래종목'].setEnabled(False)
        QGridLayout_checkbox.addWidget(self.qclist['자동거래종목'], line, 0)
        self.qclist.update({'매매딜레이': QCheckBox('매매딜레이', self)})
        self.qclist['매매딜레이'].setEnabled(False)
        QGridLayout_checkbox.addWidget(self.qclist['매매딜레이'], line, 1)

        useroptionlist = ['자동매수', '자동매도','장전매수', '손절매도', '장마무리매수']

        cnt = 0
        line += 1
        for key in useroptionlist:
            self.qclist.update({key: QCheckBox(key, self)})
            self.qclist[key].clicked.connect(lambda state, x=key: self.checkBox_clicked(x))
            QGridLayout_checkbox.addWidget(self.qclist[key], line, cnt)
            cnt += 1

        hboxlist.update({'설정제어': QHBoxLayout()})
        self.qLineEdit.update({'설정제어': QLineEdit(self)})
        self.qLineEdit['설정제어'].setFixedSize(110, 20)
        hboxlist['설정제어'].addWidget(QLabel('설정제어', self))
        hboxlist['설정제어'].addWidget(self.qLineEdit['설정제어'],Qt.AlignLeft)

        hboxlist.update({'주식기본정보': QHBoxLayout()})
        self.qLineEdit.update({'주식기본정보': QLineEdit(self)})
        self.qLineEdit['주식기본정보'].setFixedSize(110, 20)
        self.qpblist.update({'주식기본정보': QPushButton('주식기본정보', self)})
        self.qpblist['주식기본정보'].setFixedSize(110, 20)
        self.qpblist['주식기본정보'].clicked.connect(self.btn_stockinfo_Clicked)
        self.qpblist.update({'주식매도': QPushButton('주식매도', self)})
        self.qpblist['주식매도'].clicked.connect(self.btn_stocksell_Clicked)
        self.qpblist['주식매도'].setFixedSize(110, 20)
        self.qpblist.update({'주식매수': QPushButton('주식매수', self)})
        self.qpblist['주식매수'].clicked.connect(self.btn_stockbuy_Clicked)
        self.qpblist['주식매수'].setFixedSize(110, 20)
        self.qpblist.update({'주문취소': QPushButton('주문취소', self)})
        self.qpblist['주문취소'].clicked.connect(self.btn_stock_cancel_Clicked)
        self.qpblist['주문취소'].setFixedSize(110, 20)
        ql = QLabel('수동제어', self)
        ql.setFixedSize(80, 20)
        hboxlist['주식기본정보'].addWidget(ql,0,Qt.AlignLeft)
        hboxlist['주식기본정보'].addWidget(self.qLineEdit['주식기본정보'],Qt.AlignLeft)
        hboxlist['주식기본정보'].addWidget(self.qpblist['주식기본정보'],Qt.AlignLeft)
        hboxlist['주식기본정보'].addWidget(self.qpblist['주식매도'],Qt.AlignLeft)
        hboxlist['주식기본정보'].addWidget(self.qpblist['주식매수'], Qt.AlignLeft)
        hboxlist['주식기본정보'].addWidget(self.qpblist['주문취소'], Qt.AlignLeft)

        cnt = 0
        boxkey = ''
        for key, value in self.optionnamelist.items():
            if cnt % 3 == 0:
                boxkey = key
            hboxlist.update({key: QHBoxLayout()})
            ql = QLabel(value,self)
            ql.setFixedSize(100, 20)
            hboxlist[boxkey].addWidget(ql,0, Qt.AlignLeft)
            hboxlist[boxkey].addWidget(self.qllist[key],Qt.AlignRight)
            hboxlist[boxkey].addWidget(self.qpblist[key],Qt.AlignRight)
            cnt += 1

        hboxlist.update({'topbutton': QHBoxLayout()})
        for key,value in self.top_qpblist.items():
            hboxlist['topbutton'].addWidget(value)

        self.표_잔고 = QTableWidget(self)
        self.표_잔고.setColumnCount(13)
        self.표_잔고.setHorizontalHeaderLabels(['시간', '구분', '코드', '종목명', '현재가', '등락율', '체강','수익율','거래가','거래량','매매금액','평가금액','손익금'])
        self.표_잔고.setColumnWidth(Enum_표_잔고.시간.value, 50)
        self.표_잔고.setColumnWidth(Enum_표_잔고.구분.value, 40)
        self.표_잔고.setColumnWidth(Enum_표_잔고.종목코드.value, 50)
        self.표_잔고.setColumnWidth(Enum_표_잔고.현재가.value, 60)
        self.표_잔고.setColumnWidth(Enum_표_잔고.등락율.value, 45)
        self.표_잔고.setColumnWidth(Enum_표_잔고.체결강도.value, 45)
        self.표_잔고.setColumnWidth(Enum_표_잔고.수익율.value, 45)
        self.표_잔고.setColumnWidth(Enum_표_잔고.매수가.value, 60)
        self.표_잔고.setColumnWidth(Enum_표_잔고.매수량.value, 45)
        self.표_잔고.setColumnWidth(Enum_표_잔고.매매금액.value, 70)
        self.표_잔고.setColumnWidth(Enum_표_잔고.평가금액.value, 70)
        self.표_잔고.setColumnWidth(Enum_표_잔고.손익금.value, 70)
        self.표_잔고.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)

        self.표_미체결 = QTableWidget(self)
        self.표_미체결.setColumnCount(7)
        self.표_미체결.setHorizontalHeaderLabels(['시간', '구분', '코드', '종목명', '현재가', '등락율', '체강'])
        self.표_미체결.setColumnWidth(Enum_표_미체결.시간.value, 50)
        self.표_미체결.setColumnWidth(Enum_표_미체결.매매구분.value, 40)
        self.표_미체결.setColumnWidth(Enum_표_미체결.종목코드.value, 50)
        self.표_미체결.setColumnWidth(Enum_표_미체결.현재가.value, 60)
        self.표_미체결.setColumnWidth(Enum_표_미체결.등락율.value, 45)
        self.표_미체결.setColumnWidth(Enum_표_미체결.체결강도.value, 45)
        self.표_미체결.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)

        tabs_표 = QTabWidget()
        tabs_표.addTab(self.표_잔고, '잔고')
        tabs_표.addTab(self.표_미체결, '미체결')

        hboxlist.update({'tabs_표': QHBoxLayout()})
        hboxlist['tabs_표'].addWidget(tabs_표)

        vbox = QVBoxLayout()
        vbox.addLayout(hboxlist['top_myinfo'])
        vbox.addLayout(hboxlist['topbutton'])
        for key, value in self.optionnamelist.items():
            vbox.addLayout(hboxlist[key])
        vbox.addLayout(hboxlist['설정제어'])
        vbox.addLayout(hboxlist['주식기본정보'])
        vbox.addLayout(hboxlist['tabs_표'])
        # vbox.addLayout(hboxlist['표_매수'])
        # vbox.addLayout(hboxlist['표_완료'])
        vbox.addLayout(QGridLayout_checkbox)

        widget = QWidget()
        widget.setLayout(vbox)
        self.setCentralWidget(widget)

    def setUI(self):
        for key in self.optionnamelist:
            try:
                self.setText_qLineEdit_option(key,str(self.user_dict[key]))
            except:
                self.user_dict.update({key : 0})
                print('[except key:%s]' % key)
        for key in self.qclist:
            if key not in self.user_dict:
                self.user_dict.update({key: False})
            self.qclist[key].setChecked(self.user_dict[key])

    def setText_qLineEdit_myinfo(self,key,value):
        self.qLineEdit[key].setText(str(self.ConvertText(value, 'int')))

    def setText_qLineEdit_option(self,key,value):
        if 'float_' in key:
            self.qllist[key].setText(str(self.ConvertText(value, 'float')))
        elif 'int_' in key:
            self.qllist[key].setText(str(self.ConvertText(value, 'int')))
        elif 'time_' in key:
            self.qllist[key].setText('{0:06d}'.format(int(value)))
        self.save_option()

    def checkBox_clicked(self, key):
        self.user_dict.update({key: self.qclist[key].isChecked()})
        self.save_option()

    def convert_finish_price_text_colors(self, price,sell_price):
        if sell_price > price:
            txt_incom_rate = '<span style="color:blue">%s</span>' % price
        elif sell_price < price:
            txt_incom_rate = '<span style="color:red">%s</span>' % price
        else:
            txt_incom_rate = price
        return txt_incom_rate

    def btn_button_Clicked(self,button):
        for key, value in self.top_qpblist.items():
            if value == button:
                if key == '매수내역상세요청':
                    self.block_request_tr_계좌별주문체결내역상세요청(2)
                elif key == '매도내역상세요청':
                    self.block_request_tr_계좌별주문체결내역상세요청(1)
                elif key == '당일매매상위':
                    self.request_tr_당일거래량상위()  # 당일매매상위요청
                elif key == '잔고내역':
                    self.push_data("block_request_tr_계좌평가잔고내역요청")
                elif key == '미체결내역':
                    self.block_request_tr_미체결요청()
                elif key == '다른이름으로저장':
                    self.save_as_option()
                elif key == '불러오기':
                    self.load_option2()

    def test_action(self):
        # self.kiwoom.SetRealReg("1000", "005930", "20;10", 0)
        # for key in self.jango_item_dict:
        #     self.autotradingSetRealReg(self.screen_real_stock, key, "20;10;11;12;13;228", 1)
        # for key in self.realdata_stock_dict:
        #     self.autotradingSetRealReg(self.screen_real_stock, key, "20;10;11;12;13;228", 1)
        pass



    def btn_stockinfo_Clicked(self):
        self.request_tr_주식기본정보요청('결과통계',self.qLineEdit['주식기본정보'].text())

    def btn_stockbuy_Clicked(self):
        sCode = self.qLineEdit['주식기본정보'].text()
        if sCode in self.realdata_stock_dict:
            quantity = self.GetPuchaseQuantity(self.realdata_stock_dict[sCode]['현재가'])
            self.kiwoom_SendOrder_present_price_buy('수동','수동구매', sCode=sCode, quantity=quantity)

    def btn_stocksell_Clicked(self):
        sCode = self.qLineEdit['주식기본정보'].text()
        if sCode in self.contract_sell_item_dict['현금매수']:
            if self.contract_sell_item_dict['현금매수'][sCode] and self.contract_sell_item_dict['현금매수'][sCode]['상태'] == '매수':
                quantity = self.contract_sell_item_dict['현금매수'][sCode]['체결수량']
                self.kiwoom_SendOrder_present_price_sell('수동','수동판매', sCode, quantity)
            else:
                print('당일구매대기종목 확인(매도 실패)')
        else:
            if sCode in self.jango_item_dict:
                quantity = self.jango_item_dict[sCode]['매매가능수량']
                self.kiwoom_SendOrder_present_price_sell('수동','수동판매', sCode, quantity)

    def btn_stock_cancel_Clicked(self):
        sCode = self.qLineEdit['주식기본정보'].text()
        print('취소기능 없음')

    def btn_apply_Clicked(self,button):
        for key, value in self.qpblist.items():
            if value == button:
                try:
                    if 'float_' in key:
                        self.user_dict[key] = float(self.qLineEdit['설정제어'].text())
                    elif 'int_' in key:
                        self.user_dict[key] = int(self.qLineEdit['설정제어'].text())
                    elif 'time_' in key:
                        self.user_dict[key] = int(self.qLineEdit['설정제어'].text())
                    self.save_option()
                except:
                    print("잘못된 입력값입니다.")
                self.setText_qLineEdit_option(key,self.user_dict[key])

# } </editor-fold>

# { <editor-fold desc="---Main---">

    def Main(self):
        self.FIDs = FIDLIST()

        ### 스크린번호 모음
        self.screen_real_stock = "1000" #종목별 할당할 스크린 번호
        self.screen_trading_stock = "6000" #주식 거래 스크린 번호

        # { <editor-fold desc="고정변수">
        self.user_dict = {
            'account_pass': 'saya8383',
            'pass_index' : '00',
            'float_jango_buy_fluctuation' : 5, #손절
            'float_condition_lowpoint_today': 1.2,  # 손익하락감지 매도
            'int_purchase_amount' : 200000, #구매 액
            'float_default_strong_limit': 100,  # 기본 구매 체결강도
            'float_buy_strong_limit': 100, # 매수 체결강도 (기본 구매 체결강도를 무시)
            'float_sell_strong_limit': 100,  # 판매 제한 체결강도
            'float_sell_ignore_strong_limit': 100,  # 판매 제한 체결강도 (체결강도 이상시 매도 제한)
            'float_condition_fluctuations_strong_highpoint':50, #상승 체결강도 구매조건
            'float_condition_fluctuations_strong_lowpoint':10, #하락 체결강도 판매조건
            'float_condition_fluctuations_price_lowpoint': 0.3, #가격 하락 판매조건
            'int_strong_delaytime': 10,  # 체결강도 확인 시간차
            'float_fluctuation_detection':1, #등락대비가 클시 구매를 제한한다.
            'int_transaction_volume_detection': 1000000,  # 거래량변화량 감지
            'float_ignore_highpoint': 5,  # 등락율이 클시 구매를 제한한다.
            'int_transaction_volume_limit': 1000000,  # 최소 거래량 제한
            'float_sell_strong_limit_and_price_lowpoint': 0.5,  # 보류 체결강도보다 낮은 상태의 가격하락감지 판매
            'float_strong_sell' : 80, # 강도 기준치 보다 하락시 판매
            'float_jango_reg_fluctuation' : 0, # 잔고 관리 등록 기준 수익율
            'int_ignore_buy_and_off_time': 300,  # 매수 후 바로 매도 금지 시간
            'float_jango_condition_fluctuations_strong_lowpoint': 50,  # 강도하락감지 판매 (잔고 내 손해 종목의)
            'float_jango_condition_fluctuations_price_lowpoint': 10,  # 가격하락감지 판매 (잔고 내 손해 종목의)
            'int_예수금유지금액': 10000000,  # 예수금유지금액
            'float_최대수익구간': 10,  # 최대수익구간 달성시 시장가 매도
            'time_Cancel_Outstanding': 90100,  # 장 시작 후 미체결 종목 취소 시간
            'time_Long_Standby': 90500,  # 장 시작 후 대기하는 시간
            'time_Intermediate_Finish_Time': 123000,  # 장 중간 마무리 시간
        }  # 세이브데이터


        # } </editor-fold>

        ### 변수모음
        self.myAccount = None #계좌정보
        #5557845510
        #8148453411(모의)
        self.stock_state = -1 #-1:정보 없음, 0:장시작 전, 3:장 중, 2:장 종료 10분전 동시호가, 4:장 종료
        self.deposit = 0 # 예수금
        self.today_total_buy = 0 # 당일매수
        self.total_incom = 0 # 총수익
        self.total_tax = 0 # 총 세금
        self.total_commission = 0 #총 수수료
        self.jango_item_dict = {} # 잔고 내역
        self.contract_sell_item_dict = {'현금매도':{},'현금매수':{}}  # 매도매수한 종목 관리

        self.contract_complete_selling_price = {} #금일 매도한 종목의 판매금액 (손해 재매수 방지)
        self.jango_contract_stock_codelist = []  # 잔고 중 매도 타이밍 관리 종목
        self.interest_stock_codelist = [] # 관심 종목 코드 리스트

        self.realdata_stock_dict = {} #실시간 데이터관리
        self.realdata_analysis_dict = {}  # 실시간 분석 데이터관리
        self.realdata_screen_dict = {'screen_cnt':{},'screen_num':{}} #실시간 데이터 스크린 관리
        self.interest_stock_dict = {}
        self.표_잔고_관리 = {}
        self.표_매수_관리 = {}
        self.표_잔고_리스트 = []
        self.표_미체결_관리 = []

        self.hour_timer = 0
        self.highprice_timer = 0
        self.automatic_trading_Search_count = 0
        ###
        self.readyAutoTradingSystem = False
        self.readyAutoTradingStock = False
        self.readyAutoTradingStock_delay_buy = False
        self.readyStockMarket = True
        self.endOfChapterEvent = False
        self.middleChapterEvent = False
        self.장시작후미체결취소 = False

        self.timer_기준시간_stamp = 0

        self.load_option()
        self.setUI()  # UI setting
        self.mytime_today = datetime.datetime.today().strftime("%Y%m%d")

        self.kiwoom = Kiwoom()
        self.kiwoom.ocx.OnEventConnect.connect(self._handler_login)
        self.kiwoom.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.kiwoom.ocx.OnReceiveChejanData.connect(self._handler_chejan_data)
        self.kiwoom.CommConnect(block=True)

# } </editor-fold>

# { <editor-fold desc="---start---">

    def start(self):
        self.messagePrint(DEBUGTYPE.시스템.name, "---Start---")
        if self.week_check() == 1:
            self.messagePrint(DEBUGTYPE.시스템.name, "---주말입니다. 프로그램 종료---")
            self.ApplicationQuit()
        else:
            self.push_data('block_request_tr_계좌평가잔고내역요청')

    def timer_start(self):
        print('--Timer Start--')
        self.timer = QTimer(self)
        now = datetime.datetime.now()
        self.timer_기준시간_stamp = now.timestamp()
        self.timer.start(1000)
        self.timer.timeout.connect(self.timer_slot)

    def timer_slot(self):
        if self.stock_state == 3:
            now = datetime.datetime.now()
            if now.timestamp() - self.timer_기준시간_stamp > 1800:
                self.push_data('request_tr_당일거래량상위')
                self.timer_기준시간_stamp = now.timestamp()
        else:
            self.timer.stop()
            print('--Timer Stop--')


# } </editor-fold>

# { <editor-fold desc="---event handler---"> 이벤트 핸들러

    # { <editor-fold desc="-이벤트 핸들러 추가 삭제-"> 이벤트 핸들러

    def autotradingSetRealReg(self,strScreenNo,strCodeList,strFidList,strOptType): #자동매수용 실시간 이벤트 등록
        '''
        :param strScreenNo:화면번호
        :param strCodeList:종목코드 리스트
        :param strFidList:실시간 FID리스트
        :param strOptType:실시간 등록 타입, 0또는 1
        :return: 실시간 등록타입을 0으로 설정하면 등록한 종목들은 실시간 해지되고 등록한 종목만 실시간 시세가 등록됩니다.
          실시간 등록타입을 1로 설정하면 먼저 등록한 종목들과 함께 실시간 시세가 등록됩니다
        '''
        if strCodeList in self.realdata_screen_dict['screen_num']:
            self.messagePrint(DEBUGTYPE.스크린.name, "[%s][스크린번호:%s][%s][%s][%s]" % ('이미등록된종목', strScreenNo, strCodeList, strFidList, strOptType))
        elif strScreenNo in self.realdata_screen_dict['screen_cnt']:
            cnt = self.realdata_screen_dict['screen_cnt'][strScreenNo]
            if cnt < 200:
                self.realdata_screen_dict['screen_cnt'][strScreenNo] = cnt + 1
                self.realdata_screen_dict['screen_num'].update({strCodeList: strScreenNo})
                self.kiwoom.SetRealReg(strScreenNo, strCodeList, strFidList, strOptType)  # 등록
                self.messagePrint(DEBUGTYPE.스크린.name, "[%s][스크린번호:%s][%s][%s][%s]" % ('자동매수등록', strScreenNo, strCodeList, strFidList, strOptType))
            else:
                self.autotradingSetRealReg(int(strScreenNo) + 1, strCodeList, strFidList, strOptType)
        else:
            self.realdata_screen_dict['screen_cnt'].update({strScreenNo:0})
            self.realdata_screen_dict['screen_num'].update({strCodeList:strScreenNo})

    def autotradingSetRealRemove(self,strCodeList): #자동매수용 실시간 이벤트 해지
        '''
        :param strScreenNo:화면번호
        :param strCodeList:종목코드 리스트
        :return:
        '''
        if strCodeList in self.realdata_screen_dict['screen_num']:
            strScreenNo = self.realdata_screen_dict['screen_num'][strCodeList]
            self.realdata_screen_dict['screen_num'].pop(strCodeList)
            cnt = self.realdata_screen_dict['screen_cnt'][strScreenNo]
            self.realdata_screen_dict['screen_cnt'][strScreenNo] = cnt - 1
            self.kiwoom.SetRealRemove(strScreenNo, strCodeList)  # 실시간 이벤트 해지
            self.messagePrint(DEBUGTYPE.스크린.name, "[%s][스크린번호:%s][%s]" % ('자동매수삭제',strScreenNo,strCodeList))

    def kiwoomSetRealRemove(self,strScreenNo,strCodeList = "ALL"): #실시간 이벤트 해지
        '''
        :param strScreenNo:화면번호
        :param strCodeList:종목코드 리스트
        :return:
        '''
        self.kiwoom.SetRealRemove(strScreenNo, strCodeList)  # 실시간 이벤트 해지
        self.messagePrint(DEBUGTYPE.스크린.name, "[%s][스크린번호:%s][%s]" % ('리얼데이터삭제',strScreenNo,strCodeList))

    def kiwoomSetRealRemoveAll(self): #모든 실시간 이벤트 해지
        for i in range(len(self.realdata_screen_dict['screen_cnt'])):
            self.kiwoom.SetRealRemove(self.realdata_screen_dict['screen_num'][i], "ALL")  # 실시간 이벤트 해지

    # } </editor-fold>

    # { <editor-fold desc="-로그인 핸들러-"> 로그인 핸들러

    def _handler_login(self, err_code): #로그인 핸들러
        if err_code == 0:
            self.load_DataFrame()
            self.messagePrint(DEBUGTYPE.시스템.name, "---login---")
            accounts = self.kiwoom.GetLoginInfo("ACCNO")
            ServerGubun = self.kiwoom.GetLoginInfo("GetServerGubun")
            if ServerGubun == '1':
                self.messagePrint(DEBUGTYPE.내정보.name, "[서버구분 : %s]" % '모의서버')
            else:
                self.messagePrint(DEBUGTYPE.내정보.name, "[서버구분 : %s]" % ('실서버'))
            self.myAccount = accounts[0]
            self.messagePrint(DEBUGTYPE.내정보.name, "[계좌번호 : %s]" % self.myAccount)
            self.push_data('block_request_tr_계좌평가현황요청')
        else:
            self.messagePrint(DEBUGTYPE.error.name, "[ERROR][재시도][%s] %s" % (err_code, errors(err_code)))
            self.kiwoom.CommConnect(block=True)

    # } </editor-fold>

    # { <editor-fold desc="-실시간 핸들러-"> 실시간 핸들러

    def _handler_real_data(self, sCode, real_type, data): #실시간 핸들러
        if real_type == "장시작시간":
            state = self.kiwoom.GetCommRealData(sCode, 215)
            stock_time = self.kiwoom.GetCommRealData(sCode, 20)
            remained_time = self.kiwoom.GetCommRealData(sCode, 214)
            if state == '3':
                self.stock_state = 3
                self.messagePrint(DEBUGTYPE.시스템.name, "---[장시작]%s %s %s" % (state, stock_time, remained_time))
                self.push_data('timer_start')
            elif state == '8':
                self.Close_of_chapter_event()  # 장 종료 이벤트 발생
        elif real_type == "주식체결":
            if self.stock_state == 3 and self.readyAutoTradingSystem == True: #장중일때만
                myTime = int(self.kiwoom.GetCommRealData(sCode, 20))
                if self.장시작후미체결취소 == False:
                    if myTime > self.user_dict['time_Cancel_Outstanding']:  # time_Long_Standby이 넘으면
                        self.장시작후미체결취소 = True
                        self.push_data('block_request_tr_미체결요청')
                if self.readyAutoTradingStock_delay_buy == False:
                    if myTime > self.user_dict['time_Long_Standby']:  # time_Long_Standby이 넘으면
                        self.Start_of_chapter_event() #장 시작 이벤트 발생
                if self.middleChapterEvent == False:
                    if myTime > self.user_dict['time_Intermediate_Finish_Time']:  # 12시 30분이 넘으면
                        self.Middle_chapter_event() #장 중간 마무리 이벤트 발생
                if self.endOfChapterEvent == False:
                    if myTime > 150000: # 3시가 넘으면
                        self.End_of_chapter_event() #장 마감 준비 이벤트 발생

                present_price = self.emptyToZero(self.kiwoom.GetCommRealData(sCode, 10), doabs = True)
                fluctuations = self.emptyToZero(self.kiwoom.GetCommRealData(sCode, 12),1) #등락율
                previous = self.emptyToZero(self.kiwoom.GetCommRealData(sCode, 11)) #전일대비
                strong = self.emptyToZero(self.kiwoom.GetCommRealData(sCode, 228),1) #체결강도
                cumulative_transaction = self.emptyToZero(self.kiwoom.GetCommRealData(sCode, 13))  # 누적거래량
                누적거래대금 = self.emptyToZero(self.kiwoom.GetCommRealData(sCode, 14))
                sName = self.kiwoom.GetMasterCodeName(sCode)

                realdata_dict = {}
                realdata_dict.update({'종목코드': sCode})
                realdata_dict.update({'종목명': sName})
                realdata_dict.update({'현재가': present_price})
                realdata_dict.update({'등락율': fluctuations})
                realdata_dict.update({'전일대비': previous})
                realdata_dict.update({'체결강도': strong})
                realdata_dict.update({'누적거래량': cumulative_transaction})
                realdata_dict.update({'누적거래대금': 누적거래대금})

                if sCode not in self.realdata_stock_dict:
                    self.realdata_stock_dict.update({sCode: {}})
                    self.realdata_stock_dict[sCode].update({'종목명': sName})
                self.realdata_stock_dict[sCode].update({'현재가': present_price})
                self.realdata_stock_dict[sCode].update({'전일대비': previous})
                self.realdata_stock_dict[sCode].update({'현재등락율': fluctuations})
                self.realdata_stock_dict[sCode].update({'현재체결강도': strong})
                self.realdata_stock_dict[sCode].update({'현재누적거래량': cumulative_transaction})
                self.realdata_stock_dict[sCode].update({'현재누적거래대금': 누적거래대금})
                if '체결강도시간' not in self.realdata_stock_dict[sCode]:
                    self.realdata_stock_dict[sCode].update({'체결강도시간': myTime})
                    self.realdata_stock_dict[sCode].update({'체결강도': strong})
                    self.realdata_stock_dict[sCode].update({'등락율': fluctuations})
                    self.realdata_stock_dict[sCode].update({'누적거래량': cumulative_transaction})
                    self.realdata_stock_dict[sCode].update({'누적거래대금': 누적거래대금})
                    if sCode not in self.realdata_analysis_dict:
                        self.realdata_analysis_dict.update({sCode: queue.Queue()})
                else:
                    put_data = {'체결강도시간': myTime,'체결강도': strong,'등락율': fluctuations, '누적거래량':cumulative_transaction, '누적거래대금':누적거래대금}
                    self.realdata_analysis_dict[sCode].put(put_data)
                previous_time = self.realdata_stock_dict[sCode]['체결강도시간']
                while self.realdata_analysis_dict[sCode].qsize() > 0:
                    if self.ConvertTimeChange(previous_time,self.user_dict['int_strong_delaytime']) <= myTime:
                        get_data = self.realdata_analysis_dict[sCode].get()
                        self.realdata_stock_dict[sCode]['체결강도시간'] = get_data['체결강도시간']
                        self.realdata_stock_dict[sCode]['체결강도'] = get_data['체결강도']
                        self.realdata_stock_dict[sCode]['등락율'] = get_data['등락율']
                        self.realdata_stock_dict[sCode]['누적거래량'] = get_data['누적거래량']
                        self.realdata_stock_dict[sCode]['누적거래대금'] = get_data['누적거래대금']
                        previous_time = self.realdata_stock_dict[sCode]['체결강도시간']
                    else:
                        break
                previous_strong = self.realdata_stock_dict[sCode]['체결강도']
                if previous_strong == 0:
                    fluctuations_strong = 0
                    fluctuations_time = 0
                    contrast_fluctuations = 0
                    fluctuations_transaction = 0
                    누적거래대금_변화량 = 0
                else:
                    fluctuations_strong = round(strong - previous_strong, 2)  # 체결강도 등락대비
                    previous_fluctuations = self.realdata_stock_dict[sCode]['등락율']
                    contrast_fluctuations = round(fluctuations - previous_fluctuations, 2)  #현재가 등락대비
                    fluctuations_time = self.ConvertTimeChange(myTime, -previous_time) #경과시간
                    fluctuations_transaction = cumulative_transaction - self.realdata_stock_dict[sCode]['누적거래량'] #누적거래량 변화량
                    누적거래대금_변화량 = 누적거래대금 - self.realdata_stock_dict[sCode]['누적거래대금']
                self.realdata_stock_dict[sCode].update({'경과시간': fluctuations_time})
                self.realdata_stock_dict[sCode].update({'체결강도증감':fluctuations_strong})
                self.realdata_stock_dict[sCode].update({'가격등락대비': contrast_fluctuations})
                self.realdata_stock_dict[sCode].update({'거래량변화량': fluctuations_transaction})
                self.realdata_stock_dict[sCode].update({'누적거래대금_변화량': 누적거래대금_변화량})

                realdata_dict.update({'경과시간': fluctuations_time})
                realdata_dict.update({'체결강도증감': fluctuations_strong})
                realdata_dict.update({'가격등락대비': contrast_fluctuations})
                realdata_dict.update({'현재누적거래량': cumulative_transaction})
                realdata_dict.update({'거래량변화량': fluctuations_transaction})
                realdata_dict.update({'누적거래대금_변화량': 누적거래대금_변화량})

                if sCode in self.표_미체결_관리:
                    rowPosition = self.표_미체결_관리.index(sCode)
                    self.표_미체결.setItem(rowPosition, Enum_표_미체결.현재가.value, QTableWidgetItem(format(present_price,',')))
                    self.표_미체결.setItem(rowPosition, Enum_표_미체결.체결강도.value, QTableWidgetItem(str(strong)))
                    self.표_미체결.setItem(rowPosition, Enum_표_미체결.등락율.value, QTableWidgetItem(str(fluctuations)))
                    self.ConvertColorValue(self.표_미체결,fluctuations, rowPosition, Enum_표_미체결.등락율.value)

                for i, v in enumerate(self.표_잔고_리스트):
                    if v == sCode:
                        self.표_잔고.setItem(i, Enum_표_잔고.현재가.value, QTableWidgetItem(format(present_price, ',')))
                        self.표_잔고.setItem(i, Enum_표_잔고.체결강도.value, QTableWidgetItem(str(strong)))
                        self.표_잔고.setItem(i, Enum_표_잔고.등락율.value, QTableWidgetItem(str(fluctuations)))
                        self.ConvertColorValue(self.표_잔고,fluctuations, i, Enum_표_잔고.등락율.value)

                if sCode in self.jango_contract_stock_codelist:  # 잔고 중 매도 타이밍 관리 종목
                    RQName = '잔고'
                    if sCode in self.jango_item_dict:  # 잔고 종목
                        quantity = self.jango_item_dict[sCode]['매매가능수량']
                        purchase_price = self.emptyToZero(self.jango_item_dict[sCode]['매입가'])
                        incom_rate = self.get_incom_rate(present_price, purchase_price)  # 수익율 계산
                    else:
                        self.messagePrint(DEBUGTYPE.error.name, "[ERROR][잔고에 존재하지 않는 종목이다.][%s][%s]" % (sCode, sName))
                        self.jango_contract_stock_codelist.remove(sCode)
                        return
                    realdata_dict.update({'매입가': purchase_price})
                    realdata_dict.update({'수익율': incom_rate})
                    realdata_dict.update({'매매수량': quantity})
                    self.jango_item_dict[sCode].update({'수익율': incom_rate})

                    self.proceed_sell(RQName, realdata_dict)

                    if incom_rate < 0:
                        self.messagePrint(DEBUGTYPE.거래.name, "[잔고관리해제][종목코드:%s][종목명:%s][매수가:%s][현재가:%s][등락율:%s][수익율:%s][체결강도:%s][누적거래량:%s(%s)]" % (
                            sCode, sName, purchase_price, present_price, fluctuations, incom_rate, strong, cumulative_transaction, fluctuations_transaction))
                        self.jango_contract_stock_codelist.remove(sCode)
                    elif sCode in self.표_잔고_관리:
                        rowPosition = self.표_잔고_관리[sCode]
                        평가금액 = present_price * quantity
                        손익금 = 평가금액 - self.jango_item_dict[sCode]['매매금액']
                        self.jango_item_dict[sCode].update({'손익금': 손익금})
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.수익율.value, QTableWidgetItem(str(incom_rate)))
                        self.ConvertColorValue(self.표_잔고,incom_rate, rowPosition, Enum_표_잔고.수익율.value)
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.평가금액.value, QTableWidgetItem(format(평가금액, ',')))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.손익금.value, QTableWidgetItem(format(손익금, ',')))

                elif sCode in self.contract_sell_item_dict['현금매수']:
                    if self.contract_sell_item_dict['현금매수'][sCode]['상태'] != '접수':
                        if '체결단가' in self.contract_sell_item_dict['현금매수'][sCode]:
                            구분 = self.contract_sell_item_dict['현금매수'][sCode]['구분']
                            purchase_price = self.emptyToZero(self.contract_sell_item_dict['현금매수'][sCode]['체결단가'])
                            quantity = self.emptyToZero(self.contract_sell_item_dict['현금매수'][sCode]['체결수량'])
                            incom_rate = self.get_incom_rate(present_price, purchase_price)  # 수익율 계산
                            self.contract_sell_item_dict['현금매수'][sCode].update({'수익율': incom_rate})

                            realdata_dict.update({'매입가': purchase_price})
                            realdata_dict.update({'수익율': incom_rate})
                            realdata_dict.update({'매매수량': quantity})

                            self.proceed_sell(구분, realdata_dict)

                            if sCode in self.표_매수_관리:
                                rowPosition = self.표_매수_관리[sCode]
                                평가금액 = present_price * quantity
                                손익금 = 평가금액 - self.contract_sell_item_dict['현금매수'][sCode]['체결누계금액']
                                self.contract_sell_item_dict['현금매수'][sCode].update({'손익금':손익금})
                                self.표_잔고.setItem(rowPosition, Enum_표_잔고.수익율.value, QTableWidgetItem(str(incom_rate)))
                                self.ConvertColorValue(self.표_잔고,incom_rate, rowPosition, Enum_표_잔고.수익율.value)
                                self.표_잔고.setItem(rowPosition, Enum_표_잔고.평가금액.value, QTableWidgetItem(format(평가금액, ',')))
                                self.표_잔고.setItem(rowPosition, Enum_표_잔고.손익금.value, QTableWidgetItem(format(손익금, ',')))

                elif sCode in self.jango_item_dict: #잔고 종목
                    purchase_price = self.emptyToZero(self.jango_item_dict[sCode]['매입가'])
                    incom_rate = self.get_incom_rate(present_price, purchase_price)  # 수익율 계산
                    realdata_dict.update({'매입가': purchase_price})
                    realdata_dict.update({'수익율': incom_rate})
                    realdata_dict.update({'매매수량': self.GetPuchaseQuantity(present_price)})
                    self.jango_item_dict[sCode].update({'수익율': incom_rate})

                    if incom_rate > self.user_dict['float_jango_reg_fluctuation']: # 잔고 관리 종목으로 등록
                        if sCode not in self.jango_contract_stock_codelist: # 잔고 중 매도 타이밍 관리 종목
                            self.messagePrint(DEBUGTYPE.거래.name, "[잔고매도관리][종목코드:%s][종목명:%s][매수가:%s][현재가:%s][등락율:%s][수익율:%s][체결강도:%s][누적거래량:%s(%s)]" % (
                                sCode, sName, purchase_price, present_price, fluctuations, incom_rate, strong,cumulative_transaction,fluctuations_transaction))
                            self.jango_contract_stock_codelist.append(sCode)

                    elif incom_rate < -self.user_dict['float_jango_buy_fluctuation']:  # 잔고 추매 등락율 보다 낮아졌다.
                        if self.deposit - self.user_dict['int_장마무리예수금'] > self.user_dict['int_purchase_amount']:
                            self.proceed_buy('추매', realdata_dict)
                        self.jango_proceed_sell('탈출', realdata_dict)

                    if sCode in self.표_잔고_관리:
                        rowPosition = self.표_잔고_관리[sCode]
                        평가금액 = present_price * self.jango_item_dict[sCode]['매매가능수량']
                        손익금 = 평가금액 - self.jango_item_dict[sCode]['매매금액']
                        self.jango_item_dict[sCode].update({'손익금':손익금})
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.수익율.value, QTableWidgetItem(str(incom_rate)))
                        self.ConvertColorValue(self.표_잔고,incom_rate, rowPosition, Enum_표_잔고.수익율.value)
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.평가금액.value, QTableWidgetItem(format(평가금액, ',')))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.손익금.value, QTableWidgetItem(format(손익금, ',')))

                else: #잔고에도 없고 매수하지 않은 종목 (매수대상)
                    self.messagePrint(DEBUGTYPE.체결강도.name,"[REALSTRONG][구매대상][종목명:%s][현재가:%s][전일대비:%s][경과시간:%s][현재강도:%s][가격등락:%s][강도등락:%s]" %(sName, present_price, previous, fluctuations_time, strong,contrast_fluctuations, fluctuations_strong))
                    realdata_dict.update({'매매수량': self.GetPuchaseQuantity(present_price)})
                    if self.deposit - self.user_dict['int_장마무리예수금'] > self.user_dict['int_purchase_amount']:
                        self.proceed_buy('신규', realdata_dict)
                    
                self.messagePrint(DEBUGTYPE.리얼데이터.name, "[REARDATA][%s]" % (self.realdata_stock_dict[sCode]))
                if self.user_dict['DEBUGTYPE_%s' % DEBUGTYPE.종목지정.name] == True:
                    _sCode = self.qLineEdit['주식기본정보'].text()
                    if sCode == _sCode:
                        self.messagePrint(DEBUGTYPE.테스트.name, "[REARDATA][%s]" % (self.realdata_stock_dict[sCode]))
                if self.readyAutoTradingStock_delay_buy and '상한가' not in self.realdata_stock_dict[sCode]:
                    if self.highprice_timer < myTime:
                        self.push_data('request_tr_주식기본정보요청_상한가,code,%s' % sCode)
                        self.highprice_timer = self.ConvertTimeChange(myTime, 3)
        elif real_type == "주식예상체결":
            pass
        elif real_type == "주식시세":
            if sCode not in self.realdata_stock_dict:
                sName = self.kiwoom.GetMasterCodeName(sCode)
                self.realdata_stock_dict.update({sCode: {}})
                self.realdata_stock_dict[sCode].update({'종목명': sName})
                self.realdata_stock_dict[sCode].update({'현재등락율': 0})
                self.realdata_stock_dict[sCode].update({'현재체결강도': 0})
                self.realdata_stock_dict[sCode].update({'현재누적거래량': 0})
            self.realdata_stock_dict[sCode].update({'고가': self.emptyToZero(self.kiwoom.GetCommRealData(sCode, 17))})
            self.realdata_stock_dict[sCode].update({'저가': self.emptyToZero(self.kiwoom.GetCommRealData(sCode, 18))})
        elif real_type == "주식당일거래원":
            pass
        elif real_type == "주식시장외호가":
            pass
        elif real_type == "주식우선호가":
            pass
        elif real_type == "주식종목정보":
            pass
        elif real_type == "ECN주식체결":
            pass
        elif real_type == "시간외종목정보":
            pass
        elif real_type == 'ECN주식시세':
            pass
        elif real_type == '주식호가잔량':
            pass
        elif real_type == '종목프로그램매매':
            pass
        else:
            self.messagePrint(DEBUGTYPE.error.name, "[ETC][%s][%s] DATA:{%s}" % (real_type, sCode, data))

    # } </editor-fold>

    # { <editor-fold desc="-체결/잔고 핸들러-"> 체결/잔고 핸들러

    def _handler_chejan_data(self, sGubun, nItemCnt, sFIdList): #체결/잔고 핸들러
        '''
        :param sGubun: 체결구분 접수와 체결시 '0'값, 국내주식 잔고전달은 '1'값, 파생잔고 전달은 '4'\n
        :param nItemCnt:\n
        :param sFIdList:\n
        :return:
        '''
        sCode = self.kiwoom.GetChejanData('9001')[1:] #종목코드
        sName = self.kiwoom.GetChejanData('302').strip() #종목명
        시간 = self.kiwoom.GetChejanData('908')
        if sGubun == '0': #체결
            mytype = self.FIDs.REALTYPE['주문체결']
            주문번호 = int(self.kiwoom.GetChejanData('9203'))  # 주문번호
            주문상태 = self.kiwoom.GetChejanData(mytype['주문상태'])
            미체결수량 = self.emptyToZero(self.kiwoom.GetChejanData(mytype['미체결수량']))
            체결량 = self.emptyToZero(self.kiwoom.GetChejanData(mytype['체결량']))
            체결누계금액 = self.emptyToZero(self.kiwoom.GetChejanData(mytype['체결누계금액']))
            주문수량 = self.emptyToZero(self.kiwoom.GetChejanData(mytype['주문수량']))
            매도수구분 = self.kiwoom.GetChejanData(mytype['매도수구분'])
            체결가 = self.emptyToZero(self.kiwoom.GetChejanData(mytype['체결가']))
            매매구분 = self.FIDs.REALTYPE['매도수구분'][매도수구분]
            주문가격 = self.emptyToZero(self.kiwoom.GetChejanData(mytype['주문가격']))

            if 미체결수량 == 0 and 주문수량 > 체결량:  # 미체결이 없고 주문량보다 체결수량이 적을경우 취소로 판단한다.
                self.messagePrint(DEBUGTYPE.체결.name, "[취소][%s][%s][%s][주문번호:%s]]" % (매매구분, sCode, sName, 주문번호))
                if sCode in self.contract_sell_item_dict['현금%s' % 매매구분]:
                    self.contract_sell_item_dict['현금%s' % 매매구분].pop(sCode)
                if sCode in self.표_미체결_관리:
                    index = self.표_미체결_관리.index(sCode)
                    self.표_미체결.removeRow(index)
                    self.표_미체결_관리.remove(sCode)
            elif 주문상태 == '접수':
                if 주문가격 != 0:
                    self.messagePrint(DEBUGTYPE.체결.name, "[접수][%s][%s][%s][주문번호:%s][주문수량:%s][주문가격:%s]" % (매매구분, sCode, sName, 주문번호, 주문수량, 주문가격))
                    if not sCode in self.contract_sell_item_dict['현금%s' % 매매구분]:
                        self.contract_sell_item_dict['현금%s' % 매매구분].update({sCode: {}})
                        self.contract_sell_item_dict['현금%s' % 매매구분][sCode].update({'사유': '수동%s' % 매매구분})
                        self.contract_sell_item_dict['현금%s' % 매매구분][sCode].update({'구분': '수동'})
                    dict_point = self.contract_sell_item_dict['현금%s' % 매매구분][sCode]
                    dict_point.update({"종목코드": sCode})
                    dict_point.update({"미체결수량": 주문수량})
                    dict_point.update({"종목명": sName})
                    dict_point.update({"시간": 시간})
                    dict_point.update({"상태": '접수'})
                    dict_point.update({'원주문번호': self.emptyToZero(self.kiwoom.GetChejanData(mytype['원주문번호']))})

                    self.표_미체결_관리.append(sCode)
                    rowPosition = self.표_미체결.rowCount()
                    self.표_미체결.insertRow(rowPosition)
                    self.표_미체결.setItem(rowPosition, Enum_표_미체결.시간.value, QTableWidgetItem(str(시간)))
                    self.표_미체결.setItem(rowPosition, Enum_표_미체결.매매구분.value, QTableWidgetItem(매매구분))
                    self.표_미체결.setItem(rowPosition, Enum_표_미체결.종목코드.value, QTableWidgetItem(sCode))
                    self.표_미체결.setItem(rowPosition, Enum_표_미체결.종목명.value, QTableWidgetItem(sName))

            elif 미체결수량 == 0: #미체결 수량이 없을 경우.
                if 주문수량 == 체결량:  # 주문수량과 체결수량이 같을 경우 > 체결 완료
                    if 매매구분 == '매도': #매도
                        매도총액 = 체결누계금액
                        구매체결강도 = 0
                        구매사유 = ''
                        구분 = ''
                        if sCode in self.contract_sell_item_dict['현금매수']:
                            dict_point = self.contract_sell_item_dict['현금매수'][sCode]
                            구분 = dict_point['구분']
                            매입금액 = dict_point['체결단가']  # 매입금액
                            매수총액 = dict_point['체결누계금액'] #매수총액
                            구매체결강도 = dict_point['체결강도']
                            구매사유 = dict_point['사유']
                            번호 = dict_point['번호']
                            self.DataFrame_Cash_buy.loc[번호, '상태'] = '완료'
                            self.save_DataFrame()
                            self.contract_sell_item_dict['현금매수'].pop(sCode)
                            if sCode in self.표_매수_관리:
                                rowPosition = self.표_매수_관리[sCode]
                                self.표_잔고.setItem(rowPosition, Enum_표_잔고.구분.value, QTableWidgetItem('완료'))
                                self.표_잔고.item(rowPosition, Enum_표_잔고.구분.value).setBackground(QtGui.QColor(130, 240, 130))
                                self.표_매수_관리.pop(sCode)
                        elif sCode in self.jango_item_dict:
                            매입금액 = self.jango_item_dict[sCode]['매입가']  # 매입금액
                            매수총액 = self.jango_item_dict[sCode]['매매금액']  # 매매금액
                            if sCode in self.jango_contract_stock_codelist:
                                self.jango_contract_stock_codelist.remove(sCode)
                            번호 = self.jango_item_dict[sCode]['번호']
                            self.DataFrame_jango.loc[번호, '상태'] = '완료'
                            self.save_DataFrame()
                            self.jango_item_dict.pop(sCode)
                            if sCode in self.표_잔고_관리:
                                rowPosition = self.표_잔고_관리[sCode]
                                self.표_잔고.setItem(rowPosition, Enum_표_잔고.구분.value, QTableWidgetItem('완료'))
                                self.표_잔고.item(rowPosition, Enum_표_잔고.구분.value).setBackground(QtGui.QColor(130, 240, 130))
                                self.표_잔고_관리.pop(sCode)
                        else:
                            self.messagePrint(DEBUGTYPE.error.name, "[ERROR][완료][매도][%s][%s][매입금액을찾을수없다][매도금액:%s][체결누계금액:%s]" % (sCode, sName, 체결가,체결누계금액))
                            return
                        if sCode in self.표_미체결_관리:
                            index = self.표_미체결_관리.index(sCode)
                            self.표_미체결.removeRow(index)
                            self.표_미체결_관리.remove(sCode)

                        수수료 = self.emptyToZero(self.kiwoom.GetChejanData('938'))  # 당일매매수수료
                        세금 = self.emptyToZero(self.kiwoom.GetChejanData('939'))  # 당일매매세금
                        self.total_commission += 수수료
                        self.total_tax += 세금
                        차익 = 매도총액 - 매수총액
                        self.total_incom += 차익
                        self.setText_qLineEdit_myinfo('매매수익', self.total_incom)
                        self.setText_qLineEdit_myinfo('수수료+세금', self.total_commission+self.total_tax)
                        self.setText_qLineEdit_myinfo('당일실현손익', self.total_incom - (self.total_commission + self.total_tax))
                        self.deposit += 매도총액 - 세금 - 수수료
                        self.setText_qLineEdit_myinfo('예수금', self.deposit)
                        수익율 = self.get_incom_rate(매도총액, 매수총액)  # 수익율 계산
                        self.messagePrint(DEBUGTYPE.거래.name, "[완료][매도][%s][%s][주문번호:%s][수량:%s][매입금액:%s][매도금액:%s][수익율:%s][매수총액:%s][매도총액:%s][차익:%s][수수료+세금:%s][손익:%s]" % (sCode, sName, 주문번호,주문수량,매입금액, 체결가,수익율,매수총액,매도총액,차익,수수료+세금,차익-(수수료+세금)))
                        if 구분 == '장전':
                            if sCode in self.interest_stock_dict:
                                번호 = self.interest_stock_dict[sCode]['번호']
                                if 수익율 > 0:
                                    self.DataFrame_interest_stock.loc[번호,'성공'] += 1
                                    print('[%s]종목이 [%s]번째 성공을 하였다.' % (sName,self.DataFrame_interest_stock.loc[번호,'성공']))
                                    self.save_except(self.DataFrame_interest_stock, 'interest_stock', '관리종목',debugPrint=False)
                        판매사유 = ''
                        판매체결강도 = 0
                        누적거래량 = 0
                        거래량변화량 = 0
                        누적거래대금 = 0
                        if sCode in self.contract_sell_item_dict['현금매도']:
                            dict_point = self.contract_sell_item_dict['현금매도'][sCode]
                            if '체결강도' in dict_point:
                                판매체결강도 = dict_point['체결강도']
                            elif '현재체결강도' in self.realdata_stock_dict[sCode]:
                                판매체결강도 = self.realdata_stock_dict[sCode]['현재체결강도']
                            if '누적거래량' in dict_point:
                                누적거래량 = dict_point['누적거래량']
                            elif '현재누적거래량' in self.realdata_stock_dict[sCode]:
                                누적거래량 = self.realdata_stock_dict[sCode]['현재누적거래량']
                            if '거래량변화량' in dict_point:
                                거래량변화량 = dict_point['거래량변화량']
                            elif '현재거래량변화량' in self.realdata_stock_dict[sCode]:
                                거래량변화량 = self.realdata_stock_dict[sCode]['현재거래량변화량']
                            if '누적거래대금' in dict_point:
                                누적거래대금 = dict_point['누적거래대금']
                            elif '현재누적거래대금' in self.realdata_stock_dict[sCode]:
                                누적거래대금 = self.realdata_stock_dict[sCode]['현재누적거래대금']
                            판매사유 = dict_point['사유']
                            self.DataFrame_Cash_sell = self.DataFrame_Cash_sell.append({'시간':시간,'상태': '완료', '종목코드': '_%s' % sCode, '종목명': sName, '체결수량': 주문수량, '체결단가': 체결가, '체결누계금액': 체결누계금액, '체결강도': 판매체결강도, '사유': 판매사유,'누적거래량':누적거래량,'거래량변화량':거래량변화량,'누적거래대금':누적거래대금}, ignore_index=True)
                            self.contract_sell_item_dict['현금매도'].pop(sCode)
                        else:
                            if sCode in self.realdata_stock_dict and '현재체결강도' in self.realdata_stock_dict[sCode]:
                                판매체결강도 = self.realdata_stock_dict[sCode]['현재체결강도']
                            if sCode in self.realdata_stock_dict and '현재누적거래량' in self.realdata_stock_dict[sCode]:
                                누적거래량 = self.realdata_stock_dict[sCode]['현재누적거래량']
                            if sCode in self.realdata_stock_dict and '현재거래량변화량' in self.realdata_stock_dict[sCode]:
                                거래량변화량 = self.realdata_stock_dict[sCode]['현재거래량변화량']
                            if sCode in self.realdata_stock_dict and '현재누적거래대금' in self.realdata_stock_dict[sCode]:
                                누적거래대금 = self.realdata_stock_dict[sCode]['현재누적거래대금']
                            self.DataFrame_Cash_sell = self.DataFrame_Cash_sell.append({'시간': 시간, '상태': '완료', '종목코드': '_%s' % sCode, '종목명': sName, '체결수량': 주문수량, '체결단가': 체결가, '체결누계금액': 체결누계금액, '체결강도': 판매체결강도, '사유': 판매사유, '누적거래량': 누적거래량, '거래량변화량': 거래량변화량,'누적거래대금':누적거래대금}, ignore_index=True)

                        self.contract_complete_selling_price.update({sCode:체결가})
                        self.DataFrame_meme_finish = self.DataFrame_meme_finish.append({'시간':시간,'종목코드': '_%s' % sCode, '종목명': sName, '매수가': 매입금액, '매도가': 체결가, '매매량':주문수량,'매매차익': 차익,'당일매매수수료':수수료,'당일매매세금':세금, '수익율': 수익율,'구매체결강도': 구매체결강도, '판매체결강도':판매체결강도,'구매사유':구매사유,'판매사유':판매사유}, ignore_index=True)
                        self.save_DataFrame()

                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.수익율.value, QTableWidgetItem(str(수익율)))
                        self.ConvertColorValue(self.표_잔고,수익율,rowPosition,Enum_표_잔고.수익율.value)
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수가.value, QTableWidgetItem(format(체결가, ',')))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.평가금액.value, QTableWidgetItem(format(체결누계금액, ',')))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.손익금.value, QTableWidgetItem(format(차익, ',')))

                    elif 매매구분 == '매수': #매수
                        self.deposit -= 체결누계금액
                        self.setText_qLineEdit_myinfo('예수금', self.deposit)
                        self.today_total_buy += 체결누계금액
                        self.setText_qLineEdit_myinfo('당일매수', self.today_total_buy)
                        if sCode not in self.contract_sell_item_dict['현금매수']:
                            self.contract_sell_item_dict['현금매수'].update({sCode: {}})
                            self.contract_sell_item_dict['현금매수'][sCode].update({"종목코드": sCode})
                            self.contract_sell_item_dict['현금매수'][sCode].update({"시간": 시간})
                            self.contract_sell_item_dict['현금매수'][sCode].update({"종목명": sName})
                            self.contract_sell_item_dict['현금매수'][sCode].update({"누적거래량": 0})
                            self.contract_sell_item_dict['현금매수'][sCode].update({"거래량변화량": 0})
                            self.contract_sell_item_dict['현금매수'][sCode].update({"누적거래대금_변화량": 0})
                            self.contract_sell_item_dict['현금매수'][sCode].update({"체결강도": 0})
                            self.contract_sell_item_dict['현금매수'][sCode].update({"사유": ''})
                        dict_point = self.contract_sell_item_dict['현금매수'][sCode]
                        dict_point.update({'체결수량': 주문수량})
                        dict_point.update({'체결단가': 체결가})
                        dict_point.update({'체결누계금액': 체결누계금액})
                        누적거래량 = 0
                        if '누적거래량' in dict_point:
                            누적거래량 = dict_point['누적거래량']
                        체결강도 = 0
                        if '체결강도' in dict_point:
                            체결강도 = dict_point['체결강도']
                        if '구분' in dict_point:
                            if dict_point['구분'] == '장전':
                                if sCode in self.interest_stock_dict:
                                    번호 = self.interest_stock_dict[sCode]['번호']
                                    self.DataFrame_interest_stock.loc[번호,'장전매수'] += 1
                                    print('[%s]종목이 [%s]번째 장전 매수를 하였다.' % (sName,self.DataFrame_interest_stock.loc[번호,'장전매수']))
                                    self.save_except(self.DataFrame_interest_stock, 'interest_stock', '관리종목', debugPrint=False)
                        elif sCode in self.jango_item_dict:
                            dict_point.update({"구분": '추매'})
                        else:
                            dict_point.update({"구분": '신규'})

                        dict_point.update({'상태': '매수'})
                        dict_point.update({'번호': len(self.DataFrame_Cash_buy)})
                        self.DataFrame_Cash_buy = self.DataFrame_Cash_buy.append({'시간':시간, '상태':'매수','종목코드': '_%s' % sCode, '종목명': sName, '체결수량': 주문수량, '체결단가': 체결가, '체결누계금액':체결누계금액,'체결강도': 체결강도, '누적거래량': 누적거래량, '거래량변화량': dict_point['거래량변화량'], '누적거래대금_변화량': dict_point['누적거래대금_변화량'], '사유': dict_point['사유'],'구분':dict_point['구분']}, ignore_index=True)
                        self.save_DataFrame()

                        if sCode in self.표_미체결_관리:
                            index = self.표_미체결_관리.index(sCode)
                            self.표_미체결.removeRow(index)
                            self.표_미체결_관리.remove(sCode)

                        rowPosition = self.표_잔고.rowCount()
                        self.표_매수_관리.update({sCode:rowPosition})
                        self.표_잔고.insertRow(rowPosition)
                        self.표_잔고_리스트.append(sCode)
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.시간.value, QTableWidgetItem(str(시간)))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.구분.value, QTableWidgetItem(dict_point['구분']))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목코드.value, QTableWidgetItem(sCode))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목명.value, QTableWidgetItem(sName))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수가.value, QTableWidgetItem('%s' % format(체결가,',')))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수량.value, QTableWidgetItem('%s' % format(주문수량, ',')))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.매매금액.value, QTableWidgetItem('%s' % format(체결누계금액, ',')))

                        if sCode in self.realdata_stock_dict:
                            self.표_잔고.setItem(rowPosition, Enum_표_잔고.현재가.value, QTableWidgetItem('%s' % format(self.realdata_stock_dict[sCode]['현재가'], ',')))
                            if '현재체결강도' in self.realdata_stock_dict[sCode]:
                                self.표_잔고.setItem(rowPosition, Enum_표_잔고.체결강도.value, QTableWidgetItem(str(self.realdata_stock_dict[sCode]['현재체결강도'])))
                            if '현재등락율' in self.realdata_stock_dict[sCode]:
                                self.표_잔고.setItem(rowPosition, Enum_표_잔고.등락율.value, QTableWidgetItem(str(self.realdata_stock_dict[sCode]['현재등락율'])))
                        self.messagePrint(DEBUGTYPE.거래.name, "[완료][매수][%s][%s][주문번호:%s][체결수량:%s][체결금액:%s][체결누계금액:%s]" % (sCode, sName, 주문번호, 주문수량,체결가,체결누계금액))
        elif sGubun == '1': #잔고
            pass

    # } </editor-fold>

# } </editor-fold>

# { <editor-fold desc="---장 시간대별 이벤트 정리---">

    def Before_of_chapter_event(self):
        self.messagePrint(DEBUGTYPE.시스템.name, "[장시작전이벤트]")
        print('현재 가진 예수금:%s' % self.deposit)
        buy_cnt = int(self.deposit / self.user_dict['int_before_store_purchase_amount'])
        print('%s원씩 %s 계좌를 매수 하겠다.' % (self.user_dict['int_before_store_purchase_amount'],buy_cnt))
        print('관심종목수:%s' % len(self.interest_stock_codelist))
        interest_stock_codelist_length = len(self.interest_stock_codelist)
        buy_cnt = min(buy_cnt,interest_stock_codelist_length)
        rand_list = self.interest_stock_codelist
        for i in range(buy_cnt):
            rand = random.randrange(1, interest_stock_codelist_length - i)
            code = rand_list[rand]
            rand_list.remove(code)
            self.push_data('request_tr_주식기본정보요청_장시작전매수,code,%s' % code)
            self.push_data('time_wasting')
            if code in self.interest_stock_dict:
                번호 = self.interest_stock_dict[code]['번호']
                self.DataFrame_interest_stock.loc[번호, '장전매수시도'] += 1
        self.save_except(self.DataFrame_interest_stock, 'interest_stock', '관리종목', debugPrint=False)

    def Start_of_chapter_event(self): #
        self.messagePrint(DEBUGTYPE.시스템.name, "[장시작후이벤트]")
        self.readyAutoTradingStock_delay_buy = True
        self.qclist['매매딜레이'].setChecked(self.readyAutoTradingStock_delay_buy)
        for sCode, dict in self.jango_item_dict.items():
            self.autotradingSetRealReg(self.screen_real_stock, sCode, "20;10;11;12;13;228", 1)

        # self.push_data('block_request_tr_미체결요청')  # 미체결 내역을 가져온다. - 미체결 취소용


    def Middle_chapter_event(self): # 장 중간 마무리
        self.messagePrint(DEBUGTYPE.시스템.name, "[장중간마무리]")
        self.middleChapterEvent = True
        self.sell_stock_jango_and_contract_item('장중간마무리판매')

    def End_of_chapter_event(self): # 장 마무리
        self.messagePrint(DEBUGTYPE.시스템.name, "[장마무리]")
        self.endOfChapterEvent = True
        self.sell_stock_jango_and_contract_item('장마무리판매')
        if self.user_dict['장마무리매수']:
            self.buy_장마무리추가매수_jango_item()

    def sell_stock_jango_and_contract_item(self,RQName):
        for sCode in self.jango_contract_stock_codelist:
            try:
                if '수익율' in self.jango_item_dict[sCode]:
                    incom_rate = self.jango_item_dict[sCode]['수익율']
                    if incom_rate > self.user_dict['float_jango_reg_fluctuation']:
                        quantity = self.jango_item_dict[sCode]['매매가능수량']
                        self.kiwoom_SendOrder_present_price_sell('잔고',RQName, sCode, quantity)
                        print("[마무리판매][%s][%s][수익율:%s]" % (sCode, self.jango_item_dict[sCode]['종목명'], incom_rate))
            except Exception as e:
                print('[ERROR][잔고][%s][%s][사유:%s]' % (sCode, self.kiwoom.GetMasterCodeName(sCode), e))
        for key, value in self.contract_sell_item_dict['현금매수'].items():
            try:
                sCode = key
                if self.contract_sell_item_dict['현금매수'][sCode]['상태'] == '매수':
                    if '수익율' in self.contract_sell_item_dict['현금매수'][sCode]:
                        incom_rate = self.contract_sell_item_dict['현금매수'][sCode]['수익율']
                        if incom_rate > self.user_dict['float_jango_reg_fluctuation']:
                            quantity = self.contract_sell_item_dict['현금매수'][sCode]['체결수량']
                            self.kiwoom_SendOrder_present_price_sell(self.contract_sell_item_dict['현금매수'][sCode]['구분'],RQName, sCode, quantity)
                            print("[마무리판매][%s][%s][수익율:%s]" % (sCode, self.contract_sell_item_dict['현금매수'][sCode]['종목명'], incom_rate))
            except Exception as e:
                print('[ERROR][현금매수][사유:%s][value:%s]' % ( e,value))

    def buy_장마무리추가매수_jango_item(self):
        print('장마무리추가매수 시작')
        추가매수금 = 0
        for sCode, value in self.jango_item_dict.items():
            try:
                print('[%s][%s][등락율:%s]' % (sCode, self.jango_item_dict[sCode]['종목명'],self.realdata_stock_dict[sCode]['현재등락율']))
                incom_rate = round(self.realdata_stock_dict[sCode]['현재등락율'])
                if incom_rate < 0:
                    매수할금액 = min(self.user_dict['float_매수단위최고치'],abs(incom_rate)) * self.user_dict['int_장마무리매수단위']
                    if self.deposit - 추가매수금 > 매수할금액:
                        추가매수금 += 매수할금액
                        quantity = max(매수할금액 / abs(self.realdata_stock_dict[sCode]['현재가']), 1)
                        self.kiwoom_SendOrder_present_price_buy('장마무리', '추가매수(%s)' % incom_rate, sCode, quantity)
                        print("[장마무리추가매수][%s][%s][추가매수할금액:%s]" % (sCode, self.jango_item_dict[sCode]['종목명'], 매수할금액))
            except Exception as e:
                print('[ERROR][잔고추가매수][%s][%s][사유:%s]' % (sCode, self.kiwoom.GetMasterCodeName(sCode), e))

    def Close_of_chapter_event(self): #장 종료
        self.messagePrint(DEBUGTYPE.시스템.name, "[장종료]")
        self.stock_state = 4

        self.messagePrint(DEBUGTYPE.시스템.name, "---[금일 매매수익:%s]---" % self.total_incom)
        self.messagePrint(DEBUGTYPE.시스템.name, "---[금일 수수료+세금:%s]---" % (self.total_commission + self.total_tax))
        self.messagePrint(DEBUGTYPE.시스템.name, "---[금일 실현손익:%s]---" % (self.total_incom - (self.total_commission + self.total_tax)))
        self.messagePrint(DEBUGTYPE.시스템.name, "---[금일 잔고:%s]---" % self.deposit)

        ret, DataFrame_dailyOverview = self.get_load_DataFrame('daily_overview','일일개요')
        if ret:
            i = len(DataFrame_dailyOverview) - 1
            today = str(DataFrame_dailyOverview.loc[i, '날짜'])
            if today == self.mytime_today:
                DataFrame_dailyOverview.loc[i, '매매수익'] = format(self.total_incom, ',')
                DataFrame_dailyOverview.loc[i, '수수료'] = format(self.total_commission,',')
                DataFrame_dailyOverview.loc[i, '세금'] = format(self.total_tax,',')
                DataFrame_dailyOverview.loc[i, '당일실현손익'] = format(self.total_incom - (self.total_commission + self.total_tax),',')
            else:
                DataFrame_dailyOverview = DataFrame_dailyOverview.append({'날짜': self.mytime_today, '매매수익': format(self.total_incom,','), '수수료': format(self.total_commission,','), '세금': format(self.total_tax,','), '당일실현손익': format(self.total_incom - (self.total_commission + self.total_tax),',')}, ignore_index=True)
            self.save_except(DataFrame_dailyOverview,'daily_overview','일일개요')
        self.messagePrint(DEBUGTYPE.시스템.name, "---[장종료]---")
        self.push_data('block_request_tr_계좌평가잔고내역요청')
        self.push_data('time_wasting')
        self.push_data('time_wasting')
        self.push_data('time_wasting')
        self.push_data('ApplicationQuit')

    def time_wasting(self):
        pass

# } </editor-fold>

# { <editor-fold desc="---매도매수---">

    def proceed_buy(self, gubun, realdata_dict):
        self.messagePrint(DEBUGTYPE.매수정보.name, realdata_dict)
        if self.user_dict['자동매수'] and self.readyAutoTradingStock_delay_buy == True:
            if realdata_dict['체결강도'] != 500:  # 체결강도가 500인경우 무시한다.
                if realdata_dict['등락율'] < self.user_dict['float_ignore_highpoint']:  # 등락률이 높지 않다면
                    if abs(realdata_dict['가격등락대비']) < self.user_dict['float_fluctuation_detection']:  # 변동옵션 이상 오르거나 떨어지지 않았다면
                        if realdata_dict['경과시간'] < self.user_dict['int_strong_delaytime']:  # 경과시간이 딜레이 시간보다 크지 않다면
                            if realdata_dict['체결강도증감'] > self.user_dict['float_buy_strong_limit']:  # 체결강도 급상 (기본구매강도 무시)
                                # if realdata_dict['현재누적거래량'] > self.user_dict['int_transaction_volume_limit']:  # 누적거래량이 최소 조건보다 높다.
                                self.kiwoom_SendOrder_present_price_buy(gubun,'급상/기본강도무시(%s)' % realdata_dict['체결강도증감'], realdata_dict['종목코드'], realdata_dict['매매수량'])
                            elif realdata_dict['체결강도'] > self.user_dict['float_default_strong_limit']:  # 기본 구매 강도 초과시
                                if realdata_dict['체결강도증감'] > self.user_dict['float_condition_fluctuations_strong_highpoint']:  # 체결강도변화량이 구매상승조건 보다 높다
                                    self.kiwoom_SendOrder_present_price_buy(gubun, '강도상승매수(%s)' % realdata_dict['체결강도증감'], realdata_dict['종목코드'], realdata_dict['매매수량'])
                                # elif realdata_dict['거래량변화량'] > self.user_dict['int_transaction_volume_detection']: #거래량 변화량이 클시
                                #     self.kiwoom_SendOrder_present_price_buy('%s/거래량상승매수(%s)' % (RQName, realdata_dict['거래량변화량']), realdata_dict['종목코드'], realdata_dict['매매수량'])

    def proceed_sell(self, gubun, realdata_dict):
        self.messagePrint(DEBUGTYPE.매도정보.name, realdata_dict)
        if self.user_dict['자동매도'] and self.readyAutoTradingStock_delay_buy == True:
            if gubun == '장전' and realdata_dict['수익율'] > 3.5:
                self.kiwoom_SendOrder_present_price_sell(gubun, '장전목표수익율(%s)' % realdata_dict['수익율'], realdata_dict['종목코드'], realdata_dict['매매수량'])
            if realdata_dict['수익율'] > self.user_dict['float_최대수익구간']:  # 최대 수익율 달성시 매도
                self.kiwoom_SendOrder_present_price_sell(gubun, '최대수익율(%s)' % realdata_dict['수익율'], realdata_dict['종목코드'], realdata_dict['매매수량'])
            elif realdata_dict['체결강도'] > self.user_dict['float_sell_ignore_strong_limit']:  # 판매제한 체결강도가 높아 오를것으로 판단한다.(매도 안함)
                pass
            elif self.user_dict['손절매도'] or (not self.user_dict['손절매도'] and realdata_dict['수익율'] > 0):
                if self.user_dict['손절매도'] and not self.endOfChapterEvent:  # 장마무리 전
                    if realdata_dict['체결강도'] < self.user_dict['float_strong_sell']:  # 최소 체결강도보다 낮아졌다.
                        self.kiwoom_SendOrder_present_price_sell(gubun, '체결강도미만(%s)' % realdata_dict['체결강도'], realdata_dict['종목코드'], realdata_dict['매매수량'])
                    elif realdata_dict['수익율'] < -self.user_dict['float_condition_lowpoint_today']:  # 금일 수익율이 최소 수익율보다 하락했다.
                        self.kiwoom_SendOrder_present_price_sell(gubun, '수익율기준하락1(%s)' % realdata_dict['수익율'], realdata_dict['종목코드'], realdata_dict['매매수량'])
                    elif realdata_dict['체결강도'] > self.user_dict['float_sell_strong_limit']:  # 판매제한 체결강도보다 높아 오를것으로 판단한다.(매도 안함)
                        pass
                    elif realdata_dict['수익율'] > self.user_dict['float_jango_reg_fluctuation'] or self.deposit < self.user_dict['int_예수금유지금액']:  # 수익이 나고 있다. or 잔고가 10000000보다 적다.
                        if realdata_dict['가격등락대비'] < -self.user_dict['float_condition_fluctuations_price_lowpoint']:  # 가격이 급격히 하락한다.
                            self.kiwoom_SendOrder_present_price_sell(gubun, '가격급락(%s)' % realdata_dict['가격등락대비'], realdata_dict['종목코드'], realdata_dict['매매수량'])
                        elif realdata_dict['체결강도증감'] < -self.user_dict['float_condition_fluctuations_strong_lowpoint']:  # 체결강도변화량이 급격히 하락한다
                            self.kiwoom_SendOrder_present_price_sell(gubun, '체결강도급락(%s)' % realdata_dict['체결강도증감'], realdata_dict['종목코드'], realdata_dict['매매수량'])
                    else: #손해가 나고 있다.
                        if realdata_dict['수익율'] < -self.user_dict['float_sell_strong_limit_and_price_lowpoint']:  # 수익율이 판매 제한 강도보다 낮으면서 일정 수익율보다 하락했다.
                            self.kiwoom_SendOrder_present_price_sell(gubun, '수익율기준하락2(%s)' % realdata_dict['수익율'], realdata_dict['종목코드'], realdata_dict['매매수량'])
                        elif realdata_dict['가격등락대비'] < -self.user_dict['float_jango_condition_fluctuations_price_lowpoint']:  # 가격이 급격히 하락한다.
                            self.kiwoom_SendOrder_present_price_sell(gubun, '가격급락(%s)' % realdata_dict['가격등락대비'], realdata_dict['종목코드'], realdata_dict['매매수량'])
                        elif realdata_dict['체결강도증감'] < -self.user_dict['float_jango_condition_fluctuations_strong_lowpoint']:  # 체결강도변화량이 급격히 하락한다
                            self.kiwoom_SendOrder_present_price_sell(gubun, '체결강도급락(%s)' % realdata_dict['체결강도증감'], realdata_dict['종목코드'], realdata_dict['매매수량'])

    def jango_proceed_sell(self, gubun, realdata_dict):
        self.messagePrint(DEBUGTYPE.매도정보.name, realdata_dict)
        if self.readyAutoTradingStock_delay_buy and not self.endOfChapterEvent:
            if realdata_dict['등락율'] > self.user_dict['float_최대수익구간']:  # 잔고 최대 등락율 달성시 매도
                self.kiwoom_SendOrder_present_price_sell(gubun, '잔고최대등락율(%s)' % realdata_dict['수익율'], realdata_dict['종목코드'], realdata_dict['매매수량'])
            elif self.user_dict['손절매도'] and realdata_dict['가격등락대비'] < -self.user_dict['float_jango_condition_fluctuations_price_lowpoint']:  # 가격이 급격히 하락한다.
                self.kiwoom_SendOrder_present_price_sell(gubun, '가격급락(%s)' % realdata_dict['가격등락대비'], realdata_dict['종목코드'], realdata_dict['매매수량'])

    def message_meme_info(self, gubun, RQName, sCode):
        try:
            sName = self.realdata_stock_dict[sCode]['종목명']
            present_price = self.realdata_stock_dict[sCode]['현재가']
            fluctuations = self.realdata_stock_dict[sCode]['현재등락율']
            strong = self.realdata_stock_dict[sCode]['현재체결강도']
            fluctuations_time = self.realdata_stock_dict[sCode]['경과시간']
            fluctuations_strong = self.realdata_stock_dict[sCode]['체결강도증감']
            contrast_fluctuations = self.realdata_stock_dict[sCode]['가격등락대비']
            fluctuations_transaction = self.realdata_stock_dict[sCode]['거래량변화량']
            cumulative_transaction = self.realdata_stock_dict[sCode]['현재누적거래량']
            purchase_price = 0
            if sCode in self.jango_item_dict:
                purchase_price = self.jango_item_dict[sCode]['매입가']
            elif sCode in self.contract_sell_item_dict['현금매수'] and self.contract_sell_item_dict['현금매수'][sCode]['상태'] == '매수':
                purchase_price = self.contract_sell_item_dict['현금매수'][sCode]['체결단가']
            incom_rate = self.get_incom_rate(present_price, purchase_price)
            self.realdata_stock_dict[sCode].update({'수익률': incom_rate})
            self.messagePrint(DEBUGTYPE.거래.name, "[%s][%s][경과시간:%s][종목코드:%s][종목명:%s][매수가:%s][현재가:%s][등락율:%s][수익율:%s][체결강도:%s(%s)][가격등락대비:%s][누적거래량:%s(%s)]" % (
                gubun, RQName, fluctuations_time, sCode, sName, purchase_price, present_price, fluctuations, incom_rate, strong, fluctuations_strong, contrast_fluctuations, cumulative_transaction, fluctuations_transaction))
        except Exception as e:
            self.messagePrint(DEBUGTYPE.error.name, "[ERROR][message_meme_info][%s]" % e)
            self.messagePrint(DEBUGTYPE.error.name, "[ERROR][message_meme_info][%s]" % self.contract_sell_item_dict['현금매수'][sCode])
            self.messagePrint(DEBUGTYPE.error.name, "[ERROR][message_meme_info][%s]" % self.realdata_stock_dict[sCode])

    def kiwoom_SendOrder_present_price_buy(self, gubun, RQName, sCode, quantity):  # 현재가 매수
        '''
        :param RQName: 식별이름\n
        :param sCode: 종목코드\n
        :param quantity: 주문수량 \n
        :return:현재가 매수
        '''
        sName = self.realdata_stock_dict[sCode]['종목명']
        purchase_price = self.realdata_stock_dict[sCode]['현재가']
        if sCode in self.contract_complete_selling_price:
            Selling_price = self.contract_complete_selling_price[sCode]
            limit_price = int(Selling_price * 0.98)
            if limit_price < purchase_price:
                return
            else:
                print('[재매수][사유:%s][%s][%s][현재가:%s][기존매도가:%s][재매수기준가:%s]' % (RQName, sCode, sName, purchase_price, Selling_price, limit_price))
        if sCode not in self.contract_sell_item_dict['현금매수']:
            self.contract_sell_item_dict['현금매수'].update({sCode: {}})
            dict_point = self.contract_sell_item_dict['현금매수'][sCode]
            dict_point.update({'종목코드': sCode})
            dict_point.update({'종목명': sName})
            dict_point.update({'체결강도': self.realdata_stock_dict[sCode]['현재체결강도']})
            dict_point.update({'누적거래량': self.realdata_stock_dict[sCode]['현재누적거래량']})
            dict_point.update({'거래량변화량': self.realdata_stock_dict[sCode]['거래량변화량']})
            dict_point.update({'누적거래대금_변화량': self.realdata_stock_dict[sCode]['누적거래대금_변화량']})
            dict_point.update({'상태': '접수대기'})
            dict_point.update({'구분': gubun})
            dict_point.update({'사유': RQName})
            self.message_meme_info(gubun, RQName, sCode)
            self.kiwoom_SendOrder(RQName, screen=self.screen_trading_stock, order_type=1, sCode=sCode, quantity=quantity)

    def kiwoom_SendOrder_present_price_sell(self, gubun, RQName, sCode, quantity):  # 현재가 매도
        '''
        :param RQName: 식별이름\n
        :param sCode: 종목코드\n
        :param quantity: 주문수량 \n
        :return:현재가 매도
        '''
        sName = self.realdata_stock_dict[sCode]['종목명']
        if sCode not in self.contract_sell_item_dict['현금매도']:
            self.contract_sell_item_dict['현금매도'].update({sCode: {}})
            dict_point = self.contract_sell_item_dict['현금매도'][sCode]
            dict_point.update({'종목코드': sCode})
            dict_point.update({'종목명': sName})
            dict_point.update({'체결강도': self.realdata_stock_dict[sCode]['현재체결강도']})
            dict_point.update({'누적거래량': self.realdata_stock_dict[sCode]['현재누적거래량']})
            dict_point.update({'거래량변화량': self.realdata_stock_dict[sCode]['거래량변화량']})
            dict_point.update({'누적거래대금_변화량': self.realdata_stock_dict[sCode]['누적거래대금_변화량']})
            dict_point.update({'상태': '접수대기'})
            dict_point.update({'구분': gubun})
            dict_point.update({'사유': RQName})
            self.message_meme_info(gubun, RQName, sCode)
            if '하한가' in self.realdata_stock_dict[sCode]:
                price = int(self.realdata_stock_dict[sCode]['하한가'])
                print('[KIWOOM_SELL][%s][%s][하한가판매][%s]' % (sCode, self.realdata_stock_dict[sCode]['종목명'], price))
            else:
                price = self.realdata_stock_dict[sCode]['현재가'] * 0.98
                print('[KIWOOM_SELL][%s][%s][현재가판매]' % (sCode, self.realdata_stock_dict[sCode]['종목명']))
            self.kiwoom_SendOrder(RQName, screen=self.screen_trading_stock, order_type=2, sCode=sCode, quantity=quantity, price=price, hoga='00')

    def kiwoom_SendOrder_cancel_buy(self, RQName, sCode, quantity, order_no):  # 주문취소
        '''
        :param RQName: 식별이름\n
        :param sCode: 종목코드\n
        :param quantity: 주문수량 \n
        :param order_no: 원주문번호 \n
        :return:매수취소
        '''
        self.kiwoom_SendOrder(RQName, screen=self.screen_trading_stock, order_type=3, sCode=sCode, quantity=quantity, order_no=order_no)

    def kiwoom_SendOrder_correction_sell_lowprice(self, RQName, sCode, quantity, order_no):  # 정정하한가판매
        '''
        :param RQName: 식별이름\n
        :param sCode: 종목코드\n
        :param quantity: 주문수량 \n
        :param order_no: 원주문번호 \n
        :return:하한가판매
        '''
        self.kiwoom_SendOrder(RQName=RQName, screen=self.screen_trading_stock, order_type=6, sCode=sCode, quantity=quantity, price=self.realdata_stock_dict[sCode]['현재가'] * 0.8, hoga='00', order_no=order_no)

    def kiwoom_SendOrder(self, RQName, screen, order_type, sCode, quantity, price=0, hoga='03', order_no=''):  # 기본 거래
        '''
        :param RQName: 식별이름\n
        :param screen: 스크린번호\n
        :param order_type: 1: 신규매수, 2: 신규매도, 3: 매수취소, 4: 매도취소, 5: 매수정정, 6: 매도정정 \n
        :param sCode: 종목코드 \n
        :param quantity: 주문수량 \n
        :param price: 주문단가 \n
        :param hoga: 00: 지정가, 03: 시장가,  05: 조건부지정가, 06: 최유리지정가, 07: 최우선지정가, 10: 지정가IOC, 13: 시장가IOC, 16: 최유리IOC, 20: 지정가FOK, 23: 시장가FOK, 26: 최유리FOK, 61: 장전시간외종가, 62: 시간외단일가, 81: 장후시간외종가\n
        :param order_no: 원주문번호\n
        :return:
        '''
        sName = self.kiwoom.GetMasterCodeName(sCode)
        price = self.get_hoga_cal(price)
        if price == 0 and hoga == '00':
            # sName = self.kiwoom.GetMasterCodeName(sCode)
            self.messagePrint(DEBUGTYPE.error.name, "[ERROR][지정가0원][%s][%s][거래가:%s][수량:%s개]" % (RQName, sName, "현재가", quantity))
            return
        else:
            if order_type == 1:
                if hoga == '03':  # 시장가 신규구매
                    price = self.realdata_stock_dict[sCode]['현재가']
                total = price * quantity
                if price > self.user_dict['int_purchase_amount']:
                    # sName = self.kiwoom.GetMasterCodeName(sCode)
                    self.messagePrint(DEBUGTYPE.error.name, "[ERROR][옵션최대금액초과][%s][%s][거래가:%s][수량:%s개][옵션최대금액:%s]" % (RQName, sName, price, quantity, self.user_dict['int_purchase_amount']))
                    return
                else:
                    if total > self.deposit:
                        # sName = self.kiwoom.GetMasterCodeName(sCode)
                        self.messagePrint(DEBUGTYPE.error.name, "[ERROR][주문가능금액초과][%s][%s][거래가:%s][수량:%s개][주문액:%s][주문가능금액:%s]" % (RQName, sName, price, quantity, total, self.deposit))
        RQName = '%s§%s§%s§%s§%s§%s' % (RQName, sCode, price, quantity, order_type, hoga)
        if hoga == '03':
            price = 0
        ret = self.kiwoom.SendOrder(RQName, screen, self.myAccount, order_type, sCode, quantity, price, hoga, order_no)
        if ret != None:
            self.messagePrint(DEBUGTYPE.error.name,'[ERROR][SendOrder][return:%s][%s][%s][가격:%s][수량:%s]' % (ret,sCode,sName,price,quantity))

# } </editor-fold>

# { <editor-fold desc="---tr 요청---">

    def block_request_tr_계좌평가현황요청(self): #계좌평가현황요청
        df = self.kiwoom.block_request("OPW00004",
                                        계좌번호=self.myAccount,
                                        비밀번호=self.user_dict['account_pass'],
                                        상장폐지조회구분=1,
                                        비밀번호입력매체구분=00,
                                        output="계좌평가현황",
                                        next=0)
        convert_df = DataFrame(df)
        if 'D+2추정예수금' in convert_df.loc[0]:
            self.deposit = int(df['D+2추정예수금'][0])
            self.setText_qLineEdit_myinfo('예수금', self.deposit)
            self.messagePrint(DEBUGTYPE.내정보.name, "[주문가능금액 : %s]" % format(self.deposit, ','))
            self.push_data('start')

    def block_request_tr_계좌별주문체결내역상세요청(self,gubun): #계좌별주문체결내역상세요청
        '''
        :param gubun: 0:전체, 1:매도, 2:매수
        :return: 계좌별주문체결내역상세요청
        '''
        self.workerPause()
        if gubun == 0:
            구분 = '전체'
        elif gubun == 1:
            구분 = '매도'
        elif gubun == 2:
            구분 = '매수'
        df = self.kiwoom.block_request("OPW00007",
                                       주문일자=self.mytime_today,
                                       계좌번호=self.myAccount,
                                       비밀번호=self.user_dict['account_pass'],
                                       비밀번호입력매체구분=00,
                                       조회구분=4,
                                       주식채권구분=0,
                                       매도수구분=gubun,
                                       종목코드='',
                                       시작주문번호='',
                                       output="계좌별주문체결내역상세",
                                       next=0)

        convert_df = DataFrame(df)
        length = len(convert_df)
        for i in range(length):
            try:
                dict = {}
                dict.update({'종목코드': convert_df.loc[i]['종목번호'][1:]})
                dict.update({'종목명': convert_df.loc[i]['종목명']})
                dict.update({'주문구분': convert_df.loc[i]['주문구분'][:4]})
                dict.update({'체결단가': self.emptyToZero(convert_df.loc[i]['체결단가'])})
                dict.update({'체결수량': self.emptyToZero(convert_df.loc[i]['체결수량'])})
                print('[계좌별주문체결내역상세요청][%s]%s' % (i, dict))
            except Exception as e:
                print('[ERROR][계좌별주문체결내역상세요청][e:%s]' % e)
                print(convert_df.loc[i])
        self.workerStart()
        self.messagePrint(DEBUGTYPE.시스템.name, "[계좌별주문체결내역상세요청][%s][완료]" % 구분)


    def block_request_tr_미체결요청(self, sPrevNext="0", index="0", sell_index="0", conclusion_index="1"): #미체결 요청
        '''
        :param sPrevNext:\n
        :param index: # 전체종목구분 = 0:전체, 1: 종목\n
        :param sell_index: # 매매구분 = 0:전체, 1: 매도, 2: 매수\n
        :param conclusion_index: # 체결구분 = 0:전체, 2: 체결, 1: 미체결\n
        :return:
        '''
        self.workerPause()
        df = self.kiwoom.block_request("opt10075",
                                       계좌번호=self.myAccount,
                                       전체종목구분=index,
                                       매매구분=sell_index,
                                       체결구분=conclusion_index,
                                       output="계좌평가현황",
                                       next=sPrevNext)
        convert_df = DataFrame(df)
        length = len(convert_df)
        for i in range(length):
            try:
                종목코드 = convert_df.loc[i]['종목코드']
                if 종목코드 != '':
                    dict = {}
                    dict.update({'종목코드': 종목코드})
                    dict.update({'종목명': convert_df.loc[i]['종목명']})
                    매매구분 = convert_df.loc[i]['주문구분'][1:]
                    if 매매구분 != '':
                        dict.update({'매매구분': 매매구분})
                        dict.update({'미체결수량': self.emptyToZero(convert_df.loc[i]['미체결수량'])})
                        dict.update({'주문번호': self.emptyToZero(convert_df.loc[i]['주문번호'])})
                        원주문번호 = int(convert_df.loc[i]['원주문번호'])
                        if 원주문번호 == 0:
                            원주문번호 = dict['주문번호']
                        dict.update({'원주문번호': 원주문번호})
                        print('[미체결][%s]%s' % (i, dict))

                        self.kiwoom_SendOrder_cancel_buy('매수취소', dict['종목코드'], dict['미체결수량'], 원주문번호)

                        # if self.readyAutoTradingStock_delay_buy == True:
                        #     self.kiwoom_SendOrder_cancel_buy('매수취소', dict['종목코드'], dict['미체결수량'], 원주문번호)
                        # else:
                        #     self.contract_sell_item_dict['현금%s' % 매매구분].update({종목코드: {}})
                        #     dict_point = self.contract_sell_item_dict['현금%s' % 매매구분][종목코드]
                        #     dict_point.update({"종목코드": 종목코드})
                        #     dict_point.update({"종목명": dict['종목명']})
                        #     dict_point.update({'원주문번호': 원주문번호})
                        #
                        #     self.표_미체결_관리.append(종목코드)
                        #     rowPosition = self.표_미체결.rowCount()
                        #     self.표_미체결.insertRow(rowPosition)
                        #     self.표_미체결.setItem(rowPosition, Enum_표_미체결.매매구분.value, QTableWidgetItem(매매구분))
                        #     self.표_미체결.setItem(rowPosition, Enum_표_미체결.종목코드.value, QTableWidgetItem(종목코드))
                        #     self.표_미체결.setItem(rowPosition, Enum_표_미체결.종목명.value, QTableWidgetItem(dict['종목명']))
            except Exception as e:
                print('[ERROR][미체결요청][e:%s]' % e)
                print(convert_df.loc[i])
        self.workerStart()
        self.messagePrint(DEBUGTYPE.시스템.name, "[미체결요청][완료]")

    def request_tr_주식기본정보요청_상한가(self,code):
        self.request_tr_주식기본정보요청('상한가',code)

    def request_tr_주식기본정보요청_장시작전매수(self,code):
        self.request_tr_주식기본정보요청('장시작전매수',code)

    def request_tr_주식기본정보요청(self, sRQName, code): #주식기본정보요청
        '''
        :param RQName: 식별용 이름\n
        :param code: 전문 조회할 종목코드\n
        :param screen: 스크린번호
        :return:주식기본정보요청
        '''
        # self.workerPause()
        df = self.kiwoom.block_request("opt10001",
                                       종목코드=code,
                                       output="주식기본정보",
                                       next=0)
        convert_df = DataFrame(df)
        try:
            종목코드 = code
            종목명 = convert_df.loc[0]['종목명']
            돌려받은종목코드 = convert_df.loc[0]['종목코드']
            if 종목코드 != 돌려받은종목코드:
                print('요청한 종목코드[%s]와 데이터의 종목코드[%s]가 다르다.' % (종목코드,돌려받은종목코드))
                return
            if 종목명 != '':
                if 종목코드 not in self.realdata_stock_dict:
                    self.realdata_stock_dict.update({종목코드: {}})
                    self.realdata_stock_dict[종목코드].update({"종목명": 종목명})
                    self.realdata_stock_dict[종목코드].update({"종목코드": 종목코드})
                    self.realdata_stock_dict[종목코드].update({'현재등락율': 0})
                    self.realdata_stock_dict[종목코드].update({'현재체결강도': 0})
                    self.realdata_stock_dict[종목코드].update({'현재누적거래량': 0})
                상한가데이터 = False
                if '상한가' in convert_df.loc[0]:
                    상한가데이터 = True
                    상한가 = self.emptyToZero(convert_df.loc[0]['상한가'],doabs=True)
                    하한가 = self.emptyToZero(convert_df.loc[0]['하한가'],doabs=True)
                    기준가 = self.emptyToZero(convert_df.loc[0]['기준가'],doabs=True)
                    시가 = self.emptyToZero(convert_df.loc[0]['시가'],doabs=True)
                    고가 = self.emptyToZero(convert_df.loc[0]['고가'],doabs=True)
                    저가 = self.emptyToZero(convert_df.loc[0]['저가'],doabs=True)
                    현재가 = self.emptyToZero(convert_df.loc[0]['현재가'], doabs=True)
                    self.realdata_stock_dict[종목코드].update({'상한가': 상한가})
                    self.realdata_stock_dict[종목코드].update({'하한가': 하한가})
                    self.realdata_stock_dict[종목코드].update({'기준가': 기준가})
                    self.realdata_stock_dict[종목코드].update({'시가': 시가})
                    self.realdata_stock_dict[종목코드].update({'고가': 고가})
                    self.realdata_stock_dict[종목코드].update({'저가': 저가})
                    self.realdata_stock_dict[종목코드].update({'현재가': 현재가})

                if 상한가 > 0 and 하한가 > 0:
                    pass
                else:
                    print('[%s][%s] 상한가 혹은 하한가가 0이다.' % (종목코드, 종목명))
                    return

                관리종목확인 = self.kiwoom.GetMasterConstruction(종목코드)
                if 관리종목확인 == '정상':
                    pass
                else:
                    # print('%s 종목이 관리종목인것을 확인하였다.(%s)' % (종목명,관리종목확인))
                    return
                if sRQName == '결과통계':
                    self.messagePrint(DEBUGTYPE.잔고.name, "[결과통계] %s" % self.realdata_stock_dict[종목코드])
                elif sRQName == '상한가':
                    if 상한가데이터:
                        self.DataFrame_stock_info = self.DataFrame_stock_info.append({'종목코드': '_%s' % 종목코드, '종목명': 종목명, '상한가': 상한가, '하한가': 하한가}, ignore_index=True)
                        self.save_DataFrame()
                elif sRQName == '장시작전매수':
                    if not 종목코드 in self.contract_sell_item_dict['현금매수']:
                        self.contract_sell_item_dict['현금매수'].update({종목코드: {}})
                        self.contract_sell_item_dict['현금매수'][종목코드].update({'사유': '장시작전매수'})
                        self.contract_sell_item_dict['현금매수'][종목코드].update({'구분': '장전'})
                    주문단가 = int(기준가 * 0.95)
                    주문수량 = max(int(self.user_dict['int_before_store_purchase_amount'] / 주문단가), 1)
                    print('[장시작전매수][%s][%s][기준가:%s][주문단가:%s][주문수량:%s]' % (종목코드, 종목명, 기준가, 주문단가,주문수량))
                    self.kiwoom_SendOrder('장시작전매수', self.screen_trading_stock, 1, 종목코드, 주문수량, 주문단가, '00', '')
                self.autotradingSetRealReg(self.screen_real_stock, 종목코드, "20;10;11;12;13;228", 1)
            else:
                print('[주식기본정보][ERROR][%s][%s][데이터가 없다]' % (sRQName,code))
            # self.push_data('request_tr_주식기본정보요청_장시작전매수,code,%s' % code)
        except Exception as e:
            print('[ERROR][주식기본정보][%s][%s][%s]' % (e,sRQName,code))
            print('[ERROR][realdata][%s]' % self.realdata_stock_dict[종목코드])
            print('[ERROR][convert_df]\n%s' % convert_df)
        # self.workerStart()

    def block_request_tr_계좌평가잔고내역요청(self):
        '''
        :param trCode: #[opw00018: 계좌평가잔고내역요청]\n
        :param sPrevNext:\n
        :param screen_no:\n
        :return:
        '''
        self.workerPause()
        if self.stock_state != 4:
            첫잔고불러오기 = False
            length = len(self.DataFrame_jango)
            if length == 0:
                첫잔고불러오기 = True
            df = self.kiwoom.block_request("opw00018",
                                           계좌번호=self.myAccount,
                                           비밀번호=self.user_dict['account_pass'],
                                           비밀번호입력매체구분=self.user_dict['pass_index'],
                                           조회구분=2,
                                           output="계좌평가잔고개별합산",
                                           next=0)
            self.account_balance(df,첫잔고불러오기)
            while self.kiwoom.tr_remained:
                df = self.kiwoom.block_request("opw00018",
                                           계좌번호=self.myAccount,
                                           비밀번호=self.user_dict['account_pass'],
                                           비밀번호입력매체구분=self.user_dict['pass_index'],
                                           조회구분=2,
                                           output="계좌평가잔고개별합산",
                                           next=2)
                self.account_balance(df,첫잔고불러오기)

            if self.user_dict['DEBUGTYPE_%s' % DEBUGTYPE.내종목정보.name] == True:
                for i, v in enumerate(self.jango_item_dict):
                    self.messagePrint(DEBUGTYPE.내종목정보.name,'[%s]%s' % (i,self.jango_item_dict[v]))

        계좌평가결과 = self.kiwoom.block_request("opw00018",
                                       계좌번호=self.myAccount,
                                       비밀번호=self.user_dict['account_pass'],
                                       비밀번호입력매체구분=self.user_dict['pass_index'],
                                       조회구분=2,
                                       output="계좌평가결과",
                                       next=0)
        계좌평가결과변환 = DataFrame(계좌평가결과)

        try:
            총매입금액 = int(계좌평가결과변환.loc[0]['총매입금액'])
        except:
            self.messagePrint(DEBUGTYPE.error.name, "[ERROR][계좌평가잔고내역][총매입금액]")
            self.push_data('block_request_tr_계좌평가잔고내역요청')
            return
        self.messagePrint(DEBUGTYPE.내정보.name, ("[총매입금액 : %s]" % format(총매입금액, ',')))

        if 총매입금액 > 0:
            총수익률 = self.emptyToZero(계좌평가결과변환.loc[0]['총수익률(%)']) / 100
            총평가손익금액 = self.emptyToZero(계좌평가결과변환.loc[0]['총평가손익금액'])
            총평가금액 = 총매입금액 + 총평가손익금액 + self.deposit
            self.messagePrint(DEBUGTYPE.내정보.name, "[총평가손익금액 : %s]" % format(총평가손익금액, ','))
            self.messagePrint(DEBUGTYPE.내정보.name, "[총수익률 : %s%s]" % (총수익률, "%"))
            self.messagePrint(DEBUGTYPE.내정보.name, "[소지계좌갯수 : %s]" % len(self.jango_item_dict))
            self.messagePrint(DEBUGTYPE.내정보.name, "[총평가금액 : %s]" % format(총평가금액, ','))

            ret, DataFrame_dailyOverview = self.get_load_DataFrame('daily_overview', '일일개요')
            if ret:
                if len(DataFrame_dailyOverview) > 0:
                    i = len(DataFrame_dailyOverview) - 1
                    today = str(DataFrame_dailyOverview.loc[i, '날짜'])
                    if today != self.mytime_today:
                        DataFrame_dailyOverview = DataFrame_dailyOverview.append({'날짜': self.mytime_today, '총매입금액': 총매입금액, '총평가손익금액': 총평가손익금액, '총수익률': 총수익률, '소지계좌갯수': len(self.jango_item_dict), '총평가금액': 총평가금액}, ignore_index=True)
                        self.save_except(DataFrame_dailyOverview, 'daily_overview', '일일개요')
                else:
                    DataFrame_dailyOverview = DataFrame_dailyOverview.append({'날짜': self.mytime_today, '총매입금액': 총매입금액, '총평가손익금액': 총평가손익금액, '총수익률': 총수익률, '소지계좌갯수': len(self.jango_item_dict), '총평가금액': 총평가금액}, ignore_index=True)
                    self.save_except(DataFrame_dailyOverview, 'daily_overview', '일일개요')
            else:
                DataFrame_dailyOverview = DataFrame(columns=['날짜', '총매입금액', '총평가손익금액', '총수익률', '소지계좌갯수', '총평가금액', '매매수익', '수수료', '세금', '당일실현손익', '마감총평가금액', '마감예수금','전일대비'])
                DataFrame_dailyOverview = DataFrame_dailyOverview.append({'날짜': self.mytime_today, '총매입금액': 총매입금액, '총평가손익금액': 총평가손익금액, '총수익률': 총수익률, '소지계좌갯수': len(self.jango_item_dict), '총평가금액': 총평가금액}, ignore_index=True)
                self.save_except(DataFrame_dailyOverview, 'daily_overview', '일일개요')

            if self.stock_state == 4:
                ret, DataFrame_dailyOverview = self.get_load_DataFrame('daily_overview','일일개요')
                if ret:
                    i = len(DataFrame_dailyOverview) - 1
                    today = str(DataFrame_dailyOverview.loc[i, '날짜'])
                    if today == self.mytime_today:
                        DataFrame_dailyOverview.loc[i, '마감총평가금액'] = format(총평가금액, ',')
                        DataFrame_dailyOverview.loc[i, '마감예수금'] = format(self.deposit, ',')
                        전일대비 = 총평가금액 - DataFrame_dailyOverview.loc[i, '총평가금액']
                        DataFrame_dailyOverview.loc[i, '전일대비'] = format(전일대비, ',')
                        self.messagePrint(DEBUGTYPE.내정보.name, "[전일대비 : %s]" % format(전일대비, ','))
                    else:
                        DataFrame_dailyOverview = DataFrame_dailyOverview.append({'날짜': self.mytime_today, '마감총평가금액': 총평가금액, '마감예수금': self.deposit}, ignore_index=True)
                    self.save_except(DataFrame_dailyOverview,'daily_overview','일일개요')
        self.messagePrint(DEBUGTYPE.시스템.name, "[계좌평가잔고내역요청][완료]")
        self.workerStart()
        if self.readyAutoTradingSystem == False:
            self.readyAutoTradingSystem = True
            self.cheak_time()


    def account_balance(self,df,첫잔고불러오기):
        convert_df = DataFrame(df)
        length = len(convert_df)
        for i in range(length):
            # try:
            종목코드 = str(convert_df.loc[i]['종목번호'])[1:]
            if 종목코드 != '':
                종목명 = convert_df.loc[i]['종목명']
                관리종목확인 = self.kiwoom.GetMasterConstruction(종목코드)

                매매가능수량 = self.emptyToZero(convert_df.loc[i]['매매가능수량'])
                매입가 = self.emptyToZero(convert_df.loc[i]['매입가'])
                매매금액 = 매입가 * 매매가능수량

                if 종목코드 not in self.realdata_stock_dict:
                    self.realdata_stock_dict.update({종목코드: {}})
                    self.realdata_stock_dict[종목코드].update({"종목코드": 종목코드})
                    self.realdata_stock_dict[종목코드].update({"종목명": 종목명})
                    self.realdata_stock_dict[종목코드].update({'현재등락율': 0})
                    self.realdata_stock_dict[종목코드].update({'현재체결강도': 0})
                    self.realdata_stock_dict[종목코드].update({'현재누적거래량': 0})

                if 첫잔고불러오기:
                    self.jango_item_dict.update({종목코드: {}})
                    self.jango_item_dict[종목코드].update({"종목코드": 종목코드})
                    self.jango_item_dict[종목코드].update({"종목명": 종목명})
                    self.jango_item_dict[종목코드].update({"매입가": 매입가})
                    self.jango_item_dict[종목코드].update({"매매가능수량": 매매가능수량})
                    self.jango_item_dict[종목코드].update({"매매금액": 매매금액})
                    self.jango_item_dict[종목코드].update({"번호": len(self.DataFrame_jango)})

                    rowPosition = self.표_잔고.rowCount()
                    self.표_잔고_관리.update({종목코드:rowPosition})
                    self.표_잔고.insertRow(rowPosition)
                    self.표_잔고_리스트.append(종목코드)
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.시간.value, QTableWidgetItem('-'))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목코드.value, QTableWidgetItem(종목코드))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목명.value, QTableWidgetItem(종목명))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수가.value, QTableWidgetItem('%s' % format(매입가, ',')))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수량.value, QTableWidgetItem('%s' % format(매매가능수량, ',')))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매매금액.value, QTableWidgetItem('%s' % format(매입가 * 매매가능수량, ',')))

                    self.DataFrame_jango = self.DataFrame_jango.append({'상태':'잔고','종목코드': '_%s' % 종목코드, '종목명': 종목명, '매입가': 매입가, '매매가능수량': 매매가능수량, '매매금액': 매매금액}, ignore_index=True)
                    self.save_DataFrame()
                else:
                    if 종목코드 not in self.jango_item_dict and 종목코드 not in self.contract_sell_item_dict['현금매수']:
                        self.contract_sell_item_dict['현금매수'].update({종목코드: {}})
                        dict_point = self.contract_sell_item_dict['현금매수'][종목코드]
                        dict_point.update({'종목코드': 종목코드})
                        dict_point.update({'종목명': 종목명})
                        dict_point.update({'누적거래량': 0})
                        dict_point.update({'거래량변화량': 0})
                        dict_point.update({'누적거래대금_변화량': 0})
                        dict_point.update({'체결강도': 0})
                        dict_point.update({'사유': '누락'})
                        dict_point.update({'체결수량': 매매가능수량})
                        dict_point.update({'체결단가': 매입가})
                        dict_point.update({'체결누계금액': 매매금액})
                        dict_point.update({'구분': '누락'})
                        dict_point.update({'상태': '매수'})
                        dict_point.update({'번호': len(self.DataFrame_Cash_buy)})
                        self.DataFrame_Cash_buy = self.DataFrame_Cash_buy.append({'시간':'-','종목코드': '_%s' % 종목코드, '종목명': 종목명, '체결수량': 매매가능수량, '체결단가': 매입가, '체결누계금액':매매금액, '체결강도': dict_point['체결강도'], '누적거래량': dict_point['누적거래량'], '거래량변화량': dict_point['거래량변화량'],'누적거래대금_변화량': dict_point['누적거래대금_변화량'], '사유': dict_point['사유'], '구분': dict_point['구분']}, ignore_index=True)
                        self.save_DataFrame()

                        rowPosition = self.표_잔고.rowCount()
                        self.표_매수_관리.update({종목코드:rowPosition})
                        self.표_잔고.insertRow(rowPosition)
                        self.표_잔고_리스트.append(종목코드)
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.시간.value, QTableWidgetItem('-'))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.구분.value, QTableWidgetItem('누락'))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목코드.value, QTableWidgetItem(종목코드))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목명.value, QTableWidgetItem(종목명))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수가.value, QTableWidgetItem('%s' % format(매입가, ',')))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수량.value, QTableWidgetItem('%s' % format(매매가능수량, ',')))
                        self.표_잔고.setItem(rowPosition, Enum_표_잔고.매매금액.value, QTableWidgetItem('%s' % format(매매금액, ',')))
                        self.autotradingSetRealReg(self.screen_real_stock, 종목코드, "20;10;11;12;13;228", 1)
                if 관리종목확인 != '정상':
                    print('[잔고][관리종목확인][%s][%s][%s]' % (종목코드, 종목명, 관리종목확인))
                    self.kiwoom_SendOrder_present_price_sell('잔고', '관리종목(%s)' % 관리종목확인, 종목코드, 매매가능수량)
            else:
                print('[account_balance][ERROR][%s][데이터가 없다]' % i)
            # except Exception as e:
            #     print('[account_balance][ERROR]%s' % e)
            #     print('[account_balance][ERROR]%s' % convert_df.loc[i])

    def request_tr_당일거래량상위(self, gunun1="000", gunun2="1", gunun3="4",gunun4="0",gunun5="0", gunun6="0", gunun7="0", gunun8="0"):
        '''
        :param gunun1:시장구분 = 000:전체, 001:코스피, 101:코스닥
        :param gunun2:정렬구분 = 1:거래량, 2:거래회전율, 3:거래대금
        :param gunun3:관리종목포함 = 0:관리종목 포함, 1:관리종목 미포함, 3:우선주제외, 11:정리매매종목제외, 4:관리종목, 우선주제외, 5:증100제외, 6:증100마나보기, 13:증60만보기, 12:증50만보기, 7:증40만보기, 8:증30만보기, 9:증20만보기, 14:ETF제외, 15:스팩제외, 16:ETF+ETN제외
        :param gunun4:신용구분 = 0:전체조회, 9:신용융자전체, 1:신용융자A군, 2:신용융자B군, 3:신용융자C군, 4:신용융자D군, 8:신용대주
        :param gunun5:거래량구분 = 0:전체조회, 5:5천주이상, 10:1만주이상, 50:5만주이상, 100:10만주이상, 200:20만주이상, 300:30만주이상, 500:500만주이상, 1000:백만주이상
        :param gunun6:가격구분 = 0:전체조회, 1:1천원미만, 2:1천원이상, 3:1천원~2천원, 4:2천원~5천원, 5:5천원이상, 6:5천원~1만원, 10:1만원미만, 7:1만원이상, 8:5만원이상, 9:10만원이상
        :param gunun7:거래대금구분 = 0:전체조회, 1:1천만원이상, 3:3천만원이상, 4:5천만원이상, 10:1억원이상, 30:3억원이상, 50:5억원이상, 100:10억원이상, 300:30억원이상, 500:50억원이상, 1000:100억원이상, 3000:300억원이상, 5000:500억원이상
        :param gunun8:장운영구분 = 0:전체조회, 1:장중, 2:장전시간외, 3:장후시간외
        :return:
        '''
        self.workerPause()
        df = self.kiwoom.block_request("opt10030",
                                       시장구분=gunun1,
                                       정렬구분=gunun2,
                                       관리종목포함=gunun3,
                                       신용구분=gunun4,
                                       거래량구분=gunun5,
                                       가격구분=gunun6,
                                       거래대금구분=gunun7,
                                       장운영구분=gunun8,
                                       output="당일거래량상위",
                                       next=0)
        length = len(df)
        for i in range(length):
            종목코드 = df['종목코드'][i]
            관리종목확인 = self.kiwoom.GetMasterConstruction(종목코드)

            if 종목코드 not in self.interest_stock_dict:
                self.interest_stock_dict.update({종목코드: {}})
                self.interest_stock_dict[종목코드].update({'번호': len(self.DataFrame_interest_stock)})
                self.DataFrame_interest_stock = self.DataFrame_interest_stock.append({'종목코드': '_%s' % 종목코드, '종목명': self.kiwoom.GetMasterCodeName(종목코드), '성공': 0, '장전매수시도': 0, '장전매수': 0, '관리종목': 관리종목확인, '제외종목': '정상'}, ignore_index=True)
                if 관리종목확인 == '정상':
                    self.autotradingSetRealReg(self.screen_real_stock, 종목코드, "20;10;11;12;13;228", 1)
            else:
                번호 = self.interest_stock_dict[종목코드]['번호']
                self.DataFrame_interest_stock.loc[번호, '관리종목'] = 관리종목확인
                제외종목 = self.DataFrame_interest_stock.loc[번호, '제외종목']
                if 관리종목확인 == '정상':
                    if 제외종목 == '정상':
                        self.autotradingSetRealReg(self.screen_real_stock, 종목코드, "20;10;11;12;13;228", 1)
                else:
                    self.autotradingSetRealRemove(종목코드)
        self.save_except(self.DataFrame_interest_stock, 'interest_stock', '관리종목', debugPrint=False)
        self.messagePrint(DEBUGTYPE.시스템.name, "[%s][%s]" % (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), '당일거래량상위요청'))
        if not self.readyAutoTradingStock:
            self.readyAutoTradingStock = True
            self.qclist['자동거래종목'].setChecked(self.readyAutoTradingStock)
        self.workerStart()

# } </editor-fold>

# { <editor-fold desc="---tools---">

    def sum_손익계산(self):
        총손익금 = 0
        for 종목 in self.jango_item_dict:
            if '손익금' in self.jango_item_dict[종목]:
                총손익금 += self.jango_item_dict[종목]['손익금']

    def ConvertColorValue(self,표,value,rowPosition,colum):
        try:
            if value > 0:
                표.item(rowPosition, colum).setBackground(QtGui.QColor(250, 160, 160))
            elif value < 0:
                표.item(rowPosition, colum).setBackground(QtGui.QColor(Qt.cyan))
            else:
                표.item(rowPosition, colum).setBackground(QtGui.QColor(Qt.white))
        except Exception as e:
            print('[ERROR][ConvertColorValue][e:%s]' % e)
            print('[ERROR][ConvertColorValue][value:%s][rowPosition:%s][colum:%s]' % (value,rowPosition,colum))

    def ConvertTimeChange(self,originTime,changeTime):
        strOriginTime = format(originTime,'06')
        hh = self.emptyToZero(strOriginTime[0:2])
        mm = self.emptyToZero(strOriginTime[2:4])
        ss = self.emptyToZero(strOriginTime[4:6])
        minus = changeTime < 0
        strChangeTime = format(abs(changeTime),'06')
        _ss = self.emptyToZero(strChangeTime[-2:])
        _mm = self.emptyToZero(strChangeTime[-4:-2])
        _hh = self.emptyToZero(strChangeTime[-6:-4])
        if minus:
            _ss = -_ss
            _mm = -_mm
            _hh = -_hh
        new_ss = ss + _ss
        if new_ss > 60:
            mm += 1
            new_ss = new_ss - 60
        elif new_ss < 0:
            mm -= 1
            new_ss = new_ss + 60
        new_mm = mm + _mm
        if new_mm > 60:
            hh += 1
            new_mm = new_mm - 60
        elif new_mm < 0:
            hh -= 1
            new_mm = new_mm + 60
        new_hh = hh + _hh
        if new_hh > 24:
            new_hh = hh % 24
        elif new_ss < 0:
            new_hh = (24 + hh - 1) % 24
        return int(format(new_hh,'02') + format(new_mm,'02') + format(new_ss,'02'))

    def ConvertText(self,text,texttype,option = 2): #표시 변환
        strText = str(text)
        if texttype == 'int':
            try:
                result = '{0:,}'.format(int(float(strText)))
            except:
                result = 0
        elif texttype == 'float':
            try:
                result = float(strText)
            except:
                result = 0
        elif texttype == 'str':
            result = text.strip()
        else:
            result = text
        return result

    def GetCommData_dict_update_type(self, sTrCode, sRQName, index, dict, itemlist, type):
        '''
        :param sTrCode:
        :param sRQName:
        :param index:
        :param dict:
        :param itemlist:
        :param type:
        :return: int,float,str
        '''
        try:
            for item in itemlist:
                data = self.kiwoom.GetCommData(sTrCode, sRQName, index, item)
                if type == 'int':
                    dict.update({item: self.emptyToZero(data)})
                elif type == 'float':
                    dict.update({item: self.emptyToZero(data,1)})
                elif type == 'str':
                    dict.update({item: data.strip()})
                else:
                    dict.update({item: data})
        except:
            self.messagePrint(DEBUGTYPE.error.name,"[ERROR][sTrCode:%s][sRQName:%s][itemlist:%s][data:%s]" % (sTrCode,sRQName,itemlist,data))

    def GetCommData_dict_updata(self, sTrCode, sRQName, index, dict, itemlist):
        for item in itemlist:
            data = self.kiwoom.GetCommData(sTrCode, sRQName, index, item)
            dict.update({item: data})

    def GetChejanData_dict_updata_type(self, dict, fidtype, fids, type):
        '''
        :param dict:
        :param fidtype:
        :param fids:
        :param type:
        :return: int,float,str
        '''
        for fid in fids:
            data = self.kiwoom.GetChejanData(fidtype[fid])
            if type == "int":
                dict.update({fid: self.emptyToZero(data)})
            elif type == "float":
                dict.update({fid: self.emptyToZero(data,1)})
            elif type == "str":
                dict.update({fid: data.strip()})
            else:
                dict.update({fid: data})

    def GetChejanData_dict_updata(self, dict, fidtype, fids):
        for fid in fids:
            data = self.kiwoom.GetChejanData(fidtype[fid])
            dict.update({fid: data})


    def week_check(self):
        if datetime.datetime.now().weekday() >= 5:
            return 1
        else:
            return 0

    def cheak_time(self):
        now = datetime.datetime.now()
        self.messagePrint(DEBUGTYPE.시스템.name, "[타임체크]")
        hour = now.hour
        minute = now.minute
        time = (hour * 10000) + (minute * 100)
        if time > self.user_dict['time_Intermediate_Finish_Time']:
            self.middleChapterEvent = True
        if time > 150000:
            self.endOfChapterEvent = True
        if time < 90000:
            self.stock_state = 0
            if self.user_dict['장전매수']:
                self.Before_of_chapter_event()
        elif time < 150000:
            self.stock_state = 3
        else:
            self.messagePrint(DEBUGTYPE.시스템.name, "---장시간이 아닙니다.---")
            self.readyStockMarket = False
            return

    def GetPuchaseQuantity(self, unit_price):
        return max(int(self.user_dict['int_purchase_amount'] / abs(unit_price)),1)

    def emptyToZero(self, value, type = 0, doabs = False):
        '''
        :param value:
        :param type: 0:int 1:float
        :return:
        '''
        if value == '':
            value = 0
        else:
            if type == 0:
                value = int(value)
            elif type == 1:
                value = float(value)
        if doabs:
            value = abs(value)
        return value

    def messagePrint(self, key, msg):
        '''
        :param debug_lv: 1:error 2:주요메세지 3:디버그메세지 4:정보메세지
        :param msg: 메세지
        :return:
        '''
        if self.user_dict['DEBUGTYPE_%s' % DEBUGTYPE.전체.name] == True:
            if self.user_dict['DEBUGTYPE_%s' % key] == True:
                now = datetime.datetime.now()
                print("[DEBUG][%s]%s" % (now.strftime('%Y-%m-%d %H:%M:%S'),msg))

    def get_dict_value(self,dict,key):
        if key not in dict:
            dict.updata({key: 0})
        return dict[key]

    def get_incom_rate(self,price,purchase_price):
        if purchase_price == 0:
            return 0
        incom_rate = round(((price - purchase_price) / purchase_price) * 100, 2)  # 수익율 계산
        return incom_rate

    def get_hoga_unit(self, price):
        result = 1000
        price = int(price)
        if price < 1000:
            result = 1
        elif price < 5000:
            result = 5
        elif price < 10000:
            result = 10
        elif price < 50000:
            result = 50
        elif price < 100000:
            result = 100
        elif price < 500000:
            result = 500
        return result

    def get_hoga_cal(self, price):
        price = int(price)
        return int(price - (price % self.get_hoga_unit(price)))

# } </editor-fold>

# { <editor-fold desc="---데이터 관리---">

    def save_option(self):
        file = open('user2.sav', 'wb')
        pickle.dump(self.user_dict, file)
        file.close()

    def save_as_option(self):
        root = Tk()
        root.withdraw()
        root.filename = filedialog.asksaveasfilename(initialdir="./", title="Save Data files", filetypes=(("data files", "*.sav"), ("all files", "*.*")))
        if root.filename != '':
            file = open(root.filename, 'wb')
            pickle.dump(self.user_dict, file)
            file.close()

    def load_option(self):
        if os.path.exists('user2.sav'):
            file = open('user2.sav', 'rb')
            self.user_dict = pickle.load(file)
            file.close()

    def load_option2(self):
        root = Tk()
        root.withdraw()
        root.filename = filedialog.askopenfilename(initialdir="./", title="Open Data files", filetypes=(("data files", "*.sav"), ("all files", "*.*")))
        if root.filename != '':
            file = open(root.filename, 'rb')
            self.user_dict = pickle.load(file)
            file.close()
            self.setUI()

    def get_load_DataFrame(self,name,sheet_name):
        filename = 'report\\%s.xlsx' % name
        if os.path.exists(filename):
            return True, pd.read_excel(filename, sheet_name=sheet_name, index_col=0)
        else:
            return False, None

    def save_except(self,DataFrame,name,sheet_name,debugPrint = True):
        try:
            filename = 'report\\%s.xlsx' % name
            with pd.ExcelWriter(filename) as writer:
                DataFrame.to_excel(writer, sheet_name=sheet_name)
            if debugPrint:
                self.messagePrint(DEBUGTYPE.시스템.name,'[저장완료][%s.xlsx]' % name)
        except:
            filename = 'report\\%s_bak.xlsx' % name
            with pd.ExcelWriter(filename) as writer:
                DataFrame.to_excel(writer, sheet_name=sheet_name)
            self.messagePrint(DEBUGTYPE.시스템.name,'[임시저장][%s_bak.xlsx]' % name)

    def save_DataFrame(self):
        filename = 'report\\report_%s.xlsx' % self.mytime_today
        try:
            with pd.ExcelWriter(filename) as writer:
                self.DataFrame_jango.to_excel(writer, sheet_name='잔고')
                self.DataFrame_Cash_buy.to_excel(writer, sheet_name='현금매수')
                self.DataFrame_Cash_sell.to_excel(writer, sheet_name='현금매도')
                self.DataFrame_meme_finish.to_excel(writer, sheet_name='매도완료')
                self.DataFrame_stock_info.to_excel(writer, sheet_name='주식기본정보')
        except:
            filename = 'report\\report_%s_bak.xlsx' % self.mytime_today
            with pd.ExcelWriter(filename) as writer:
                self.DataFrame_jango.to_excel(writer, sheet_name='잔고')
                self.DataFrame_Cash_buy.to_excel(writer, sheet_name='현금매수')
                self.DataFrame_Cash_sell.to_excel(writer, sheet_name='현금매도')
                self.DataFrame_meme_finish.to_excel(writer, sheet_name='매도완료')
                self.DataFrame_stock_info.to_excel(writer, sheet_name='주식기본정보')

    def load_DataFrame(self):
        self.messagePrint(DEBUGTYPE.시스템.name,'[데이터불러오기]')
        filename = 'report\\interest_stock.xlsx'
        if os.path.exists(filename):
            self.DataFrame_interest_stock = pd.read_excel(filename, sheet_name='관리종목', index_col=0)
            length = len(self.DataFrame_interest_stock)
            for i in range(length):
                종목코드 = self.DataFrame_interest_stock.loc[i, '종목코드'][1:]
                construction = self.kiwoom.GetMasterConstruction(종목코드)
                self.DataFrame_interest_stock.loc[i, '관리종목'] = construction
                제외종목 = self.DataFrame_interest_stock.loc[i, '제외종목']
                if construction == '정상' and 제외종목 == '정상':
                    self.interest_stock_codelist.append(종목코드)
                    self.autotradingSetRealReg(self.screen_real_stock, 종목코드, "20;10;11;12;13;228", 1)
                self.interest_stock_dict.update({종목코드:{}})
                self.interest_stock_dict[종목코드].update({'번호': i})
        else:
            self.DataFrame_interest_stock = DataFrame(columns=['종목코드','종목명','장전매수시도','장전매수','성공','관리종목','제외종목'])
        self.save_except(self.DataFrame_interest_stock, 'interest_stock', '관리종목')

        filename = 'report\\report_%s.xlsx' % self.mytime_today
        if os.path.exists(filename):
            self.DataFrame_Cash_sell = pd.read_excel(filename, sheet_name='현금매도', index_col=0)

            self.DataFrame_stock_info = pd.read_excel(filename, sheet_name='주식기본정보', index_col=0)
            length = len(self.DataFrame_stock_info)
            for i in range(length):
                sCode = self.DataFrame_stock_info.loc[i, '종목코드'][1:]
                if sCode not in self.realdata_stock_dict:
                    self.realdata_stock_dict.update({sCode: {}})
                self.realdata_stock_dict[sCode].update({'종목코드': sCode})
                self.realdata_stock_dict[sCode].update({'종목명': self.DataFrame_stock_info.loc[i, '종목명']})
                self.realdata_stock_dict[sCode].update({'상한가': self.emptyToZero(self.DataFrame_stock_info.loc[i, '상한가'],doabs=True)})
                self.realdata_stock_dict[sCode].update({'하한가': self.emptyToZero(self.DataFrame_stock_info.loc[i, '하한가'],doabs=True)})

            self.DataFrame_jango = pd.read_excel(filename, sheet_name='잔고', index_col=0)
            length = len(self.DataFrame_jango)
            for i in range(length):
                상태 = self.DataFrame_jango.loc[i, '상태']
                if 상태 != '완료':
                    종목코드 = self.DataFrame_jango.loc[i, '종목코드'][1:]
                    종목명 = self.DataFrame_jango.loc[i, '종목명']
                    매입가 = self.emptyToZero(self.DataFrame_jango.loc[i, '매입가'])
                    매매가능수량 = self.emptyToZero(self.DataFrame_jango.loc[i, '매매가능수량'])
                    매매금액 = self.emptyToZero(self.DataFrame_jango.loc[i, '매매금액'])
                    self.jango_item_dict.update({종목코드: {}})
                    self.jango_item_dict[종목코드].update({"번호": i})
                    self.jango_item_dict[종목코드].update({"종목코드": 종목코드})
                    self.jango_item_dict[종목코드].update({"종목명": 종목명})
                    self.jango_item_dict[종목코드].update({"매입가": 매입가})
                    self.jango_item_dict[종목코드].update({"매매가능수량": 매매가능수량})
                    self.jango_item_dict[종목코드].update({"매매금액": 매매금액})

                    rowPosition = self.표_잔고.rowCount()
                    self.표_잔고_관리.update({종목코드:rowPosition})
                    self.표_잔고.insertRow(rowPosition)
                    self.표_잔고_리스트.append(종목코드)
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.시간.value, QTableWidgetItem('-'))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목코드.value, QTableWidgetItem(종목코드))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목명.value, QTableWidgetItem(종목명))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.구분.value, QTableWidgetItem('잔고'))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수가.value, QTableWidgetItem('%s' % format(매입가, ',')))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수량.value, QTableWidgetItem('%s' % format(매매가능수량, ',')))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매매금액.value, QTableWidgetItem('%s' % format(매매금액, ',')))

            self.DataFrame_Cash_buy = pd.read_excel(filename, sheet_name='현금매수',index_col=0)
            length = len(self.DataFrame_Cash_buy)
            for i in range(length):
                상태 = self.DataFrame_Cash_buy.loc[i, '상태']
                if 상태 != '완료':
                    sCode = self.DataFrame_Cash_buy.loc[i, '종목코드'][1:]
                    sName = self.DataFrame_Cash_buy.loc[i, '종목명']
                    시간 = self.DataFrame_Cash_buy.loc[i, '시간']
                    체결수량 = self.emptyToZero(self.DataFrame_Cash_buy.loc[i, '체결수량'])
                    체결단가 = self.emptyToZero(self.DataFrame_Cash_buy.loc[i, '체결단가'])
                    누적거래량 = self.emptyToZero(self.DataFrame_Cash_buy.loc[i, '누적거래량'])
                    거래량변화량 = self.emptyToZero(self.DataFrame_Cash_buy.loc[i, '거래량변화량'])
                    체결누계금액 = self.emptyToZero(self.DataFrame_Cash_buy.loc[i, '체결누계금액'])
                    누적거래대금_변화량 = self.emptyToZero(self.DataFrame_Cash_buy.loc[i, '누적거래대금_변화량'])
                    구분 = self.DataFrame_Cash_buy.loc[i, '구분']
                    사유 = self.DataFrame_Cash_buy.loc[i, '사유']
                    체결강도 = self.emptyToZero(self.DataFrame_Cash_buy.loc[i, '체결강도'])

                    self.contract_sell_item_dict['현금매수'].update({sCode: {}})
                    dict_point = self.contract_sell_item_dict['현금매수'][sCode]
                    dict_point.update({"종목코드": sCode})
                    dict_point.update({"상태": 상태})
                    dict_point.update({"시간": 시간})
                    dict_point.update({"종목명": sName})
                    dict_point.update({"누적거래량": 누적거래량})
                    dict_point.update({"거래량변화량": 거래량변화량})
                    dict_point.update({"누적거래대금_변화량": 누적거래대금_변화량})
                    dict_point.update({"체결강도": 체결강도})
                    dict_point.update({"사유": 사유})
                    dict_point.update({'체결수량': 체결수량})
                    dict_point.update({'체결단가': 체결단가})
                    dict_point.update({'체결누계금액': 체결누계금액})
                    dict_point.update({"구분": 구분})
                    dict_point.update({"번호": i})

                    rowPosition = self.표_잔고.rowCount()
                    self.표_매수_관리.update({sCode:rowPosition})
                    self.표_잔고.insertRow(rowPosition)
                    self.표_잔고_리스트.append(sCode)
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.시간.value, QTableWidgetItem(str(시간)))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.구분.value, QTableWidgetItem(구분))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목코드.value, QTableWidgetItem(sCode))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목명.value, QTableWidgetItem(sName))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수가.value, QTableWidgetItem('%s' % format(체결단가, ',')))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수량.value, QTableWidgetItem('%s' % format(체결수량, ',')))
                    self.표_잔고.setItem(rowPosition, Enum_표_잔고.매매금액.value, QTableWidgetItem('%s' % format(체결누계금액, ',')))
                    self.autotradingSetRealReg(self.screen_real_stock, sCode, "20;10;11;12;13;228", 1)

            self.DataFrame_meme_finish = pd.read_excel(filename, sheet_name='매도완료', index_col=0)
            length = len(self.DataFrame_meme_finish)
            for i in range(length):
                sCode = self.DataFrame_meme_finish.loc[i, '종목코드'][1:]
                sName =  self.DataFrame_meme_finish.loc[i, '종목명']
                시간 =  self.DataFrame_meme_finish.loc[i, '시간']
                매수가 = self.emptyToZero(self.DataFrame_meme_finish.loc[i, '매수가'])
                매도가 = self.emptyToZero(self.DataFrame_meme_finish.loc[i, '매도가'])
                매매량 = self.emptyToZero(self.DataFrame_meme_finish.loc[i, '매매량'])
                수익율 = self.DataFrame_meme_finish.loc[i, '수익율']
                매매차익 = self.emptyToZero(self.DataFrame_meme_finish.loc[i, '매매차익'])
                self.total_incom += 매매차익
                self.total_tax += self.emptyToZero(self.DataFrame_meme_finish.loc[i, '당일매매세금'])
                self.total_commission += self.emptyToZero(self.DataFrame_meme_finish.loc[i, '당일매매수수료'])

                rowPosition = self.표_잔고.rowCount()
                self.표_잔고.insertRow(rowPosition)
                self.표_잔고_리스트.append(sCode)
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.시간.value, QTableWidgetItem(str(시간)))
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.구분.value, QTableWidgetItem('완료'))
                self.표_잔고.item(rowPosition, Enum_표_잔고.구분.value).setBackground(QtGui.QColor(130, 240, 130))
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목코드.value, QTableWidgetItem(sCode))
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.종목명.value, QTableWidgetItem(sName))
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.수익율.value, QTableWidgetItem(str(수익율)))
                self.ConvertColorValue(self.표_잔고,수익율, rowPosition, Enum_표_잔고.수익율.value)
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수가.value, QTableWidgetItem('%s' % format(매도가, ',')))
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.매수량.value, QTableWidgetItem('%s' % format(매매량, ',')))
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.매매금액.value, QTableWidgetItem('%s' % format(매수가*매매량, ',')))
                self.표_잔고.setItem(rowPosition, Enum_표_잔고.손익금.value, QTableWidgetItem('%s' % format(매매차익, ',')))
        else:
            self.DataFrame_jango = DataFrame(columns=['상태', '종목코드', '종목명', '매입가', '매매가능수량','매매금액'])
            self.DataFrame_Cash_buy = DataFrame(columns=['시간','상태','종목코드', '종목명','체결수량','체결단가','체결누계금액','체결강도','누적거래량','거래량변화량','누적거래대금_변화량','사유','구분'])
            self.DataFrame_Cash_sell = DataFrame(columns=['시간','상태','종목코드', '종목명', '체결수량', '체결단가','체결누계금액', '체결강도','누적거래량','거래량변화량','누적거래대금_변화량','사유'])
            self.DataFrame_stock_info = DataFrame(columns=['종목코드', '종목명', '상한가', '하한가'])
            self.DataFrame_meme_finish = DataFrame(columns=['시간','종목코드', '종목명', '매수가', '매도가','매매량','매매차익','당일매매수수료','당일매매세금','수익율','구매체결강도','판매체결강도','구매사유','판매사유'])
            self.save_DataFrame()
        self.setText_qLineEdit_myinfo('수수료+세금', self.total_commission + self.total_tax)
        self.setText_qLineEdit_myinfo('매매수익', self.total_incom)
        self.setText_qLineEdit_myinfo('당일실현손익', self.total_incom - (self.total_commission + self.total_tax))

# } </editor-fold>

    def closeEvent(self, event):
        self.kiwoomSetRealRemoveAll()
        # self.workerStop()

    def ApplicationQuit(self):
        QCoreApplication.instance().quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    data_queue = Queue()
    order_queue = Queue()
    window = MyWindow(data_queue, order_queue)
    window.show()
    app.exec_()
