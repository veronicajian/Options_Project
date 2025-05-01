import sys
import os
import csv
import shioaji as sj
from datetime import datetime
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QRadioButton,
    QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QComboBox,
    QFrame, QGroupBox, QTableWidget, QDialog, QTableWidgetItem, 
    QHeaderView, QMessageBox
    )
from shioaji import BidAskFOPv1, Exchange
from PyQt6.QtCore import QObject, pyqtSignal, QFileSystemWatcher, QTimer
from PyQt6.QtGui import QColor
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt
import logging
import numpy as np


class OptionDataManager(QObject):
    bidask_received = pyqtSignal(object)    # 五檔報價
    pnl_bidask_received = pyqtSignal(object)    # 計算損益
    monitor_bidask_received = pyqtSignal(object)  # 價格監控


    def __init__(self):
        super().__init__()

        self.api = sj.Shioaji()
        # 建立 logger 物件
        self.logger = logging.getLogger(__name__)
        self.api.quote.on_event(self.event_callback)

        self.api.login(
            api_key = "", 
            secret_key ="", 
        )

        Result = self.api.activate_ca(  
            ca_path=r"C:\Users\veron\Desktop\永豐api\Sinopac.pfx",
            ca_passwd="",
            person_id="",
        )

        # 註冊回調函式
        self.api.quote.set_on_bidask_fop_v1_callback(self.quote_callback)


    def event_callback(self, resp_code: int, event_code: int, info: str, event: str):
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.logger.info(f"[{now}] Event code: {event_code} | Event: {event}")
        if event_code == 1 or event_code == 2 or event_code == 12 or event_code == 13:  # SOLCLIENT_SESSION_EVENT_DOWN_ERROR
            sys.exit()


    def quote_callback(self, exchange:Exchange, bidask:BidAskFOPv1):
        self.bidask_received.emit(bidask)


class ProfitCheckThread(QThread):
    """獨立線程用於持續檢查達標條件"""
    stage_reached = pyqtSignal(int)  # 信號：傳遞達標的階段 
    error_occurred = pyqtSignal(str)  # 信號：傳遞錯誤訊息

    def __init__(self, parent=None):
        super().__init__(parent)
        self.running = False
        self.stage_to_check = None
        self.option_manager = parent.option_manager  # parent 是主視窗

    def set_stage(self, stage):
        """設定要檢查的階段 (1 或 2)"""
        self.stage_to_check = stage

    def run(self):
        """線程主邏輯：持續檢查達標條件"""
        self.running = True
        while self.running:
            try:
                if self.stage_to_check == 1:
                    # 檢查 buy_2 和 sell_2 是否達標
                    if self.check_stage_1():
                        self.stage_reached.emit(1)
                elif self.stage_to_check == 2:
                    # 檢查 buy_1 是否達標
                    if self.check_stage_2():
                        self.stage_reached.emit(2)
                
                self.msleep(5000)  # 每 5 秒檢查一次
            except Exception as e:
                self.error_occurred.emit(str(e))


    def check_stage_1(self):

        try:
            # 檢查必要的輸入欄位是否為空
            required_fields = [
                self.parent().sell_1_input.text(),
                self.parent().buy_1_input.text(),
                self.parent().outter_price_diff_input.text(),
            ]
            
            # 檢查是否有空值
            if any(field.strip() == "" for field in required_fields):
                raise ValueError("請填寫所有必要的輸入欄位！")
            
            delivery_month = '202505'
            code = 'TX1'
            sell_1_premium = float(self.parent().sell_1_input.text())
            buy_1_premium = float(self.parent().buy_1_input.text())
            price_diff = float(self.parent().outter_price_diff_input.text())

            # 1. 獲取 buy_2 最新報價
            buy_2_strike = self.parent().buy_2_strike_input.text().strip()
            buy_2_type = 'P' if 'Put' in self.parent().buy_2_type_input.currentText() else 'C'
            buy_2_contract = self.option_manager.api.Contracts.Options[code].get(
                f"{code}{delivery_month}{int(buy_2_strike)}{buy_2_type}"
            )
            buy_2_snapshot = self.option_manager.api.snapshots([buy_2_contract])[0]
            buy_2_premium = buy_2_snapshot.sell_price

            # 2. 獲取 sell_2 最新報價
            sell_2_strike = self.parent().sell_2_strike_input.text().strip()
            sell_2_type = 'P' if 'Put' in self.parent().sell_2_type_input.currentText() else 'C'
            sell_2_contract = self.option_manager.api.Contracts.Options[code].get(
                f"{code}{delivery_month}{int(sell_2_strike)}{sell_2_type}"
            )
            sell_2_snapshot = self.option_manager.api.snapshots([sell_2_contract])[0]
            sell_2_premium = sell_2_snapshot.buy_price

            return (sell_1_premium + sell_2_premium - buy_1_premium - buy_2_premium - price_diff) > 0

        except Exception as e:
            self.error_occurred.emit(str(e))
            return False
        


    def check_stage_2(self):
        try:
            # 檢查必要的輸入欄位是否為空
            required_fields = [
                self.parent().sell_1_input.text(),
                self.parent().buy_1_input.text(),
                self.parent().buy_2_input.text(),
                self.parent().outter_price_diff_input.text(),
                self.parent().sell_2_strike_input.text(),
            ]
        
            # 檢查是否有空值
            if any(field.strip() == "" for field in required_fields):
                raise ValueError("請填寫所有必要的輸入欄位！")
            
            delivery_month = '202505'
            code = 'TX1'
            sell_1_premium = float(self.parent().sell_1_input.text())
            buy_1_premium = float(self.parent().buy_1_input.text())
            buy_2_premium = float(self.parent().buy_2_input.text())
            price_diff = float(self.parent().outter_price_diff_input.text())

            # 取得sell_2的報價
            sell_2_strike = self.parent().sell_2_strike_input.text().strip()
            sell_2_type = 'P' if 'Put' in self.parent().sell_2_type_input.currentText() else 'C'
            sell_2_contract = self.option_manager.api.Contracts.Options[code].get(
                f"{code}{delivery_month}{int(sell_2_strike)}{sell_2_type}"
            )
            sell_2_snapshot = self.option_manager.api.snapshots([sell_2_contract])[0]
            sell_2_premium = sell_2_snapshot.buy_price

            print(sell_2_premium)
            return (sell_1_premium + sell_2_premium - buy_1_premium - buy_2_premium - price_diff) > 0
        
        except Exception as e:
            self.error_occurred.emit(str(e))
            return False


    def stop(self):
        """安全停止線程"""
        self.running = False
        self.wait()


class OptionTQuoteApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("無風險計算機")
        self.setGeometry(100, 50, 650, 180)  
        self.option_manager = OptionDataManager()
        self.current_selected_group = None
        self.price_range = np.array([]) # 繪圖的座標範圍
        self.option_manager.bidask_received.connect(self.update_bidask_table)

        self.init_ui()
        '''將視窗置於螢幕中央'''
        # 取得視窗的建議尺寸
        self.resize(self.sizeHint())
        
        # 取得主螢幕
        screen = self.screen().geometry()
        window = self.geometry()
        
        self.move(
            (screen.width() - window.width()) // 2,
            (screen.height() - window.height()) // 2
        )

        # 創建利潤檢查線程
        self.profit_check_thread = ProfitCheckThread(self)
        self.profit_check_thread.stage_reached.connect(self.on_stage_reached)
        self.profit_check_thread.error_occurred.connect(self.on_error_received)


    def create_separator(self):
        """創建水平分隔線"""
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        return line

    def create_option_group(self, title, option_name):
        """創建統一的選擇權輸入群組"""
        group = QGroupBox(title)
        font = QFont('Arial', 12)
        group.setFont(font)
        
        layout = QVBoxLayout()
        
        # 類型與履約價
        type_layout = QHBoxLayout()
        type_label = QLabel(f"{option_name}類型:")
        type_label.setFont(font)
        
        type_combo = QComboBox()
        type_combo.addItems(["Call", "Put"])
        type_combo.setFont(font)
        setattr(self, f"{option_name}_type_input", type_combo)

        # 紀錄下單時大盤的成交價
        setattr(self, f"{option_name}_market_price", 0) 
        
        strike_label = QLabel("履約價:")
        strike_label.setFont(font)
        
        strike_input = QLineEdit("0")
        strike_input.setPlaceholderText("輸入履約價")
        strike_input.setFont(font)
        setattr(self, f"{option_name}_strike_input", strike_input)
        
        type_layout.addWidget(type_label)
        type_layout.addWidget(type_combo)
        type_layout.addWidget(strike_label)
        type_layout.addWidget(strike_input)
        
        # 成交價
        price_layout = QHBoxLayout()
        price_label = QLabel("成交價:")
        price_label.setFont(font)
        
        price_input = QLineEdit("0")
        price_input.setPlaceholderText("輸入成交價")
        price_input.setFont(font)
        setattr(self, f"{option_name}_input", price_input)

        current_pnl_label = QLabel("當前損益:")
        current_pnl_label.setFont(font)
        
        current_pnl_input = QLineEdit("0")
        current_pnl_input.setFont(font)
        setattr(self, f"{option_name}_current_pnl_input", current_pnl_input)

        # 新增紀錄成交的按鈕
        save_record_button = QPushButton("紀錄")
        save_record_button.setFont(font)
        save_record_button.setFixedWidth(80)  # 設定固定寬度
        setattr(self, f"{option_name}_finish_button", save_record_button)
        setattr(self, f"{option_name}_is_recorded", False)
        save_record_button.clicked.connect(lambda _, name=option_name: self.save_record(name))

        # 新增報價按鈕
        quote_button = QPushButton("上下五檔")
        quote_button.setFont(font)
        quote_button.setFixedWidth(80)  # 設定固定寬度
        setattr(self, f"{option_name}_quote_button", quote_button)
        quote_button.clicked.connect(lambda _, name=option_name: self.show_quote_popup(name))

        price_layout.addWidget(price_label)
        price_layout.addWidget(price_input)
        price_layout.addWidget(current_pnl_label)
        price_layout.addWidget(current_pnl_input)
        price_layout.addWidget(save_record_button)
        price_layout.addWidget(quote_button)
        
        layout.addLayout(type_layout)
        layout.addLayout(price_layout)
        group.setLayout(layout)
        
        return group

    def init_ui(self):
        main_layout = QVBoxLayout()
        font = QFont('Arial', 12)

        # ==== 模式選擇 ====
        mode_frame = QWidget()
        mode_layout = QHBoxLayout()
        
        self.mode_diff_radio = QRadioButton("價差模式")
        self.mode_diff_radio.setFont(font)
        self.mode_atm_radio = QRadioButton("價平模式")
        self.mode_atm_radio.setFont(font)
        self.mode_atm_radio.setChecked(True) # 預設是價平模式

        inner_price_diff_label = QLabel("內部價差:")
        inner_price_diff_label.setFont(font)
        self.inner_price_diff_input = QLineEdit("0")
        self.inner_price_diff_input.setPlaceholderText("內部價差")
        self.inner_price_diff_input.setFont(font)
        self.inner_price_diff_input.setReadOnly(True) # 預設是0，不能被更動

        outter_price_diff_label = QLabel("外部價差:")
        outter_price_diff_label.setFont(font)
        self.outter_price_diff_input = QLineEdit("50")
        self.outter_price_diff_input.setPlaceholderText("外部價差")
        self.outter_price_diff_input.setFont(font)

        self.mode_atm_radio.toggled.connect(self.update_input_fields)
        self.mode_diff_radio.toggled.connect(self.update_input_fields)

        # ==== 新增重置按鈕 ====
        reset_button = QPushButton("重置")
        reset_button.setFont(font)
        reset_button.setFixedWidth(60)
        reset_button.clicked.connect(self.reset_ui)

        # ==== 新增繪畫總損益按鈕 ====
        total_pnl_plot_button = QPushButton("匯出總損益圖")
        total_pnl_plot_button.setFont(font)
        total_pnl_plot_button.setFixedWidth(120)
        # total_pnl_plot_button.clicked.connect(self.reset_ui)

        mode_layout.addWidget(self.mode_atm_radio)
        mode_layout.addWidget(self.mode_diff_radio)
        mode_layout.addSpacing(20)
        mode_layout.addWidget(inner_price_diff_label)
        mode_layout.addWidget(self.inner_price_diff_input)
        mode_layout.addWidget(outter_price_diff_label)
        mode_layout.addWidget(self.outter_price_diff_input)
        mode_layout.addWidget(reset_button)
        mode_layout.addWidget(total_pnl_plot_button)
        mode_layout.addStretch()
        mode_frame.setLayout(mode_layout)

        # ==== 按鈕區 ====
        top_button_layout = QHBoxLayout()
        self.bullish_button = QPushButton("看漲")
        self.bearish_button = QPushButton("看跌")


        button_font = QFont('Arial', 12, QFont.Weight.Bold)
        self.bullish_button.setFont(button_font)
        self.bearish_button.setFont(button_font)
        
        self.bullish_button.setStyleSheet("background-color: #FF0000;")
        self.bearish_button.setStyleSheet("background-color: #228B22;")
        
        self.bullish_button.clicked.connect(self.on_bullish_clicked)
        self.bearish_button.clicked.connect(self.on_bearish_clicked)
        
        top_button_layout.addWidget(self.bullish_button)
        top_button_layout.addWidget(self.bearish_button)

        # ==== 主內容 ====
        content_layout = QHBoxLayout()
        left_layout = QVBoxLayout()

        # ==== 創建四個群組 ====
        buy_1_group = self.create_option_group("買1 設定", "buy_1")
        sell_1_group = self.create_option_group("賣1 設定", "sell_1")
        buy_2_group = self.create_option_group("買2 設定", "buy_2")
        sell_2_group = self.create_option_group("賣2 設定", "sell_2")

        # ==== 賣1與賣2同步事件 ====
        self.sell_1_strike_input.textChanged.connect(self.sync_sell_2_strike_input)

        # ==== 計算按鈕 ====
        self.calculate_button = QPushButton("計算")
        self.calculate_button.setFont(font)
        self.calculate_button.clicked.connect(self.calculate)

        # ==== 紀錄按鈕 ====
        self.record_button = QPushButton("輸出紀錄")
        self.record_button.setFont(font)
        self.record_button.clicked.connect(self.output_data)

        # ==== 左側組裝 ====
        left_layout.addWidget(buy_1_group)
        left_layout.addWidget(self.create_separator())
        left_layout.addWidget(sell_1_group)
        left_layout.addWidget(self.create_separator())
        left_layout.addWidget(buy_2_group)
        left_layout.addWidget(self.create_separator())
        left_layout.addWidget(sell_2_group)
        left_layout.addWidget(self.create_separator())
        left_layout.addWidget(self.calculate_button)

        # ==== 右側輸出 ====
        right_layout = QVBoxLayout()

        # 輸出行情變化區域
        # analysis_output_label = QLabel("行情數據分析")
        # analysis_output_label.setFont(font)
        # self.analysis_output_text = QTextEdit("")
        # self.analysis_output_text.setFont(font)

        # 輸出計算結果區域
        cal_output_label = QLabel("計算結果")
        cal_output_label.setFont(font)
        self.cal_output_text = QTextEdit("")
        self.cal_output_text.setFont(font)

        # right_layout.addWidget(analysis_output_label)
        # right_layout.addWidget(self.analysis_output_text)
        right_layout.addWidget(cal_output_label)
        right_layout.addWidget(self.cal_output_text)
        right_layout.addWidget(self.create_separator())
        right_layout.addWidget(self.record_button)

        # ==== 組合整體畫面 ====
        content_layout.addLayout(left_layout, stretch=1)
        content_layout.addLayout(right_layout, stretch=1)

        main_layout.addWidget(mode_frame)
        main_layout.addLayout(top_button_layout)
        main_layout.addLayout(content_layout)

        self.setLayout(main_layout)


    def update_input_fields(self):
        if  self.mode_atm_radio.isChecked():
            self.inner_price_diff_input.setText("0")
            self.inner_price_diff_input.setReadOnly(True)
        elif self.mode_diff_radio.isChecked():
            self.inner_price_diff_input.setText("100")
            self.inner_price_diff_input.setReadOnly(False)


    def reset_ui(self):
        """只重置輸入框內容"""
        self.current_selected_group = None
        self.outter_price_diff_input.setText("50")
        self.buy_1_strike_input.setText("0")
        self.sell_1_strike_input.setText("0")
        self.buy_2_strike_input.setText("0")
        self.sell_2_strike_input.setText("0")
        self.buy_1_input.setText("0")
        self.sell_1_input.setText("0")
        self.buy_2_input.setText("0")
        self.sell_2_input.setText("0")
        self.cal_output_text.setText("")
        self.bullish_button.setStyleSheet("background-color: #FF0000;")
        self.bearish_button.setStyleSheet("background-color: #228B22;")
        self.mode_atm_radio.setChecked(True)
        self.update_input_fields()

        for option_name in ["buy_1", "buy_2", "sell_1", "sell_2"]:
            setattr(self, f"{option_name}_market_price", None)
            setattr(self, f"{option_name}_is_recorded", False)

            button = getattr(self, f"{option_name}_finish_button")
            button.setStyleSheet("")
            button.setText("紀錄")

        """關閉事件，確保安全停止線程"""
        if hasattr(self, 'profit_check_thread') and self.profit_check_thread.isRunning():
            self.profit_check_thread.stop()  # 停止線程
            self.profit_check_thread.wait()  # 等待線程完全結束


    def save_record(self, option_name):

        strike = getattr(self, f"{option_name}_strike_input").text()
        premium = getattr(self, f"{option_name}_input").text()
        is_recorded = getattr(self, f"{option_name}_is_recorded")
        button = getattr(self, f"{option_name}_finish_button")

        if strike == '0' or premium == '0':
            # 顯示警告視窗
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("輸入錯誤")
            msg.setText("請先輸入完整履約價或成交價")
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.exec()
            return

        if not is_recorded:
            index_contract = self.option_manager.api.Contracts.Futures.TXF.TXFR1 # 找台指期的指數
            index_snapshot = self.option_manager.api.snapshots([index_contract])[0]
            current_index_price = index_snapshot.close
            setattr(self, f"{option_name}_market_price", current_index_price)
            setattr(self, f"{option_name}_is_recorded", True)

            button.setStyleSheet("background-color: lightgray; color: black;")
            button.setText("已紀錄")

            # 特別紀錄buy_1的 market price
            if option_name=='buy_1':
                pass
                # 啟動一個線程去偵測目前市價跟buy_1的 market price關係

        else:
            setattr(self, f"{option_name}_market_price", None)
            setattr(self, f"{option_name}_is_recorded", False)

            button.setStyleSheet("")  
            button.setText("紀錄")

        # 根據記錄的 option_name 啟動對應檢查

        # if option_name == 'buy_1':
        #     self.update_plot(0) # 畫損益圖

        if option_name == 'sell_1':
            self.start_profit_check(1)  # 啟動階段1檢查

        if option_name == 'buy_2':
            self.start_profit_check(2)  # 啟動階段2檢查

        # elif option_name == 'sell_2':
        #     pass  # 啟動階段1檢查



    def start_profit_check(self, stage):
        if not self.profit_check_thread.isRunning():
            self.profit_check_thread.set_stage(stage)
            self.profit_check_thread.start()
        else:
            # 若線程已在運行，僅更新檢查階段
            self.profit_check_thread.set_stage(stage)


    def update_plot(self, stage):

        self.ax.clear()
        try:
            if stage == 0:
                # 獲取輸入值
                buy_1_premium = float(self.buy_1_input.text())
                buy_1_type = self.buy_1_type_input.currentText()
                buy_1_strike = float(self.buy_1_strike_input.text())
                
                # 計算買方損益
                if buy_1_type == "Call":
                    payoff = np.where(
                        self.price_range <= buy_1_strike, 
                        -buy_1_premium, 
                        self.price_range - buy_1_strike - buy_1_premium
                    )
                else:  # Put
                    payoff = np.where(
                        self.price_range >= buy_1_strike, 
                        -buy_1_premium, 
                        buy_1_strike - self.price_range - buy_1_premium
                    )
                
                # 繪製損益曲線
                self.ax.plot(self.price_range, payoff, 'r-', linewidth=2.5, label='損益曲線')
                
                # 標記關鍵點
                self.ax.plot(buy_1_strike, -buy_1_premium, 'bo', markersize=6, label='履約價')
                
                # 計算盈虧平衡點
                if buy_1_type == "Call":
                    breakeven = buy_1_strike + buy_1_premium
                    self.ax.plot(breakeven, 0, 'go', markersize=6, label='盈虧平衡點')
                    self.ax.annotate(f'{breakeven:.2f}', 
                                  xy=(breakeven, 0), 
                                  xytext=(breakeven+2, 5))
                else:  # Put
                    breakeven = buy_1_strike - buy_1_premium
                    self.ax.plot(breakeven, 0, 'go', markersize=6, label='盈虧平衡點')
                    self.ax.annotate(f'{breakeven:.2f}', 
                                  xy=(breakeven, 0), 
                                  xytext=(breakeven-10, 5))
                
                # 設置適當的y軸範圍
                y_min = min(min(payoff), -buy_1_premium * 1.5)
                y_max = max(max(payoff), buy_1_premium * 3)
                self.ax.set_ylim(y_min, y_max)

            
            # if stage == 1:
            #     # 獲取輸入值
            #     buy_1_premium = float(self.buy_1_input.text())
            #     buy_1_type = self.buy_1_type_input.currentText()
            #     buy_1_strike = float(self.buy_1_strike_input.text())
            #     sell_1_premium = float(self.sell_1_input.text())
            #     sell_1_type = self.sell_1_type_input.currentText()
            #     sell_1_strike = float(self.sell_1_strike_input.text())

            #     combined_premium = sell_1_premium-buy_1_premium

            #     # 計算買方損益
            #     if buy_1_type == "Call":
            #         payoff = np.where(
            #             self.price_range <= buy_1_strike, 
            #             -buy_1_premium, 
            #             self.price_range - buy_1_strike - buy_1_premium
            #         )
            #     else:  # Put
            #         payoff = np.where(
            #             self.price_range >= buy_1_strike, 
            #             -buy_1_premium, 
            #             buy_1_strike - self.price_range - buy_1_premium
            #         )
                
            # 添加基本元素
            self.ax.axhline(0, color='black', linestyle='-', linewidth=1)  # 零損益線
            self.ax.grid(True, linestyle='--', alpha=0.6)
            
            # 設置標題和標籤
            self.ax.set_title('選擇權損益曲線', fontsize=12, fontweight='bold')
            self.ax.set_xlabel('標的價格', fontsize=10)
            self.ax.set_ylabel('損益 (點)', fontsize=10)
            
            # 應用緊湊佈局
            self.figure.tight_layout()
            
            # 重新繪製畫布
            self.canvas.draw()
            
        except Exception as e:
            print(f"繪圖錯誤: {str(e)}")
            # 發生錯誤時顯示一個空圖
            self.ax.set_title('損益曲線 (請檢查輸入值)')
            self.ax.set_xlabel('標的價格')
            self.ax.set_ylabel('損益')
            self.cal_output_text.setText(f"錯誤: {str(e)}\n請檢查輸入值")
            self.canvas.draw()



    def show_quote_popup(self, option_name):

        contract = None
        strike = getattr(self, f"{option_name}_strike_input").text()
        option_type = 'P' if 'Put' in getattr(self, f"{option_name}_type_input").currentText() else 'C'
        delivery_month = '202505'
        code = 'TX1'

        if strike == '0' :
            # 顯示警告視窗
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("輸入錯誤")
            msg.setText("請先輸入履約價")
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.exec()
            return
        
        
        if option_type == 'P':
            contract = self.option_manager.api.Contracts.Options[code].get(f"{code}{delivery_month}{int(strike)}P")
        else:
            contract = self.option_manager.api.Contracts.Options[code].get(f"{code}{delivery_month}{int(strike)}C")

        if contract == None:
            # 顯示警告視窗
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("輸入錯誤")
            msg.setText("請先輸入履約價")
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.exec()
            return


        self.bidask_dialog = QDialog(self)
        self.bidask_dialog.setWindowTitle(f"{option_name} 上下五檔報價")
        self.bidask_dialog.resize(400, 350)
        layout = QVBoxLayout()

        # ================= 新增商品資訊列 =================
        self.product_info_label = QLabel(f"台指選 {code}{delivery_month}_{strike}{option_type}")  # 替換為實際商品ID
        self.product_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.product_info_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: bold;
                padding: 8px;
            }
        """)
        layout.addWidget(self.product_info_label)

        # 建立 10 行 4 欄的表格
        self.bidask_table = QTableWidget(10, 8)
        self.bidask_table.setHorizontalHeaderLabels(["刪買", "買進", "買量", "買價", "賣價", "賣量", "賣出", "刪賣"])

        # 設定每個標題欄位的背景顏色
        header_items = [
            QTableWidgetItem("刪買"), 
            QTableWidgetItem("買進"), 
            QTableWidgetItem("買量"),
            QTableWidgetItem("買價"),
            QTableWidgetItem("賣價"),
            QTableWidgetItem("賣量"),
            QTableWidgetItem("賣出"),  
            QTableWidgetItem("刪賣")  
        ]

        # 設紅色背景（買量、買價）
        header_items[0].setBackground(QColor(255, 0, 0))  
        header_items[1].setBackground(QColor(255, 0, 0))
        header_items[2].setBackground(QColor(255, 0, 0))  
        header_items[3].setBackground(QColor(255, 0, 0))

        # 設綠色背景（賣價、賣量）
        header_items[4].setBackground(QColor(34, 139, 34))  # 淺綠
        header_items[5].setBackground(QColor(34, 139, 34))
        header_items[6].setBackground(QColor(34, 139, 34))  # 淺綠
        header_items[7].setBackground(QColor(34, 139, 34))

        # 指派回去
        for i in range(8):
            self.bidask_table.setHorizontalHeaderItem(i, header_items[i])


        # 填入賣方（上面 5 行）
        for i in range(5):
            self.bidask_table.setItem(i, 4, QTableWidgetItem(""))
            self.bidask_table.setItem(i, 5, QTableWidgetItem(""))

        # 填入買方（下面 5 行）
        for i in range(5):
            row = 9 - i  # 由下往上填入
            self.bidask_table.setItem(row, 2, QTableWidgetItem(""))
            self.bidask_table.setItem(row, 3, QTableWidgetItem(""))

        # 對齊設定
        for row in range(10):
            for col in range(8):
                self.bidask_table.setItem(row, col, QTableWidgetItem(""))

        # 表格外觀設定
        self.bidask_table.verticalHeader().setVisible(False)
        self.bidask_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.bidask_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        layout.addWidget(self.bidask_table)
        self.bidask_dialog.setLayout(layout)

        self.option_manager.api.quote.subscribe(
            contract,
            quote_type = sj.constant.QuoteType.BidAsk,
            version=sj.constant.QuoteVersion.v1
        )

        # 監聽 Dialog 關閉事件，並在關閉時取消訂閱
        self.bidask_dialog.finished.connect(lambda: self.on_bidask_dialog_closed(contract))
        self.bidask_dialog.show()

    def update_bidask_table(self, bidask):
        if not hasattr(self, 'bidask_table') or not self.bidask_table:
            return  # 確保表格已初始化
        
        # 取得買賣五檔資料
        bid_prices = [str(price) for price in bidask.bid_price]  # 買價（低→高）
        bid_prices.reverse()  
        bid_volumes = [str(vol) for vol in bidask.bid_volume]    # 買量
        bid_volumes.reverse()  
        ask_prices = [str(price) for price in bidask.ask_price]  # 賣價（高→低）
        ask_prices.reverse()
        ask_volumes = [str(vol) for vol in bidask.ask_volume]    # 賣量
        ask_volumes.reverse()


        # 更新賣方
        for i in range(5):
            row = i  # 賣方從第0行開始
            self.bidask_table.setItem(row, 4, QTableWidgetItem(ask_prices[i]))  # 賣價
            self.bidask_table.setItem(row, 5, QTableWidgetItem(ask_volumes[i])) # 賣量

        # 更新買方
        for i in range(5):
            row = 9 - i  # 買方從第9行往上填
            self.bidask_table.setItem(row, 2, QTableWidgetItem(bid_volumes[i])) # 買量
            self.bidask_table.setItem(row, 3, QTableWidgetItem(bid_prices[i]))  # 買價

        # 刷新表格顯示
        self.bidask_table.viewport().update()


    def on_bidask_dialog_closed(self, contract):
        """當 BidAsk Dialog 關閉時，取消訂閱該合約的報價"""
        self.option_manager.api.quote.unsubscribe(contract, quote_type = sj.constant.QuoteType.BidAsk, version=sj.constant.QuoteVersion.v1)


    def on_stage_reached(self, stage):
        try:
            if stage == 1:
                # 階段1邏輯
                msg = "階段1達標：整體策略已鎖定利潤！"
                self.cal_output_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
                
            elif stage == 2:
                # 階段2邏輯
                msg = "階段2達標：整體策略已鎖定利潤！"
                self.cal_output_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
                
                # 自動儲存 + 停止檢查
                self.profit_check_thread.stop()
                
        except Exception as e:
            self.cal_output_text.append(f"錯誤: {str(e)}")



    def on_error_received(self, msg):
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            self.cal_output_text.append(f"[ERROR {timestamp}] {msg}")

            print(msg)
            
            QMessageBox.warning(
                self,
                "監測線程錯誤",
                f"達標檢查時發生錯誤:\n\n{msg}",
                QMessageBox.StandardButton.Ok
            )
            
            self.profit_check_thread.stop()
        
        except Exception as e:
            print(f"{str(e)}")


    def sync_sell_2_strike_input(self):
        if self.mode_atm_radio.isChecked():
            try:
                sell_1_value = int(self.sell_1_strike_input.text())
                self.sell_2_strike_input.setText(str(sell_1_value))
            except ValueError:
                pass  # 忽略非數字輸入

    def find_cloest_spot(self):
        future_contract = self.option_manager.api.Contracts.Futures.TXF.TXFR1
        snapshot = self.option_manager.api.snapshots([future_contract])[0]
        future_price = snapshot.close # 期貨報價
                                                                                
        contracts = self.option_manager.api.Contracts.Options['TX1']
        contracts_filtered = [contract for contract in contracts]

        strike_prices = sorted(set(contract.strike_price for contract in contracts_filtered))

        closest_strike = min(strike_prices, key=lambda x: abs(x - future_price)) # 價平
        
        return closest_strike


    def on_bearish_clicked(self):
        try:
            self.current_selected_group = "bearish"
            inner_price_diff = int(self.inner_price_diff_input.text())
            outter_price_diff = int(self.outter_price_diff_input.text())
            atm_strike=self.find_cloest_spot()

            if  self.mode_atm_radio.isChecked():
                sell_1 = int(atm_strike)
                sell_2 = int(atm_strike)
                buy_1 = int(sell_1 - outter_price_diff)
                buy_2 = int(sell_2 + outter_price_diff)

            else:
                sell_1 = int(atm_strike+inner_price_diff)
                sell_2 = int(atm_strike-inner_price_diff)
                buy_1 = int(sell_2 - outter_price_diff)
                buy_2 = int(sell_1 + outter_price_diff)

            self.sell_1_strike_input.setText(f"{sell_1}")
            self.sell_2_strike_input.setText(f"{sell_2}")
            self.buy_1_strike_input.setText(f"{buy_1}")
            self.buy_2_strike_input.setText(f"{buy_2}")

            # 設定選項類型（反向）
            self.buy_1_type_input.setCurrentText("Put")
            self.buy_2_type_input.setCurrentText("Call")
            self.sell_1_type_input.setCurrentText("Call")
            self.sell_2_type_input.setCurrentText("Put")

            # 更新按鈕樣式
            self.bullish_button.setStyleSheet("background-color: gray;")
            self.bearish_button.setStyleSheet("background-color: #228B22;")  # 恢復紅色

            # 找繪圖的範圍
            min_strike = min(int(self.buy_1_strike_input.text()), int(self.buy_2_strike_input.text()))
            max_strike = max(int(self.buy_1_strike_input.text()), int(self.buy_2_strike_input.text()))
            buffer = 400  # 上下100點
            self.price_range = np.linspace(
                min_strike - buffer,
                max_strike + buffer,
                100
            )

        except ValueError:
            self.cal_output_text.setText("輸入錯誤（賣1或價差）")


    def on_bullish_clicked(self):
        try:
            self.current_selected_group = "bullish"
            inner_price_diff = int(self.inner_price_diff_input.text())
            outter_price_diff = int(self.outter_price_diff_input.text())
            atm_strike=self.find_cloest_spot()

            if  self.mode_atm_radio.isChecked():
                sell_1 = int(atm_strike)
                sell_2 = int(atm_strike)
                buy_1 = int(sell_1 + outter_price_diff)
                buy_2 = int(sell_2 - outter_price_diff)

            else:
                sell_1 = int(atm_strike-inner_price_diff)
                sell_2 = int(atm_strike+inner_price_diff)
                buy_1 = int(sell_2+outter_price_diff)
                buy_2 = int(sell_1-outter_price_diff)

            self.sell_1_strike_input.setText(f"{sell_1}")
            self.sell_2_strike_input.setText(f"{sell_2}")
            self.buy_1_strike_input.setText(f"{buy_1}")
            self.buy_2_strike_input.setText(f"{buy_2}")

            # 設定選項類型
            self.buy_1_type_input.setCurrentText("Call")
            self.buy_2_type_input.setCurrentText("Put")
            self.sell_1_type_input.setCurrentText("Put")
            self.sell_2_type_input.setCurrentText("Call")

            # 更新按鈕樣式
            self.bearish_button.setStyleSheet("background-color: gray;")
            self.bullish_button.setStyleSheet("background-color: #FF0000;")  

            # 找繪圖的範圍
            min_strike = min(int(self.buy_1_strike_input.text()), int(self.buy_2_strike_input.text()))
            max_strike = max(int(self.buy_1_strike_input.text()), int(self.buy_2_strike_input.text()))
            buffer = 400  # 上下100點
            self.price_range = np.linspace(
                min_strike - buffer,
                max_strike + buffer,
                100
            )

        except ValueError:
            self.cal_output_text.setText("輸入錯誤（賣1或價差）")


    def calculate(self):
        try:
            sell_1 = float(self.sell_1_input.text())
            sell_2 = float(self.sell_2_input.text())
            buy_1 = float(self.buy_1_input.text())
            buy_2 = float(self.buy_2_input.text())
            price_diff = float(self.outter_price_diff_input.text())

            # Get option types
            buy_1_type = self.buy_1_type_input.currentText()
            sell_1_type = self.sell_1_type_input.currentText()
            buy_2_type = self.buy_2_type_input.currentText()
            sell_2_type = self.sell_2_type_input.currentText()

            self.result = sell_1 + sell_2 - buy_1 - buy_2 - price_diff

            if self.result >= 0:
                self.cal_output_text.setText(f"已達標，至少獲利 {self.result:.1f} 點\n")
            else:
                required_sell_2 = buy_1 + buy_2 + price_diff - sell_1
                self.cal_output_text.setText(f"未達標，至少損失 {self.result:.1f} 點\n"
                                         f"賣二至少需要大於 {required_sell_2:.1f} 才能達標\n")
        except ValueError:
            self.cal_output_text.setText("輸入錯誤，請檢查數值")


    def isSumProfitable(self, stage):

        delivery_month = '202505'
        code = 'TX1'
        buy_1_premium = float(self.buy_1_input.text())
        sell_1_premium = float(self.sell_1_input.text())
        price_diff = float(self.outter_price_diff_input.text())

        if stage == 1:
            # 取得buy_2的報價
            buy_2_strike = getattr(self, "buy_2_strike_input").text()
            buy_2_option_type = 'P' if 'Put' in getattr(self, "buy_2_type_input").currentText() else 'C'
            buy_2_contract = self.option_manager.api.Contracts.Options[code].get(f"{code}{delivery_month}{int(buy_2_strike)}P")
            buy_2_snapshot = self.option_manager.api.snapshots([buy_2_contract])[0]
            buy_2_premium = buy_2_snapshot.close

            # 取得sell_2的報價
            sell_2_strike = getattr(self, "sell_2_strike_input").text()
            sell_2_option_type = 'P' if 'Put' in getattr(self, "sell_2_type_input").currentText() else 'C'
            sell_2_contract = self.option_manager.api.Contracts.Options[code].get(f"{code}{delivery_month}{int(sell_2_strike)}P")
            sell_2_snapshot = self.option_manager.api.snapshots([sell_2_contract])[0]
            sell_2_premium = sell_2_snapshot.close

            # 計算是否已達標，若有就可以下單
            if sell_1_premium+sell_2_premium-buy_1_premium-buy_2_premium-price_diff>0:
                print("已達標!")

        else:
            # 取得sell_2的報價
            sell_2_strike = getattr(self, "sell_2_strike_input").text()
            sell_2_option_type = 'P' if 'Put' in getattr(self, "sell_2_type_input").currentText() else 'C'
            sell_2_contract = self.option_manager.api.Contracts.Options[code].get(f"{code}{delivery_month}{int(sell_2_strike)}P")

            #記得用移動停利


    def output_data(self):
        try:
            # 收集資料
            strategy = "看漲" if self.current_selected_group == "bullish" else "看跌"
            
            sell_1_premium = float(self.sell_1_input.text())
            sell_2_premium = float(self.sell_2_input.text())
            buy_1_premium = float(self.buy_1_input.text())
            buy_2_premium = float(self.buy_2_input.text())

            buy_1_type = self.buy_1_type_input.currentText()
            buy_2_type = self.buy_2_type_input.currentText()
            sell_1_type = self.sell_1_type_input.currentText()
            sell_2_type = self.sell_2_type_input.currentText()

            buy_1_strike = int(self.buy_1_strike_input.text())
            buy_2_strike = int(self.buy_2_strike_input.text())
            sell_1_strike = int(self.sell_1_strike_input.text())
            sell_2_strike = int(self.sell_2_strike_input.text())

            buy_1_market_price = int(getattr(self, "buy_1_market_price"))
            buy_2_market_price = int(getattr(self, "buy_2_market_price"))
            sell_1_market_price = int(getattr(self, "sell_1_market_price"))
            sell_2_market_price = int(getattr(self, "sell_2_market_price"))

            P_and_L = self.result * 50

            # 固定檔名（同一個檔案）
            csv_filename = "option_strategy_results.csv"
            
            # 檢查檔案是否存在，決定是否要寫入標題
            file_exists = os.path.isfile(csv_filename)
            
            # 寫入 CSV 檔案（使用 'a' 模式追加，newline='' 避免空行）
            with open(csv_filename, mode='a', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file)
                
                # 如果檔案不存在，寫入標題列
                if not file_exists:
                    writer.writerow([
                        "時間戳記",
                        "策略類型", 
                        "買權1類型", "買權1履約價", "買權1權利金", "買權1市價",
                        "賣權1類型", "賣權1履約價", "賣權1權利金", "賣權1市價",
                        "買權2類型", "買權2履約價", "買權2權利金", "買權2市價",
                        "賣權2類型", "賣權2履約價", "賣權2權利金", "賣權2市價",
                        "損益 (P&L)"
                    ])
                
                # 寫入資料列（加入時間戳記）
                writer.writerow([
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    strategy,
                    buy_1_type, buy_1_strike, buy_1_premium, buy_1_market_price,
                    sell_1_type, sell_1_strike, sell_1_premium, sell_1_market_price,
                    buy_2_type, buy_2_strike, buy_2_premium, buy_2_market_price,
                    sell_2_type, sell_2_strike, sell_2_premium, sell_2_market_price,
                    P_and_L
                ])

            # 顯示成功訊息
            self.cal_output_text.setText(f"資料已成功追加至 CSV: {os.path.abspath(csv_filename)}")

            """關閉事件，確保安全停止線程"""
            if hasattr(self, 'profit_check_thread') and self.profit_check_thread.isRunning():
                self.profit_check_thread.stop()  # 停止線程
                self.profit_check_thread.wait()  # 等待線程完全結束

        except ValueError:
            self.cal_output_text.setText("錯誤：請檢查輸入是否為有效數字")
        except Exception as e:
            self.cal_output_text.setText(f"匯出失敗: {str(e)}")


    def closeEvent(self, event):
        """關閉事件，確保安全停止線程"""
        if hasattr(self, 'profit_check_thread') and self.profit_check_thread.isRunning():
            self.profit_check_thread.stop()  # 停止線程
            self.profit_check_thread.wait()  # 等待線程完全結束
        
        # 確保 Shioaji API 登出
        if hasattr(self, 'option_manager') and hasattr(self.option_manager, 'api'):
            self.option_manager.api.logout()
        
        event.accept()  # 允許關閉視窗


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OptionTQuoteApp()
    window.show()
    sys.exit(app.exec())