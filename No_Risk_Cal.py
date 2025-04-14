import sys
import shioaji as sj
from datetime import datetime, timedelta, time
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QRadioButton,
    QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QComboBox,
    QFrame, QGroupBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import pyqtgraph as pg
import logging


# class OptionDataManager:
#     def __init__(self):

#         self.api = sj.Shioaji()
#         # 建立 logger 物件
#         self.logger = logging.getLogger(__name__)
#         self.api.quote.on_event(self.event_callback)

#         self.api.login(
#             api_key = "", 
#             secret_key ="", 
#         )

#         Result = self.api.activate_ca(  
#             ca_path=r"C:\Users\user\Desktop\公司帳戶\Sinopac.pfx",
#             ca_passwd="83167949",
#             person_id="83167949",
#         )


#     def event_callback(self, resp_code: int, event_code: int, info: str, event: str):
#         now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#         self.logger.info(f"[{now}] Event code: {event_code} | Event: {event}")
#         if event_code == 1 or event_code == 2 or event_code == 12 or event_code == 13:  # SOLCLIENT_SESSION_EVENT_DOWN_ERROR
#             sys.exit()


class OptionTQuoteApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("無風險計算機")
        self.setGeometry(100, 50, 600, 180)  
        # self.option_manager = OptionDataManager()
        self.init_ui()

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

        # 新增完成按鈕
        confirm_button = QPushButton("完成")
        confirm_button.setFont(font)
        confirm_button.setFixedWidth(60)  # 設定固定寬度
        setattr(self, f"{option_name}_confirm_button", confirm_button)
        
        price_layout.addWidget(price_label)
        price_layout.addWidget(price_input)
        price_layout.addWidget(confirm_button)
        
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
        self.mode_atm_radio = QRadioButton("價平模式")
        self.mode_atm_radio.setChecked(True)
        self.mode_diff_radio.setFont(font)
        self.mode_atm_radio.setFont(font)

        price_diff_label = QLabel("價差點數:")
        price_diff_label.setFont(font)
        self.price_diff_input = QLineEdit("100")
        self.price_diff_input.setPlaceholderText("價差點數")
        self.price_diff_input.setFont(font)

        # 新增重置按鈕
        reset_button = QPushButton("重置")
        reset_button.setFont(font)
        reset_button.setFixedWidth(60)
        reset_button.clicked.connect(self.reset_ui)

        mode_layout.addWidget(self.mode_atm_radio)
        mode_layout.addWidget(self.mode_diff_radio)
        mode_layout.addSpacing(20)
        mode_layout.addWidget(price_diff_label)
        mode_layout.addWidget(self.price_diff_input)
        mode_layout.addWidget(reset_button)
        mode_layout.addStretch()
        mode_frame.setLayout(mode_layout)

        # ==== 按鈕區 ====
        top_button_layout = QHBoxLayout()
        self.bullish_button = QPushButton("看漲")
        self.bearish_button = QPushButton("看跌")
        
        button_font = QFont('Arial', 12, QFont.Weight.Bold)
        self.bullish_button.setFont(button_font)
        self.bearish_button.setFont(button_font)
        
        self.bullish_button.setStyleSheet("background-color: #FF9999;")
        self.bearish_button.setStyleSheet("background-color: #99FF99;")
        
        self.bullish_button.clicked.connect(self.on_bullish_clicked)
        self.bearish_button.clicked.connect(self.on_bearish_clicked)
        
        top_button_layout.addWidget(self.bullish_button)
        top_button_layout.addWidget(self.bearish_button)

        # ==== 主內容 ====
        content_layout = QHBoxLayout()
        left_layout = QVBoxLayout()

        # 創建四個群組
        buy_1_group = self.create_option_group("買1 設定", "buy_1")
        sell_1_group = self.create_option_group("賣1 設定", "sell_1")
        buy_2_group = self.create_option_group("買2 設定", "buy_2")
        sell_2_group = self.create_option_group("賣2 設定", "sell_2")

        # 賣1與賣2同步事件
        self.sell_1_strike_input.textChanged.connect(self.sync_sell_2_strike_input)

        # ==== 計算按鈕 ====
        self.calculate_button = QPushButton("計算")
        self.calculate_button.setFont(font)
        self.calculate_button.clicked.connect(self.calculate)

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
        output_label = QLabel("輸出結果")
        output_label.setFont(font)
        self.output_text = QTextEdit("0")
        self.output_text.setPlaceholderText("結果會顯示在這裡")
        self.output_text.setFont(font)

        right_layout.addWidget(output_label)
        right_layout.addWidget(self.output_text)

        # ==== 組合整體畫面 ====
        content_layout.addLayout(left_layout, stretch=1)
        content_layout.addLayout(right_layout, stretch=1)

        main_layout.addWidget(mode_frame)
        main_layout.addLayout(top_button_layout)
        main_layout.addLayout(content_layout)

        self.setLayout(main_layout)


    def reset_ui(self):
        """只重置輸入框內容"""
        self.price_diff_input.setText("100")
        self.buy_1_strike_input.setText("0")
        self.sell_1_strike_input.setText("0")
        self.buy_2_strike_input.setText("0")
        self.sell_2_strike_input.setText("0")
        self.buy_1_input.setText("0")
        self.sell_1_input.setText("0")
        self.buy_2_input.setText("0")
        self.sell_2_input.setText("0")
        self.output_text.setText("0")
        self.bullish_button.setStyleSheet("background-color: #FF9999;")
        self.bearish_button.setStyleSheet("background-color: #99FF99;")
        self.mode_atm_radio.setChecked(True)


    def sync_sell_2_strike_input(self):
        if self.mode_atm_radio.isChecked():
            try:
                sell_1_value = int(self.sell_1_strike_input.text())
                self.sell_2_strike_input.setText(str(sell_1_value))
            except ValueError:
                pass  # 忽略非數字輸入


    def on_bearish_clicked(self):
        try:
            sell_1 = float(self.sell_1_strike_input.text())
            sell_2 = float(self.sell_2_strike_input.text())
            price_diff = float(self.price_diff_input.text())

            buy_1 = int(sell_1 - price_diff)
            buy_2 = int(sell_2 + price_diff)

            self.buy_1_strike_input.setText(f"{buy_1}")
            self.buy_2_strike_input.setText(f"{buy_2}")

            # 設定選項類型（反向）
            self.buy_1_type_input.setCurrentText("Put")
            self.buy_2_type_input.setCurrentText("Call")
            self.sell_1_type_input.setCurrentText("Call")
            self.sell_2_type_input.setCurrentText("Put")

            # 更新按鈕樣式
            self.bullish_button.setStyleSheet("background-color: gray;")
            self.bearish_button.setStyleSheet("background-color: #99FF99;")  # 恢復紅色
        except ValueError:
            self.output_text.setText("輸入錯誤（賣1或價差）")


    def on_bullish_clicked(self):
        try:
            sell_1 = float(self.sell_1_strike_input.text())
            sell_2 = float(self.sell_2_strike_input.text())
            price_diff = float(self.price_diff_input.text())

            buy_1 = int(sell_1 + price_diff)
            buy_2 = int(sell_2 - price_diff)

            self.buy_1_strike_input.setText(f"{buy_1}")
            self.buy_2_strike_input.setText(f"{buy_2}")

            # 設定選項類型
            self.buy_1_type_input.setCurrentText("Call")
            self.buy_2_type_input.setCurrentText("Put")
            self.sell_1_type_input.setCurrentText("Put")
            self.sell_2_type_input.setCurrentText("Call")

            # 更新按鈕樣式
            self.bearish_button.setStyleSheet("background-color: gray;")
            self.bullish_button.setStyleSheet("background-color: #FF9999;")  # 恢復綠色
        except ValueError:
            self.output_text.setText("輸入錯誤（賣1或價差）")


    def calculate(self):
        try:
            sell_1 = float(self.sell_1_input.text())
            sell_2 = float(self.sell_2_input.text())
            buy_1 = float(self.buy_1_input.text())
            buy_2 = float(self.buy_2_input.text())
            price_diff = float(self.price_diff_input.text())

            # Get option types
            buy_1_type = self.buy_1_type_input.currentText()
            sell_1_type = self.sell_1_type_input.currentText()
            buy_2_type = self.buy_2_type_input.currentText()
            sell_2_type = self.sell_2_type_input.currentText()

            result = sell_1 + sell_2 - buy_1 - buy_2 - price_diff

            if result >= 0:
                self.output_text.setText(f"已達標，至少獲利 {result:.1f} 點\n")
            else:
                required_sell_2 = buy_1 + buy_2 + price_diff - sell_1
                self.output_text.setText(f"賣二至少需要大於 {required_sell_2:.1f} 才能達標\n")
        except ValueError:
            self.output_text.setText("輸入錯誤，請檢查數值")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OptionTQuoteApp()
    window.show()
    sys.exit(app.exec())