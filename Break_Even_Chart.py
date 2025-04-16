import shioaji as sj
import numpy as np
from datetime import datetime, timedelta, time
import sys
from PyQt6 import QtWidgets
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QBrush, QColor
from PyQt6 import QtCore
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QScrollArea, 
    QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, 
    QTableWidget, QTableWidgetItem, QHeaderView, QGroupBox, QFrame
)
import pyqtgraph as pg
from pyqtgraph.Qt import QtWidgets  



class OptionDataManager:
    def __init__(self):
        self.api = sj.Shioaji()
        self.api.login(
            api_key="", 
            secret_key="", 
        )

        self.option_codes = ['TXO', 'TX1', 'TX2', 'TX4', 'TX5']
        self.expirations = self.get_option_expirations()

        Result = self.api.activate_ca(  
            ca_path=r"C:\Users\veron\Desktop\永豐api\Sinopac.pfx",
            ca_passwd="",
            person_id="",
        )

    def get_option_expirations(self):
        expirations_dict = {}  

        for code in self.option_codes:

            if hasattr(self.api.Contracts.Options, code):
                code_contracts = getattr(self.api.Contracts.Options, code)

                code_contracts_sorted = sorted(
                    list(code_contracts),
                    key=lambda c: datetime.strptime(c.delivery_date, "%Y/%m/%d")
                )

                for contract in code_contracts_sorted:

                    if contract.category not in expirations_dict:
                        expirations_dict[contract.category] = []

                    if contract.delivery_date not in expirations_dict[contract.category]:
                        expirations_dict[contract.category].append(contract.delivery_date)
                    else:
                        continue


        # 將字典轉換為 (商品, 日期) 的元組列表
        expiration_list = []
        for product, dates in expirations_dict.items():
            for date in dates:
                expiration_list.append((product, date))

        # 按照日期排序
        expiration_list.sort(key=lambda x: x[1])

        return expiration_list


class OptionAnalyzerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.option_manager = OptionDataManager()
        self.setWindowTitle("選擇權分析工具")
        self.setGeometry(100, 100, 1170, 900)
        # 取得螢幕大小
        screen = QtWidgets.QApplication.primaryScreen()
        screen_rect = screen.availableGeometry()
        
        # 計算螢幕中心 - 視窗大小的一半
        x = (screen_rect.width() - self.width()) // 2
        y = (screen_rect.height() - self.height()) // 2
        
        # 移動視窗到計算後的位置
        self.move(x, y)

        self.init_ui()


    def init_ui(self):
        # 主畫布
        main_layout = QVBoxLayout()

        # 上方區域
        control_layout = QHBoxLayout()
        control_layout.setSpacing(10)

        control_layout.addStretch(1)

        # 現價顯示
        self.current_price_label = QLabel("現價：取得中...")
        control_layout.addWidget(self.current_price_label)
        self.update_current_price()

        # 下單類型選擇
        self.strategy_type_combo = QComboBox()  # 改名為 strategy_type_combo
        self.strategy_type_combo.addItems(['Short Call', 'Short Put', 'Long Call', 'Long Put'])
        control_layout.addWidget(QLabel("選擇下單類型："))
        control_layout.addWidget(self.strategy_type_combo)

        # 到期日選擇 (改名為 expiration_combo)
        self.expiration_combo = QComboBox()
        expiration_strings = [f"{product} ({date})" for (product, date) in self.option_manager.expirations]
        self.expiration_combo.addItems(expiration_strings)
        control_layout.addWidget(QLabel("選擇合約(到期日)："))
        control_layout.addWidget(self.expiration_combo)

        # 履約價選擇
        self.strikes_combo = QComboBox()
        control_layout.addWidget(QLabel("選擇履約價："))
        control_layout.addWidget(self.strikes_combo)

        control_layout.addStretch(1)

        # 初始化履約價列表
        self.update_strikes_combo()

        # 連接信號
        self.strategy_type_combo.currentTextChanged.connect(self.update_strikes_combo)
        self.expiration_combo.currentTextChanged.connect(self.update_strikes_combo)

        main_layout.addLayout(control_layout)
        self.init_profit_charts(main_layout)
        self.setLayout(main_layout)

    def update_strikes_combo(self):
        """更新履約價下拉選單內容"""
        # 清除現有內容
        self.strikes_combo.clear()
        
        # 獲取當前選擇的下單類型和到期日
        strategy_type = self.strategy_type_combo.currentText()
        selected_text = self.expiration_combo.currentText()
        
        try:
            code, delivery_date = selected_text.split(" (")
            delivery_date = delivery_date[:-1]  # 移除末尾的 ")"
            
            # 根據策略類型確定買賣權
            option_right = 'P' if 'Put' in strategy_type else 'C'
            
            if hasattr(self.option_manager.api.Contracts.Options, code):
                code_contracts = getattr(self.option_manager.api.Contracts.Options, code)
                
                # 獲取並排序履約價
                self.sorted_strikes = sorted(
                    [int(c.strike_price) for c in code_contracts 
                    if c.delivery_date == delivery_date 
                    and c.option_right.name[0] == option_right]
                )
                
                # 添加履約價到下拉選單
                self.strikes_combo.addItems([str(strike) for strike in self.sorted_strikes])
                
                # 找到與 current_price 最接近的履約價
                if hasattr(self, 'current_price') and self.sorted_strikes:
                    closest_strike = min(self.sorted_strikes, key=lambda p: abs(p - self.current_price))
                    index = self.strikes_combo.findText(str(closest_strike))
                    
                    if index >= 0:
                        # 設置背景色
                        model = self.strikes_combo.model()
                        item = model.index(index, 0)
                        model.setData(item, QBrush(QColor(211, 211, 211)), QtCore.Qt.ItemDataRole.BackgroundRole)
                        
                        popup = self.strikes_combo.view()
        
                        visible_items = popup.height() // popup.sizeHintForRow(0)
                        
                        start = max(0, index - (visible_items // 2))
                        end = min(self.strikes_combo.count(), start + visible_items)
                        
                        if end == self.strikes_combo.count():
                            start = max(0, end - visible_items)
                        
                        popup.setMinimumHeight(popup.sizeHintForRow(0) * visible_items)
                        popup.setMaximumHeight(popup.sizeHintForRow(0) * visible_items)
                        
                        popup.scrollTo(self.strikes_combo.model().index(start, 0), QtWidgets.QAbstractItemView.ScrollHint.PositionAtTop)

                        self.strikes_combo.setCurrentIndex(index)

        except Exception as e:
            print(f"更新履約價時發生錯誤: {e}")


    def init_profit_charts(self, parent_layout):
        """初始化損益平衡圖並添加到布局"""
        # 創建主圖形容器（水平布局）
        chart_container = QHBoxLayout()
        parent_layout.addLayout(chart_container)
        
        # ===== 左側區域（垂直排列四個小圖）=====
        left_panel = QVBoxLayout()
        chart_container.addLayout(left_panel, stretch=1)  # stretch=1 表示左側佔1份
        
        # 創建左側的圖形容器並放入滾動區域
        self.left_graph = pg.GraphicsLayoutWidget()
        
        # 將 GraphicsLayoutWidget 放入 QScrollArea
        scroll_area = QScrollArea()
        scroll_area.setWidget(self.left_graph)
        scroll_area.setWidgetResizable(True)  # 允許小部件隨滾動區域調整大小
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)  # 顯示水平滾動條
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)   # 禁用垂直滾動條
        left_panel.addWidget(scroll_area)
        
        # 初始化4個垂直排列的子圖
        self.plot1 = self.left_graph.addPlot(title="Sell Put", row=0, col=0)
        self.plot2 = self.left_graph.addPlot(title="Sell Call", row=1, col=0)
        self.plot3 = self.left_graph.addPlot(title="Buy Call", row=2, col=0)
        self.plot4 = self.left_graph.addPlot(title="Buy Put", row=3, col=0)

        # 設定刻度
        if self.current_price not in self.sorted_strikes:
            self.sorted_strikes.append(self.current_price)
            self.sorted_strikes.sort()
            index_position = self.sorted_strikes.index(self.current_price)  
            self.sorted_strikes.remove(self.current_price) 
        else:
            index_position = self.sorted_strikes.index(self.current_price) 

        lower_bound = max(index_position - 10, 0)
        upper_bound = min(index_position + 10, len(self.sorted_strikes))

        filtered_strikes = self.sorted_strikes[lower_bound:upper_bound]
        
        # 生成所有刻度（完整分佈）
        all_ticks = [(strike, str(strike)) for strike in self.sorted_strikes]

        # 設置統一樣式
        for i, plot in enumerate([self.plot1, self.plot2, self.plot3, self.plot4]):
            plot.setLabel('left', "Profit")
            plot.setLabel('bottom', "Underlying Price" if i == 3 else None)  # 只在最下方顯示x軸標籤
            plot.addLegend()
            plot.setMinimumHeight(200)  # 設定最小高度

            # 設置初始可見範圍，但允許滾動查看完整分佈
            plot.setXRange(min(filtered_strikes), max(filtered_strikes), padding=0)
            plot.getAxis('bottom').setTicks([all_ticks])  # 使用完整刻度

            # 固定Y軸範圍 (-100, 100)
            plot.setYRange(-100, 100)

            # 禁用滑鼠滾輪縮放 X 軸
            view_box = plot.getViewBox()
            view_box.wheelEvent = lambda event: None  # 禁用滾輪事件，防止縮放
            
            # 在Y軸0的位置添加參考線
            zero_line = pg.InfiniteLine(
                pos=0, 
                angle=0,  # 水平線
                pen=pg.mkPen(color='r', width=1, style=Qt.PenStyle.DashLine)
            )
            plot.addItem(zero_line)
            
            # 設置Y軸刻度（可選）
            y_axis = plot.getAxis('left')
            y_axis.setTicks([
                [(0, "0"), (50, "50"), (100, "100")],  # 正半軸
                [(-50, "-50"), (-100, "-100")]         # 負半軸
            ])

                
            
        # ===== 右側區域（單個大圖）=====
        right_panel = QVBoxLayout()
        chart_container.addLayout(right_panel, stretch=2)  # stretch=2 表示右側佔2份

        # 創建右側的疊加圖並放入滾動區域
        self.combined_graph = pg.GraphicsLayoutWidget()

        # 將 GraphicsLayoutWidget 放入 QScrollArea
        scroll_area_right = QScrollArea()  # 為右側圖表創建獨立的滾動區域
        scroll_area_right.setWidget(self.combined_graph)
        scroll_area_right.setWidgetResizable(True)  # 允許小部件隨滾動區域調整大小
        scroll_area_right.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)  # 顯示水平滾動條
        scroll_area_right.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)   # 禁用垂直滾動條
        right_panel.addWidget(scroll_area_right)

        # 初始化疊加圖
        self.combined_plot = self.combined_graph.addPlot(title="Combined Strategies")
        self.combined_plot.setLabel('left', "Profit")
        self.combined_plot.setLabel('bottom', "Underlying Price")
        self.combined_plot.setXRange(min(filtered_strikes), max(filtered_strikes), padding=0)
        self.combined_plot.getAxis('bottom').setTicks([all_ticks])

        # 固定Y軸範圍 (-100, 100)
        self.combined_plot.setYRange(-100, 100)

        # 禁用滑鼠滾輪縮放 X 軸
        view_box = self.combined_plot.getViewBox()
        view_box.wheelEvent = lambda event: None  # 禁用滾輪事件，防止縮放

        # 在Y軸0的位置添加參考線
        zero_line = pg.InfiniteLine(
            pos=0, 
            angle=0,  # 水平線
            pen=pg.mkPen(color='r', width=1, style=Qt.PenStyle.DashLine)
        )
        self.combined_plot.addItem(zero_line)

        # 設置Y軸刻度（可選）
        y_axis = self.combined_plot.getAxis('left')
        y_axis.setTicks([
            [(0, "0"), (50, "50"), (100, "100")],  # 正半軸
            [(-50, "-50"), (-100, "-100")]         # 負半軸
        ])
        self.combined_plot.addLegend()

        # 設定布局比例
        chart_container.setStretchFactor(left_panel, 1)
        chart_container.setStretchFactor(right_panel, 2)

        # 連接信號：當選擇變更時更新圖表
        self.strategy_type_combo.currentTextChanged.connect(self.update_charts)
        self.strikes_combo.currentTextChanged.connect(self.update_charts)


    def update_charts(self):
        try:
            strike = float(self.strikes_combo.currentText())
            premium = 10  # 假设权利金固定为10
            
            # 計算損益
            x = np.linspace(min(self.sorted_strikes), max(self.sorted_strikes), 500)
            y_sell_put = np.where(x >= strike, premium, premium - (strike - x))
            y_sell_call = np.where(x <= strike, premium, premium - (x - strike))
            y_buy_call = np.where(x <= strike, -premium, (x - strike) - premium)
            y_buy_put = np.where(x >= strike, -premium, (strike - x) - premium)

            plots = [self.plot1, self.plot2, self.plot3, self.plot4]
            y_data = [y_sell_put, y_sell_call, y_buy_call, y_buy_put]
            colors = ['r', 'g', 'b', 'm']
            names = ["Sell Put", "Sell Call", "Buy Call", "Buy Put"]
            
            for plot, y, color, name in zip(plots, y_data, colors, names):
                plot.clear()
                plot.plot(x, y, pen=color, name=name)

                # 在Y軸0的位置添加參考線
                zero_line = pg.InfiniteLine(
                    pos=0, 
                    angle=0,  # 水平線
                    pen=pg.mkPen(color='r', width=1, style=Qt.PenStyle.DashLine)
                )
                plot.addItem(zero_line)
                
                # 自動將strike滾動到圖中央，並保持一樣的x刻度範圍）
                current_view = plot.getViewBox().viewRange()[0]
                view_width = current_view[1] - current_view[0]
                plot.setXRange(strike - view_width/2, strike + view_width/2, padding=0)

                
            # # 添加垂直标记线
            # strike_line = pg.InfiniteLine(
            #     pos=strike, angle=90,
            #     pen=pg.mkPen(color=(0, 0, 0), width=1.5,
            #     label=f"Strike: {strike}",
            #     labelOpts={'position': 0.9, 'color': 'k', 'fill': (255,255,255,150)}
            # ))
            # self.combined_plot.addItem(strike_line)
            
        except Exception as e:
            print(f"更新图表时出错: {e}")


    
    def update_current_price(self):
        try:
            future_contract = self.option_manager.api.Contracts.Futures.TXF.TXFR1
            snapshot = self.option_manager.api.snapshots([future_contract])[0]
            self.current_price = snapshot.close
            self.current_price_label.setText(f"現價：{self.current_price}")

        except Exception as e:
            print(f"無法取得當前價格。錯誤：{e}")
            self.current_price_label.setText("現價：無法取得")


# strike = 100
# premium = 10
# x = np.linspace(50, 150, 500)  # 設定標的價格範圍

# # 損益計算公式：
# y_sell_put = np.where(x >= strike, premium, x + (premium - strike))
# y_sell_call = np.where(x <= strike, premium, premium - (x - strike))
# y_buy_call = np.where(x <= strike, -premium, (x - strike) - premium)
# y_buy_put = np.where(x >= strike, -premium, (strike - x) - premium)

# # 建立 Qt 應用程式與視窗
# app = QtWidgets.QApplication([])
# win = pg.GraphicsLayoutWidget(show=True, title="選擇權損益平衡圖")
# win.resize(1000, 600)
# win.setWindowTitle('PyQtGraph - Options Profit Diagrams')

# # 在 2x2 的網格中繪製各圖

# # 賣出 Put 損益平衡圖 (左上)
# p1 = win.addPlot(title="Sell Put Profit Diagram", row=0, col=0)
# p1.plot(x, y_sell_put, pen='r')
# p1.setLabel('left', "Profit")
# p1.setLabel('bottom', "Underlying Price")

# # 賣出 Call 損益平衡圖 (右上)
# p2 = win.addPlot(title="Sell Call Profit Diagram", row=0, col=1)
# p2.plot(x, y_sell_call, pen='g')
# p2.setLabel('left', "Profit")
# p2.setLabel('bottom', "Underlying Price")

# # 買進 Call 損益平衡圖 (左下)
# p3 = win.addPlot(title="Buy Call Profit Diagram", row=1, col=0)
# p3.plot(x, y_buy_call, pen='b')
# p3.setLabel('left', "Profit")
# p3.setLabel('bottom', "Underlying Price")

# # 買進 Put 損益平衡圖 (右下)
# p4 = win.addPlot(title="Buy Put Profit Diagram", row=1, col=1)
# p4.plot(x, y_buy_put, pen='m')
# p4.setLabel('left', "Profit")
# p4.setLabel('bottom', "Underlying Price")



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OptionAnalyzerApp()
    window.showMaximized()
    sys.exit(app.exec())
