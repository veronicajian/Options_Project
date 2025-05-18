import sys
import os
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QSplitter, QLineEdit, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import shioaji as sj

# 設定中文字型
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

class OptionDataManager:
    def __init__(self):

        self.api = sj.Shioaji()

        self.api.login(
            api_key="", 
            secret_key="", 
        )

        Result = self.api.activate_ca(  
            ca_path=r"",
            ca_passwd="",
            person_id="",
        )

    def get_positions(self):
        positions = self.api.list_positions(self.api.futopt_account)
        positions_data = []
        if positions:
            for pos in positions:
                if pos.direction.value not in ['Sell', 'Buy']:
                    continue
                try:
                    contract = self.api.Contracts.Options.get(pos.code)
                except Exception as e:
                    print(f"無法取得合約 {pos.code}: {e}")
                    continue
                if not contract:
                    continue
                expiration_date = contract.delivery_date
                option_type = 'Call' if contract.symbol.endswith('C') else 'Put'
                strike_price = int(contract.strike_price)
                positions_data.append({
                    'expiration': expiration_date,
                    'strike': strike_price,
                    'type': option_type,
                    'action': pos.direction.value,
                    'quantity': pos.quantity,
                    'price': pos.price,
                    'source': '原始'
                })
        return positions_data

    def get_current_price(self):
        try:
            future_contract = self.api.Contracts.Futures.TXF.TXFR1
            snapshot = self.api.snapshots([future_contract])[0]
            future_price = snapshot.close
            return future_price
        except Exception as e:
            print(f"無法取得當前價格: {e}")
            return None

    def get_all_expirations(self):
        expirations = set()
        today = datetime.now().date()
        weekly_codes = ['TX1', 'TX2', 'TX4', 'TX5']
        for code in weekly_codes:
            if hasattr(self.api.Contracts.Options, code):
                code_contracts = getattr(self.api.Contracts.Options, code)
                for contract in code_contracts:
                    expiration_date = datetime.strptime(contract.delivery_date, '%Y/%m/%d').date()
                    if expiration_date >= today:
                        expirations.add(contract.delivery_date)
        txo_contracts = self.api.Contracts.Options.TXO
        for contract in txo_contracts:
            expiration_date = datetime.strptime(contract.delivery_date, '%Y/%m/%d').date()
            if expiration_date >= today:
                expirations.add(contract.delivery_date)
        formatted_expirations = sorted(expirations)
        return formatted_expirations

    def get_strike_prices_for_expiration(self, expiration):
        strike_prices = set()
        weekly_codes = ['TX1', 'TX2', 'TX4', 'TX5']
        for code in weekly_codes:
            if hasattr(self.api.Contracts.Options, code):
                code_contracts = getattr(self.api.Contracts.Options, code)
                for contract in code_contracts:
                    if contract.delivery_date == expiration:
                        strike_prices.add(int(contract.strike_price))
        txo_contracts = self.api.Contracts.Options.TXO
        for contract in txo_contracts:
            if contract.delivery_date == expiration:
                strike_prices.add(int(contract.strike_price))
        if strike_prices:
            return sorted(strike_prices)
        else:
            return []

    def get_contract_price(self, expiration, strike, opt_type):
        target_symbol = 'C' if opt_type == 'Call' else 'P'
        candidates = []
        weekly_codes = ['TX1', 'TX2', 'TX4', 'TX5']
        for code in weekly_codes:
            if hasattr(self.api.Contracts.Options, code):
                code_contracts = getattr(self.api.Contracts.Options, code)
                for c in code_contracts:
                    if c.delivery_date == expiration and int(c.strike_price) == strike and c.symbol.endswith(target_symbol):
                        candidates.append(c)
        txo_contracts = self.api.Contracts.Options.TXO
        for c in txo_contracts:
            if c.delivery_date == expiration and int(c.strike_price) == strike and c.symbol.endswith(target_symbol):
                candidates.append(c)
        if not candidates:
            return None
        contract = candidates[0]
        snapshot = self.api.snapshots([contract])[0]
        bid = snapshot.buy_price
        ask = snapshot.sell_price
        if bid and ask:
            return (bid+ask)/2.0
        else:
            return snapshot.close if snapshot.close else 0.0

    def close(self):
        self.api.logout()
        print("已登出Shioaji API")

class ProfitChartCanvas(FigureCanvasQTAgg):
    def __init__(self, parent=None):
        fig = Figure(figsize=(8,6), dpi=100)
        self.axes = fig.add_subplot(111)
        super().__init__(fig)
        self.setParent(parent)
        self.axes.set_title("損益圖")
        self.axes.set_xlabel("標的價格")
        self.axes.set_ylabel("盈虧 (NTD)")
        self.axes.grid(True)

    def plot_profit_curve(self, curves, current_price, max_profit, max_loss, breakeven_points):
        self.axes.clear()
        self.axes.set_title("損益圖")
        self.axes.set_xlabel("標的價格")
        self.axes.set_ylabel("盈虧 (NTD)")
        self.axes.grid(True)

        if not curves:
            self.draw()
            return

        # 繪製曲線
        for expiry, data in curves.items():
            self.axes.plot(data['x'], data['y'], label=expiry, color=data.get('color','blue'))

        # 標示目前價格的紅色點（不畫垂直線）
        if current_price is not None and 'Total' in curves:
            cp_y = np.interp(current_price, curves['Total']['x'], curves['Total']['y'])
            self.axes.plot(current_price, cp_y, 'ro')
            # 在紅色點附近加上文字
            self.axes.text(current_price, cp_y, f"目前:{current_price:.2f}", fontsize=10, ha='left', va='bottom', bbox=dict(facecolor='white', alpha=0.7))

        # 損益平衡點
        if breakeven_points:
            for bp in breakeven_points:
                self.axes.plot(bp, 0, 'go')
                self.axes.text(bp, 0, f"BE:{bp:.2f}", fontsize=9, color='green', ha='center', va='bottom', bbox=dict(facecolor='white', alpha=0.7))

        # 填充正負收益區域
        # 假設有'Total'曲線為主參考
        if 'Total' in curves:
            x = curves['Total']['x']
            y = curves['Total']['y']
            self.axes.fill_between(x, y, where=(y>=0), color='green', alpha=0.2)
            self.axes.fill_between(x, y, where=(y<=0), color='red', alpha=0.2)

        # 在圖的右下角顯示最大獲利/虧損資訊
        self.axes.text(0.95, 0.05,
                       f"最大獲利: {'無限' if np.isinf(max_profit) else f'{max_profit:.2f}'}\n"
                       f"最大虧損: {'無限' if np.isinf(max_loss) else f'{max_loss:.2f}'}",
                       transform=self.axes.transAxes,
                       ha='right', va='bottom', fontsize=10,
                       bbox=dict(facecolor='white', alpha=0.7))

        # 調整legend
        self.axes.legend(loc='upper left', fontsize=10)

        self.draw()

class OptionAnalyzerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("選擇權投組分析器")
        self.setGeometry(100, 100, 1600, 900)

        self.option_manager = OptionDataManager()
        self.original_positions = self.option_manager.get_positions()
        self.virtual_positions = []
        self.current_price = self.option_manager.get_current_price()
        self.expirations = self.option_manager.get_all_expirations()

        self.max_profit = 0
        self.max_loss = 0
        self.breakeven_points = []

        self.init_ui()

        # 初始繪製
        self.update_original_chart()
        self.update_adjusted_chart()

    def init_ui(self):
        main_layout = QHBoxLayout(self)

        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        main_layout.addWidget(splitter)

        # 左邊：原始投組圖 + 原始部位表
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        self.original_canvas = ProfitChartCanvas(left_widget)
        left_layout.addWidget(QLabel("原始投組損益圖"))
        left_layout.addWidget(self.original_canvas)

        self.original_table = QTableWidget()
        self.original_table.setColumnCount(6)
        self.original_table.setHorizontalHeaderLabels(["到期日", "類型", "動作", "履約價", "數量", "價格"])
        self.original_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.load_original_positions()
        left_layout.addWidget(QLabel("原始部位"))
        left_layout.addWidget(self.original_table)

        splitter.addWidget(left_widget)

        # 右邊：調整後圖 + 虛擬部位控制
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        # 選擇到期日顯示
        top_control_layout = QHBoxLayout()
        top_control_layout.addWidget(QLabel("選擇到期日圖表:"))
        self.expiry_filter_combo = QComboBox()
        self.expiry_filter_combo.addItem("總圖")
        for exp in self.expirations:
            self.expiry_filter_combo.addItem(exp)
        self.expiry_filter_combo.currentIndexChanged.connect(self.update_adjusted_chart)
        top_control_layout.addWidget(self.expiry_filter_combo)
        right_layout.addLayout(top_control_layout)

        self.adjusted_canvas = ProfitChartCanvas(right_widget)
        right_layout.addWidget(QLabel("調整後投組損益圖"))
        right_layout.addWidget(self.adjusted_canvas)

        # 控制面板 (新增/刪除虛擬部位)
        control_panel = QWidget()
        control_layout = QHBoxLayout(control_panel)

        control_layout.addWidget(QLabel("到期日:"))
        self.expiry_combo = QComboBox()
        self.expiry_combo.addItems(self.expirations)
        self.expiry_combo.currentIndexChanged.connect(self.update_strike_prices)
        control_layout.addWidget(self.expiry_combo)

        control_layout.addWidget(QLabel("類型:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Call", "Put"])
        self.type_combo.currentIndexChanged.connect(self.update_price_field)
        control_layout.addWidget(self.type_combo)

        control_layout.addWidget(QLabel("動作:"))
        self.action_combo = QComboBox()
        self.action_combo.addItems(["Sell", "Buy"])
        control_layout.addWidget(self.action_combo)

        control_layout.addWidget(QLabel("履約價:"))
        self.strike_combo = QComboBox()
        self.strike_combo.currentIndexChanged.connect(self.update_price_field)
        control_layout.addWidget(self.strike_combo)

        control_layout.addWidget(QLabel("價格:"))
        self.sell_price_input = QLineEdit()
        self.sell_price_input.setPlaceholderText("自動取得")
        control_layout.addWidget(self.sell_price_input)

        control_layout.addWidget(QLabel("數量:"))
        self.qty_input = QLineEdit()
        self.qty_input.setPlaceholderText("數量")
        control_layout.addWidget(self.qty_input)

        add_button = QPushButton("新增部位")
        add_button.clicked.connect(self.add_virtual_position)
        control_layout.addWidget(add_button)

        remove_button = QPushButton("刪除所選部位")
        remove_button.clicked.connect(self.remove_selected_virtual_position)
        control_layout.addWidget(remove_button)

        right_layout.addWidget(control_panel)

        self.virtual_table = QTableWidget()
        self.virtual_table.setColumnCount(7)
        self.virtual_table.setHorizontalHeaderLabels(["到期日", "類型", "動作", "履約價", "數量", "價格", "選擇"])
        self.virtual_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        right_layout.addWidget(QLabel("虛擬部位"))
        right_layout.addWidget(self.virtual_table)

        splitter.addWidget(right_widget)
        self.setLayout(main_layout)

        self.update_strike_prices()

    def load_original_positions(self):
        positions = self.original_positions
        self.original_table.setRowCount(0)
        for pos in positions:
            row = self.original_table.rowCount()
            self.original_table.insertRow(row)
            self.original_table.setItem(row, 0, QTableWidgetItem(pos['expiration']))
            self.original_table.setItem(row, 1, QTableWidgetItem(pos['type']))
            self.original_table.setItem(row, 2, QTableWidgetItem(pos['action']))
            self.original_table.setItem(row, 3, QTableWidgetItem(str(pos['strike'])))
            self.original_table.setItem(row, 4, QTableWidgetItem(str(pos['quantity'])))
            self.original_table.setItem(row, 5, QTableWidgetItem(f"{pos['price']:.2f}"))
            self.original_table.setRowHeight(row, 25)

    def load_virtual_positions(self):
        positions = self.virtual_positions
        self.virtual_table.setRowCount(0)
        for idx, pos in enumerate(positions):
            row = self.virtual_table.rowCount()
            self.virtual_table.insertRow(row)
            self.virtual_table.setItem(row, 0, QTableWidgetItem(pos['expiration']))
            self.virtual_table.setItem(row, 1, QTableWidgetItem(pos['type']))
            self.virtual_table.setItem(row, 2, QTableWidgetItem(pos['action']))
            self.virtual_table.setItem(row, 3, QTableWidgetItem(str(pos['strike'])))
            self.virtual_table.setItem(row, 4, QTableWidgetItem(str(pos['quantity'])))
            self.virtual_table.setItem(row, 5, QTableWidgetItem(f"{pos['price']:.2f}"))
            select_button = QPushButton("選擇")
            select_button.clicked.connect(lambda checked, i=idx: self.select_virtual_position(i))
            self.virtual_table.setCellWidget(row, 6, select_button)
            self.virtual_table.setRowHeight(row, 25)

    def update_strike_prices(self):
        expiration = self.expiry_combo.currentText()
        strikes = self.option_manager.get_strike_prices_for_expiration(expiration)
        self.strike_combo.clear()
        for s in strikes:
            self.strike_combo.addItem(str(s))
        self.update_price_field()

    def update_price_field(self):
        if self.strike_combo.count() == 0:
            return
        expiration = self.expiry_combo.currentText()
        opt_type = self.type_combo.currentText()
        strike_text = self.strike_combo.currentText()
        if not strike_text.isdigit():
            return
        strike = int(strike_text)
        price = self.option_manager.get_contract_price(expiration, strike, opt_type)
        if price is not None:
            self.sell_price_input.setText(f"{price:.2f}")

    def add_virtual_position(self):
        expiry = self.expiry_combo.currentText()
        opt_type = self.type_combo.currentText()
        action = self.action_combo.currentText()
        strike_text = self.strike_combo.currentText().strip()
        price_text = self.sell_price_input.text().strip()
        qty_text = self.qty_input.text().strip()

        if not strike_text.isdigit():
            QMessageBox.warning(self, "輸入錯誤", "請輸入正確的履約價。")
            return
        if not self.is_float(price_text):
            QMessageBox.warning(self, "輸入錯誤", "請輸入正確的價格。")
            return
        if not qty_text.isdigit():
            QMessageBox.warning(self, "輸入錯誤", "請輸入正確的數量。")
            return

        strike = int(strike_text)
        price = float(price_text)
        qty = int(qty_text)

        new_pos = {
            'expiration': expiry,
            'strike': strike,
            'type': opt_type,
            'action': action,
            'quantity': qty,
            'price': price,
            'source': '虛擬'
        }

        self.virtual_positions.append(new_pos)
        self.update_adjusted_chart()

    def remove_selected_virtual_position(self):
        selected_indexes = self.virtual_table.selectionModel().selectedRows()
        if not selected_indexes:
            QMessageBox.warning(self, "未選擇", "請先選擇要刪除的虛擬部位。")
            return
        for index in sorted(selected_indexes, reverse=True):
            row = index.row()
            self.virtual_positions.pop(row)
        self.update_adjusted_chart()

    def select_virtual_position(self, index):
        self.virtual_table.selectRow(index)

    def is_float(self, value):
        try:
            float(value)
            return True
        except:
            return False

    def calculate_pnl_curve(self, positions, price_range=None):

        # 新的X軸計算方式：依據所有部位的最高最低履約價決定
        all_strikes = [p['strike'] for p in positions] if positions else []
        if not all_strikes:
            return {}
        min_strike = min(all_strikes)
        max_strike = max(all_strikes)
        # x範圍
        low = min_strike * 0.99
        high = max_strike * 1.01
        price_range = np.linspace(low, high, 500)

        grouped = {}
        for pos in positions:
            expiry = pos['expiration']
            if expiry not in grouped:
                grouped[expiry] = []
            grouped[expiry].append(pos)

        contract_size = 50
        curves = {}
        total_y = np.zeros_like(price_range, dtype=float)
        # global_y_values = []
        colors = plt.cm.tab10(np.linspace(0,1,len(grouped)))

        for i, (expiry, pos_list) in enumerate(grouped.items()):
            y = np.zeros_like(price_range, dtype=float)
            for pos in pos_list:
                strike = pos['strike']
                qty = pos['quantity']
                price = pos['price']
                action = pos['action']
                if pos['type'] == 'Call':
                    if action == 'Sell':
                        payoff = (price*qty*contract_size) - np.maximum(price_range - strike, 0)*qty*contract_size
                    else:
                        payoff = (-price*qty*contract_size) + np.maximum(price_range - strike, 0)*qty*contract_size
                else: # Put
                    if action == 'Sell':
                        payoff = (price*qty*contract_size) - np.maximum(strike - price_range, 0)*qty*contract_size
                    else:
                        payoff = (-price*qty*contract_size) + np.maximum(strike - price_range, 0)*qty*contract_size
                y += payoff
            curves[expiry] = {'x': price_range, 'y': y, 'color': colors[i]}
            total_y += y
            # global_y_values.extend(y)


        if grouped:
            curves['Total'] = {'x': price_range, 'y': total_y, 'color': 'black'}
            # global_y_values.extend(total_y)

        # if global_y_values:
        #     max_profit = max(global_y_values)
        #     min_loss = min(global_y_values)

        if 'Total' in curves:
            max_profit = np.max(total_y)
            min_loss = np.min(total_y)

        else:
            max_profit = 0
            min_loss = 0

        # 無限判斷(簡化)
        if len(price_range) > 2 and 'Total' in curves:
            if total_y[-1] > total_y[-2] + 1000:
                max_profit = float('inf')
            if total_y[0] < total_y[1] - 1000:
                min_loss = float('-inf')

        # 損益平衡點
        breakevens = []
        if 'Total' in curves:
            total_arr = total_y
            for i in range(len(total_arr)-1):
                if total_arr[i]*total_arr[i+1] < 0:
                    x0 = price_range[i]
                    x1 = price_range[i+1]
                    y0 = total_arr[i]
                    y1 = total_arr[i+1]
                    bp = x0 - (y0*(x1 - x0)/(y1 - y0))
                    breakevens.append(bp)
        else:
            breakevens = []

        self.max_profit = max_profit
        self.max_loss = min_loss
        self.breakeven_points = breakevens

        return curves

    def update_original_chart(self):
        curves = self.calculate_pnl_curve(self.original_positions)
        self.original_canvas.plot_profit_curve(curves, self.current_price, self.max_profit, self.max_loss, self.breakeven_points)

    def update_adjusted_chart(self):
        combined_positions = self.original_positions + self.virtual_positions
        all_curves = self.calculate_pnl_curve(combined_positions)

        selected_exp = self.expiry_filter_combo.currentText()
        if selected_exp == "總圖":
            filtered_curves = all_curves
            show_max_profit = self.max_profit
            show_max_loss = self.max_loss
            show_bes = self.breakeven_points
        else:
            if selected_exp in all_curves:
                x = all_curves[selected_exp]['x']
                y = all_curves[selected_exp]['y']
                max_p = max(y) if len(y)>0 else 0
                min_p = min(y) if len(y)>0 else 0
                if len(x)>2:
                    if y[-1] > y[-2] + 1000:
                        max_p = float('inf')
                    if y[0] < y[1] - 1000:
                        min_p = float('-inf')
                bes = []
                for i in range(len(y)-1):
                    if y[i]*y[i+1]<0:
                        x0,x1 = x[i],x[i+1]
                        y0,y1 = y[i],y[i+1]
                        bp = x0 - (y0*(x1 - x0)/(y1 - y0))
                        bes.append(bp)
                filtered_curves = {selected_exp: all_curves[selected_exp]}
                show_max_profit = max_p
                show_max_loss = min_p
                show_bes = bes
            else:
                filtered_curves = {}
                show_max_profit = 0
                show_max_loss = 0
                show_bes = []

        self.adjusted_canvas.plot_profit_curve(filtered_curves, self.current_price, show_max_profit, show_max_loss, show_bes)
        self.load_virtual_positions()

    def closeEvent(self, event):
        self.option_manager.close()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OptionAnalyzerApp()
    window.show()
    sys.exit(app.exec())
