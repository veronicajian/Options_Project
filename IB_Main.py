#!/usr/bin/env python3
import sys
import time
import threading
import mysql.connector
import pandas as pd
from datetime import datetime, timedelta, date

# ---- PyQt6 ----
from PyQt6.QtCore import Qt, QThread, QEventLoop, pyqtSignal, pyqtSlot, QTimer, QSize
from PyQt6.QtGui import QColor, QPalette, QFont ,QMovie, QPixmap
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPlainTextEdit, QDialog,
    QVBoxLayout, QHBoxLayout, QGridLayout, QFrame, QMessageBox,
    QComboBox, QLineEdit, QPushButton, QLabel,
    QSpinBox, QTableWidget, QTableWidgetItem, QSplitter, QGroupBox, QHeaderView
)

# ---- IB API ----
from ibapi.client import EClient
from ibapi.wrapper import EWrapper
from ibapi.contract import Contract
from ibapi.order import Order


def now_str():
    return time.strftime("[%H:%M:%S]")


def sep_line():
    return "-" * 50


# =============================================================================
#   ==========【新增】抓取 & 儲存選擇權合約函式 (移除60天限制)==========
# =============================================================================
def fetch_and_store_option_chain(symbol):
    """
    與 IB TWS 連線，抓取該 symbol 的選擇權合約(全部到期日 & 履約價)，
    並插入至 MySQL 資料庫(若無該symbol資料表，則自動建立)。
    """

    from ibapi.client import EClient
    from ibapi.wrapper import EWrapper
    from ibapi.contract import Contract
    import sys

    class OptionOrderDatabaseManager:
        def __init__(self, host='localhost', user='root', password='', database='IBOption_DATA'):
            self.host = host
            self.user = user
            self.password = password
            self.database = database
            self.conn = None
            self.cursor = None
            self.connect_db()

        def connect_db(self):
            try:
                self.conn = mysql.connector.connect(
                    host=self.host,
                    user=self.user,
                    password=self.password
                )
                self.cursor = self.conn.cursor(dictionary=True)
                self.cursor.execute("SHOW DATABASES LIKE %s", (self.database,))
                result = self.cursor.fetchone()
                if not result:
                    self.cursor.execute(f"CREATE DATABASE `{self.database}`")
                    print(f"[DB] 資料庫 `{self.database}` 已建立。")
                self.conn.database = self.database
                print(f"[DB] 已連線到資料庫 `{self.database}`。")
            except mysql.connector.Error as err:
                print(f"[DB Error] 連線失敗: {err}")
                self.conn = None
                self.cursor = None

        def close_db(self):
            if self.cursor:
                self.cursor.close()
            if self.conn:
                self.conn.close()
                print("[DB] 資料庫連線已關閉。")

        def format_table_name(self, symbol):
            return symbol.upper()

        def create_option_chain_table_if_not_exists(self, symbol):
            table_name = self.format_table_name(symbol)
            create_table_query = f"""
            CREATE TABLE IF NOT EXISTS `{table_name}` (
                id INT AUTO_INCREMENT PRIMARY KEY,
                expiry DATE,
                strike DECIMAL(10,2),
                UNIQUE(expiry, strike)
            );
            """
            try:
                self.cursor.execute(create_table_query)
                self.conn.commit()
                print(f"[DB] 確認資料表 `{table_name}` 已存在或已建立。")
            except mysql.connector.Error as err:
                print(f"[DB Error] 建立資料表 `{table_name}` 失敗: {err}")

        def get_existing_expiries(self, symbol):
            table_name = self.format_table_name(symbol)
            query = f"SELECT DISTINCT expiry FROM `{table_name}`"
            try:
                self.cursor.execute(query)
                rows = self.cursor.fetchall()
                return {row['expiry'] for row in rows}
            except mysql.connector.Error as err:
                print(f"[DB Error] 查詢已存在到期日失敗: {err}")
                return set()

        def insert_option_chain_data(self, symbol, expiry, strikes):
            table_name = self.format_table_name(symbol)
            insert_query = f"""
            INSERT IGNORE INTO `{table_name}` (expiry, strike)
            VALUES (%s, %s)
            """
            data = sorted([(expiry, strike) for strike in strikes])
            try:
                self.cursor.executemany(insert_query, data)
                self.conn.commit()
                print(f"[DB] 插入 {len(data)} 條資料到 `{table_name}`（到期日: {expiry}）。")
            except mysql.connector.Error as err:
                print(f"[DB Error] 插入資料到 `{table_name}` 失敗: {err}")


    class FetchAndStoreApp(EWrapper, EClient):
        def __init__(self, symbol, db_manager):
            EClient.__init__(self, self)
            self.symbol = symbol
            self.db_manager = db_manager
            self.contract_details_data = []
            self.done = False

        def error(self, reqId, errorCode, errorString, advancedOrderRejectJson=""):
            if errorCode not in (2104, 2106, 2158, 2108):
                print(f"[Error] reqId={reqId}, code={errorCode}, msg={errorString}")

        def nextValidId(self, orderId: int):
            contract = Contract()
            contract.symbol = self.symbol
            contract.secType = "OPT"
            contract.currency = "USD"
            contract.exchange = "SMART"
            self.reqContractDetails(1, contract)

        def contractDetails(self, reqId, contractDetails):
            self.contract_details_data.append(contractDetails)

        def contractDetailsEnd(self, reqId):
            print(f"[Info] contractDetailsEnd for reqId={reqId}")
            self.handle_option_chain()
            self.done = True
            self.disconnect()

        def handle_option_chain(self):
            chain_map = {}
            for cd in self.contract_details_data:
                c = cd.contract
                exp_str = c.lastTradeDateOrContractMonth
                strike = c.strike
                if exp_str and strike > 0:
                    try:
                        exp_date = datetime.strptime(exp_str, "%Y%m%d").date()

                        # === 這裡移除原先 "只抓未來60天" 的限制 ===
                        # if exp_date <= datetime.now().date() + timedelta(days=60):
                        chain_map.setdefault(exp_date, set()).add(strike)

                    except ValueError:
                        print(f"[Warning] 無法解析到期日格式: {exp_str}")

            if not chain_map:
                print(f"[Info] {self.symbol}: 沒有符合條件的期權合約資料。")
                return

            # 建立資料表（若不存在）
            self.db_manager.create_option_chain_table_if_not_exists(self.symbol)
            # 查詢已存在的到期日
            existing_exps = self.db_manager.get_existing_expiries(self.symbol)
            print(f"[DB] 已存在的到期日: {sorted(existing_exps)}")

            # 插入新的到期日和履約價
            for exp_date, strikes in sorted(chain_map.items()):
                if exp_date not in existing_exps:
                    self.db_manager.insert_option_chain_data(
                        symbol=self.symbol,
                        expiry=exp_date,
                        strikes=strikes
                    )
                else:
                    print(f"[DB] 到期日 {exp_date} 已存在，跳過。")


    # 初始化資料庫管理器
    db_manager = OptionOrderDatabaseManager(
        host='192.168.1.107',
        user='root',
        password='!Yuchen8888',
        database='IBOption_DATA'
    )

    if db_manager.conn is None:
        print("[Error] 無法連線到資料庫，終止程式。")
        sys.exit(1)

    # 建立 App 實例
    app = FetchAndStoreApp(symbol, db_manager)

    print(f"[Info] 正在連線 IB TWS...")
    app.connect("127.0.0.1", 7497, clientId=18)

    start_time = time.time()
    try:
        app.run()
    except KeyboardInterrupt:
        print("\n[Info] 手動中止程式")
        app.done = True
        app.disconnect()
    end_time = time.time()

    # 關閉資料庫連線
    db_manager.close_db()

    print(f"[Info] {symbol} 的Option Chain處理結束，總耗時 {end_time - start_time:.2f} 秒")


# =============================================================================
#   IB Worker
# =============================================================================
class IBWrapper(EWrapper):
    def __init__(self):
        super().__init__()
        self.accounts = []
        self.accountValues = {}
        self._workerRef = None

        self.interested_keys = {
            "NetLiquidation",
            "TotalCashValue",
            "AccruedCash",
            "BuyingPower",
            "InitMarginReq",
            "MaintMarginReq",
            "AvailableFunds",
            "ExcessLiquidity",
            "SMA"
        }

    def managedAccounts(self, accountsList: str):
        self.accounts = [a.strip() for a in accountsList.split(',') if a.strip()]
        if self._workerRef:
            self._workerRef.onManagedAccountsReceived(self.accounts)
            for acct in self.accounts:
                self._workerRef.client.reqAccountUpdates(True, acct)

    def updateAccountValue(self, key, val, currency, accountName):
        if key in self.interested_keys:
            self.accountValues[key] = val
            if self._workerRef:
                self._workerRef.onAccountValueUpdate(self.accountValues)

    def nextValidId(self, orderId: int):
        if self._workerRef:
            self._workerRef.onNextValidId(orderId)


class IBClient(EClient):
    def __init__(self, wrapper):
        super().__init__(wrapper)


class IBApiWorker(QThread):
    accountsSignal = pyqtSignal(list)
    accountValueSignal = pyqtSignal(dict)

    def __init__(self, host='127.0.0.1', port=7497, clientId=999):
        super().__init__()
        self.host = host
        self.port = port
        self.clientId = clientId
        self.wrapper = IBWrapper()
        self.wrapper._workerRef = self
        self.client = IBClient(self.wrapper)
        self.nextOrderId = None

    def run(self):
        self.client.connect(self.host, self.port, self.clientId)
        self.client.run()

    def stop(self):
        self.client.disconnect()

    def onManagedAccountsReceived(self, accounts):
        self.accountsSignal.emit(accounts)

    def onAccountValueUpdate(self, acctVals: dict):
        self.accountValueSignal.emit(acctVals)

    def onNextValidId(self, oid):
        self.nextOrderId = oid

    def place_mkt_order(self, contract: Contract, action: str, qty: float):
        if self.nextOrderId is None:
            time.sleep(1)
        order_id = self.nextOrderId
        self.nextOrderId += 1

        order = Order()
        order.action = action
        order.orderType = "MKT"
        order.totalQuantity = qty
        order.transmit = True  # 真正下單
        order.eTradeOnly = False
        order.firmQuoteOnly = False
        self.client.placeOrder(order_id, contract, order)
        print(f"已發送 {action}市價單{qty}口")
        return order_id


# =============================================================================
#   FetchCurrentPrice
# =============================================================================
class PriceCalcAPI(EWrapper, EClient):
    def __init__(self):
        EWrapper.__init__(self)
        EClient.__init__(self, wrapper=self)
        self.latest_price = None
        self.lock = threading.Lock()
        self.symbol_map = {}

    def tickPrice(self, reqId, tickType, price, attrib):
        if tickType in [4, 68, 75, 76] and price > 0:
            print(f"[tickPrice] reqId: {reqId}, TickType: {tickType}, Price: {price}")
            with self.lock:
                if self.latest_price is None or self.latest_price <= 0:
                    self.latest_price = price


class PriceFetchThread(QThread):
    priceFetched = pyqtSignal(float)

    def __init__(self, host, port, symbol, timeout=10, parent=None):
        super().__init__(parent)
        self.host = host
        self.port = port
        self.symbol = symbol
        self.timeout = timeout

    def run(self):
        try:
            price = self.get_latest_price(self.host, self.port, self.symbol, self.timeout)
            if price is not None:
                self.priceFetched.emit(price)
        except Exception as e:
            print(f'取得價格失敗或超時:{e}')

    def get_latest_price(self, host, port, symbol, timeout):
        app = PriceCalcAPI()
        try:
            app.connect(host, port, '88')
        except Exception as e:
            print(f"連線失敗: {e}")
            return None

        def run_loop():
            app.run()

        api_thread = threading.Thread(target=run_loop, daemon=True)
        api_thread.start()

        # 等待連線建立
        time.sleep(1)

        # 建立合約資料
        contract = Contract()
        contract.symbol = symbol
        contract.secType = "STK"
        contract.exchange = 'SMART'
        contract.currency = "USD"

        req_id = 1
        app.symbol_map[req_id] = symbol

        # 設定市場數據類型，3 為實時數據
        app.reqMarketDataType(3)
        app.reqMktData(req_id, contract, '', False, False, [])

        # 等待價格回傳
        start_time = time.time()
        price = None
        while True:
            with app.lock:
                if app.latest_price is not None:
                    price = app.latest_price
                    break
            if time.time() - start_time > timeout:
                price = None
                break
            time.sleep(0.1)

        app.disconnect()
        api_thread.join()
        return price
    
# =============================================================================
#   取得未平倉資訊
# =============================================================================

class PositionAPI(EWrapper, EClient):
    def __init__(self):
        EWrapper.__init__(self)
        EClient.__init__(self, wrapper=self)
        self.positions = [] 
        self.lock = threading.Lock()
        self.positions_complete = False 
        
    def position(self, account: str, contract, position: float, avgCost: float):
        with self.lock:
            if contract.secType == "OPT":
                self.positions.append((contract, position))

    def positionEnd(self):
        with self.lock:
            self.positions_complete = True
        
class PositionFetchThread(QThread):
        positionsFetched = pyqtSignal(list)  # 發射包含合約物件的列表

        def __init__(self, host, port, timeout=10, parent=None):
            super().__init__(parent)
            self.host = host
            self.port = port
            self.timeout = timeout

        def run(self):
            try:
                positions = self.get_latest_positions(self.host, self.port, self.timeout)
                if positions:
                        self.positionsFetched.emit(positions) # 發射合約列表
            except Exception as e:
                print(f'取得合約失敗或超時: {e}')

        def get_latest_positions(self, host, port, timeout):
            app = PositionAPI()
            try:
                app.connect(host, port, 88)
            except Exception as e:
                print(f"連線失敗: {e}")
                return []

            def run_loop():
                app.run()

            api_thread = threading.Thread(target=run_loop, daemon=True)
            api_thread.start()

            time.sleep(1)  

            # 查詢帳戶持倉
            app.reqPositions()

            start_time = time.time()
            while True:
                with app.lock:
                    if app.positions_complete: 
                        break
                if time.time() - start_time > timeout:
                    print("超時未收到所有合約回應")
                    break
                time.sleep(0.1)

            app.disconnect()
            api_thread.join()
            return app.positions  

# =============================================================================
#   DB Manager
# =============================================================================
class OptionOrderDatabaseManager:
    def __init__(self, host='localhost', user='root', password='', database='IBOption_DATA'):
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        self.conn = None
        self.cursor = None
        self.connect_db()

    def connect_db(self):
        try:
            self.conn = mysql.connector.connect(host=self.host, user=self.user, password=self.password)
            self.cursor = self.conn.cursor(dictionary=True)
            self.cursor.execute("SHOW DATABASES LIKE %s", (self.database,))
            if not self.cursor.fetchone():
                self.cursor.execute(f"CREATE DATABASE `{self.database}`")
            self.conn.database = self.database
            self.create_option_orders_table_if_not_exists()
        except mysql.connector.Error as err:
            print(f"[DB Error] 連線失敗: {err}")
            self.conn = None
            self.cursor = None

    def close_db(self):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()

    def format_table_name(self, symbol):
        return symbol.upper()

    def table_exists(self, symbol):
        if not self.conn or not self.cursor:
            return False
        t = self.format_table_name(symbol)
        self.cursor.execute(f"SHOW TABLES LIKE '{t}'")
        return bool(self.cursor.fetchone())

    def fetch_expiries(self, symbol):
        t = self.format_table_name(symbol)
        if not self.conn or not self.cursor:
            return []
        self.cursor.execute(f"SHOW TABLES LIKE '{t}'")
        if not self.cursor.fetchone():
            return []
        q = f"SELECT DISTINCT expiry FROM `{t}` ORDER BY expiry ASC"
        self.cursor.execute(q)
        rows = self.cursor.fetchall()
        return [r['expiry'] for r in rows] if rows else []

    def fetch_strikes(self, symbol, expiry):
        t = self.format_table_name(symbol)
        if not self.conn or not self.cursor:
            return []
        self.cursor.execute(f"SHOW TABLES LIKE '{t}'")
        if not self.cursor.fetchone():
            return []
        q = f"SELECT DISTINCT strike FROM `{t}` WHERE expiry=%s ORDER BY strike ASC"
        self.cursor.execute(q, (expiry,))
        rows = self.cursor.fetchall()
        return [r['strike'] for r in rows] if rows else []
    
    def fetch_active_orders(self):
        q = """
        SELECT * FROM option_orders
        WHERE status = 'Active' 
        ORDER BY order_time DESC
        """
        self.cursor.execute(q)
        orders = self.cursor.fetchall()

        today = datetime.now().date()  

        for od in orders:
            expiry = od.get("expiry")
            order_id = od.get("id")

            if expiry:
                if isinstance(expiry, datetime):
                    expiry_date = expiry.date()  
                elif isinstance(expiry, str):
                    expiry_date = datetime.strptime(expiry, "%Y-%m-%d").date() 
                else:  
                    expiry_date = expiry  

                if expiry_date < today:
                    update_q = """
                    UPDATE option_orders
                    SET status = 'Closed'
                    WHERE id = %s
                    """
                    self.cursor.execute(update_q,(order_id,))
                    self.conn.commit()
        return orders

    def create_option_orders_table_if_not_exists(self):
        # 移除 price 欄位
        q = """
        CREATE TABLE IF NOT EXISTS option_orders (
            id INT AUTO_INCREMENT PRIMARY KEY,
            order_time DATETIME,
            expiry DATE,
            symbol VARCHAR(50),
            strike FLOAT,
            action VARCHAR(10),
            totalQuantity FLOAT,
            win_rate FLOAT,
            prob_less FLOAT,
            prob_greater FLOAT,
            exec_prob FLOAT,
            status VARCHAR(20),
            usagePct FLOAT
        )"""
        self.cursor.execute(q)
        self.conn.commit()

    def insert_option_order(self, order_data: dict):
        q = """
        INSERT INTO option_orders (
            order_time, expiry, symbol, action, totalQuantity, strike,
            win_rate, prob_less, prob_greater, exec_prob, status, usagePct
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        vals = (
            order_data.get("order_time"),
            order_data.get("expiry"),
            order_data.get("symbol"),
            order_data.get("action"),
            order_data.get("totalQuantity"),
            order_data.get("strike"),
            order_data.get("win_rate"),
            order_data.get("prob_less"),
            order_data.get("prob_greater"),
            order_data.get("exec_prob"),
            order_data.get("status"),
            order_data.get("usagePct")
        )
        self.cursor.execute(q, vals)
        self.conn.commit()

    def update_exec_prob(self, id, new_prob):
        try:
            # 查詢是否存在對應的 id
            select_query = "SELECT * FROM option_orders WHERE id = %s"
            self.cursor.execute(select_query, (id,))
            record = self.cursor.fetchone()

            if not record:
                print(f"[警告] ID 為 {id} 的資料不存在。")
                return

            # 更新 exec_prob 欄位
            update_query = "UPDATE option_orders SET exec_prob = %s WHERE id = %s"
            self.cursor.execute(update_query, (new_prob, id))
            self.conn.commit()
            print(f"成功更新 ID 為 {id} 的 exec_prob 為 {new_prob}。")
        except mysql.connector.Error as err:
            print(f"[DB Error] 無法更新 exec_prob: {err}")

    def update_option_order(self, target):
        select_query = """
        SELECT * FROM option_orders
        WHERE symbol = %s AND status = 'Active'
        ORDER BY id ASC
        LIMIT 1
        """
        self.cursor.execute(select_query, (target,))
        result = self.cursor.fetchone()  # 取得第一筆資料

        if result:
            order_id = result['id']
            update_query = """
            UPDATE option_orders
            SET status = 'Closed'
            WHERE id = %s
            """
            self.cursor.execute(update_query, (order_id,))
            self.conn.commit()
            print(f"成功更新訂單 ID: {order_id} 的狀態為 'Closed'")
        else:
            print(f"沒有找到符合條件的訂單 (target={target}, status=Active)")

    def update_multiple_fields(self, symbol, action, qty, expiry, updated_data):
        try:
            set_clause = ", ".join([f"{key} = %s" for key in updated_data.keys()])
            query = f"""
            UPDATE option_orders
            SET {set_clause}
            WHERE symbol = %s AND action = %s AND totalQuantity = %s AND expiry = %s
            """
            values = list(updated_data.values()) + [symbol, action, qty, expiry]
            self.cursor.execute(query, values)
            self.conn.commit()
            print(f"成功更新資料庫: {updated_data}, Symbol = {symbol}, Action = {action}, Qty = {qty}, Expiry = {expiry}")
        except mysql.connector.Error as err:
            print(f"[DB Error] 更新多個欄位時發生錯誤: {err}")


# =============================================================================
#   StrategyDialog
# =============================================================================
class StrategyDialog(QDialog):
    def __init__(self, parent, symbol, expiry, callStrike, putStrike,
                 winRate, probLo, probHi, usageAmt,
                 call_contract_value, call_allowed, put_contract_value, put_allowed):
        super().__init__(parent)
        self.setWindowTitle("策略建倉試算")
        self.resize(700, 500)

        self.parentWin = parent
        self.symbol = symbol
        self.expiry = expiry
        self.callStrike = callStrike
        self.putStrike = putStrike
        self.winRate = winRate
        self.probLo = probLo
        self.probHi = probHi
        self.usageAmtForCall = usageAmt / 2
        self.usageAmtForPut = usageAmt / 2

        self.call_contract_value = call_contract_value
        self.call_allowed = call_allowed
        self.put_contract_value = put_contract_value
        self.put_allowed = put_allowed

        mainLayout = QVBoxLayout(self)

        self.infoLabel = QLabel(
            f"<h2>{self.symbol} 到期 {self.expiry}</h2>"
            f"<h3>CALL Strike: {self.callStrike}   PUT Strike: {self.putStrike}</h3>"
            f"<h3>勝率: {self.winRate}%   機率範圍: < {self.probLo}%, > {self.probHi}%</h3>"
        )
        mainLayout.addWidget(self.infoLabel)

        # CALL 區塊
        callBox = QFrame()
        callLayout = QVBoxLayout(callBox)
        lbl_callTitle = QLabel("<b>[CALL] 賣出看漲期權</b>")
        callLayout.addWidget(lbl_callTitle)

        self.lbl_callInfo = QLabel(
            f"<font size='4'>合約價值: {self.call_contract_value:.2f}，可下口數: {self.call_allowed}</font>"
        )
        callLayout.addWidget(self.lbl_callInfo)

        self.spin_callLot = QSpinBox()
        self.spin_callLot.setRange(1, self.call_allowed)
        self.spin_callLot.setValue(self.call_allowed)
        btn_call = QPushButton("建倉(CALL)")
        btn_call.clicked.connect(self.onCallBuild)
        hlayout_call = QHBoxLayout()
        hlayout_call.addWidget(QLabel("口數:"))
        hlayout_call.addWidget(self.spin_callLot)
        hlayout_call.addWidget(btn_call)
        callLayout.addLayout(hlayout_call)

        # PUT 區塊
        putBox = QFrame()
        putLayout = QVBoxLayout(putBox)
        lbl_putTitle = QLabel("<b>[PUT] 賣出看跌期權</b>")
        putLayout.addWidget(lbl_putTitle)

        self.lbl_putInfo = QLabel(
            f"<font size='4'>合約價值: {self.put_contract_value:.2f}，可下口數: {self.put_allowed}</font>"
        )
        putLayout.addWidget(self.lbl_putInfo)

        self.spin_putLot = QSpinBox()
        self.spin_putLot.setRange(1, self.put_allowed)
        self.spin_putLot.setValue(self.put_allowed)
        btn_put = QPushButton("建倉(PUT)")
        btn_put.clicked.connect(self.onPutBuild)
        hlayout_put = QHBoxLayout()
        hlayout_put.addWidget(QLabel("口數:"))
        hlayout_put.addWidget(self.spin_putLot)
        hlayout_put.addWidget(btn_put)
        putLayout.addLayout(hlayout_put)

        mainLayout.addWidget(callBox)
        mainLayout.addWidget(putBox)

    def updateData(self, callStrike, putStrike,
                   call_contract_value, call_allowed,
                   put_contract_value, put_allowed):
        self.callStrike = callStrike
        self.putStrike = putStrike
        self.call_contract_value = call_contract_value
        self.call_allowed = call_allowed
        self.put_contract_value = put_contract_value
        self.put_allowed = put_allowed

        self.infoLabel.setText(
            f"<h2>{self.symbol} 到期 {self.expiry}</h2>"
            f"<h3>CALL Strike: {self.callStrike}   PUT Strike: {self.putStrike}</h3>"
            f"<h3>勝率: {self.winRate}%   機率範圍: < {self.probLo}%, > {self.probHi}%</h3>"
        )

        self.lbl_callInfo.setText(
            f"<font size='4'>合約價值: {self.call_contract_value:.2f}，可下口數: {self.call_allowed}</font>"
        )
        self.spin_callLot.setRange(1, self.call_allowed)
        self.spin_callLot.setValue(self.call_allowed)

        self.lbl_putInfo.setText(
            f"<font size='4'>合約價值: {self.put_contract_value:.2f}，可下口數: {self.put_allowed}</font>"
        )
        self.spin_putLot.setRange(1, self.put_allowed)
        self.spin_putLot.setValue(self.put_allowed)

    def onCallBuild(self):
        lot = self.spin_callLot.value()
        msg = f"[CALL下單] {self.symbol}, {self.expiry}, strike={self.callStrike}, 口數={lot}"
        self.parentWin.logTextEdit.appendPlainText(f"{now_str()} {msg}\n{sep_line()}")

        c = Contract()
        c.symbol = self.symbol
        c.secType = "OPT"
        c.exchange = "SMART"
        c.currency = "USD"
        c.lastTradeDateOrContractMonth = self.expiry.strftime("%Y%m%d")
        c.strike = float(self.callStrike)
        c.right = "C"

        order_id = self.parentWin.ibWorker.place_mkt_order(c, "SELL", lot)
        self.parentWin.logTextEdit.appendPlainText(f"{now_str()} [已送出市價單(Call)] OrderId={order_id}\n{sep_line()}")

        od = {
            "order_time": datetime.now(),
            "expiry": self.expiry,
            "symbol": f"{self.symbol} C {self.callStrike}",
            "action": "SELL",
            "totalQuantity": lot,
            "strike": float(self.callStrike),
            "win_rate": float(self.parentWin.line_winRate.text().strip() or 0),
            "prob_less": float(self.parentWin.line_probLo.text().strip() or 0),
            "prob_greater": float(self.parentWin.line_probHi.text().strip() or 0),
            "exec_prob": self.parentWin.call_exec_prob if hasattr(self.parentWin, 'call_exec_prob') else 0,
            "status": 'Active',
            "usagePct": float(self.parentWin.line_usagePct.text().strip() or 0)
        }
        self.parentWin.dbmgr.insert_option_order(od)
        self.parentWin.refreshOrderTables()

    def onPutBuild(self):
        lot = self.spin_putLot.value()
        msg = f"[PUT下單] {self.symbol}, {self.expiry}, strike={self.putStrike}, 口數={lot}"
        self.parentWin.logTextEdit.appendPlainText(f"{now_str()} {msg}\n{sep_line()}")

        c = Contract()
        c.symbol = self.symbol
        c.secType = "OPT"
        c.exchange = "SMART"
        c.currency = "USD"
        c.lastTradeDateOrContractMonth = self.expiry.strftime("%Y%m%d")
        c.strike = float(self.putStrike)
        c.right = "P"

        order_id = self.parentWin.ibWorker.place_mkt_order(c, "SELL", lot)
        self.parentWin.logTextEdit.appendPlainText(f"{now_str()} [已送出市價單(PUT)] OrderId={order_id}\n{sep_line()}")

        od = {
            "order_time": datetime.now(),
            "expiry": self.expiry,
            "symbol": f"{self.symbol} P {self.putStrike}",
            "action": "SELL",
            "totalQuantity": lot,
            "strike": float(self.putStrike),
            "win_rate": float(self.parentWin.line_winRate.text().strip() or 0),
            "prob_less": float(self.parentWin.line_probLo.text().strip() or 0),
            "prob_greater": float(self.parentWin.line_probHi.text().strip() or 0),
            "exec_prob": self.parentWin.put_exec_prob if hasattr(self.parentWin, 'put_exec_prob') else 0,
            "status": 'Active',
            "usagePct": float(self.parentWin.line_usagePct.text().strip() or 0)
        }
        self.parentWin.dbmgr.insert_option_order(od)
        self.parentWin.refreshOrderTables()


# =============================================================================
#   ==========【NEW】執行 fetch_and_store_option_chain 的執行緒==========
# =============================================================================
class FetchOptionChainThread(QThread):
    finishedSignal = pyqtSignal(bool)  # 用 True/False 表示是否有新到期日

    def __init__(self, symbol, parent=None):
        super().__init__(parent)
        self.symbol = symbol

    def run(self):
        # 1) 先取舊的 expiry
        old_exps = set(self.get_current_expiries(self.symbol))

        # 2) 執行抓取 (這裡會去抓「全部到期日 & 履約價」)
        try:
            fetch_and_store_option_chain(self.symbol)
        except Exception as e:
            print("執行 fetch_and_store_option_chain 發生錯誤:", e)
            self.finishedSignal.emit(False)
            return

        # 3) 再取新的 expiry
        new_exps = set(self.get_current_expiries(self.symbol))

        # 如果 new_exps - old_exps 有資料 => 表示有新增到期日
        if len(new_exps - old_exps) > 0:
            self.finishedSignal.emit(True)
        else:
            self.finishedSignal.emit(False)

    def get_current_expiries(self, sym):
        db = OptionOrderDatabaseManager(
            host='192.168.1.107',
            user='root',
            password='!Yuchen8888',
            database='IBOption_DATA'
        )
        exps = db.fetch_expiries(sym)
        db.close_db()
        return exps


# =============================================================================
#   MainWindow
# =============================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IB")
        self.resize(1400, 800)
        self.initializing = True  # 初始化標誌

        # IB Worker
        self.ibWorker = IBApiWorker(clientId=999)
        self.ibWorker.accountsSignal.connect(self.onAccountsReady)
        self.ibWorker.accountValueSignal.connect(self.onAccountValueUpdate)
        self.ibWorker.start()

        # DB
        self.dbmgr = OptionOrderDatabaseManager(
            host='192.168.1.107',
            user='root',
            password='!Yuchen8888',
            database='IBOption_DATA'
        )

        self.important_tags = {
            "NetLiquidation": "",
            "TotalCashValue": "",
            "AccruedCash": "",
            "BuyingPower": "",
            "InitMarginReq": "",
            "MaintMarginReq": "",
            "AvailableFunds": "",
            "ExcessLiquidity": "",
            "SMA": ""
        }
        self.threads = []
        self.qqq_daily_data = None

        try:
            df = pd.read_csv(
                r"C:\Users\user\Desktop\folder\QQQ_daily_data.csv",
                parse_dates=["Date"]
            )
            df.sort_values("Date", inplace=True)
            df.reset_index(drop=True, inplace=True)
            if "daily_return" not in df.columns:
                df["daily_return"] = df["Close"].pct_change()
            self.qqq_daily_data = df
            print("已載入 QQQ 日線資料，共", len(self.qqq_daily_data), "筆")
        except Exception as e:
            print("讀取QQQ失敗:", e)

        # 偵測功能相關屬性
        self.detect_timer = QTimer(self)
        self.countdown = 30 # 30秒倒數
        self.detect_timer.setInterval(1000)
        self.detect_timer.timeout.connect(self.onTimerTick)
        self.detection_running = False

        self._initUI()
        self._fillAccountInfo()
        self.refreshOrderTables()

        # 取得現貨價
        self.priceThread = PriceFetchThread('127.0.0.1', 7497, 'QQQ', timeout=10)
        self.priceThread.priceFetched.connect(self.onPriceFetched_update_label)
        self.priceThread.start()
        
        # 完成初始化
        self.initializing = False  

        # 執行持倉部位線程
        self.positionThread = PositionFetchThread('127.0.0.1', 7497, timeout=10)
        self.positionThread.positionsFetched.connect(self.onPositionsFetched)
        self.positionThread.start()

    def _initUI(self):
        central = QWidget()
        self.setCentralWidget(central)
        mainLayout = QVBoxLayout(central)

        topFrame = QFrame()
        topLayout = QHBoxLayout(topFrame)
        mainLayout.addWidget(topFrame, stretch=2)

        bottomFrame = QFrame()
        bottomLayout = QHBoxLayout(bottomFrame)
        mainLayout.addWidget(bottomFrame, stretch=8)

        self.topLeft = QFrame()
        self.topLeftLayout = QVBoxLayout(self.topLeft)

        self.topCenter = QFrame()
        self.topCenterLayout = QVBoxLayout(self.topCenter)

        self.topRight = QFrame()
        self.topRightLayout = QVBoxLayout(self.topRight)

        topLayout.addWidget(self.topLeft, stretch=2)
        topLayout.addWidget(self.topCenter, stretch=5)
        topLayout.addWidget(self.topRight, stretch=3)

        lbl_account = QLabel("請選擇帳戶：")
        self.combo_account = QComboBox()
        f1 = QFrame()
        l1 = QVBoxLayout(f1)
        l1.addWidget(lbl_account)
        l1.addWidget(self.combo_account)
        self.topLeftLayout.addWidget(f1)

        lbl_symbol = QLabel("請輸入標的：")
        self.line_symbol = QLineEdit()
        self.line_symbol.setText("QQQ")

        self.btn_search = QPushButton("搜尋")
        self.btn_search.clicked.connect(self.onSearchClicked)

        # === NEW === 新增按鈕「太晶化噴火龍」
        self.btn_special = QPushButton("太晶化噴火龍")
        self.btn_special.clicked.connect(self.onSpecialFetchClicked)

        f2 = QFrame()
        l2 = QVBoxLayout(f2)
        l2.addWidget(lbl_symbol)

        row_symbol = QHBoxLayout()
        row_symbol.addWidget(self.line_symbol)
        row_symbol.addWidget(self.btn_search)

        # === NEW === 把「太晶化噴火龍」按鈕加到右邊
        row_symbol.addWidget(self.btn_special)

        l2.addLayout(row_symbol)
        self.topLeftLayout.addWidget(f2)

        lbl_expiry = QLabel("請選擇到期日：")
        self.combo_expiry = QComboBox()

        self.btn_refreshExpiry = QPushButton("更新")
        self.btn_refreshExpiry.setFixedSize(60, 28)
        self.btn_refreshExpiry.setStyleSheet("")

        self.btn_refreshExpiry.clicked.connect(self.onRefreshExpiryClicked)

        f3 = QFrame()
        l3 = QVBoxLayout(f3)
        l3.addWidget(lbl_expiry)

        row_expiry = QHBoxLayout()
        row_expiry.addWidget(self.combo_expiry)
        row_expiry.addWidget(self.btn_refreshExpiry)
        l3.addLayout(row_expiry)

        self.topLeftLayout.addWidget(f3)

        self.lbl_status = QLabel("狀態：就緒")
        self.topLeftLayout.addWidget(self.lbl_status)
        self.topLeftLayout.addStretch()

        self.grid_account_info = QGridLayout()
        self.topCenterLayout.addLayout(self.grid_account_info)
        self.topCenterLayout.addStretch()

        self.topRightLayout.addStretch()

        self.bottomLeft = QFrame()
        self.bottomLeftLayout = QVBoxLayout(self.bottomLeft)

        self.bottomCenter = QFrame()
        self.bottomCenterLayout = QVBoxLayout(self.bottomCenter)

        self.bottomRight = QFrame()
        self.bottomRightLayout = QVBoxLayout(self.bottomRight)

        bottomLayout.addWidget(self.bottomLeft, stretch=2)
        bottomLayout.addWidget(self.bottomCenter, stretch=7)
        bottomLayout.addWidget(self.bottomRight, stretch=1)

        detectFrame = QFrame()
        detectLayout = QHBoxLayout(detectFrame)
        self.btn_startDetect = QPushButton("開始偵測")
        self.btn_startDetect.clicked.connect(self.onStartDetect)
        self.lbl_countdown = QLabel("倒數: 30")
        self.lbl_currentPrice = QLabel("目前標的價格: ???")
        detectLayout.addWidget(self.btn_startDetect)
        detectLayout.addWidget(self.lbl_countdown)
        detectLayout.addWidget(self.lbl_currentPrice)
        self.bottomCenterLayout.addWidget(detectFrame)

        self.splitterForOrders = QSplitter(Qt.Orientation.Vertical)
        self.callGrp = QGroupBox("CALL 訂單")
        self.putGrp = QGroupBox("PUT 訂單")

        self.callTable = QTableWidget()
        self.callTable.setColumnCount(9)
        self.callTable.setHorizontalHeaderLabels([
            "到期日", "Symbol", "Action", "Qty", "勝率", "機率(小)", "機率(大)", "現在的機率", "資金運用百分比"
        ])
        self.callTable.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        callLayout = QVBoxLayout(self.callGrp)
        callLayout.addWidget(self.callTable)

        self.putTable = QTableWidget()
        self.putTable.setColumnCount(9)
        self.putTable.setHorizontalHeaderLabels([
            "到期日", "Symbol", "Action", "Qty", "勝率", "機率(小)", "機率(大)", "現在的機率", "資金運用百分比"
        ])
        self.putTable.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        # 綁定表格的數據更改信號
        self.callTable.itemChanged.connect(self.on_item_changed)
        self.putTable.itemChanged.connect(self.on_item_changed)

        putLayout = QVBoxLayout(self.putGrp)
        putLayout.addWidget(self.putTable)

        self.splitterForOrders.addWidget(self.callGrp)
        self.splitterForOrders.addWidget(self.putGrp)
        self.bottomCenterLayout.addWidget(self.splitterForOrders)

        self.logTextEdit = QPlainTextEdit()
        self.logTextEdit.setReadOnly(True)
        self.bottomRightLayout.addWidget(self.logTextEdit)

        lbl_title = QLabel("策略建倉")
        lbl_title.setStyleSheet("QLabel { font-size: 16px; font-weight: bold; }")
        self.bottomLeftLayout.addWidget(lbl_title)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("勝率:"))
        self.line_winRate = QLineEdit()
        self.line_winRate.setText("90")
        row1.addWidget(self.line_winRate)
        self.bottomLeftLayout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("當機率小於:"))
        self.line_probLo = QLineEdit()
        self.line_probLo.setText("2.5")
        row2.addWidget(self.line_probLo)
        self.bottomLeftLayout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("當機率大於:"))
        self.line_probHi = QLineEdit()
        self.line_probHi.setText("12.5")
        row3.addWidget(self.line_probHi)
        self.bottomLeftLayout.addLayout(row3)

        row4 = QHBoxLayout()
        row4.addWidget(QLabel("占用總現金價值(%):"))
        self.line_usagePct = QLineEdit("")
        row4.addWidget(self.line_usagePct)
        self.lbl_usageAmt = QLabel("對應金額= ???")
        row4.addWidget(self.lbl_usageAmt)
        self.bottomLeftLayout.addLayout(row4)

        self.btn_build = QPushButton("開倉")
        self.btn_build.clicked.connect(self.onBuildClicked)
        self.bottomLeftLayout.addWidget(self.btn_build)

        gifGrid = QGridLayout()

        label_tl = QLabel()
        label_tl.setFixedSize(150, 150)
        label_tl.setScaledContents(True) 
        movie_tl = QMovie(r"D:\Tim work station\IB雙賣\lee-chaeyoung-fromis-9 (1).gif")
        movie_tl.setScaledSize(QSize(150, 150))

        label_tl.setMovie(movie_tl)
        movie_tl.start()
        gifGrid.addWidget(label_tl, 0, 0)

        label_tr = QLabel()
        label_tr.setFixedSize(150, 150)
        label_tr.setScaledContents(True)
        movie_tr = QMovie(r"D:\Tim work station\IB雙賣\lee-chaeyoung-fromis-9 (1).gif")
        movie_tr.setScaledSize(QSize(150, 150))

        label_tr.setMovie(movie_tr)
        movie_tr.start()
        gifGrid.addWidget(label_tr, 0, 1)

        label_bl = QLabel()
        label_bl.setFixedSize(150, 150)
        label_bl.setScaledContents(True)
        movie_bl = QMovie(r"D:\Tim work station\IB雙賣\lee-chaeyoung-fromis-9 (1).gif")
        movie_bl.setScaledSize(QSize(150, 150))

        label_bl.setMovie(movie_bl)
        movie_bl.start()
        gifGrid.addWidget(label_bl, 1, 0)

        label_br = QLabel()
        label_br.setFixedSize(150, 150)
        label_br.setScaledContents(True)
        movie_br = QMovie(r"D:\Tim work station\IB雙賣\lee-chaeyoung-fromis-9 (1).gif")
        movie_br.setScaledSize(QSize(150, 150))

        label_br.setMovie(movie_br)
        movie_br.start()
        gifGrid.addWidget(label_br, 1, 1)

        self.bottomLeftLayout.addLayout(gifGrid)
        self.bottomLeftLayout.addStretch()

        self.line_usagePct.textChanged.connect(self.onUsagePctChanged)

    def onRefreshExpiryClicked(self):
        self.btn_refreshExpiry.setStyleSheet("")

        self.onSearchClicked()

    def _fillAccountInfo(self):
        row = 0
        for k in self.important_tags.keys():
            label_key = QLabel(f"{k}：")
            label_val = QLabel("")
            label_key.setStyleSheet("QLabel { font-size: 14px; }")
            label_val.setStyleSheet("QLabel { font-size: 14px; color: #F0F0A0; }")
            self.grid_account_info.addWidget(label_key, row, 0)
            self.grid_account_info.addWidget(label_val, row, 1)
            row += 1

    def refreshOrderTables(self):
        self.callTable.blockSignals(True)
        self.putTable.blockSignals(True)

        orders = self.dbmgr.fetch_active_orders()
        self.callTable.setRowCount(0)
        self.putTable.setRowCount(0)
        if not orders:
            self.callTable.blockSignals(False)
            self.putTable.blockSignals(False)
            return
        for od in orders:
            self.addOrderRow(od)

        self.callTable.blockSignals(False)
        self.putTable.blockSignals(False)

    def addOrderRow(self, od: dict):
        if " C " in od.get("symbol", ""):
            table = self.callTable
        else:
            table = self.putTable
        row = table.rowCount()
        table.insertRow(row)
        expiry_val = od.get("expiry")
        if isinstance(expiry_val, (date, datetime)):
            expiry_str = expiry_val.strftime("%Y-%m-%d")
        else:
            expiry_str = ""
        table.setItem(row, 0, QTableWidgetItem(expiry_str))
        table.setItem(row, 1, QTableWidgetItem(str(od.get("symbol", ""))))
        table.setItem(row, 2, QTableWidgetItem(str(od.get("action", ""))))
        table.setItem(row, 3, QTableWidgetItem(str(od.get("totalQuantity", 0))))
        wval = od.get("win_rate", 0)
        table.setItem(row, 4, QTableWidgetItem(f"{wval:.2f}%"))
        lo = od.get("prob_less", 0)
        hi = od.get("prob_greater", 0)
        table.setItem(row, 5, QTableWidgetItem(f"{lo:.2f}%"))
        table.setItem(row, 6, QTableWidgetItem(f"{hi:.2f}%"))
        exec_prob = od.get("exec_prob", 0)
        # 如果機率超過 1，顯示 100%
        if exec_prob > 1:
            exec_prob = 1.0
        table.setItem(row, 7, QTableWidgetItem(f"{exec_prob*100:.2f}%"))
        # 另外一欄可以留空或作其他用途
        table.setItem(row, 8, QTableWidgetItem("5%"))

    @pyqtSlot(list)
    def onAccountsReady(self, accounts: list):
        self.combo_account.clear()
        self.combo_account.addItems(accounts)
        self.log_append(f"帳戶清單載入: {accounts}")

    @pyqtSlot(dict)
    def onAccountValueUpdate(self, vals: dict):
        self.log_append(f"帳戶資訊更新: {vals}")
        for k in self.important_tags.keys():
            if k in vals:
                self.important_tags[k] = vals[k]
        for i in reversed(range(self.grid_account_info.count())):
            w = self.grid_account_info.itemAt(i).widget()
            if w:
                w.setParent(None)
        row = 0
        for k, v in self.important_tags.items():
            label_key = QLabel(f"{k}：")
            label_val = QLabel(v)
            label_key.setStyleSheet("QLabel { font-size: 14px; }")
            label_val.setStyleSheet("QLabel { font-size: 14px; color: #F0F0A0; }")
            self.grid_account_info.addWidget(label_key, row, 0)
            self.grid_account_info.addWidget(label_val, row, 1)
            row += 1

    @pyqtSlot(QTableWidgetItem)
    def on_item_changed(self, item):
        if self.initializing:
            return
        row = item.row()
        column = item.column()

        relevant_columns = [4, 5, 6]  # "勝率"、"機率(小)"、"機率(大)"
        if column not in relevant_columns:
            return

        table = item.tableWidget()
        expiry = table.item(row, 0).text()
        symbol = table.item(row, 1).text()
        action = table.item(row, 2).text()
        qty = float(table.item(row, 3).text())

        updated_data = {}
        for col_index, field_name in zip([4, 5, 6], ["win_rate", "prob_less", "prob_greater"]):
            cell = table.item(row, col_index)
            if cell is not None and cell.text():
                text = cell.text().strip().rstrip('%')
                try:
                    value = float(text)
                    updated_data[field_name] = value
                except ValueError:
                    QMessageBox.warning(self, "輸入錯誤", f"無效的 {field_name}: {cell.text()}")
                    return

        try:
            self.dbmgr.update_multiple_fields(symbol, action, qty, expiry, updated_data)
        except Exception as e:
            print(f"更新資料庫失敗: {e}")

    def onSearchClicked(self):
        sym = self.line_symbol.text().strip().upper()
        if not sym:
            self.log_append("[警告] 請輸入標的!")
            return
        exps = self.dbmgr.fetch_expiries(sym)
        today = datetime.now().date()
        valid = [e for e in exps if (e - today).days >= 0]
        self.combo_expiry.clear()
        for e in sorted(valid):
            d = (e - today).days
            self.combo_expiry.addItem(f"{e} (剩{d}天)")
        if valid:
            self.log_append(f"已顯示舊資料 (未過期): {len(valid)} 筆")
            self.lbl_status.setText("狀態：就緒")
        else:
            self.log_append("[提醒] DB 無此標的期權資料，需另建。")
            self.lbl_status.setText("狀態：無資料")

    # === NEW === 「太晶化噴火龍」按鈕的事件
    def onSpecialFetchClicked(self):
        sym = self.line_symbol.text().strip().upper()
        if not sym:
            self.log_append("[警告] 請先輸入標的 Symbol")
            return
        self.log_append(f"[INFO] 開始執行抓取全部到期日(含所有履約價)，Symbol={sym}")

        self.fetchThread = FetchOptionChainThread(symbol=sym)
        self.fetchThread.finishedSignal.connect(self.onFetchThreadFinished)
        self.fetchThread.start()

    # === NEW === 線程完成後，根據有無新到期日改按鈕顏色
    def onFetchThreadFinished(self, has_new):
        if has_new:
            self.btn_refreshExpiry.setStyleSheet("background-color: orange;")
            self.log_append("有新的到期日！【更新】按鈕已變橙色。")
        else:
            self.btn_refreshExpiry.setStyleSheet("")
            self.log_append("抓取結束，沒有新增到期日。")
        # 呼叫一次搜尋來刷新下拉選單
        self.onSearchClicked()

    def onUsagePctChanged(self):
        pct_str = self.line_usagePct.text().strip()
        if not pct_str:
            self.lbl_usageAmt.setText("對應金額= ???")
            return
        try:
            pct_val = float(pct_str)
        except:
            pct_val = 0
        cv_str = self.important_tags.get("TotalCashValue", "0")
        try:
            cv = float(cv_str)
        except:
            cv = 0
        usageAmt = cv * (pct_val / 100.0)
        self.lbl_usageAmt.setText(f"對應金額= {usageAmt:.2f}")

    def count_trading_days(self, df: pd.DataFrame, start_date: date, end_date: date) -> int:
        if df is None or df.empty:
            return 0
        start_ts = pd.Timestamp(start_date)
        end_ts = pd.Timestamp(end_date)
        mask = (df['Date'] >= start_ts) & (df['Date'] <= end_ts)
        sub = df.loc[mask]
        return len(sub)

    def cal_data(self, df: pd.DataFrame, days=1):
        df2 = df.copy()
        df2['pct'] = df2['Close'].pct_change(periods=days)
        df2.dropna(subset=['pct'], inplace=True)
        return df2

    def compute_execution_probability(self, strike, days_left, current_price, side="C"):
        if self.qqq_daily_data is None or len(self.qqq_daily_data) == 0:
            return 0.0
        if days_left <= 0:
            return 0.0
        if len(self.qqq_daily_data) < days_left:
            days_left = len(self.qqq_daily_data)
        df_calc = self.cal_data(self.qqq_daily_data, days=days_left)
        if df_calc.empty:
            return 0.0
        needed_pct = (float(strike) / current_price) - 1.0
        total_count = len(df_calc)
        if total_count == 0:
            return 0.0
        if side.upper() == "C":
            count = len(df_calc[df_calc["pct"] >= needed_pct])
        else:
            count = len(df_calc[df_calc["pct"] <= needed_pct])
        prob = count / total_count
        if prob > 1:
            prob = 1.0
        print(f"[Debug] side={side}, strike={strike}, days={days_left}, "
              f"needed_pct={needed_pct*100:.2f}%, prob={prob*100:.2f}%")
        return prob

    def onBuildClicked(self):
        try:
            self.winRate = float(self.line_winRate.text().strip())
        except:
            self.winRate = 0
        try:
            self.probLo = float(self.line_probLo.text().strip())
        except:
            self.probLo = 0
        try:
            self.probHi = float(self.line_probHi.text().strip())
        except:
            self.probHi = 0
        try:
            self.usagePct = float(self.line_usagePct.text().strip())
        except:
            self.usagePct = 0

        bp_str = self.important_tags.get("BuyingPower", "0")
        try:
            bp = float(bp_str)
        except:
            bp = 0
        self.usageAmt = bp * (self.usagePct / 100.0)

        self.sym = self.line_symbol.text().strip().upper()
        if not self.sym:
            self.log_append("[警告] 尚未輸入標的")
            return
        cur_item = self.combo_expiry.currentText().strip()
        if not cur_item:
            self.log_append("[警告] 尚未選擇到期日")
            return
        expiry_date_str = cur_item.split()[0]
        try:
            self.expiry_date = datetime.strptime(expiry_date_str, "%Y-%m-%d").date()
        except:
            self.log_append("[警告] 到期日格式有誤")
            return

        raw_trading_days = self.count_trading_days(self.qqq_daily_data, date.today(), self.expiry_date)
        if raw_trading_days < 1:
            if self.expiry_date > date.today():
                fallback_days = (self.expiry_date - date.today()).days
                
                if fallback_days < 1:
                    self.log_append("[警告] 到期日已過或不足1天")
                    return
                else:
                    self.log_append(f"[提醒] 改用純日曆天={fallback_days}")
                    days_to_expiry = fallback_days
            else:
                self.log_append("[警告] 到期日已過或剩餘交易日不足1天")
                return
        else:
            days_to_expiry = raw_trading_days

        if days_to_expiry < 1:
            self.log_append("[警告] 到期日已過或剩餘天數不足1天")
            return

        self.exec_prob_input = (100 - self.winRate) / 100 / 2
        self.log_append(f"目標勝率={self.winRate:.2f} => 輸入執行機率={self.exec_prob_input*100:.2f}% (days={days_to_expiry})")

        strikes = self.dbmgr.fetch_strikes('QQQ', self.expiry_date)

        if not strikes:
            self.log_append("[警告] 找不到 strikes 資料")
            return
        self.strikes = sorted(strikes, key=lambda x: float(x))

        self.strategyDialog = StrategyDialog(
            parent=self,
            symbol='QQQ',
            expiry=self.expiry_date,
            callStrike="等待中...",
            putStrike="等待中...",
            winRate=self.winRate,
            probLo=self.probLo,
            probHi=self.probHi,
            usageAmt=self.usageAmt,
            call_contract_value=0,
            call_allowed=0,
            put_contract_value=0,
            put_allowed=0
        )
        self.strategyDialog.show()

        self.priceThread = PriceFetchThread('127.0.0.1', 7497, 'QQQ', timeout=10)
        self.priceThread.priceFetched.connect(self.onPriceFetched_UI)
        self.priceThread.start()

    def onPriceFetched_UI(self, latest_price):
        if latest_price is None:
            current_price = float(self.qqq_daily_data["Close"].iloc[-1])
        else:
            current_price = float(latest_price)

        strikes = self.dbmgr.fetch_strikes('QQQ', self.expiry_date)

        if not strikes:
            self.log_append("[警告] 找不到 strikes 資料")
            return
        strikes = sorted(strikes, key=lambda x: float(x))

        call_strike, self.call_exec_prob = self.calculate_call_strike(current_price, self.expiry_date, self.exec_prob_input, strikes)
        put_strike, self.put_exec_prob = self.calculate_put_strike(current_price, self.expiry_date, self.exec_prob_input, strikes)

        multiplier = 100
        contract_value = current_price * multiplier

        call_usage = self.usageAmt / 2
        put_usage = self.usageAmt / 2
        call_allowed = int(call_usage // contract_value)
        if call_allowed < 1:
            call_allowed = 1
        put_allowed = int(put_usage // contract_value)
        if put_allowed < 1:
            put_allowed = 1

        self.strategyDialog.updateData(
            callStrike=call_strike,
            putStrike=put_strike,
            call_contract_value=contract_value,
            call_allowed=call_allowed,
            put_contract_value=contract_value,
            put_allowed=put_allowed
        )

    def onPositionsFetched(self, positions):
        if not positions:
            print("未取得合約資訊")
        else:
            print(positions)
            # 建立持倉字典，key = (symbol, right, strike)，value = position
            # position_dict = {
            #     (contract.symbol, contract.right, round(float(contract.strike), 2)): position
            #     for contract, position in positions if position != 0  
            # }

            # orders = self.dbmgr.fetch_active_orders()
            # unmatched_orders = [] 

            # for od in orders:
            #     id = od['id']
            #     db_symbol = od['symbol']
            #     parts = db_symbol.split()
            #     symbol = parts[0]
            #     option_type = parts[1]  # 'P' 或 'C'
            #     strike_price = float(parts[2])

            #     key = (symbol, option_type, round(strike_price, 2))

            #     if key not in position_dict:
            #         unmatched_orders.append((id, db_symbol)) 

            # unmatched_orders.sort(key=lambda x: x[0])

            # # 如果有未匹配的訂單，彈出 QMessageBox 提醒用戶
            # if unmatched_orders:
            #     msg = QMessageBox()
            #     msg.setIcon(QMessageBox.Icon.Warning)
            #     msg.setWindowTitle("持倉異常")
            #     unmatched_list = "\n".join([f"ID: {order_id}, Symbol: {symbol}" for order_id, symbol in unmatched_orders])
            #     msg.setText(f"以下資料庫中訂單未匹配 IB 持倉:\n\n{unmatched_list}")
            #     msg.exec()

    def onStartDetect(self):
        if not self.detection_running:
            self.detection_running = True
            self.btn_startDetect.setText("停止偵測")
            self.countdown = 30
            self.detect_timer.start()
        else:
            self.detection_running = False
            self.btn_startDetect.setText("開始偵測")
            self.detect_timer.stop()

    def onTimerTick(self):
        self.countdown -= 1
        self.lbl_countdown.setText(f"倒數: {self.countdown}")
        if self.countdown <= 0:
            self.detect_timer.stop()
            self.startPriceFetch() 
            self.countdown = 30
            self.detect_timer.start()

    def onPriceFetched_update_label(self, latest_price):
        """
        當 PriceFetchThread 抓取價格完成時的回調。
        """
        if latest_price is None:
            current_price = float(self.qqq_daily_data["Close"].iloc[-1]) 
        else:
            current_price = float(latest_price)
        self.lbl_currentPrice.setText(f"目前標的價格: {current_price:.2f}")

    def onPriceFetched(self, latest_price):
        """
        當 PriceFetchThread 抓取價格完成時的回調。
        """
        if latest_price is None:
            current_price = float(self.qqq_daily_data["Close"].iloc[-1])
        else:
            current_price = float(latest_price)
        self.performDetection(current_price)

    def startPriceFetch(self):
        """
        啟動 PriceFetchThread 抓取報價，並在完成後觸發 performDetection。
        """
        if self.qqq_daily_data is not None and not self.qqq_daily_data.empty:
            self.priceThread = PriceFetchThread('127.0.0.1', 7497, 'QQQ', timeout=10)
            self.priceThread.priceFetched.connect(self.onPriceFetched)
            self.priceThread.start()
        else:
            self.performDetection(0.0)

    def performDetection(self, current_price):
        self.adj = {'C': [], 'P': []}
        self.lbl_currentPrice.setText(f"目前標的價格: {current_price:.2f}")
        orders = self.dbmgr.fetch_active_orders()

        today = date.today()

        # # 執行持倉部位線程
        # self.positionThread = PositionFetchThread('127.0.0.1', 7497, timeout=10)
        # self.positionThread.positionsFetched.connect(self.onPositionsFetched)
        # self.positionThread.start()

        for od in orders:
            self.adjust_signal = False
            if od.get("expiry") and od.get("strike"):
                days_left = (od["expiry"] - today).days
                # 鎖住剩餘日期為2天
                if days_left <= 2 :  
                    days_left = 2
                id = od.get("id")
                strike_val = od.get("strike", 0.0)
                expiry_date = od.get("expiry")
                quantity = od.get("totalQuantity")
                side = "C" if " C " in od.get("symbol", "") else "P"
                prob_lo = od.get("prob_less", 0.0)
                prob_hi = od.get("prob_greater", 0.0)
                win_rate = od.get("win_rate", 0.0)
                usagePct = od.get("usagePct", 0.0)
                new_prob = self.compute_execution_probability(strike_val, days_left, current_price, side=side)
                od["exec_prob"] = new_prob  # 存到資料庫
                self.dbmgr.update_exec_prob(id, new_prob)
                # 執行機率判斷調倉邏輯
                self.compare_execution_probability(side, new_prob, strike_val, quantity, expiry_date, prob_lo, prob_hi, win_rate, current_price, usagePct)

        if self.adj:
            for side, adjustments in self.adj.items():
                for adj in adjustments:
                    expiry_date = str(adj.get('expiry_date')).replace("-", "")
                    lot = int(adj.get('quantity'))
                    c = Contract()
                    c.symbol = 'QQQ'
                    c.secType = "OPT"
                    c.exchange = "SMART"
                    c.currency = "USD"
                    c.lastTradeDateOrContractMonth = expiry_date
                    c.strike = int(adj.get('original_strike'))
                    c.right = side
                    order_id = self.ibWorker.place_mkt_order(c, "Buy", lot)
                    self.logTextEdit.appendPlainText(f"{now_str()} [已送出市價單平倉] OrderId={order_id}\n{sep_line()}")

                    target = f"{c.symbol} {c.right} {c.strike:.2f}"
                    self.dbmgr.update_option_order(target)

                    usagePct = int(adj.get('usagePct'))
                    lot = 1
                    new_c = Contract()
                    new_c.symbol = "QQQ"
                    new_c.secType = "OPT"
                    new_c.exchange = "SMART"
                    new_c.currency = "USD"
                    new_c.lastTradeDateOrContractMonth = expiry_date
                    new_c.strike = int(adj.get('new_strike'))
                    new_c.right = side
                    order_id = self.ibWorker.place_mkt_order(new_c, "Sell", lot)
                    self.logTextEdit.appendPlainText(f"{now_str()} [已送出市價單開倉] OrderId={order_id}\n{sep_line()}")

                    od = {
                        "order_time": datetime.now(),
                        "expiry": adj.get('expiry_date'),
                        "symbol": f"{new_c.symbol} {side} {new_c.strike:.2f}",
                        "action": "SELL",
                        "totalQuantity": lot,
                        "strike": float(new_c.strike),
                        "win_rate": float(adj.get('win_rate')),
                        "prob_less": float(adj.get('prob_lo')),
                        "prob_greater": float(adj.get('prob_hi')),
                        "exec_prob": adj.get('exec_prob'),
                        "status": 'Active',
                        "usagePct": float(adj.get('usagePct'))
                    }
                    self.dbmgr.insert_option_order(od)

        self.refreshOrderTables()
        self.log_append("偵測更新完成！")

    def compare_execution_probability(self, side, new_prob, strike_val, quantity, expiry_date, prob_lo, prob_hi, win_rate, current_price, usagePct):
        if new_prob*100 < prob_lo:
            self.adjust_signal = True
            print('低於閾值')
        elif new_prob*100 > prob_hi:
            self.adjust_signal = True
            print('超過閾值')
        else:
            print("執行機率適中，保持現有部位")

        if self.adjust_signal == True:
            # 分call/put將adjustment加入adj
            exec_win_rate = (100 - win_rate) / 100 / 2
            strikes = self.dbmgr.fetch_strikes('QQQ', expiry_date)
            if side =='C':
                if not strikes:
                    self.log_append("[警告] 找不到 strikes 資料")
                    return
                strikes = sorted(strikes, key=lambda x: float(x))
                call_strike, call_exec_prob = self.calculate_call_strike(current_price, expiry_date, exec_win_rate, strikes)
                print(f'找QQQ過期日{expiry_date}履約價的合約')
                adjustment = {
                    'original_strike': strike_val,
                    'side': 'C',
                    'expiry_date': expiry_date,
                    'quantity': quantity,
                    'new_strike': call_strike,
                    'new_prob': new_prob,
                    'current_price': current_price,
                    'win_rate': win_rate,
                    'prob_lo': prob_lo,
                    'prob_hi': prob_hi,
                    'exec_prob': call_exec_prob,
                    'usagePct': usagePct
                }
            else:
                if not strikes:
                    self.log_append("[警告] 找不到 strikes 資料")
                    return
                strikes = sorted(strikes, key=lambda x: float(x))
                put_strike, put_exec_prob = self.calculate_put_strike(current_price, expiry_date, exec_win_rate, strikes)
                print(f'找QQQ過期日{expiry_date}履約價的合約')
                adjustment = {
                    'original_strike': strike_val,
                    'side': 'P',
                    'expiry_date': expiry_date,
                    'quantity': quantity,
                    'new_strike': put_strike,
                    'new_prob': new_prob,
                    'current_price': current_price,
                    'win_rate': win_rate,
                    'prob_lo': prob_lo,
                    'prob_hi': prob_hi,
                    'exec_prob': put_exec_prob,
                    'usagePct': usagePct
                }
            self.adj[side].append(adjustment)

    def calculate_call_strike(self, current_price, expiry_date, winrate, strikes): #  計算call合約機率
        call_strike = None
        prev_strike = None
        today = date.today()
        days_left = (expiry_date - today).days
        # 鎖住剩餘日期為2天
        if days_left <=2:
            days_left = 2

        for st in reversed(strikes):
            p = self.compute_execution_probability(st, days_left, current_price, side="C")
            if p > winrate:
                call_strike = prev_strike if prev_strike is not None else st
                break
            prev_strike = st
        if call_strike is None:
            call_strike = strikes[0]

        # 計算選取履約價的機率
        call_exec_prob = self.compute_execution_probability(call_strike, days_left, current_price, side="C")
        self.log_append(f"選取的CALL strike={call_strike}, 執行機率={call_exec_prob*100:.2f}%")
        return call_strike, call_exec_prob
    
    def calculate_put_strike(self, current_price, expiry_date, winrate, strikes):  # 計算put合約機率
        put_strike = None
        prev_strike = None
        today = date.today()
        days_left = (expiry_date - today).days
        # 鎖住剩餘日期為2天
        if days_left <=2:
            days_left = 2

        for st in strikes:
            p = self.compute_execution_probability(st, days_left, current_price, side="P")
            if p > winrate:
                put_strike = prev_strike if prev_strike is not None else st
                break
            prev_strike = st
        if put_strike is None:
            put_strike = strikes[-1]

        # 計算選取的 PUT 履約價執行機率，並記錄到日誌
        put_exec_prob = self.compute_execution_probability(put_strike, days_left, current_price, side="P")
        self.log_append(f"選取的PUT strike={put_strike}, 執行機率={put_exec_prob*100:.2f}%")
        return put_strike, put_exec_prob
        
    def closeEvent(self, event):
        self.ibWorker.stop()
        self.ibWorker.quit()
        self.ibWorker.wait()
        for th in self.threads:
            try:
                th.quit()
                th.wait()
            except:
                pass
        self.dbmgr.close_db()
        super().closeEvent(event)

    def log_append(self, msg: str):
        final_msg = f"{now_str()} {msg}\n{sep_line()}"
        self.logTextEdit.appendPlainText(final_msg)


# =============================================================================
#   main
# =============================================================================
def main():
    app = QApplication(sys.argv)

    darkPalette = QPalette()
    darkPalette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
    darkPalette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    darkPalette.setColor(QPalette.ColorRole.Base, QColor(35, 35, 35))
    darkPalette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
    darkPalette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    darkPalette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    darkPalette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    darkPalette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    darkPalette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    darkPalette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    darkPalette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    darkPalette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    darkPalette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
    app.setPalette(darkPalette)

    font = QFont("Microsoft YaHei", 11)
    app.setFont(font)

    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
