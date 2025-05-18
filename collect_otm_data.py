import sys
from datetime import datetime
from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.QtCore import QTimer
import shioaji as sj
import json
import os

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

        self.option_codes = ['TXO', 'TX1', 'TX2', 'TX4', 'TX5']
        self.monthly_code = self.get_monthly_option_code()

    def get_weekly_option_codes(self):
        all_contracts = dir(self.api.Contracts.Options)
        weekly_codes = [code for code in all_contracts if code.startswith("TX") and not code.startswith("TXO")]

        if 'TXO' in all_contracts:
            weekly_codes.append('TX3')  # 視為第三週選擇權

        return sorted(weekly_codes)
    
    def get_monthly_option_code(self):
        current_date = datetime.now()
        current_month_code = f"TXO{current_date.strftime('%Y%m')}"
        return current_month_code

    def get_monthly_snapshots(self):
        contracts = self.api.Contracts.Options.TXO

        contracts_filtered = [contract for contract in contracts]

        delivery_months = sorted(set(contract.delivery_month for contract in contracts_filtered))
        delivery_month=delivery_months[0]

        if not contracts_filtered:
            return [], 0.0

        index_contract = self.api.Contracts.Futures.TXF.TXFR1 # 找台指期的指數
        index_snapshot = self.api.snapshots([index_contract])[0]
        current_index_price = index_snapshot.close

        strike_prices = sorted(set(contract.strike_price for contract in contracts_filtered))

        if current_index_price not in strike_prices:
            strike_prices.append(current_index_price)
            strike_prices.sort()
            index_position = strike_prices.index(current_index_price)  
            strike_prices.remove(current_index_price) 
        else:
            index_position = strike_prices.index(current_index_price) 

        lower_bound = max(index_position - 25, 0)
        upper_bound = min(index_position + 25, len(strike_prices))

        put_filtered_strikes = strike_prices[lower_bound:index_position]
        call_filtered_strikes = strike_prices[index_position:upper_bound]

        filtered_put_contracts = [self.api.Contracts.Options['TXO'].get(f"{'TXO'}{delivery_month}{int(strike)}P") for strike in put_filtered_strikes]
        filtered_call_contracts = [self.api.Contracts.Options['TXO'].get(f"{'TXO'}{delivery_month}{int(strike)}C") for strike in call_filtered_strikes]

        call_snapshots = []
        put_snapshots = []

        for put_contract, call_contract in zip(filtered_put_contracts, filtered_call_contracts):
            try:
                # 取得 put 和 call 的快照
                snapshot_put = self.api.snapshots([put_contract]) if put_contract else None
                snapshot_call = self.api.snapshots([call_contract]) if call_contract else None

                if snapshot_call and all(snap.close is not None for snap in snapshot_call):
                    call_snapshots.append((call_contract, snapshot_call))
                else:
                    print(f"獲取合約 {call_contract} 的快照失敗")

                if snapshot_put and all(snap.close is not None for snap in snapshot_put):
                    put_snapshots.append((put_contract, snapshot_put))
                else:
                    print(f"獲取合約 {put_contract} 的快照失敗")

            except Exception as e:
                print(f"獲取合約 {call_contract} 或 {put_contract} 的快照時出錯: {e}")

        return call_snapshots, put_snapshots
    


    def get_filtered_snapshots(self, code):
        current_date = datetime.now()
        year_month = current_date.strftime("%Y%m")

        if code == 'TXO':
            return self.get_monthly_snapshots()

        contracts = self.api.Contracts.Options[code]

        contracts_filtered = [contract for contract in contracts]

        delivery_month=contracts_filtered[0].delivery_month

        if not contracts_filtered:
            return [], 0.0

        index_contract = self.api.Contracts.Futures.TXF.TXFR1 # 找台指期的指數
        index_snapshot = self.api.snapshots([index_contract])[0]
        current_index_price = index_snapshot.close

        strike_prices = sorted(set(contract.strike_price for contract in contracts_filtered))

        if current_index_price not in strike_prices:
            strike_prices.append(current_index_price)
            strike_prices.sort()
            index_position = strike_prices.index(current_index_price)  
            strike_prices.remove(current_index_price) 
        else:
            index_position = strike_prices.index(current_index_price) 

        lower_bound = max(index_position - 25, 0)
        upper_bound = min(index_position + 25, len(strike_prices))

        put_filtered_strikes = strike_prices[lower_bound:index_position]
        call_filtered_strikes = strike_prices[index_position:upper_bound]

        filtered_put_contracts = [self.api.Contracts.Options[code].get(f"{code}{delivery_month}{int(strike)}P") for strike in put_filtered_strikes]
        filtered_call_contracts = [self.api.Contracts.Options[code].get(f"{code}{delivery_month}{int(strike)}C") for strike in call_filtered_strikes]

        call_snapshots = []
        put_snapshots = []

        for put_contract, call_contract in zip(filtered_put_contracts, filtered_call_contracts):
            try:
                # 取得 put 和 call 的快照
                snapshot_put = self.api.snapshots([put_contract]) if put_contract else None
                snapshot_call = self.api.snapshots([call_contract]) if call_contract else None

                if snapshot_call and all(snap.close is not None for snap in snapshot_call):
                    call_snapshots.append((call_contract, snapshot_call))
                else:
                    print(f"獲取合約 {call_contract} 的快照失敗")

                if snapshot_put and all(snap.close is not None for snap in snapshot_put):
                    put_snapshots.append((put_contract, snapshot_put))
                else:
                    print(f"獲取合約 {put_contract} 的快照失敗")

            except Exception as e:
                print(f"獲取合約 {call_contract} 或 {put_contract} 的快照時出錯: {e}")

        return call_snapshots, put_snapshots
    

    def get_latest_contract(self):
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

        expirations = []
        nearest_category = None
        nearest_date = None

        for category, dates in expirations_dict.items():
            first_expiration = datetime.strptime(dates[0], '%Y/%m/%d')
            if nearest_date is None or first_expiration < nearest_date:
                nearest_date = first_expiration
                print(nearest_date)
                nearest_category = category

        expirations.append((nearest_category, nearest_date.strftime("%Y/%m/%d")))
        print(expirations)

        return expirations

    def get_contract_by_tv_ratio(self):

        sorted_expirations = self.get_latest_contract()
        target = sorted_expirations[0][0]
        call_snapshots, put_snapshots = self.get_filtered_snapshots(target)

        otm_dic = {}

        call_otm_sum = 0
        put_otm_sum = 0
        otm_sum = 0

        for put_snap, call_snap in zip(put_snapshots, call_snapshots):

            put_strike = put_snap[0].strike_price
            put_snap_list = put_snap[1]
            put_buy_price = put_snap_list[0].buy_price if put_snap else 0
            put_sell_price = put_snap_list[0].sell_price if put_snap else 0
            put_mean_price = (put_buy_price+put_sell_price)/2
            otm_dic[put_strike] = put_mean_price
            put_otm_sum += put_mean_price

            call_strike = call_snap[0].strike_price
            call_snap_list = call_snap[1]
            call_buy_price = call_snap_list[0].buy_price if call_snap_list else 0
            call_sell_price = call_snap_list[0].sell_price if call_snap_list else 0
            call_mean_price = (call_buy_price+call_sell_price)/2
            otm_dic[call_strike] = call_mean_price
            call_otm_sum += call_mean_price

        otm_sum = put_otm_sum + call_otm_sum

        print(f"put的價外25檔總和:{put_otm_sum}")
        print(f"call的價外25檔總和:{call_otm_sum}")
        print(f"價外上下25檔總和:{otm_sum}")

        otm_dic = dict(sorted(otm_dic.items()))

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        today = datetime.now().strftime("%Y-%m-%d")
        date = datetime.now().strftime("%Y-%m")
        base_path = r"C:\Users\user\Desktop\公司帳戶"
        folder_name = f"{date}_{target}" 
        base_folder = os.path.join(base_path, folder_name)
        file_name = f"{today}_otm.json"
        file_path = os.path.join(base_folder, file_name)

        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        # 讀取舊資料
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                existing_data = json.load(f)
        else:
            existing_data = {}

        existing_data[timestamp] = {
            "otm_data": otm_dic,
            "otm_sum": otm_sum,
            "put_otm_sum": put_otm_sum,
            "call_otm_sum": call_otm_sum
        }
        
        with open(file_path, "w", encoding="utf-8") as json_file:
            json.dump(existing_data, json_file, indent=4)

        print(f"{timestamp}資料儲存成功!")
    

class OptionAnalyzerApp(QWidget):

    def __init__(self):
        super().__init__()

        self.option_manager = OptionDataManager()

        self.api = self.option_manager.api

        self.option_manager.get_contract_by_tv_ratio()

        # self.refresh_timer = QTimer(self)
        # self.refresh_timer.timeout.connect(self.option_manager.get_contract_by_tv_ratio)
        # self.refresh_timer.start(300000)  # 每5分鐘更新一次

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OptionAnalyzerApp()
    sys.exit(app.exec())

