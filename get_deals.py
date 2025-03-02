import MetaTrader5 as mt5
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill
from PySide6.QtWidgets import QFileDialog, QInputDialog
import configparser

def get_pip_size(symbol):
    info = mt5.symbol_info(symbol)
    return info.point * 10 if info else 0.0001

def load_config():
    config = configparser.ConfigParser()
    config.read("config.ini")
    return float(config.get("Settings", "qty", fallback=1.0))

def download_real_transactions(start_date, end_date, include_opening=True):
    if not mt5.initialize():
        print("MT5 initialization failed.", mt5.last_error())
        return []

    qty_threshold = load_config()
    start_date = datetime.combine(start_date, datetime.min.time())
    end_date = datetime.combine(end_date, datetime.max.time())

    deals = mt5.history_deals_get(start_date, end_date)
    if deals is None:
        print("No deals found.", mt5.last_error())
        return []

    positions = {}
    transactions = []
    
    for deal in deals:
        print(f"Processing deal: Ticket={deal.ticket}, Position ID={deal.position_id}, Type={'Open' if deal.entry == 0 else 'Close'}, Price={deal.price}, Volume={deal.volume}")
        
        if deal.position_id not in positions:
            positions[deal.position_id] = {
                "entry_price": 0,
                "volume": 0,
                "sl_prices": [],
                "tp_prices": [],
                "entry_time": None,
                "close_price": 0,
                "side": ""
            }
        
        pos = positions[deal.position_id]
        pip_size = get_pip_size(deal.symbol)

        # Get SL/TP from all historical orders with non-zero SL/TP
        order_data = mt5.history_orders_get(position=deal.position_id)
        if order_data:
            for order in order_data:
                if order.sl > 0:
                    pos["sl_prices"].append(order.sl)
                if order.tp > 0:
                    pos["tp_prices"].append(order.tp)
            print(f"Order Data Found: SLs={pos['sl_prices']}, TPs={pos['tp_prices']}")
        else:
            print(f"No historical order data found for Position ID: {deal.position_id}")

        # Only consider entry trades for these calculations
        if deal.entry == 0:  # Opening or increasing position
            pos["entry_price"] += deal.price * deal.volume
            pos["volume"] += deal.volume
            pos["side"] = "Buy" if deal.type == 0 else "Sell"
            if pos["entry_time"] is None or deal.time < pos["entry_time"]:
                pos["entry_time"] = deal.time

        # Only consider closing trades for close price calculation
        if deal.entry == 1:  # Closing or reducing position
            pos["close_price"] = deal.price

    summary = {
        "total_result": 0,
        "total_count": 0,
        "win_result": 0,
        "lose_result": 0,
        "zero_result": 0,
        "win_count": 0,
        "lose_count": 0,
        "zero_count": 0
    }

    for position_id, pos in positions.items():
        pos["entry_price"] /= pos["volume"] if pos["volume"] > 0 else 1
        pos["sl_price"] = sum(pos["sl_prices"]) / len(pos["sl_prices"]) if pos["sl_prices"] else None
        pos["tp_price"] = sum(pos["tp_prices"]) / len(pos["tp_prices"]) if pos["tp_prices"] else None
        print(f"Final calculated values for Position ID {position_id}: Entry Price={pos['entry_price']}, SL Price={pos['sl_price']}, TP Price={pos['tp_price']}, Close Price={pos['close_price']}")
    
    for deal in deals:
        if not include_opening and deal.entry == 0:
            continue

        entry_price = positions[deal.position_id]["entry_price"]
        sl_price = positions[deal.position_id]["sl_price"]
        tp_price = positions[deal.position_id]["tp_price"]
        close_price = positions[deal.position_id]["close_price"]
        side = positions[deal.position_id]["side"] if not include_opening else "Buy" if deal.type == 0 else "Sell"
        entry_time = datetime.fromtimestamp(positions[deal.position_id]["entry_time"]).strftime("%H:%M:%S") if positions[deal.position_id]["entry_time"] else ""
        pip_size = get_pip_size(deal.symbol)
        sl_pips = abs(entry_price - sl_price) / pip_size if sl_price else ""
        tp_pips = abs(tp_price - entry_price) / pip_size if tp_price else ""
        result_pips = (close_price - entry_price) / pip_size if close_price and entry_price else ""
        result_usd = deal.profit

        category = "Zero"
        if result_usd > qty_threshold:
            category = "Win"
            summary["win_count"] += 1
            summary["win_result"] += result_usd
        elif result_usd < -qty_threshold:
            category = "Lose"
            summary["lose_count"] += 1
            summary["lose_result"] += result_usd
        else:
            summary["zero_count"] += 1
            summary["zero_result"] += result_usd

        summary["total_count"] += 1
        summary["total_result"] += result_usd
        
        transactions.append({
            "date": datetime.fromtimestamp(deal.time).strftime("%d.%m.%Y"),
            "ticket": str(deal.ticket),
            "position_id": deal.position_id,
            "action": "Open" if deal.entry == 0 else "Close",
            "instrument": deal.symbol,
            "entry_time": entry_time,
            "side": side,
            "entry_price": entry_price,
            "close_price": close_price,
            "sl_price": sl_price,
            "sl_pips": sl_pips,
            "tp_price": tp_price,
            "tp_pips": tp_pips,
            "volume": deal.volume,
            "result_pips": result_pips,
            "result_usd": result_usd,
            "category": category
        })
    
    return transactions, summary
