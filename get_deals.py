import MetaTrader5 as mt5
from datetime import datetime
from PySide6.QtWidgets import QFileDialog, QInputDialog

def download_real_transactions(start_date, end_date, include_opening=True):
    if not mt5.initialize():
        print("MT5 initialization failed.", mt5.last_error())
        return []

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
                "close_price": 0
            }
        
        pos = positions[deal.position_id]

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
            if pos["entry_time"] is None or deal.time < pos["entry_time"]:
                pos["entry_time"] = deal.time

        # Only consider closing trades for close price calculation
        if deal.entry == 1:  # Closing or reducing position
            pos["close_price"] = deal.price

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
        entry_time = datetime.fromtimestamp(positions[deal.position_id]["entry_time"]).strftime("%H:%M:%S") if positions[deal.position_id]["entry_time"] else ""
        sl_pips = abs(entry_price - sl_price) * 10000 if sl_price else ""
        tp_pips = abs(tp_price - entry_price) * 10000 if tp_price else ""
        sl_usd = sl_pips * deal.volume * 10 if sl_pips else ""
        rr_plan = round(tp_pips / sl_pips, 2) if sl_pips and tp_pips else "N/A"
        rr_fact = round(deal.profit / sl_usd, 2) if sl_usd else "N/A"
        result_pips = (close_price - entry_price) * 10000 if close_price and entry_price else ""
        
        print(f"Exporting Deal Ticket={deal.ticket}: Entry Price={entry_price}, SL Price={sl_price}, TP Price={tp_price}, Close Price={close_price}")

        trans = {
            "date": datetime.fromtimestamp(deal.time).strftime("%d.%m.%Y"),
            "ticket": str(deal.ticket),
            "position_id": deal.position_id,
            "action": "Open" if deal.entry == 0 else "Close",
            "instrument": deal.symbol,
            "entry_time": entry_time,
            "side": "Buy" if deal.type == 0 else "Sell",
            "entry_price": entry_price,
            "close_price": close_price,
            "sl_price": sl_price,
            "sl_pips": sl_pips,
            "sl_usd": sl_usd,
            "tp_price": tp_price,
            "tp_pips": tp_pips,
            "volume": deal.volume,
            "result_pips": result_pips,
            "result_usd": deal.profit,
            "rr_plan": rr_plan,
            "rr_fact": rr_fact
        }
        transactions.append(trans)
    
    return transactions
