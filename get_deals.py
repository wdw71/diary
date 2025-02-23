import MetaTrader5 as mt5
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill
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
        print(f"Processing deal: Ticket={deal.ticket}, Position ID={deal.position_id}")
        
        if deal.position_id not in positions:
            positions[deal.position_id] = {
                "entry_price": 0,
                "volume": 0,
                "sl_price": 0,
                "tp_price": 0,
                "count": 0,
                "entry_time": None,
                "close_price": 0
            }
        
        pos = positions[deal.position_id]

        # Only consider entry trades for these calculations
        if deal.entry == 0:  # Opening or increasing position
            pos["entry_price"] += deal.price * deal.volume
            pos["volume"] += deal.volume
            pos["sl_price"] += deal.price if deal.type == 0 else 0  # Buy orders
            pos["tp_price"] += deal.price if deal.type == 0 else 0  # Sell orders
            pos["count"] += 1
            if pos["entry_time"] is None or deal.time < pos["entry_time"]:
                pos["entry_time"] = deal.time

        # Only consider closing trades for close price calculation
        if deal.entry == 1:  # Closing or reducing position
            pos["close_price"] = deal.price
            # pos["close_price"] += deal.price * deal.volume

    for position_id, pos in positions.items():
        pos["entry_price"] /= pos["volume"] if pos["volume"] > 0 else 1
        pos["sl_price"] /= pos["count"] if pos["count"] > 0 else 1
        pos["tp_price"] /= pos["count"] if pos["count"] > 0 else 1
        # pos["close_price"] /= pos["volume"] if pos["volume"] > 0 else 1

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

def export_transactions_to_excel(transactions):
    if not transactions:
        print("No transactions to export.")
        return

    file_path, _ = QFileDialog.getOpenFileName(None, "Select Excel File", "", "Excel Files (*.xlsx)")
    if not file_path:
        return

    try:
        wb = openpyxl.load_workbook(file_path)
    except Exception as e:
        print(f"Error opening file: {e}")
        return

    sheet_names = wb.sheetnames
    sheet_name, ok = QInputDialog.getItem(None, "Select Sheet", "Choose a sheet:", sheet_names, 0, False)
    if not ok or not sheet_name:
        return

    ws = wb[sheet_name]

    existing_tickets = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1]:
            existing_tickets.add(str(row[1]))

    row_to_write = ws.max_row + 1
    for row in range(2, ws.max_row + 1):
        if not ws.cell(row=row, column=2).value:
            row_to_write = row
            break

    # Color fills
    yellow_fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
    green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

    for trans in transactions:
        if trans["ticket"] in existing_tickets:
            continue

        row = [
            trans["date"],       # A - Transaction Date (DD.MM.YYYY)
            trans["ticket"],     # B - Ticket
            trans["position_id"],# C - Position ID
            trans["action"],     # D - Opening/Closing
            trans["instrument"], # E - Instrument
            trans["entry_time"], # F - Entry Time (HH:MM:SS)
            "",                  # G (empty)
            trans["side"],       # H - Buy/Sell
            trans["entry_price"],# I - Entry Price
            trans["close_price"],# J - Entry Price
            trans["sl_price"],   # K - SL Price
            trans["sl_pips"],    # L - SL in Pips
            trans["sl_usd"],     # M - SL in Pips
            trans["tp_price"],   # N - TP Price
            trans["tp_pips"],    # O - TP in Pips
            trans["volume"],     # P - Volume
            trans["result_pips"],# Q - Result in Pips
            trans["result_usd"], # R - Result in USD
            trans["rr_plan"],    # S - Winrate Plan
            trans["rr_fact"]     # T - Winrate Fact
        ]

        ws.append(row)
        current_row = ws.max_row

        # Apply conditional formatting
        if trans["action"] == "Open":
            for col in range(1, len(row) + 1):
                ws.cell(row=current_row, column=col).fill = yellow_fill
        elif trans["action"] == "Close":
            fill = green_fill if float(trans["result_usd"]) > 0 else red_fill
            for col in range(1, len(row) + 1):
                ws.cell(row=current_row, column=col).fill = fill

    try:
        wb.save(file_path)
        print(f"Transactions successfully added to sheet '{sheet_name}' in {file_path}")
    except Exception as e:
        print(f"Error saving file: {e}")