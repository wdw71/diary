import MetaTrader5 as mt5

def connect_to_mt5(login, password, server):
    try:
        login_int = int(login)
    except ValueError:
        print("Invalid login: must be an integer.")
        return False

    if not mt5.initialize(login=login_int, password=password, server=server):
        print(f"Failed to connect to MT5: {mt5.last_error()}")
        return False

    print("Connected to MT5 successfully.")
    return True
