import tkinter as tk
from tkinter import filedialog
import pandas as pd
import re
from collections import defaultdict
import requests
from currency_converter import CurrencyConverter
import os
import time
import sys
from colorama import init, Fore, Style

# Init colorama
init(autoreset=True)

# Initialize CurrencyConverter without cache
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CURRENCY_DATA_DIR = os.path.join(BASE_DIR, 'currency_converter')
CURRENCY_DATA_PATH = os.path.join(CURRENCY_DATA_DIR, 'eurofxref-hist.zip')
LAST_UPDATE_FILE_PATH = os.path.join(BASE_DIR, 'last_update_time.txt')
# CMC
CMC_API_KEY = "3e5a1c85-1d9c-4a4a-b955-e1d2e827be32" 

def initialize_currency_converter():
    if not os.path.exists(CURRENCY_DATA_DIR):
        os.makedirs(CURRENCY_DATA_DIR)

    if not os.path.exists(CURRENCY_DATA_PATH):
        print("Downloading eurofxref-hist.zip...")
        download_currency_data()

    if os.path.exists(CURRENCY_DATA_PATH):
        last_update_time = load_last_update_time()
        if time.time() - last_update_time > 8 * 3600:
            print("Refreshing currency data...")
            os.remove(CURRENCY_DATA_PATH)
            download_currency_data()

    return CurrencyConverter()

def download_currency_data():
    url = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.zip"
    try:
        response = requests.get(url)
        with open(CURRENCY_DATA_PATH, 'wb') as file:
            file.write(response.content)
        save_last_update_time()
    except Exception as e:
        print(f"Error downloading currency data: {e}")

def load_last_update_time():
    try:
        with open(LAST_UPDATE_FILE_PATH, 'r') as file:
            last_update_time = float(file.read())
    except (FileNotFoundError, ValueError):
        last_update_time = 0
    return last_update_time

def save_last_update_time():
    with open(LAST_UPDATE_FILE_PATH, 'w') as file:
        file.write(str(time.time()))

def select_file():
    root = tk.Tk()
    root.withdraw()
    root.update_idletasks()
    root.call('wm', 'attributes', '.', '-topmost', True)
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    root.update()
    root.destroy()
    return file_path

def extract_amount_and_currency(value):
    match = re.match(r"([-+]?\d*\.?\d+)\s*(\w+)", str(value))
    if match:
        amount = abs(float(match.group(1)))
        currency = match.group(2)
        if currency.endswith('FIAT'):
            currency = currency[:-4]
        return amount, currency
    return 0, ""

def get_price_from_coinmarketcap(symbol):
    if symbol == 'BCD':
        return 1
    elif symbol == 'JB':
        return 0
    elif symbol == 'BCL':
        return 0.1

    url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest"
    parameters = {"symbol": symbol.upper()}
    headers = {"Accepts": "application/json", "X-CMC_PRO_API_KEY": CMC_API_KEY}

    response = requests.get(url, params=parameters, headers=headers)
    data = response.json()

    if response.status_code == 200 and data['status']['error_code'] == 0:
        symbol_upper = symbol.upper()
        if symbol_upper in data['data']:
            btc_price = data['data'][symbol_upper]['quote']['USD']['price']
            return btc_price
    return None

def print_currency_totals(totals, converter, unrecognized_currencies):
    total_usd_sum = 0
    for currency, total in totals.items():
        if currency in converter.currencies:
            total_usd = converter.convert(total, currency, 'USD')
            total_usd_sum += total_usd
            print(f"Total for {currency}: {total:.2f} {currency} (converted to {total_usd:.2f} USD)")
        else:
            price_from_coinmarketcap = get_price_from_coinmarketcap(currency)
            if price_from_coinmarketcap is not None:
                total_usd = total * price_from_coinmarketcap
                total_usd_sum += total_usd
                print(f"Total for {currency}: {total:.2f} {currency} (converted to {total_usd:.2f} USD)")
            else:
                unrecognized_currencies.append((currency, total))
                print(f"Total for {currency}: {total:.2f} {currency}")

    print("------------------------------------------------------------------")
    print(Fore.CYAN + f"Total USD value: {total_usd_sum:.2f} USD")
    if unrecognized_currencies:
        print(Fore.RED + "Unrecognized currencies:")
        for currency, total in unrecognized_currencies:
            print(Fore.RED + f"{total} {currency}")
    print("------------------------------------------------------------------")

def main():
    file_path = select_file()
    if file_path:
        df = pd.read_excel(file_path)
        df['Create Date'] = pd.to_datetime(df['Create Date'], errors='coerce')

        if 'Create Date' in df.columns:
            min_logs_date = df['Create Date'].min()
            max_logs_date = df['Create Date'].max()

            if all(col in df.columns for col in ['real money change amount', 'Description', 'UID', 'Create Date']):
                uids = df['UID'].unique()
                if len(uids) == 1:
                    uid_value = uids[0]
                    filtered_df = df[df['Description'].isin(['Original Bet', 'Original War', 'Third Party Bet', 'Trade Bet-Contest', 'Trade Bet-Contract', 'Trade Bet-Order', 'Sports Bet', 'Horse Bet', 'Lottery Lotter Purchase'])]

                    min_transaction_date = filtered_df['Create Date'].min()
                    max_transaction_date = filtered_df['Create Date'].max()

                    totals = defaultdict(float)
                    for value in filtered_df['real money change amount']:
                        amount, currency = extract_amount_and_currency(value)
                        if currency:
                            totals[currency] += amount

                    print("******************************************************************")
                    print(f"File successfully read: {file_path}")
                    print("******************************************************************")
                    print(Fore.GREEN + f"UID: {uid_value}")
                    print("------------------------------------------------------------------")
                    print(f"Time frame (Overall): {min_logs_date} - {max_logs_date}")
                    print(f"Time frame (Filtered): {min_transaction_date} - {max_transaction_date}")
                    print("------------------------------------------------------------------")

                    unrecognized_currencies = []

                    converter = initialize_currency_converter()
                    # Process Third Party Win
                    third_party_wins_df = df[df['Description'] == 'Third Party Win']
                    third_party_totals = defaultdict(float)
                    for value in third_party_wins_df['real money change amount']:
                        amount, currency = extract_amount_and_currency(value)
                        if currency:
                            third_party_totals[currency] += amount

                    print("Third Party Win:\n")
                    print_currency_totals(third_party_totals, converter, unrecognized_currencies)
                    print("Total Wagering:\n")
                    print_currency_totals(totals, converter, unrecognized_currencies)

                    while True:
                        choice = input("Manually calculate the timeframe or recalculate (y/n/r)? ")

                        if choice.lower() in ['y', 'ㄗ']:
                            print("------------------------------------------------------------------")
                            start_time = input("Please enter start time (YYYY-MM-DD): ")
                            end_time = input("Please enter end time (YYYY-MM-DD): ")
                            try:
                                start_time = pd.to_datetime(start_time)
                                end_time = pd.to_datetime(end_time)
                                filtered_df = df[(df['Create Date'] >= start_time) & (df['Create Date'] <= end_time) & df['Description'].isin(['Original Bet', 'Original War', 'Third Party Bet', 'Trade Bet-Contest', 'Trade Bet-Contract', 'Trade Bet-Order', 'Sports Bet', 'Horse Bet', 'Lottery Lotter Purchase'])]

                                totals = defaultdict(float)
                                for value in filtered_df['real money change amount']:
                                    amount, currency = extract_amount_and_currency(value)
                                    if currency:
                                        totals[currency] += amount

                                # Process Third Party Win
                                third_party_wins_df = df[df['Description'] == 'Third Party Win']
                                third_party_totals = defaultdict(float)
                                for value in third_party_wins_df['real money change amount']:
                                    amount, currency = extract_amount_and_currency(value)
                                    if currency:
                                        third_party_totals[currency] += amount

                                unrecognized_currencies.clear()  # Clear unrecognized currencies list for each new calculation
                                print("------------------------------------------------------------------")
                                print("Third Party Win:\n")
                                print_currency_totals(third_party_totals, converter, unrecognized_currencies)
                                print("Total Wagering:\n")
                                print_currency_totals(totals, converter, unrecognized_currencies)


                            except Exception as e:
                                print(f"Error parsing dates: {e}")

                        elif choice.lower() in ['n', 'ㄙ']:
                            print("Exiting...")
                            sys.exit()

                        elif choice.lower() in ['r', 'ㄐ']:
                            main()
                            return
                        else:
                            print("Invalid choice, please enter y/n/r.")
                else:
                    print("Multiple UIDs found, expected only one UID.")
            else:
                print("Columns 'real money change amount', 'Description', 'UID', or 'Create Date' not found")
        else:
            print("Column 'Create Date' not found")
    else:
        print("No file selected")

if __name__ == "__main__":
    main()
