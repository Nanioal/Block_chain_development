# -*- coding: utf-8 -*-
"""
@author: Rediet Bekele
CoinGecko API: Fetch live data into Excel
"""
import pandas as pd
from pycoingecko import CoinGeckoAPI
import time

# Create a CoinGecko API client
cg = CoinGeckoAPI()

def fetch_and_save_data():
    # Fetch the top 50 cryptocurrencies by market capitalization
    try:
        coinsMktDataByIds = cg.get_coins_markets(vs_currency='usd', 
                                                 per_page=50, 
                                                page=1,  
                                                order='market_cap_desc')
        print("Data fetched sucessfully.")

    # Convert the data to a pandas DataFrame
        df = pd.DataFrame(coinsMktDataByIds)
        print("Data converted to Data frame")
    # Select the required fields
        selected_columns = df[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]
        # Display the data
        print("Top 50 Cryptocurrencies:")
        print(selected_columns.head(50))  # Display the top 50 cryptocurrencies
    # Save the DataFrame to an Excel file
        selected_columns.to_excel('top_50_cryptos_live_data.xlsx', index=False)
        print("Data saved to Excel!")
        top_5_by_market_cap = selected_columns.sort_values(by='market_cap', ascending=False).head(5)
        print("\nTop 5 Cryptocurrencies by Market Capitalization:")
        print(top_5_by_market_cap[['name', 'market_cap']])

    # 2. Calculate the average price of the top 50 cryptocurrencies
        average_price = selected_columns['current_price'].mean()
        print(f"\nAverage Price of the Top 50 Cryptocurrencies: ${average_price:.2f}")

# 3. Analyze the highest and lowest 24-hour percentage price change among the top 50
        highest_24h_change = selected_columns.loc[selected_columns['price_change_percentage_24h'].idxmax()]
        lowest_24h_change = selected_columns.loc[selected_columns['price_change_percentage_24h'].idxmin()]

        print(f"\nHighest 24-hour Percentage Price Change: {highest_24h_change['name']} ({highest_24h_change['symbol']}) - {highest_24h_change['price_change_percentage_24h']:.2f}%")
        print(f"Lowest 24-hour Percentage Price Change: {lowest_24h_change['name']} ({lowest_24h_change['symbol']}) - {lowest_24h_change['price_change_percentage_24h']:.2f}%")

# -----------------------------------------
# Optional: Save the results to an Excel file
# -----------------------------------------
        selected_columns.to_excel('top_50_cryptos_analysis.xlsx', index=False)

    except Exception as e:
        print(f"An error occured: {e}")

if __name__ == '__main__':
    while True:
        fetch_and_save_data()
        print("Waiting for the next update...")
        time.sleep(300)  # Wait for 5 minutes (300 seconds)
    