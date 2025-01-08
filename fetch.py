import requests
import pandas as pd
import time
from openpyxl import Workbook

# Function to fetch live cryptocurrency data
def fetch_crypto_data():
    url = 'https://api.coingecko.com/api/v3/coins/markets'
    params = {
        'vs_currency': 'usd',
        'order': 'market_cap_desc',
        'per_page': 50,
        'page': 1,
        'sparkline': False
    }
    response = requests.get(url, params=params)
    data = response.json()
    
    # Create a DataFrame
    df = pd.DataFrame(data)
    df = df[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]
    df.columns = ['Cryptocurrency Name', 'Symbol', 'Current Price (USD)', 'Market Capitalization', '24h Trading Volume', '24h Price Change (%)']
    
    return df

# Function to analyze data
def analyze_data(df):
    # Print the top 50 cryptocurrencies
    print("Top 50 Cryptocurrencies:")
    print(df[['Cryptocurrency Name', 'Symbol', 'Current Price (USD)', 'Market Capitalization', '24h Trading Volume', '24h Price Change (%)']])
    
    top_5 = df.nlargest(5, 'Market Capitalization')
    average_price = df['Current Price (USD)'].mean()
    highest_change = df['24h Price Change (%)'].max()
    lowest_change = df['24h Price Change (%)'].min()
    
    # Print analysis results to terminal
    print("\nTop 5 Cryptocurrencies by Market Cap:")
    print(top_5[['Cryptocurrency Name', 'Symbol', 'Market Capitalization']])
    print(f"\nAverage Price of Top 50 Cryptocurrencies: ${average_price:.2f}")
    print(f"Highest 24h Price Change: {highest_change:.2f}%")
    print(f"Lowest 24h Price Change: {lowest_change:.2f}%\n")
    
    analysis = {
        'Top 5 Cryptocurrencies by Market Cap': top_5,
        'Average Price of Top 50 Cryptocurrencies': average_price,
        'Highest 24h Price Change (%)': highest_change,
        'Lowest 24h Price Change (%)': lowest_change
    }
    
    return analysis

# Function to save data and analysis to Excel
def save_to_excel(df, analysis):
    with pd.ExcelWriter('crypto_data.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Crypto Data')
        
        # Save analysis results
        analysis_summary = pd.DataFrame({
            'Metric': ['Average Price', 'Highest 24h Price Change (%)', 'Lowest 24h Price Change (%)'],
            'Value': [analysis['Average Price of Top 50 Cryptocurrencies'], 
                      analysis['Highest 24h Price Change (%)'], 
                      analysis['Lowest 24h Price Change (%)']]
        })
        
        analysis_summary.to_excel(writer, index=False, sheet_name='Analysis')
        
        # Save top 5 cryptocurrencies
        analysis['Top 5 Cryptocurrencies by Market Cap'].to_excel(writer, index=False, sheet_name='Top 5')

# Main function to run the process
def main():
    while True:
        df = fetch_crypto_data()
        analysis = analyze_data(df)
        save_to_excel(df, analysis)
        print("Data and analysis saved to 'crypto_data.xlsx'. Updating in 5 minutes...\n")
        time.sleep(300)  # Wait for 5 minutes

if __name__ == "__main__":
    main()