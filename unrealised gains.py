import yfinance as yf
import pandas as pd
import numpy as np

# Load the csv file
portfolio_df = pd.read_csv('PortfolioReport-Equities-504883-202307141154.csv')

# Create an empty list to store the latest prices
latest_prices = []

# Loop over each row in the dataframe
for index, row in portfolio_df.iterrows():
    # Fetch the latest price from Yahoo Finance
    ticker_info = yf.Ticker(row['Security Code']).info
    latest_price = ticker_info['regularMarketPrice'] if 'regularMarketPrice' in ticker_info else np.nan
    # Append the latest price to the list
    latest_prices.append(latest_price)

# Add the latest prices to the dataframe
portfolio_df['Latest Price $'] = latest_prices

# Calculate the starting value (Quantity * Average Cost $)
portfolio_df['Starting Value $'] = portfolio_df['Quantity'] * portfolio_df['Average Cost $']

# Calculate the unrealised gains and losses (Quantity * Latest Price $ - Starting Value $)
portfolio_df['Unrealised Gain/Loss $'] = portfolio_df['Quantity'] * portfolio_df['Latest Price $'] - portfolio_df['Starting Value $']

# Display the dataframe
print(portfolio_df)
