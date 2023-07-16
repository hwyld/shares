import yfinance as yf
import pandas as pd
import numpy as np

# Load the csv file
portfolio_df = pd.read_csv('PortfolioReport-Equities-504883-202307141154.csv')

# Create empty lists to store the Yahoo Finance tickers, latest prices, and exception tickers
yahoo_finance_tickers = []
latest_prices = []
yahoo_finance_tickers_exceptions = []

# Loop over each row in the dataframe to create a list of Yahoo Finance tickers
for index, row in portfolio_df.iterrows():
    yahoo_finance_tickers.append(row['Security Code'] + '.AX')

# Print the list of Yahoo Finance tickers
print(yahoo_finance_tickers)

# Loop over the list of Yahoo Finance tickers to fetch the latest prices
for ticker in yahoo_finance_tickers:
    try:
        ticker_info = yf.Ticker(ticker).info
        latest_price = ticker_info['regularMarketPreviousClose'] if 'regularMarketPreviousClose' in ticker_info else np.nan
        latest_prices.append(latest_price)
    except:
        yahoo_finance_tickers_exceptions.append(ticker)
        latest_prices.append(np.nan)

# Print the list of exception tickers
print(yahoo_finance_tickers_exceptions)

# Load the exception prices from the separate csv file
#exception_prices_df = pd.read_csv('ticker_exceptions_market_prices.csv')

# Loop over the exception tickers to replace the corresponding latest prices
#for ticker in yahoo_finance_tickers_exceptions:
    # Find the index of the ticker in the list of Yahoo Finance tickers
#    index = yahoo_finance_tickers.index(ticker)
    # Replace the corresponding latest price with the price from the exception prices dataframe
#    latest_prices[index] = exception_prices_df.loc[exception_prices_df['Ticker'] == ticker, 'regularMarketPreviousClose'].values[0]

print(latest_prices)


# Add the Yahoo Finance tickers and the latest prices to the dataframe
portfolio_df['Yahoo Finance Ticker'] = yahoo_finance_tickers
portfolio_df['Latest Price $'] = latest_prices

# Create a new column to indicate whether each ticker is an exception
portfolio_df['Is Exception'] = portfolio_df['Yahoo Finance Ticker'].isin(yahoo_finance_tickers_exceptions)

# Calculate the starting value (Quantity * Average Cost $)
portfolio_df['Starting Value $'] = portfolio_df['Quantity'] * portfolio_df['Average Cost $']

# Calculate the unrealised gains and losses (Quantity * Latest Price $ - Starting Value $)
portfolio_df['Unrealised Gain/Loss $'] = portfolio_df['Quantity'] * portfolio_df['Latest Price $'] - portfolio_df['Starting Value $']

# Display the dataframe
print(portfolio_df)

# Export the dataframe to a CSV file
portfolio_df.to_csv('updated_portfolio.csv', index=False)