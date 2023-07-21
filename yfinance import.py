import pandas as pd         # Pandas for data manipulation
import yfinance as yf       # Yahoo Finance API
import numpy as np          # Numpy for numerical computing
import openpyxl as open     # Openpyxl for reading excel files


## Fetch prices from Yfinance ##

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

# Add an empty column to the dataframe to store the market prices for the exception tickers


# Export the list of exception tickers to a CSV file
#
portfolio_df.to_csv('ticker_exceptions.csv', index=False)

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

# Calculate the unrealised gains and losses (Quantity * Latest Price $ - Starting Value $)
portfolio_df['Unrealised Gain/Loss $'] = portfolio_df['Quantity'] * portfolio_df['Latest Price $'] - portfolio_df['Starting Value $']

# Calculate the current market value (Quantity * Latest Price $)
portfolio_df['Current Market Value $'] = portfolio_df['Quantity'] * portfolio_df['Latest Price $']

# Calcualte current weights using Market Value $ / Total Market Value $ 
portfolio_df['Current Weights'] = portfolio_df['Current Market Value $'] / portfolio_df['Current Market Value $'].sum()

#### DIVIDENDS FROM YFINANCE ####
# Create an empty DataFrame to store the dividends data
dividends_df_yfinance = pd.DataFrame()

# Loop over each row in the dataframe to fetch the dividends from Yahoo Finance
for index, row in portfolio_df.iterrows():
    try:
        ticker = yf.Ticker(row['Yahoo Finance Ticker'])
        dividends = ticker.dividends.loc[start_date:end_date]
        temp_df = pd.DataFrame(dividends)
        temp_df['Yahoo Finance Ticker'] = row['Yahoo Finance Ticker']
        dividends_df_yfinance = pd.concat([dividends_df_yfinance, temp_df])
    except:
        pass

# Rename the columns
dividends_df_yfinance.columns = ['Dividends', 'Yahoo Finance Ticker']

# Export the dataframe to a CSV file
dividends_df_yfinance.to_csv('dividends_portfolio.csv', index=False)

# Sum the dividends for each Yahoo Finance Ticker
dividends_df_yfinance = dividends_df.groupby('Yahoo Finance Ticker')['Dividends'].sum().reset_index()

# Set 'Yahoo Finance Ticker' as the index in both dataframes
portfolio_df.set_index('Yahoo Finance Ticker', inplace=True)
dividends_df_yfinance.set_index('Yahoo Finance Ticker', inplace=True)

# Merge the dividends_df_yfinance with portfolio_df
portfolio_df = pd.merge(portfolio_df, dividends_df_yfinance, left_index=True, right_index=True, how='left')

# Reset the index
portfolio_df.reset_index(inplace=True)

# Calculate the dividends (Quantity * Dividends)
portfolio_df['Dividends'] = portfolio_df['Quantity'] * portfolio_df['Dividends']

# Fill NaN values with 0
portfolio_df['Dividends'] = portfolio_df['Dividends'].fillna(0)

# Merge the dividends dataframe with the portfolio dataframe where yfinance doesn't cover the dividends
# portfolio_df = pd.merge(portfolio_df, dividends_df, on='Security Code', how='left')

# Now merge the 'Dividends' from dividends_df into portfolio_df where 'Dividends' in portfolio_df is 0
portfolio_df.loc[portfolio_df['Dividends'] == 0, 'Dividends'] = portfolio_df['Security Code'].map(dividends_df.set_index('Security Code')['Dividends'])

# Fill NaN values with 0
portfolio_df['Dividends'] = portfolio_df['Dividends'].fillna(0)


#### Export the Portfolio ####

# Add a row with the sum of 'Starting Value $', 'Unrealised Gain/Loss $', 'Market Value $', 'Current Weights', 'Dividends'
portfolio_df.loc['Total'] = pd.Series(portfolio_df[['Starting Value $','Cost Value $', 'Unrealised Gain/Loss $', 'Market Value $', 'Current Weights','Dividends']].sum())

# Display the dataframe
print(portfolio_df)

# Export the dataframe to a CSV file
portfolio_df.to_csv('updated_portfolio.csv', index=False)


# HISTORIC DATA FROM YFINANCE #

# Create empty dataframes to store the historical data
close_df = pd.DataFrame()
dividends_df = pd.DataFrame()
beta_df = pd.DataFrame()
volume_df = pd.DataFrame()

# Loop over each ticker in the dataframe
for ticker in portfolio_df['Yahoo Finance Ticker']:
    # Fetch the historical data
    data = yf.download(ticker, start=start_date, end=end_date)
    
    # Append the data to the respective dataframe
    close_df[ticker] = data['Close']
    dividends_df[ticker] = data['Dividends']
    beta_df[ticker] = data['Beta (5Y Monthly)']
    volume_df[ticker] = data['Volume']

# Export the dataframes to CSV files
close_df.to_csv('close.csv')
dividends_df.to_csv('dividends.csv')
beta_df.to_csv('beta.csv')
volume_df.to_csv('volume.csv')