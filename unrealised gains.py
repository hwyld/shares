
import pandas as pd         # Pandas for data manipulation
import yfinance as yf       # Yahoo Finance API
import numpy as np          # Numpy for numerical computing
import openpyxl as open     # Openpyxl for reading excel files


## Creating the Portfolio ##

# Convert the Excel file to CSV
trades_df = pd.read_excel('AllTradesReport 0107022 30062023.xlsx', sheet_name='Combined')
trades_df.to_csv('AllTradesReport.csv', index=False)

# Load the trades data from the new CSV file
trades_df = pd.read_csv('AllTradesReport.csv')

# Convert the 'Date' column to datetime format and store it as 'Trade Date'
trades_df['Trade Date'] = pd.to_datetime(trades_df['Date'])

# Get the start and end date
start_date = trades_df['Date'].min().strftime('%Y-%m-%d')
end_date = trades_df['Date'].max().strftime('%Y-%m-%d')

# Load the portfolio data
portfolio_df = pd.read_csv('PortfolioReport-Equities.csv')
portfolio_df['Realised Gain/Loss $'] = 0  # Initialize the new column with zeros
portfolio_df['Starting Value $'] = portfolio_df['Quantity'] * portfolio_df['Average Cost $']  # Define the starting value

# Create a new DataFrame to store the new trades
new_trades_df = pd.DataFrame(columns=['Security Code', 'Quantity', 'Average Cost $', 'Cost Value $', 'Brokerage', 'Purchase Date'])

# Prepare an empty DataFrame to store the sales data
sales_df = pd.DataFrame(columns=['Security Code', 'Realised Gain/Loss $', 'Sale Date', 'Return %'])

# Iterate through each trade
for i, trade in trades_df.iterrows():
    # Find the corresponding holding in the portfolio
    holding = portfolio_df.loc[portfolio_df['Security Code'] == trade['Code']]
    
    if holding.empty:  # If the holding doesn't exist in the portfolio, add it to new_trades_df
        new_trade = pd.DataFrame({
            'Security Code': [trade['Code']],
            'Quantity': [trade['Qty']] if trade['Qty'] > 0 else [0],
            'Average Cost $': [trade['Price']],
            'Cost Value $': [trade['Qty'] * trade['Price']] if trade['Qty'] > 0 else [0],
            'Brokerage': [trade['Brokerage']] if trade['Qty'] > 0 else [0],
            'Purchase Date': [trade['Trade Date']]
        })
        new_trades_df = pd.concat([new_trades_df, new_trade])
    else:  # If the holding exists in the portfolio, adjust its size
        index = portfolio_df.loc[portfolio_df['Security Code'] == trade['Code']].index[0]
        if trade['Qty'] > 0:  # If it's a purchase, increase the quantity and adjust the average cost
            total_cost = portfolio_df.loc[index, 'Quantity'] * portfolio_df.loc[index, 'Average Cost $'] + trade['Qty'] * trade['Price']
            total_qty = portfolio_df.loc[index, 'Quantity'] + trade['Qty']
            portfolio_df.loc[index, 'Average Cost $'] = total_cost / total_qty
            portfolio_df.loc[index, 'Quantity'] = total_qty
            portfolio_df.loc[index, 'Starting Value $'] = portfolio_df.loc[index, 'Quantity'] * portfolio_df.loc[index, 'Average Cost $']  # Update the starting value
        else:  # If it's a sale, decrease the quantity and calculate the return %
            starting_value = -trade['Qty'] * portfolio_df.loc[index, 'Average Cost $']  # Starting value is the cost of the shares sold
            sale_amount = -trade['Qty'] * trade['Price']                        # Sale amount is the sale price of the shares sold
            realised_gain_loss = sale_amount - starting_value                  # Realised gain/loss is the difference between the sale amount and the starting value
            return_pct = realised_gain_loss / starting_value               # Return % is the realised gain/loss divided by the starting value
            sale = pd.DataFrame({
                'Security Code': [trade['Code']],
                'Sale Amount': [sale_amount],
                'Purchase Amount': [starting_value],
                'Realised Gain/Loss $': [realised_gain_loss],
                'Sale Date': [trade['Trade Date']],
                'Return %': [return_pct]
            })
            sales_df = pd.concat([sales_df, sale])
            portfolio_df.loc[index, 'Realised Gain/Loss $'] += realised_gain_loss
            portfolio_df.loc[index, 'Quantity'] += trade['Qty']  # Qty is negative for sales
            if portfolio_df.loc[index, 'Quantity'] <= 0:  # If the quantity is zero or less, remove the holding
                portfolio_df.drop(index, inplace=True)

# Add a row with the sum of 'Realised Gain/Loss $', 'Sale Amount' and 'Purchase Amount'
sales_df.loc['Total'] = pd.Series(sales_df[['Realised Gain/Loss $','Sale Amount','Purchase Amount']].sum())


# Export the sales data to a CSV file
sales_df.to_csv('sales_data.csv', index=False)

print(new_trades_df)
print(portfolio_df)
print(trades_df)

#PORTFOLIO AS AT DATE#




## TRACK CASH BALANCE , DIVIDENDS , INTEREST PAID , WITHHOLDING TAX##

# Load the cash transactions data
cash_df = pd.read_csv('CashTransactionSummary.csv')

# Convert the 'Date' column to datetime format
cash_df['Date'] = pd.to_datetime(cash_df['Date'], dayfirst=True)


# Initialize the new columns
cash_df['Interest'] = 0
cash_df['Dividends'] = 0
cash_df['Dividends Security Code'] = ''
cash_df['Tax'] = 0
cash_df['Bank Transfers'] = 0

# Iterate through each cash transaction
for i, row in cash_df.iterrows():
    if 'MACQUARIE CMA INTEREST PAID' in row['Description']:
        cash_df.loc[i, 'Interest'] = row['Credit $']
    elif 'DIV' in row['Description']:
        cash_df.loc[i, 'Dividends'] = row['Credit $']
        cash_df.loc[i, 'Dividends Security Code'] = row['Description'].split(" ")[0]
    elif 'AUSTRALIAN RESIDENT WITHHOLDING TAX' in row['Description']:
        cash_df.loc[i, 'Tax'] = row['Debit $']
    elif 'TRANSACT FUNDS TFR' in row['Description']:
        cash_df.loc[i, 'Bank Transfers'] = row['Credit $'] - row['Debit $']

# Sum the dividends for each security code
dividends_sum = cash_df.groupby('Dividends Security Code')['Dividends'].sum()

# Convert the Series to a DataFrame
dividends_df = dividends_sum.reset_index()

# Rename the columns to match the portfolio dataframe
dividends_df.columns = ['Security Code', 'Dividends']

print(cash_df)
print('hello')



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