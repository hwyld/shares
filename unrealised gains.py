
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

## Fetch prices from Yfinance ##

#### WANT TO UPDATE THIS TO GET 1 YEAR OF DATA FOR EACH TICKER ####
#### ALSO WANT TO RUN A MTM ON THE PORTFOLIO ####

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
portfolio_df.to_csv('ticker_exceptions_market_prices.csv', index=False)

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
#portfolio_df['Starting Value $'] = portfolio_df['Quantity'] * portfolio_df['Average Cost $']

# Calculate the unrealised gains and losses (Quantity * Latest Price $ - Starting Value $)
portfolio_df['Unrealised Gain/Loss $'] = portfolio_df['Quantity'] * portfolio_df['Latest Price $'] - portfolio_df['Starting Value $']

# Calculate the current market value (Quantity * Latest Price $)
portfolio_df['Current Market Value $'] = portfolio_df['Quantity'] * portfolio_df['Latest Price $']

# Calcualte current weights using Market Value $ / Total Market Value $ 
portfolio_df['Current Weights'] = portfolio_df['Current Market Value $'] / portfolio_df['Current Market Value $'].sum()

# Add a row with the sum of 'Starting Value $', 'Unrealised Gain/Loss $', 'Market Value $', and 'Current Weights'
portfolio_df.loc['Total'] = pd.Series(portfolio_df[['Starting Value $','Cost Value $', 'Unrealised Gain/Loss $', 'Market Value $', 'Current Weights']].sum())

# Display the dataframe
print(portfolio_df)

# Export the dataframe to a CSV file
portfolio_df.to_csv('updated_portfolio.csv', index=False)