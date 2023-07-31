
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

# Convert 'Date' column to datetime format before calling strftime
trades_df['Date'] = pd.to_datetime(trades_df['Date'])

# Get the start and end date
start_date = trades_df['Date'].min().strftime('%Y-%m-%d')
end_date = trades_df['Date'].max().strftime('%Y-%m-%d')

# Load the portfolio data
# Specify the columns to be read
columns_to_read = ["Security Code", "Company Name", "Quantity", "Last Price", "Average Cost $", "Market Value $"]
portfolio_df = pd.read_csv('PortfolioReport-Equities.csv', usecols=columns_to_read)
#portfolio_df = pd.read_csv('PortfolioReport-Equities.csv')
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
print(sales_df)
print(portfolio_df)
print(trades_df)
print('Hello')

#PORTFOLIO AS AT DATE#

# Add 'Starting Value $' column to new_trades_df
new_trades_df['Starting Value $'] = new_trades_df['Cost Value $'] + new_trades_df['Brokerage']
print(new_trades_df)
# Add 'Cost Value $' and 'Brokerage' columns to portfolio_df and fill them with zeros
portfolio_df['Cost Value $'] = 0
portfolio_df['Brokerage'] = 0

# Merge portfolio_df and new_trades_df
merged_portfolio_df = pd.concat([portfolio_df, new_trades_df], ignore_index=True)
print(merged_portfolio_df)

# Replace 'Starting Value $' column in merged_portfolio_df with 'Cost Value $' column
#merged_portfolio_df['Starting Value $'] = merged_portfolio_df['Cost Value $']

# Print the merged dataframe
#print(merged_portfolio_df)

## REMOVING DUPLICATES ##
# Convert 'Purchase Date' to datetime if it's not already
merged_portfolio_df['Purchase Date'] = pd.to_datetime(merged_portfolio_df['Purchase Date'])

# Group by 'Security Code', sum numerical columns, calculate weighted average of 'Average Cost $', and keep the earliest 'Purchase Date'
grouped_df = merged_portfolio_df.groupby('Security Code').apply(lambda x: pd.Series({
    'Quantity': x['Quantity'].sum(),
    'Average Cost $': (x['Average Cost $'] * x['Quantity']).sum() / x['Quantity'].sum() if x['Quantity'].sum() != 0 else 0,
    'Cost Value $': x['Cost Value $'].sum(),
    'Brokerage': x['Brokerage'].sum(),
    'Starting Value $': x['Starting Value $'].sum(),
    'Purchase Date': x['Purchase Date'].min()
}))

# Reset the index to turn 'Security Code' back into a column
grouped_df = grouped_df.reset_index()

print(grouped_df)

# YAHOO FINANCE TICKERS #

# Create empty dataframes to store the historical data
#yahoo_tickers_df = pd.DataFrame()
# Append '.AX' to each 'Security Code' in the dataframe to create 'Yahoo Finance Ticker'
grouped_df['Yahoo Finance Ticker'] = grouped_df['Security Code'] + '.AX'

# HISTORIC DATA FROM YFINANCE #

# Create empty dataframes to store the historical data
close_df = pd.DataFrame()
dividends_df = pd.DataFrame()
beta_df = pd.DataFrame()
volume_df = pd.DataFrame()
ticker_exceptions = []

# Loop over each ticker in the dataframe
for ticker in grouped_df['Yahoo Finance Ticker']:
    # Fetch the historical data
    data = yf.download(ticker, start=start_date, end=end_date)

     # Check if the data is empty
    if data.empty:
        ticker_exceptions.append(ticker)
        continue
    
    # Append the data to the respective dataframe
    close_df[ticker] = data['Close']
    #dividends_df[ticker] = data['Dividends']
    #beta_df[ticker] = data['Beta (5Y Monthly)']
    volume_df[ticker] = data['Volume']

# Export the ticker exceptions to a CSV file
pd.Series(ticker_exceptions).to_csv('ticker_exceptions.csv', index=False)
# Export the dataframes to CSV files
close_df.to_csv('close.csv')
#dividends_df.to_csv('dividends.csv')
#beta_df.to_csv('beta.csv')
volume_df.to_csv('volume.csv')


## NEED TO THEN ADD HOW TO SOURCE PRICES FOR MISSING TICKERS ##

## CHARTING ##

## MINIMUM VARIANCE PORTFOLIO ##

from scipy.optimize import minimize

# Calculate daily returns
returns = close_df.pct_change()

# Define the objective function
def portfolio_variance(weights, returns):
    # Weights should sum to 1
    weights = np.array(weights)
    portfolio_return = np.sum(returns.mean() * weights) * 252
    portfolio_volatility = np.sqrt(np.dot(weights.T, np.dot(returns.cov() * 252, weights)))
    return portfolio_volatility

# Constraints
cons = ({'type': 'eq', 'fun': lambda x: np.sum(x) - 1})

# Bounds
bounds = tuple((0, 1) for x in range(len(returns.columns)))

# Initial guess (equal distribution)
init_guess = [1/len(returns.columns) for x in range(len(returns.columns))]

# Run the optimizer
opts = minimize(portfolio_variance, init_guess, args=(returns), method='SLSQP', bounds=bounds, constraints=cons)

# Get the optimal weights
min_var_weights = opts.x
print(min_var_weights)

# Add the weights to the grouped_df DataFrame
grouped_df['MinVar Weights'] = min_var_weights

print(grouped_df)