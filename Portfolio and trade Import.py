
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





