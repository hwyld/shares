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


