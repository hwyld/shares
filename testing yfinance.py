
# Import yfinance module in the program  
import yfinance as yahooFin  
# Using ticker for the Facebook in yfinance function  
retrFBInfo = yahooFin.Ticker("FB")  
# Printing Facebook financial information in the output  
print(retrFBInfo.info)  