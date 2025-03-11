Project Overview:
This project automates tracking of portfolio transactions and dividend updates using Google Apps Script and Python in Google Sheets.


Purpose:
The purpose of making this project is to real time monitoring of client's portfolio across different region , multi asset classes and different currencies. I have created a dashboard in google sheet with portfolio attributes of diffrent client. The system is built using google sheets with backend coding is done in google app script. The portfolio attributes for different clients are displayed in a dashboard of google sheet having overall view of investment, asset allocation, transaction summary's performance metrics, portfolio performance metrics.

Data Source and Processing:

1. Input.csv:
   • Contains transaction details for each client.
   • Data is taken from the client's brokerage account.
   • you can refer the sheet for the data attributes and the last attributes which is live price is directly 
     incorporated from the security symbol using google sheet formula.


2. Dividend Summary.csv:
   • Stores dividend payout details for each client.
   • Data is generated by Dividend_Summary.gs 


Google App ScripT Implementation 

1. Portfolio_ Summary.gs:
   
   • Fetches transaction data and dividend data from Input.csv and Dividend Summary.csv.
   • Performs backend calculation for key portfolio metrics which could be seen in Portfolio Transa.
   




