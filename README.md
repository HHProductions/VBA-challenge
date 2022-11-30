# VBA-challenge
Week 2 assignment
The script created performs the required task on the provided multiyear stock data file in approximately 1 minute. (processing power dependent)
For the script to work all tickers need to be sorted, as it will register each change in value (row to row) as a new ticker.  
The script could have been modified to also look through the newly created unique ticker list, before adding a new unique ticker, but teh assumption was that the tickers are sorted.
The bonus script was written so that in the unlikely event that  multiple tickers share the same "maximum increase", "maxiumu decrease" value or  the same maximum trading  volume, the valeu/volume will be displayed,  but instead of specifying a ticker, it will state 
    # tickers shared the greatest % increase
    # tickers shared the greatest % decrease
    # tickers shared the greatest total volume
where # is the number of tickers.
A message box appears at the end of the main script and another at the end of bonus test.
