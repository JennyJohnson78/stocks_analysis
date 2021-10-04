# Stocks Analysis
## Overview
Steve is a financial analyst who is seeking help to analyze stocks to determine which ones yield the greatest returns. There are 12 stocks he would like to have analyzed for the years of 2017 and 2018. The VBA code that was written for him will need to be refractored to allow for a more complex analysis that won't take a long time to run. 

## Results

![image](https://user-images.githubusercontent.com/67409852/135783121-0c9f54fc-1713-4d53-b865-fe69d0f4febd.png)

The Total Daily Volume is the number of trades of the given stock for a day, over the course of a year (2017 or 2018). The Return column is the stock values at the end of the year, divided by its price at the beginning of the year,
```
Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
```
showing if the stock gained value or lost value. Green signifies an increase in value, while red signifies a decrease in value. 
