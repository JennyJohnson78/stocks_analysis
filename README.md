# Stocks Analysis
## Overview
Steve is a financial analyst who is seeking help to analyze stocks for his clients to determine which ones yield the greatest returns. There are 12 stocks he would like to have analyzed for the years of 2017 and 2018. The VBA code that was written for him will need to be refractored to allow for a more complex analysis that won't take a long time to run. 

## Results

![image](https://user-images.githubusercontent.com/67409852/135783121-0c9f54fc-1713-4d53-b865-fe69d0f4febd.png)

The Total Daily Volume is the number of trades of the given stock for a day, over the course of a year (2017 or 2018). The Return column is the stock values at the end of the year, divided by its price at the beginning of the year,
```
Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
```
showing if the stock gained value or lost value. Green signifies an increase in value, while red signifies a decrease in value. From the data, it can be determined that more stocks increased in value in 2017, with only the stock "TERP" losing value. The top stocks for 2017 were "DQ" and "SEDG". However, in 2018, most stocks lost value and are "in the red". Only "ENPH" and "RUN" gained value, with "DQ", the stock Steve's clients were originally interested in. Due to our finidings, it would be beneficial to share this information with his clients and suggest a different stock to invest in, perhaps "RUN", which is the only stock that gained value from 2017 to 2018.

The original code, consisting of nested loops, took over a second to iterate through all of the data Steve wanted analyzed. By using refractored code, the time the code ran for was cut drastically.

Instead of nested loops, an array of the stocks was created and iterated over:

![image](https://user-images.githubusercontent.com/67409852/135784805-314ad767-c3e3-4eb6-97d7-d911c758b8fe.png)

Because the nested loops have now been replaced, these are the run times for each year's code:

![image](https://user-images.githubusercontent.com/67409852/135784508-aeccdf58-4fa0-4241-bfac-e1a7ad5a993d.png)

## Summary
