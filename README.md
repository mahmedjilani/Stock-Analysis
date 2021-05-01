
# Stock Analysis with VBA for Excel

The stock analysis and refactoring the code

---
## Overview of the project	
The purpose behind the project is to help the end-user to process the stock data and determine whether they are worth investing in by calculating total volume traded for the year and yearly return for each stock.
The data set consist of stock tickers with their daily volumes, high & lows, open & closing prices, and adjusted closing price for both **2017**  & **2018**.

In the previously written code, there was room for code improvement to make the code run faster, efficiently (low memory utilization), and making the code easy to understand for anyone who want to refactor it even further and that’s what I have tried to achieve in this project.

## Results

### Processing time 

After refactoring the code, the processing time was reduced to 0.179 seconds and saved us **2.761 seconds** for 2017 and **0.488 seconds** for 2018. (see below)

#### Original VBA Script

![Old_Code_2017](https://user-images.githubusercontent.com/82117986/116796676-5c9fb780-aaac-11eb-8acd-6a010a24c184.png)
![Old_Code_2018](https://user-images.githubusercontent.com/82117986/116796680-66c1b600-aaac-11eb-9cb0-4bbc899634e1.png)

#### Processing time – After Refactoring

![VBA_Challenge_2017](https://user-images.githubusercontent.com/82117986/116796696-953f9100-aaac-11eb-8f82-b1968c8e3753.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/82117986/116796697-9a044500-aaac-11eb-9deb-07e6dfd848c9.png)

### Code
#### Original VBA Script

The previous code used nested for loop to calculate our desired parameter’s i.e., starting price, closing price, and subsequently calculate yearly volumes & return. This was making the code redundant, and it was taking more time to process i.e. 2.941 seconds for 2017

![Old_VBA_Script](https://user-images.githubusercontent.com/82117986/116796705-b43e2300-aaac-11eb-828d-438cd49b80eb.png)

### Refactored Code.
The code was refactored to create arrays for all of them to save time (see below). Moreover, the whole code used single variable “I” for all the **FOR loops** and made it easier for any person to understand the true flow of the code. It is now replaced with unique codes for every FOR loop. (Refer to the code in vba) in the repository.


![New_Script_1](https://user-images.githubusercontent.com/82117986/116796713-c7e98980-aaac-11eb-8795-cefc8125f4ac.png)
![New_Script_2](https://user-images.githubusercontent.com/82117986/116796716-d041c480-aaac-11eb-847d-8702a7bd9ebc.png)

### Stock Performances (2017 vs 2018)
The dataset is available for 2 years only and for the scrips in consideration. It looks like 2018 has been a bad years for the stock market as a whole but only two scrips remained in bullish in 2018. i.e., ENPH and RUN.
Both the scrips didn’t perform as good as preceding year (2017) but at least remained strong when the whole market was performing bad. This makes my stock recommendation as buy for both **ENPH & RUN.**

![Stock_Performance_2017+2018png](https://user-images.githubusercontent.com/82117986/116796719-d9cb2c80-aaac-11eb-8421-24eb85acf3e7.png)

## Summary
### Refactoring Advantages
1.	Avoids reinventing the wheel and saves time.
You do not have to spend a lot of time thinking about the main flow of the code from scratch and gives you to opportunity to search for and implement the efficiencies only, hence saves time.
2.	Performance and Design enhancements
Refactoring gives you the chance to make the changes to design of the code without changing its core behavior or expected outcome.

### Refactoring Disadvantages.
The above advantages can be disadvantages at the same time but it solely depends how the code was written in the first place. 
1.	Can take time to make the code work	
If the engine of the code is poorly written, it can be difficult to navigate through the code and make changes. It can take more time than to write a new code from start. 
2.	Can convince you to think in the same way  as the previous code was written and hence can block the mind to think in a different way.

### Pros & Cons of Refactoring to original VBA Script
The original vba script provided me with the basic structure of the code so I did not have to make the logical flow of the code from start. The con was that the code used the same variables repeatedly for FOR Loop and it took me time to understand that part. Had the code was written with unique variables, it would have taken even lesser time for me to make the changes. I just had to replace the basic flow of the code to compute the results in slightly different way.

