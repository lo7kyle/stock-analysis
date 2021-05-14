# VBA of Wall Street

## Overview of Project
In this challenge we are given a scenario where we have a friend named Steve whose parents are interested in investing in Green Energy Stocks. We are tasked with analyzing a stock of their choice DAQO and other Green energy stocks to see if the stocks are a good investment or not. Using some basic VBA macros and stock analysis for yearly return we have created and optimized a Macro for analyzing stocks given specific parameters of a stock. 

### Purpose
The main purpose of this challenge is to develop a deeper understanding of VBA and the utilization of Macros to perform simple analysis of stock prices. With a better understanding of statistics and stock market indicators (doji,EPS,P/E,etc.) VBA can become a more powerful tool with even greater analysis than what we did. In the challenge specifically we had to optimize our code to have a faster runtime. 

## Analysis and Challenges
The challenge of this week was to refactor our code that we created from the module. We had to optimize the runtime which is important for real life scenarios where we will have many more data points than in our module. 

### Analysis of Original vs. Refactored
In the module we utilized a nested for-loop to loop through an array of cells in the worksheet to be analyzed. We had a defined array of tickers we were looking for, then we looped through each ticker and each row until the end of the data set to add the daily volumes and find the starting and last closing price to calculate the yearly return. From college I learned that the BigO notation for for-loops is O(n) where n is the number of iterations needed to be calculated. Assuming 1 sequence is 0(1), we have to iterate N times and if we have another for-loop, say M, the time complexity is exponential N * M times N^2. This gives you O(N^2). See below for the original VBA code:

[Original VBA code](https://github.com/user/repo/blob/branch/other_file.md)

In the challenge we must take the gained knowledge and refactor our code to lessen the time. As explained previously a for-loop has time complexity of O(n). For the challenge we had to create multiple for-loops and store the data into arrays. I would have never thought of this method and was extremely intrigued and excited to see such a cool work around for increasing performance. I've ran into this common algorithm problem before in the binomial coefficient problem. In this problem, if we have nth terms, our runtime will be n^n which is poorly written code. From binomial expansion we see that our next coefficient will always be (n-1) meaning we have already calculated the term previously. With our VBA code, we took the values from our for-loop and stored them into an array that was already created. In recursive programming you take the previous value and add it to the new value having to recalculate the first value then adding it to the new calculated value. This repetition is the opposite of what coding is supposed to be for. There is a term called dynamic programming where have the previous data is stored and can be used in the current calculation to reduce repetition. In our case we created empty arrays and stored our single for-loop values into that specific array. Since single for loops are O(n) having multiple for-loops would only be N * O(n) which would give us a shorter time than something exponential. See below for the array implementation. 

[Refactored VBA code](https://github.com/user/repo/blob/branch/other_file.md)

### Challenges and Difficulties Encountered
There were many challenges and difficulties that I created for myself. I initially wanted to make this Macro dynamic and scale with nth amount of ticker symbols. I tried implementing a tickers function that kept giving me an overflow error, so I completely removed that function. I also ran into multiple overflow errors when trying to keep the tickerVolumes, tickerStartingPrices, tickerEndingPrices dynamic. This ended up working by using a separate loop and using the ReDim function that allows you to resize the dimensions of your array. Another issue I ran into was trying to create a function to store a user's input. I was trying to create a button that asked for the year and when inputting the year, you can click the analysis button and get our analysis for that current year. For some reason I had trouble passing the function's stored data value into another Macro. In the original I created a function that called itself but that is just repetitive. All in all, to make this modular to work for any stocks and any amount of data was the limitation of my knowledge although I learned more than I needed researching these different functions and errors.  

## Results
From the results we can clearly see that the Refactored code is a lot faster than the nested for-loop code. See below for a comparison. The main takeaway from these results is that we want to avoid any code that will take an exponential amount of time to iterate through the loops, but at the same time make it as robust as possible. 

[Original 2017 Runtime](https://github.com/user/repo/blob/branch/other_file.md)
[Refactored 2017 Runtime](https://github.com/user/repo/blob/branch/other_file.md)

[Original 2018 Runtime](https://github.com/user/repo/blob/branch/other_file.md)
[Refactored 2018 Runtime](https://github.com/user/repo/blob/branch/other_file.md)

## Summary
### Advantages Vs Disadvantages Original
There are many advantages and disadvantages to both methods. An advantage would be that a nested for-loop is easier to implement. It is quite simple to follow, and you know what you are getting by following the iterations. This is an advantage over the arrays method because it is scalable at the cost of runtime. What I mean by scalable is we can know as little about our dataset and get the result we want. The disadvantage of the refactored code is that we need to know our variables. We had to create an additional index to keep count of and I felt like there can be more room for errors in terms of overflow because of the arrays. Another advantage would be that for a smaller dataset where runtime is minimal the nested for-loop method can be used and would save you debug time if you end up with more bugs by trying to refactor your code. 

### Advantages Vs Disadvantages Original
The advantage of the refactored code is that the runtime is only a multiple (N*O(n)) and not an exponential. When reading many more thousand lines of code or having a larger dataset (1GB+) it will take an exceptionally long time to calculate. The disadvantage to all this is that you must understand your dataset a bit more. You must know what you want from your data set and it cannot be as robust and modular as you want. In other words, you must write that code for a specific case like for our challenge, 12 green energy stocks. By creating multiple arrays, you are prone to more overflow errors and scalability by having to Redim the arrays. 