**Deliverable 2 Requirements**

1. **Overview of the Project**

The purpose of refactoring this code was to analyze the performance of the given stocks over the years 2017 and 2018. Additionally, we also refactored the code that we previously used so that it may run in a more efficient manner.

2. **Results**

The below links showed the Code times for each year before and after they were refactored.
*2017 Original Code Time*
(https://github.com/JoePizz/stock-analysis/blob/main/2017%20Original%20Code%20Time.png)
*2018 Original Code Time*
(https://github.com/JoePizz/stock-analysis/blob/main/2018%20Original%20Code%20Time.png)

*2017 Refactored Time*
(https://github.com/JoePizz/stock-analysis/blob/main/VBA_Challenge_2017.png)

*2018 Refactored Time*
(https://github.com/JoePizz/stock-analysis/blob/main/VBA_Challenge_2018.png)

As you can see the code after it was refactored ran more than 4 times quicker than the original code. Additionally, in refactoring the code more comments were left making it easier to read for the next person that needs to use it.
Using the tickerIndex variable to access the correct index across the four arrays that were created, we were able to refactor the original VBA Code to make it more efficient.

**2017 and 2018 Comparison**

[Link to the Volume and Results for 2017 and 2018]
2017: (https://github.com/JoePizz/stock-analysis/blob/main/Charts%202017.png)
2018: (https://github.com/JoePizz/stock-analysis/blob/main/Charts%202018.png)
In comparing the years 2017 and 2018 we can see that the stocks we analyzed have been analyzing as a whole performed worse than they did overall in the year 2017. There were only 2 stocks that were positive in both 2017 and 2018 

**ENPH**

2017 Return: 129.5%

2018 Return: 81.9%


**RUN**

2017 Return: 5.5%
2018 Return: 84%


3. **Advantages and Disadvantages of Refactoring Code**

3 - 1) Refactoring code can be advantageous as it saves us time in our analysis and presenting our results. Additionally we are able to work on the given code to make it more efficient and easier to follow or read. We can see this present in the Module code versus the challenge code. The module code ran in 0.945 seconds for the 2018 analysis and in 0.891 seconds for the 2017 analysis.
When we refactored the code though we were able to run each year in significantly less time. 2018 ran in 0.223 seconds and 2017 was able to run in just 0.395 seconds.
The disadvantages of refactoring code are that it can bring about new bugs or errors in the code, it may create oversights in your analysis that would have been realized had you created the code on your own and in the long run there is a possibility that it creates bad habits in the coder's process.

3 - 2) As mentioned above these pros do apply to our original VBA code as we were able to reduce the run time in the code significantly, making our code run more efficiently. 
In regards to disadvantages, I ran into a situation where I was receiving an error "subscript out of range" for quite a while. I was analyzing the code I wrote and could not find any issues. In the end it was due to the original code spelling a sheet slightly different than I had. Although refactoring can be more efficient, this small oversight costed me a lot of time.
