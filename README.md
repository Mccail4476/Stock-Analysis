# Stock-Analysis
Module 2 of the UCSD data science bootcamp

# Overview (purpose and background)

***Purpose***
The goal of this project was to analyze multiple stocks in a quick & efficient process using VBA starting from a half-finished code. Since code is not usually efficient during its first draft, it's important that coders come back to optimize their code to reduce time & memory consumption of their computer(s). 

***Background***
Using a template of the macro used to analyze stock, we had to refactor the code so that it is an optimzed reversion of itself that displays the data correctly. By utlizing the skills taught in module 2, a macro can be not only be made but take existing ones and refactor them to improve effieicncy & output. The macro that was created throughout the module was only designed to account for a dozen of stocks for a given year. However, Steve wants to have the ability to test multiple timeframes for the same stocks in order to make a better prediction. It's very handy for any user to be able to know how the code works in order to make proper edits if need be. An example would be if there is a new stock that is to be considered in the data set. If Steve knows how to use VBA, he can not only add the new stock of interest, but he can modify the code or add to include other criteria that may be of interest when deciding which stock(s) to choose from. 

# results (screenshots and code)

***Inside look of the code***

![refactor code1](https://user-images.githubusercontent.com/99565016/156870408-e37c1bac-a388-4803-862f-527476d4f5ff.png)

![refactor code2](https://user-images.githubusercontent.com/99565016/156870498-b4dd6dde-d0a6-4294-a34f-fcad63ffe64f.png)

The code goes through each ticker that's assigned a value (for i = 0 to 11) and goes through each row of the entire raw data set of the respective year that the user is asking. For each row, the code is adding the volume for each data point to find the total volume. Furthermore, by going row-by-row, the coe is looking for the starting price & ending price in order to determine returns, which can be seen near the bottom of the screenshot. Afterward, it formats the data correctly before color coding the returns to make the data visually presentable. If the returns are below 0% for the selected year, it is displayed red. If the returns are above 0% for the selected year, it is displayed green.

***Displaying Results***

![original_challenge_2017](https://user-images.githubusercontent.com/99565016/156870855-84305d42-a19e-4f80-a9f3-94cf4300d6e9.PNG)

**The original code output for 2017**

![original_challenge_2018](https://user-images.githubusercontent.com/99565016/156870864-9fe8f0b0-052c-430e-9e7a-90f6afc142de.PNG)

**The original code output for 2018***

We can see from the above images of the original code that they do not display any data on the excel sheet, rendering the macro useless. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/99565016/156870825-9486eeb9-10ec-44b4-973a-75648d939b3c.png)

**The refactored code output for 2017**

![VBA_Challenge_2018](https://user-images.githubusercontent.com/99565016/156870842-a163e98a-3af9-4480-aa34-695ee0980f16.png)

**The refactored code output for 2018**

We can see from the above images of the refactored code that the excel sheet has actual data & is colorized for easy viewing.

Based off the data received after running the macro, we can see that of all stocks analyzed, "ENPH" had the most success for both years (2017 & 2018). Steve would be wise to recommend to his parents that they should consider investing into "ENPH". A follow up stock to consider is "RUN", though this stock does not seem to have as high returns. It is best to use historical data to make better predictions of how a stock will perform rather than using abbreviations similiar to Dairy Queen. 


# Summary 

***Refactoring Code from a general viewpoint***
An advantage of refactoring code is that one can correct the mistakes, rather by themselves or their peers. Since coding can be team projects, it is important that a coder can be able to not only fix their previous code, but be able to fix others' mistakes. Another advantage is that refactoring code can help improve new objectives on the projects. Since the idea of refactoring is taking existing code, most of the work has already been done. Simply add more lines or modify existing lines to get the desired result(s). 

A disadvantage is that refactoring code is a time-consuming process. Unless the coder is very famliar with the code, they can take hours or days to optimize the code to running condition. It may even be better to simply restart the code from scratch depending on the length of the inefficient code & how many errors are present. Another disadvantage to consider is increasing the output of the code can lead waiting longer for the output.

***Original versus Refactored VBA***
We saw from the above images that the original VBA code that Steve was using didn't output any data due omission of code. By refacotring the original code, Steve can now see the desired results. Another advantage is that coders can optimize the data output to be more efficient, such as color coding or formatting the output correctly. As a data analyst, it is important to display output data so anyone can quickly get the story & be able to make a decision.

A disadvantage to refactoring this original code is that since our refactored code has a higher output, it needs more time. Code needs to be efficient, and a method to determine if a code is efficient is how fast does it take the computer to run the program. A longer timeframe to run the code may be consider inefficient! 
