# stock-analysis

## Overview of Project
We have analyzed and organzied data on stocks to help Steve help his parents with investing.  The goal of this project was to expand the dataset from focusing on a dozen stocks to  incorporating the entire dataset over the last few years.  

### Purpose
In this challenge, we practiced the skill of refactoring.  We wanted our original analysis to run more efficiently.  Our goal was to loop through all the data at once for 2017 and 2018 to collect the same information in a quicker time and make it easier to read and analyze.  

## Analysis and Challenges
 Using a refactored code, we are quickly able to get a break down of all the data for 2017 and 2018.  See below:
 ![image](https://user-images.githubusercontent.com/64279232/124983608-2d567b00-e006-11eb-83be-e886c3057252.png)
![image](https://user-images.githubusercontent.com/64279232/124983839-7e666f00-e006-11eb-8e42-8f55c59e1fed.png)

We can see through this analysis that 2017 had a much better rate of return compared to 2018. 


### Challenges and Difficulties Encountered
 I enjoyed trying different things with the code this time around compared to the first time we ran the code.  I did face some challenges that I think will benefit me in the long run.  For example, the for loops are still complicated for me to fully grasp and execute.  It takes a lot of trial and error and research until I get it to flow correctly.  Once I had the code completed, I kept getting different types of error messages when I tried running the code.  My biggest issued ended being a mixup between the sheets I activated.  I switched the yearValue and "All Stocks Analysis" up which caused the hearder row to end up on the third row of the "2017" sheet rather than the "All Stocks Analysis" sheet. 
 
 See below:
 ![image](https://user-images.githubusercontent.com/64279232/124983214-adc8ac00-e005-11eb-9b68-f5dbb08db592.png)
 
## Results
Below are the execution times before refactoring the code with the original script:
![image](https://user-images.githubusercontent.com/64279232/124984073-c71e2800-e006-11eb-8d96-d8dcea6e5e8b.png)
![image](https://user-images.githubusercontent.com/64279232/124984096-ceddcc80-e006-11eb-962d-e10ecbdf7363.png)
It took 0.65625 seconds for 2017 to run and 0.6757813 seconds for 2018 to run. 

Below are the execution times for the refactored script:
![image](https://user-images.githubusercontent.com/64279232/124984384-254b0b00-e007-11eb-83d9-c5c43c0ebcb7.png)
![image](https://user-images.githubusercontent.com/64279232/124984439-385ddb00-e007-11eb-8c61-5a8904df2209.png)
It took 0.0625 seconds for 2017 and 0.0546875 seconds for 2018. 

Based on this data and the color codes were positive and negative rate of returns, it is very clear to see that 2017 was a much more successful year for stocks than 2018.  In 2017, only one stock (TERP) had a negative rate of return.  The rest were all positive even with 4 stocks over 100%.  In 2018, only two stocks had a positive rate of return (ENPH and RUN).  

##Summary
-What are the advantages or disadvantages of refactoring code?
Refactoring takes a lot of time and it can cause errors along the way that you might not have had to start with.  However, it makes the code run a lot more efficiently and quicker with fewer steps. 

-How do these pros and cons apply to refactoring the original VBA script?
The pros and cons of refactoring also applied to this specific VBA script.  I did have some errors while refactoring that mixed up a few things in the spreadsheet for a bit.  I had to redo parts, and this could have caused inaccurate results if I did not realize it earler.  With the refactoring, all we had to do was run the code and enter the year we wanted to run the analysis on in the MsgBox.  See below: 
![image](https://user-images.githubusercontent.com/64279232/124986048-2c731880-e009-11eb-927d-1e4bc025ee0a.png)


