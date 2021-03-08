# Stock-Analysis with VBA

## Overview of Project
Steve is a recent graduate with a finance degree.  He has to give financial investment advice to his parents who are his first clients. They want to invest in a green energy stock for a solar panel company called DAQO.  The tickerindex ID for DAQO is DQ.  

Steve is aware that his parents are passionate about green energy since they believe that alternative energy is the future.  Steve is also aware that his parents have not done their research about DQ.  He would like to evaluate the performance over the year 2017 and 2018 for all solar panel stocks. 

Steve would also like a VBA code that would process faster than the original script so he can analyze a larger dataset incase he chooses to expand his research about what stocks are good for investment.  In general, the process of changing the current code is called **refactoring**. 

### Purpose
The purpose of this project is to:

1. To create a VBA script so that Steve could analyze data for DQ stock as compared to other solar panel stocks over the year 2017 and 2018. This is necessary to make an informed decision about investment in the company. 

2. Refactor the original VBA Script so that it would work faster.  This would allow Steve to execute a larger dateset more efficiently.

## Analysis of Results of 2017 and 2018 stocks
Images *2017-Green Energy Returns.png* and *2018-Green Energy Returns.png* in the *Resources* folder are used for this part of the analysis.  Looking at an the data between 2017 and 2018, it is clear that even thought DQ stock had a striking 199% return in 2017, it has dropped drastically to -62.6% in 2018.  This would not be a good investment for Steve's parents.

Looking closely at the other options for solar panel stocks presented to Steve, the only two stocks making a profitable return in 2018 had the tickers ENPH and RUN.  

It is important to note that daily volume for ENPH and RUN was 2.7 and 1.8 times more respectively, in 2018 than in 2017.  ENPH stock had a return of 129.5% and 81.4% in 2017 and 2018 respectively. RUN's stock had returns of 5.5% and 84.0% for years 2017 and 2018 respectively.  

## Results and Summary of Refactoring 
Images *"AllStockAnalysis"-OriginalVBA-NestedForLoop.png* and *"All Stock Analysis"-Refactor VBA -For Loops Only.png* in the *Resources* folder show the differences between the codes.  In the original code there is one for loop and a nested loop, the code runs through 3012 lines before going back into another ticker index in the original code.  In the refactored code there are two For Loops that are used to run the code.

### Advantages and disadvantage of refactoring code in general
According to *Quora*, the advantages of refactoring a code is that it can boost system performance because of the simplicity of the code, the processing time would be faster. (1).  You can also refactor the code to fit the changing needs of a business (1).

The disadvantages of refactoring is that it takes time to do so and sometimes it is easier to new code from scratch than to refactor it (1).

### Pros and cons apply to refactoring the original VBA script
Pros for refactoring the current VBA code was that the code can run more efficiently.  The original code had nested loops which means the loop goes through the data set for all the rows, which is 3012 rows for the current data set.  When refactored code had only two For Loops which reduces the processing time drastically.  When compairing images *VBA_Original_2017* and *VBA_Challenge_2017* it can be seen that the refactored code is 8.7 times faster than the original code.  Similarly it is evident from images *VBA_Original_2018* and *VBA_Challenge_2018* in the *Resources* folder, that the refactored code is 8.5 times faster than the orginal code.

Cons of refactoring the original code was that in the case that Steve does not need to use it for other stocks, that time could have been saved.

## Recommendations
Upon analysis it is recommended that if Steve needs to pick a solar energy stock it should be the stock RUN, as there is a steady increase in their daily volume and returns from 2017 to 2018 which shows growth of a company.  Even though the rate of returned shows a decline for ENPH between the year 2017 to 2018, the daily volume of that stock has increased.  Steve may need to do more research on the policies and investores within the company to make a more informed decision.

If Steve decides to run a larger dataset, he should definitely use the refactored VBA script.

It is recommended that Steve's parents broaden their portfolio with other stocks using the new refractored VBA script.  

## Reference
(1)What are the pros and cons of refactoring?.*Quora*.  Accessed on March 6, 2021.  https://www.quora.com/What-are-the-pros-and-cons-of-refactoring
