# Refactoring Stock Workbook VBA Code

## Overview of Project
For Module 2, a solution code was prepared and given to Steve to help better analyze the stock his parents had invested into. The workbook created allowed for Steve to analyze the dataset with the press of one button created in the workbook. However, the dataset only contained data for only around a dozen or so stocks and Steve wants to expand it to include the entire stock market over the last few years. The solution code created may have worked efficiently for around a couple of stocks however, it may not work for thousands of stock options. Our purpose for this challenge is to refactor the code to ensure it will run in a shorter amount of time and compare the time differences with the old Module 2 solution to allow for faster analysis outputs.

## Results

Upon refactoring code to allow for an increased ticker volume, we were able to analyze the data in a fraction of time according to the messagebox displaying the time it took to run the code.
### 2017 (Solution Code)
![](Resources/green_stock_analysis_2017.PNG)

### 2017 (Refactored Code)
![](Resources/VBA_Challenge_2017.png)

Running the refactored code showed that prior to refactoring, the 2017 analysis code took around .84 seconds to complete analysis. The 2017 refactored code meanwhile, took around .16 seconds to run. The difference is significantly noticeable as the refactored code managed to perform the analysis at 1/5 the speed it took the original solution code to perform.  

### 2018 (Solution Code)
![](Resources/green_stock_analysis_2018.PNG)

### 2018 (Refactored Code)
![](Resources/VBA_Challenge_2018.png)

Upon comparison of the two codes when it comes to the 2018 stock analysis data, the images shows that the original code took roughly .83 seconds to complete whereas the refactored code came in at roughly 1/5 of that time at .17 seconds again.

##Summary

The difference in speed alone proves that this new edited version of the code is more efficient espeically time efficient in analyzing stock value. In other terms, the new code ran the anaysis five times faster. In the world of cliche economics, time certainly is money and any business would certainly agree that having the same amount of work performed five times faster is to the betterment of the company, its interests, and to the consumers it caters to. This statement serves as an advantage to refactoring a code for an efficient analysis overall. Also, by incorporating a greater index, it allows for the code to analyze a larger dataset thus leading to an overall boost in performance, quality control, and less time spent on to debug the code. 

As we can see in the refactored code, we see a speed improvement by 5 times in comparison to the original code thus being an advantage of the refactoring process in general. Another advantage of refactoring the code is that we can create an index capable of factoring a greater amount of data, thus leading to a more streamlined and concise VBA code. Since the overall VBA code is shorter, it can be considered more efficient in debugging as well. Fewer lines of code leads to fewer amounts of error and can lead to a greater amount of applicatin to multiple datasets thus leading to a more reusable code with a broader range of applications.

Refactoring a code is not without its disadvantages as well. It can create more errors than the previous one due to an increase/ decrease in the dataset, can be less time efficient and may be more difficult to interpret. It can be moer time consuming in the sense the refactored code may seem streamlined but is actually pulling a greater amount of data thus leading to an increase in analysis time. The idea of a refactored code can seem appealing in its reusability aspect but it may not always apply in a similar analysis due to a slight difference in data i.e. one dataset may contain a few more filters which can lead to the script not accounting for those factors. It can also be a disadvantage in the sense it may introduce more errors that needs to be debugged since, as mentioned before, the newwer, streamlined version may not account for all factors in a certain worksheet/ dataset. All in all, in this specific scenario, the refactored code allows for Steve to process data five times faster than before, much to his betterment.


