# Stock-analysis
week 2 module over VBA

## Overview: 
The goal of this analysis was to use excels macros developed through VBA, to provide and accurate analysis of the the difference is stock values over the course of 2017 and 2018 to provide the clients with a better view of how their DQ stocks are preforming. 

## Results: 
The results of the stocks analysis showed that the performance of all the stocks in 2017, except for RUN, decreased their returns over the course of 2018, and that the stock DQ, has a 262% decrease over the course of the year, the most significant decrease out of all the stock’s return. The results of the refactored all stocks analysis macro vs. the original showed a significant increase in speed by going from approximately .72 seconds per run in the original code to .12 seconds for the refactored code


This improvement in run time was accomplished by removing the nested for loops and instead using arrays to bypass the nested for loops.  
(For loop script pictures here) 
The result leads to less complicated processes and pushes the codes optimization significantly.  


## Advantages and disadvantages of refactoring: 

Refactoring is a crucial part of coding for the fact that an initial script or program will never be as optimized as it can be at its first draft. You have to set up the variables and the logic so it flows in a concise and straightforward manner. However, our logic can be flawed or inefficient so one of the main advantages of refactoring is that it offers a chance to refine the code to be more efficient and easier on the computer. Inefficient code can cause simple tasks to take a needlessly long amount of time and resources during the process. Refactoring allows for optimization, reducing run time and simplifying task. This can be seen in the  difference in all stocks analysis and all stocks analysis refactored code. This refactoring has also helped with the readability. 
A disadvantage of refactoring code is that changing the initial arguments can cause the original skeleton of the code to not function. Refactoring can mean reconstructing the script from scratch and making the original code obsolete. A majority of the disadvantages of coding revolve around the fact that a lot of its original functionality has to be reworked to accommodate the coding. This can be seen with removing the nest for loops and replacing them with arrays.

