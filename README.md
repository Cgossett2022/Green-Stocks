# Green-Stocks-Analysis

## Overview

For this project, I used VBA to help Steve perform an analysis for a dataset of stock. I started by creating Macros, which iterate through the 2017 and 2018 datasets and produce outcomes on the “All Stocks Analysis” page. These outcomes include the "Year" being analyzed, the "Ticker Id", "Total Daily Volume", and "Return". While the code does work, I am helping Steve refactor it so that it will run faster and increase efficiency if adding additional Macros. 

![VBA_Results](https://user-images.githubusercontent.com/111243284/188976994-aeb27d50-ec21-43da-a2d4-48183dc824b7.PNG)

![VBA_Results 2017](https://user-images.githubusercontent.com/111243284/188986507-1ea95fce-a822-43ed-902f-4d4ca2c4ac51.PNG)


## Results

The initial code I wrote had a nested for loop, which was used to first determine the number of tickers to loop through, and then the number of rows in the data. When I first ran the code, I had a run time of about 1.90 seconds for 2018 and 1.80 seconds for 2017. I used the following Macro to perform this analysis, but I knew there had to be a way to make the code run more efficiently.

![VBA_Challenge_2018 Original](https://user-images.githubusercontent.com/111243284/188982275-490472c2-5999-4410-ba89-e581a7a22d2d.png)

![image](https://user-images.githubusercontent.com/111243284/188986658-fe0ad133-8d7c-4e44-b5fe-07235bebe757.png)



As part of the refactoring, I created output arrays for tickerVolumes, tickerStartingPrices, and tickerEndingPrices. This allowed me to insert the “tickerIndex” as an input for each array and loop through all twelve tickers and row numbers one time. This proved to be a more efficient process since I got the same result, and my run time was cut down to about 0.32 seconds for 2017 and 0.21 seconds for 2018.

![VBA_Challenge_2018 Refactor](https://user-images.githubusercontent.com/111243284/188977504-89602c5e-2fa3-41b3-8500-10d17b98bc3d.png)

![image](https://user-images.githubusercontent.com/111243284/188986762-b9bfe1a1-b382-4dff-afa1-c0d53471b20e.png)



## Summary

### Advantages and Disadvantages of Refactoring Code
One advantage of refactoring code is that it improves efficiency by reducing run times for Macros. The main disadvantage with refactoring code is the time it takes to find a quicker solution. It took me several days to refactor the code for this project and I found it frustrating at times. I would suggest getting more comfortable with this skill before using it as a first step.

### Advantages and Disadvantages of Original and Refactored VBA script
I would echo what I mentioned in the previous question. This project proved to me that run times can make a difference in the efficiency of a project. However, I would really want to improve my refactoring skills, before taking on additional projecs.
