# Green-Stocks

## Overview

For this project, I used VBA to help Steve perform an analysis for a dataset of stock. I started by creating Macros, which iterate through the 2017 and 2018 datasets and produce outcomes on the “All Stocks Analysis” page. These outcomes include the year being analyzed, the ticker id, total daily volume, and return. While the code does work, I am helping Steve refactor it so that it will run faster and increase efficiency if adding additional Macros. For this analysis, I used 2018 to compare my run times.

![VBA_Results](https://user-images.githubusercontent.com/111243284/188976994-aeb27d50-ec21-43da-a2d4-48183dc824b7.PNG)

## Results

The initial code I wrote had a nested for loop, which was used to first determine the number of tickers to loop through, and then the number of rows in the data. When I first ran the code, I had a run time of about 1.85 seconds. I used the following Macro to perform this analysis, but I knew there had to be a way to make the code run more efficiently.

![VBA_Challenge_2018 Original](https://user-images.githubusercontent.com/111243284/188977246-bbcaa090-3106-42cd-83ec-1112c09102c9.png)

As part of the refactoring, I created output arrays for tickerVolumes, tickerStartingPrices, and tickerEndingPrices. This allowed me to insert the “tickerIndex” as an input for each array and loop through all twelve tickers and row numbers one time. This proved to be a more efficient process since I got the same result, and my run time was cut down to about 0.21 seconds.

![VBA_Challenge_2018 Refactor](https://user-images.githubusercontent.com/111243284/188977504-89602c5e-2fa3-41b3-8500-10d17b98bc3d.png)


## Summary

 •	Advantages and Disadvantages of Refactoring Code
One advantage of refactoring code is that it improves efficiency by reducing run times for Macros. The main disadvantage with refactoring code is the time it takes to find a quicker solution. It took me several days to refactor the code for this project and I found it frustrating at times. I would suggest getting more comfortable with this skill before using it as a first step.

 •	Advantages and Disadvantages of Original and Refactored VBA script
I would reiterate what I mentioned in the previous question. This project proved to me that run times can make a difference in the efficiency of a project. However, I see lots of areas where I could improve my skills in debugging and allowing my self to get better at refactoring.
