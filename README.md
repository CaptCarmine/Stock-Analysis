# Stock-Analysis

## Overview of Project
 <!-- Explain the purpose of this analysis. -->
Steve needs to analyse the yearly stock data. Hoping to find the total volume of each stock ticker, as well as find the % return. He wants this process of retrieving volume and return to be completed autonomously. We are looking to create a VBA macro, that will complete the process based on a yearly spreadsheet, procured to us by Steven.

This macro will have two requirements

 1. It will have an button that accepts a year input. That way steven can adds additional years in the future.
 2. Two the macro will be a single run of the data, so that as steven adds more data it will run smoothly.

## Results

<!-- Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script. -->
![intialrefactoredcode2017](https://github.com/CaptCarmine/Stock-analysis/blob/main/Resources/2017_timegraph.png?raw=true)

We were able to complete the project in the span of a week. See above for initial run of the refactored code. We brought the time it took from over a second to .35 seconds. Getting it to further reduce down to .138 seconds. We were able to change this by making an important change to the code.

![VBA_Challenge2017](https://github.com/CaptCarmine/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)

```VBA
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
        End If
    Next i
```

As seen above, The final code used a variable for the index that changed the logic as we were going through the data. Initially we had to loop through the entire dataset for each ticker. Using this logic we only needed to loop through the entire dataset once, rather than as many times as there are tickers. Allowing for great impovement in the time taken, which as the sheets steve provides get larger, the change will get more noticable.

## Summary

<!-- In a summary statement, address the following questions.
1.What are Advantages/Disadvantages of Refactoring code?
  -the Advantages are?
2.How do these pros and cons apply to refactoring the original VBA script?-->

In this project we learned the importance and many advantages of refactoring code to be more efficient, as well as possible problems.

### Pros

- The most important thing we learned about refactoring code, is the improvement in efficiency.
- The code we ran only had to loop through the entire dataset once, instead of 11 times for this data set. if the data set had 30 tickers the refactored code would have still only needed to run through the dataset once, Rather than 30 times , which could have made the macro significantly slower, or caused excel to freeze.

### Cons

- The large con to refactoring code is that it only works in this example if the data is cleanly organized. If the data set wasnt organized it would have capped at the 12 tickers for the array and stopped before running through all the data.
- with the original, non-refactored script it will find all the data for volume, but would fail for starting and ending values, if the dataset was not organized.

## Final thoughts

I think this is a great project to show how much value can be added by refactoring code, and as well as shows how important it is to clean the data before applying logic.
