  ![Bull](Images/charging-bull.jpg)

## Summary

I will use VBA scripting to analyze stock market data using this [test data](Resources/alphabetical_testing.xlsx) while developing my scripts and this [stock data](Resources/Multiple_year_stock_data.xlsx) to run my scripts to generate the final report.

### Stock market analyst

  ![Stock Market](Images/stockmarket.jpg)

I created a script that looped through all the stocks for one year and output the following:

  * The ticker symbol.
  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The total stock volume of the stock.

I also have conditional formatting that will highlight positive changes in green and negative changes in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

### CHALLENGES

1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:

![hard_solution](Images/hard_solution.png)

2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

### Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

## Submission

* To submit please upload the following to Github:

  * A screen shot for each year of your results on the Multi Year Stock Data.

  * VBA Scripts as separate files.

* Ensure your repository has regular commits (i.e. 20+ commits), a thorough README.md file

* After everything has been saved, create a sharable link and submit that to <https://bootcampspot-v2.com/>.

## Rubric

[Unit 2 Rubric - VBA Homework - The VBA of Wall Street](https://docs.google.com/document/d/1OjDM3nyioVQ6nJkqeYlUK7SxQ3WZQvvV3T9MHCbnoWk/edit?usp=sharing)

- - -

© 2021 Trilogy Education Services, LLC, a 2U, Inc. brand. Confidential and Proprietary. All Rights Reserved.
