# VBA-challenge

Write a VBA module to complete a basic analysis of stock year over year. The module shall meet the following requirements:

 * Return the ticker symbol.
 
 * The yearly change over the given year.
 
 * The percent change over the same period.
 
 * The total volume of activity over the same period.
 
 * Use conditional formatting that will highlight the yearly positive change in green and yearly negative change in red.
 
 * The data will be displayed on the same worksheet as the data as an adjacent table.

Bonus requirements, provide a view into the yearly data that shall meet the following requirements:

 * The stock with the Greatest % increase.
 
 * The stock with the Greatest % decrease.
 
 * The stock with the Greatest total volume.

## Files in the VBA-challange repo

The following files are in the VBA-challenge repo

 * Starting files:
 
  * alphabetical_testing.xlsx
   
    * This is the initial development workbook with a small subset of the actual data to be used for development and testing.
   
  * Multiple_year_stock_data.xlsx
    
    * This is the initial production workbook to be used after development has been completed.
  
  * alphabetical_testing.xlsm
   
    * This is the development workbook after development and testing of the module.
   
  * Multiple_year_stock_data.xlsx
    
    * This is the production workbook after development and testing of the module.
  
  * StockSummary.bas
  
    * The module developed to meet the requirements detailed above.
  
  * Screenshot StockSummary 2014.png
    
    * A screenshot showing a small portion of the raw data, a small portion of the summary data, and the bonus table for 2014.

  * Screenshot StockSummary 2015.png
    
    * A screenshot showing a small portion of the raw data, a small portion of the summary data, and the bonus table for 2015.

  * Screenshot StockSummary 2016.png
    
    * A screenshot showing a small portion of the raw data, a small portion of the summary data, and the bonus table for 2016.

## Formulas used

The two formulas used in this module are the following.

 * Yearly change = Closing Value - Opening Value
 
 * Percent change = (Closing Value - Opening Value) / Opening Value
 
   * The above formula is reduced in the module to Yearly Change / Opening Value.

## Design decisions

In writing this module it was decided to use a While condition WEnd loop construct instead of the For Next construct. This was based more on experience having used the While contruct throughout most of my T-SQL development when looping contruct was deemed the appropriate choice.

## The design process

In designing the module, it was appropriate to write the code to work on a single sheet. This allowed for the development of the process that would then be implemented over all the sheets in an Excel Workbook.

### Basic algorithm used

1. Initalize two variables to be used in the while loop: InputRow = 2, OutputRow = 1

2. Capture the number of rows in the spreadsheet in the variable LastRow, this represents the end of the data to be processed.

3. Capture the Ticker Symbol, the Opening Value, and the Stock Volume from the first row of data. The Stock Volume will be the initial value in the TotalStockVolume accumulator.

4. Begin the While loop with the condition InputRow <= LastRow
   
   1. The first statement in the loop will be IF ELSE construct.
      
      If the ticker value in the current row equals the ticker value in the next row, capture the closing value of the stock in the next row and add the stock volume to the value currently in the TotalStockVolume accumulator, else we have accumulated all the data needed to complete the aggregation of the summary data.
      
      The yearly change and percent change will be calculated at this time, the TotalStockVolume has already been accumulated.  The OutputRow will be incremented by 1 (one) to insert the summary data in the first data row of the summary table.
      
      After writing the summary data the Stock Ticker, Opening Value, and Stock Volume (reinitializing the TotalStockVolume) will be captured from the next row (InputRow + 1) of data.

   2. The InputRow variable is incremented by 1 (one).

### Adding additional processing.

After initial testing it appeared that the algorithm worked as designed. It was time to add additional code to meet the requirements of the program.

1. Added the headers for the summary data before the While loop.

2. Added the appropriate formatting to the summary data when it was written to the summary table, this included the color coding of the yearly change in green or red, displaying the percent change as a percentage to 2 (two) decimal places.

### Final testing.

After completing the changes final testing for single sheet processing was then completed and everything still looked good.

At this point a For Each Next loop was added around the existing code.  This was necessary as all processing done for the first sheet needed to be accomplished on each subsequent sheet.

### Problem!

While testing the new code it was discovered that there was a problem as the code aborted while processing the data in the final sheet.  Making use of carefully placed conditionals and message boxes the offending data was quickly found.  It turns out the data for one of the stocks listed was all 0 (zero).  This resulted in an illegal operation, 0/0.

Once this found, a conditional statement was placed around the percent change calculation, if the opening value was equal to 0 (zero) set the percent change to 0, else make the normal calculation.

### On, another problem!

While working the divide by zero issue, it was also determined that there was another problem, the original algorithm had a flaw.  We were only capturing the Stock Ticker, Opening Value, and the Stock Volume.  All the sample data had multiple rows of data for each stock.  What if there was only a single row of data for a particular stock.  The algoritm had a flaw.

To resolve this flaw the Closing Value of the first row of data had to be captured as well.  This was added prior to the While loop as well as in the While loop.

### Final testing.

After making the needed changes, final testing was then completed successfully.

## Bonus work.

At this time it was decided to add the bonus work.  How was the the additional data going to be generated.  After considering adding another loop to the module to process the Summary data, a different approach was decided.

Instead of looping through the summary data, we determine each of the bonus category data as the summary data was written.  One of the decisions that also needed to be made at this time was what data would be captured: the first Stock that met the criteria, the last Stock that met the criteria, all of the Stock items if there were a tie.  The last option was discarded at this time as it would require additional testing and work to insure the data was correctly captured. Capturing the first or last was a fairly simple process and determined only be the conditional used.

To accomplish this we first create another table on the spreadsheet with Column Headers Ticker and Value and row headers of Greatest % Increase, Greatest % Decrease, and Greatest Total Volume.

In the While loop after we write the summary data for a Stock item, we then test the current summary data against the bonus table.  This resulted in 3 (three) simple if constructs. Yes, two of them could be combined but this was simpler.

The tests were simple, compare the PercentChange or TotalStockVolume against the data in the bonus table. If either of the values in the summary data where greater than the value in the bonus table for Greatest % or Greatest Total write the Ticker and appropriate value to the Bonus table.  For the Greatest % Decrease the conditional was reversed, check if PercentChange was less than the value in the table, and if so write the Ticker and the value to the table.

This allowed the Greatest Value to "bubble up" to the top of each each category.  No additional looping was needed to meet the additional requirement.

During this additional code was also added after the While loop that formatted the columns of the summary and bonus data sothat the data and headers fit properly in the columns.

## Final coding and testing

After adding the bonus work code, final testing was then done.  Everything looked good.  The StockSummary module was exported from the test spreadsheet.

## Production Run.

The StockSummary module was imported to the production spreadsheet,  After this was completed, the module was then run.  Again, everything ran without error.  Spot checking of data found no problems.

Screen shots of each sheet were then taken and saved.

## Moral of the story.

There really isn't one.  This was actually an easier excercise to complete as it used skills that I already had and little to do with working in Excel itself.  I had fun working this assignment.
