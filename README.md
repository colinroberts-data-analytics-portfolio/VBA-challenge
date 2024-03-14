# VBA-challenge
VBA-challenge2

Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol

Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock. The result should match the following image:

Moderate solution

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

Hard solution

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

NOTE
Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

Requirements
Retrieval of Data (20 points)
The script loops through one year of stock data and reads/ stores all of the following values from each row:

ticker symbol (5 points)

volume of stock (5 points)

open price (5 points)

close price (5 points)

Column Creation (10 points)
On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:

ticker symbol (2.5 points)

total stock volume (2.5 points)

yearly change ($) (2.5 points)

percent change (2.5 points)

Conditional Formatting (20 points)
Conditional formatting is applied correctly and appropriately to the yearly change column (10 points)

Conditional formatting is applied correctly and appropriately to the percent change column (10 points)

Calculated Values (15 points)
All three of the following values are calculated correctly and displayed in the output:

Greatest % Increase (5 points)

Greatest % Decrease (5 points)

Greatest Total Volume (5 points)

Looping Across Worksheet (20 points)
The VBA script can run on all sheets successfully.
GitHub/GitLab Submission (15 points)
All three of the following are uploaded to GitHub/GitLab:
Screenshots of the results (5 points)

Separate VBA script files (5 points)

README file (5 points)

Reference:
-	Loop through worksheets 
https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop
-	Start Variables 
https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop
-	Set headers for summary table  
https://stackoverflow.com/questions/62975110/vba-script-to-format-cells-within-a-column-range-only-formats-the-first-sheet-in
-	Loop through each row in the sheet
https://www.bing.com/search?q=%27+Check+if+the+current+row+has+a+different+ticker+symbol+++++++++++++If+ws.Cells%28i+%2B+1%2C+1%29.Value+%3C%3E+ws.Cells%28i%2C+1%29.Value+Then+++++++++++++++++%27+Set+ticker+symbol+++++++++++++++++ticker+%3D+ws.Cells%28i%2C+1%29.Value+++++++++++++++++&form=ANNTH1&refig=ee330c174ca24736a0e455c4c0322639&pc=U531
-	Closing price
-	https://stackoverflow.com/questions/76548179/dont-know-how-to-fix
-	Percent change
https://money.stackexchange.com/questions/84534/what-is-the-correct-answer-for-percent-change-when-the-start-amount-is-zero-doll
-	Clear variables
https://stackoverflow.com/questions/42980386/how-to-reset-variables-or-declarations-vba
-	Set max values
https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.cells
-	Greatest percent decrease
https://www.mashupmath.com/blog/calculating-percent-decrease
-	greatest total volume               
https://www.exceldome.com/solutions/if-a-cell-is-greater-than-a-specific-value/  
-	greatest percent increase, decrease, and total volume       https://www.exceldome.com/solutions/if-a-cell-is-greater-than-a-specific-value/
   
