# VBA-challenge
Repository for VBA Homework in Week 2 of Monash Bootcamp

This repository holds the scripting for the VBA Homework, it can be located in both WallStreetLoop_Final.txt file and the WallStreetLoop_Final.vbs file in case of issues loading or reading.

This repository also holds screenshots of the output from the macro which has been run on the multi-year data document provided for the homework.

The VBA scripts contain the following:

WallStreetSingle is the required portion of the homework, running the script in a single page to create a summary table which creates a summary table showing each stock ticker, its price change from start of year to end of year, the percentage price change and the total stock volume.

WallStreetLoop_Final creates the original summary table for yearly data and a second summary table which it populates with the ticker which showed greatest % increase in price, greatest % decrease in price and the ticker with the highest stock volume. This script is also configured to run on every sheet in the workbook by running the macro once.

This homework was challenging as someone new to VBA, with the creation of my macro requiring a number of attempts to ensure correct formatting and syntax. The most difficult part for me was gettiung it to loop on every worksheet, and then to get my ordering of variable storage correct to ensure my values calculated correctly.

With more practice, I believe I can clean the code up substantially and discover new paths to reach the same output.
