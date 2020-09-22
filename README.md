# VBA-challenge
VBA challenge

General:

Hello! My homework submission is uploaded as per the instructions.
There are three folders one for each year that have screen shots of first few rows & last few rows along with the challenge solution
VBAStocks-Final.vbs is the final script that has the entire solution - challenge included.

I approached the homework by revisiting the recordings (credit card checker & WellsFargo) & started writing & testing my script on an even smaller dataset (~200 rows of single worksheet) of alphabetical_testing.xlsx Once successful, I added the worsheets reference and last row to for loop, tested it on alphabetical_testing.xlsx, followed by Multiple_year_stock_data.xlsx

Code Overview: 

I took an approach similar to the credit card checker activity for my code & I've added comments as/where necessary.

The key part was handling the challenges as I'd two approaches: 
1- Create a second for loop that would run through the o/p table and generate the values for greatest stock volume, % increase & % decrease
2- Generate the values for greatest stock volume, % increase & % decrease within the main for loop. I went with this approach that can be seen in the else part of the loop

For every ith value, once total stock volume is calculated, it gets compared to greatest stock volume (initialized to 0). If this condition is true, 
then total stock volume = greatest stock volume. 
Similarly for greatest % increase & greatest % decrease (initialized to 0), I'm comparing these values to the percent change value generated for each ith value & setting the appropriate value as that percent change & o/p

A part where I got error an overflow error for my % change calculation as I had not factored in divisible by 0. Watch window & study group helped debug this as my initial thought when I saw the error was incorrect variable Dim (i had it as integer).

Conclusion:

I would say I spent maximum time (over 8 hrs) just getting my script right & functioning for the small (~200 rows) dataset, alongwith revisiting the recordings & this gave me a good grasp to implement it on larger dataset.
