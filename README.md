# Stock-analysis
## Green energy project

### Overview of the project
 - The purpose of this project was to evaluate the performance of different stocks -in the green power market- throghout the years 2017 and 2018.  This analysis should help our client  to choose the most profitable company in wich to invest.
 - To achieve this, we will use VBA to create a macro that iterates through a list of tickers and retrieves the data concerning total volume (wich reflects trading activity) and  yearly return of each stock.
 - In order to work with a larger dataset, our code should be refactor and bring a solution in less time than the original one.


### Results
 - Original code running time for both years
<Img src="Images/Original%202017-2018%20Timer.png" width="650">
 
 - Refactored code running time for 2017
<Img src="Images/2017%20Analysis%20Timer.png" width="650">
 
 - Refactored code running time for 2018
<Img src="Images/2018%20Analysis%20Timer.png" width="650">
 
 - Creating a list for tickers:
   - The image on the right shows how an array was created to work with this particular dataset. Twelve tickers were involved in the project, so we manually assigned an index to each stock.
   - On the left, the refactored code is displayed. An array (or list) is created by iterating through all the rows containing ticker names, and adding every new value to our list while preserving all the data previously stored. We declare a tickers count variable as well, wich will hold the index of each sotck. 
   - Refactored code can work with 'n' number of tickers 
<Img src="Images/Array%20for%20tickers.png" width="750">
 
 ### Summary
 - Refactore code : In general
   - Advantajes: 
     1. It can adapt to any dataset that needs to be dealt with. 
     2. It is also a way to avoid 'magic numbers', a common source of erros if we try to aply our code to a similar problem, but with more o fewer values in it
   
   - Disadvantages: 
     1. Since we avoid magic numbers,  more variables need to be declared in the code and new formulas as well
     2. Adding extra steps to our routine will automatically increase the running time of our program 
  

 - Refactore code: 
   - Refactoring code, in this scenario, was introduced as a necessity if we are to calculate yearly return and total volume on any given number of tickers. 
   - Advantages:
     1. The original VBA script just added indexes manually to each ticker, wich is time consuming in a larger dataset
     2. Our refactored code eliminates this drag with the "Redim Preserv" statement.
  
   - Disadvantages:
     1. Since we avoid magic numbers, VBA will have to look for this values by iterating through all the rows on our spreadsheet, thus taking more time to run the program.
     2. Stil, I see no mayor disadvantages on refactoring code in this case.

 


                                                 








