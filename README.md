# Green_stocks_VBA_challenge
In this challenge, we are helping Steve to determine the most appropriate stock to invest in by analysing data from 2017 and 2018. We then refactor our code to make it run faster and efficiently.

Overview and Purpose of the project:

In this analysis, we determine which stocks are good to invest money in. We use VBA and codes to extract all stocks' volumes and yearly returns for 2017 and 2018 respectively and figure out which ones were successful.  We refactor the code with VBA. After collecting each stock data from 2017 and 2018 respectively, and refactoring, we will determine if refactoring the code helped VBA script run faster. We compare both run times. Also, we ensure that we make our data, code more efficient and easy to comprehend. 

Refactoring Code:

We followed the following steps to refactor our code and measured the final performance.

1.	We created a header row in our output sheet and initialized arrays of all 12 tickers and assign variable
	<img width="264" alt="Create header row and arrays" src="https://user-images.githubusercontent.com/86980240/132613471-0a485734-84fd-4bb1-8f09-845f5c953f02.png">

3.	After counting rows in datasheets, we set ticker index and three output arrays, assigning variables.
	<img width="313" alt="Set ticker Index and three output arrays" src="https://user-images.githubusercontent.com/86980240/132613528-774743ba-a9dd-4cc6-a657-94768b2a7a05.png">

5.	We used for loops function, set ticker volumes to zero and looked for tickers and volumes through out our datasheets. We checked first and last row for ticker index and increased ticker volumes
	<img width="705" alt="Create for loops and increase ticker index" src="https://user-images.githubusercontent.com/86980240/132613563-8202c96b-7141-48ad-8b95-44d1e7d80ff0.png">

7.	Finally, we post our outputs in our output sheet and then set formatting. We assign number formats in cell B and C. Also, in cell C we assign green and red color to the percentage returns. We assign our respective macros.
	<img width="604" alt="Formatting worksheet" src="https://user-images.githubusercontent.com/86980240/132613598-9886ff83-a118-43c5-85ad-3c34f191ca91.png">


Results:
After refactoring the code, our code runs efficiently and faster. It looks neat and readble. As we can see in the following images and compare the difference in running times.

<img width="230" alt="All Stocks 2017" src="https://user-images.githubusercontent.com/86980240/132613703-5c5b20be-6bb3-4e88-b531-211884113734.png">

<img width="241" alt="VBA_challenge_2017" src="https://user-images.githubusercontent.com/86980240/132613738-63cfb6bf-c2a7-4fea-ba2d-e48185118228.png">

<img width="232" alt="All Stocks 2018" src="https://user-images.githubusercontent.com/86980240/132613765-20554829-ff15-4357-8960-f50f2914ffd9.png">

<img width="232" alt="VBA_challenge_2018" src="https://user-images.githubusercontent.com/86980240/132613798-a4900dd0-f01f-41b7-a19e-a739c810ad33.png">


What are disadvantages of refactoring the code:

1)	Refactoring process may affect outcome of the data
2)	Refactoring the code can be time consuming and also quite difficult for bigger and complicated codes
What are advantages of refactoring the code:

1)	Refactoring the code helped to reduce run time of macros
2)	It makes the code look sorted out better, efficient and easy to comprehend.

How do these pros and cons of apply to refactoring original VBA:

Refactoring the code makes our code look clearer and easy to read. It helps to sort out  various factors. It improves programming and debugging and makes our code run faster. This technique would help to understand the complexity of codes, takes few steps thus saves systems memory and executes faster. On the other hand, we are refactoring a code which is already in use and working. If we don’t use correct syntax and refactor it correctly, our code will not run efficiently. 

