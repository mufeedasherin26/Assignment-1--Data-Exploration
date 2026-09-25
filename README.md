# Assignment-1--Data-Exploration
An Excel-based data exploration assignment using basic Excel functions such as SUM, COUNT, AVERAGE, MIN, MAX, IF, SUMIF, COUNTIF, and text formatting.
Total price of all products in the dataset  : =SUM(G2:G35)
To Find How Many Products Are  In The Dataset   :  =COUNTA(E2:E35)
average price of all products   :  =AVERAGE(G2:G35)
Mnimum price among all products  :  =MIN(G2:G35)
Maximum price among all products :  =MAX(G2:G35)
Using an IF function, create a new column named Price Range to categorize products with a price greater than or equal to $500 as 'High Price' and others as 'Standard Price    :   =IF(G2>=500,"High Price","Standard Price")
Total Price For Products In The Electronic Category :  =SUMIF(J2:J35,J2,G2:G35)
Count Of Products With A Price Less Than 100  :  =COUNTIF(G2:G35,"<100")
Column named Day with the first 2 characters of each Product ID using the LEFT function  :  =LEFT(A2,2)
column named Country Code by extracting the last 2 characters from the Product ID using RIGHT function :  =RIGHT(A2,2)
column named Month by extracting 4th to 6th characters from the Product ID column using the MID function :  =MID(A2,4,3)




