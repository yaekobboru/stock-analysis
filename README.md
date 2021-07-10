# Overview of Project

Steve is analyzing green energy stocks data to provide best investment decicion for his client. 
This project support Steve's decision with automated solution using VBA(Visual Basic Application).
Steve's cleint's are interested to know the stock performance for ticker "DQ" but this solution
provides 11 more tickers performance for the year 2017 and 2018. 
This analysis conducts refactoring to calculate its effect on code execution times.


## Results

The data collected from Steve's excel file got automated using VBA scripts and user can analyze corresponding ticker's
total daily volume and return percentages of each year with single click. The following are major summary of results.
  
   * ENPH and RUN has positive returns in 2017 and 2018.
   * Only TERP has negative return in 2017 and 2018.
   * All stocks except TERP has positive returns in 2017.
   * DQ , ENPH ,FSLR and SEDG have high positive percentage of return(more than 100%) in 2017.
   * Only ENPH and RUN have positive return percentage which is around 80% in 2018.
   * DQ and JKS have lowest negative percentage around 60% in 2018.

while conducting automation VBA codes got refactored to analyze performance of code excution. The results show original code takes 0.52 seconds
and the refactored code 0.08 seconds. This shows a significant performance improvement by refactoring.  


## Summary

Refactoring a code increases its readability (readable codes saves developer's time of debugging) and also improves performance(reduced excution time).
reduced excution time gives a developer flexibility for multi tasking, saves CPU usage (free CPU for peers) 
and ultimately satisfies final user(with very short response time from application or User Interface end.)

Ticker index and output arrays got introduced on refactoring the code and excution time 
reduced ,this will help Steve to increase the data size by increasing the number of tickers. Increasing the
data volume will help the results to be more close to the required output. If the excution time takes longer 
Steve will be forced to focus on limited data set to make his decision.




