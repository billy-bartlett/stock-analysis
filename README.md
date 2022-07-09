# Project Overview

This project set out to analyze the stock market! I performed analysis on a handful of alternative energy stocks for the years 2017 and 2018. I set out to determine the return for these stocks and used VBA Scripting to automate the analysis. I created two sets of code (original and refactored) and coded in a timer to display the run time for both sets of code. I then took my first draft of code (original) and refactored it to decrease run time to make the analysis quicker! This analysis will show you which stocks were a great investment option and which stocks you would have been wise to avoid back in 2017 and 2018.

# Results

My initial code ran much slower than my refactored code so refactoring was a success! As you can see in the images below, for 2017, my original code ran in .671875 seconds and ran in .609375 seconds for 2018.

![VBA_Challenge_2_Old_Code_2017](https://user-images.githubusercontent.com/101153516/178094738-5006f335-2816-49f3-86d6-d5f1c0355a85.jpg)

![VBA_Challenge_2_Old_Code_2018](https://user-images.githubusercontent.com/101153516/178094744-96f294f0-5887-4f03-99d3-a1ee575f31ec.jpg)

Below you will see two images of my refactored code. For 2017, I was able to run my analysis in .1149979 seconds. For 2018, I was able to run the analysis in .1180038 seconds. You may even be able to run this code even faster on your machine!

![VBA_Challenge_2017](https://user-images.githubusercontent.com/101153516/178094755-bdc155a1-d4e6-4560-8cf6-1d8cf20d4b29.jpg)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/101153516/178094758-6a7cd69d-40d8-48b2-b25c-23b296f236ab.jpg)

## Time For Some Stock Analysis! 
As you can see in the images below, I was able to use VBA Scripting to analyze stocks for 2017 and 2018. 2017 was a better year to invest in these 12 stocks than 2018. All but one stock provided a positive return in 2017! What a year for green energy! You had a 91.7% chance to get a positive return on your investment in these 12 stocks back in 2017. Nice! If you invested in DQ in 2017, your return on investment was 199.4%! Following closely, SEDG returned 184.5% on your investment. In total, 4 stocks provided a return of over 100%!

![Investment_Return_2017](https://user-images.githubusercontent.com/101153516/178094765-56e47f55-f020-406a-a20c-fc190a372c23.jpg)

Unfortunately, 2018 was a tough year for green energy stocks. Only two stocks that I analyzed for 2018 provided a positive return. However, the two stocks that did provide a positive return, provided a great return: ENPH at 81.92% and RUN at 83.95%! You only had a 16.7% chance to receive a positive return in 2018 but when you invested right, you made money! 

![Investment_Return_2018](https://user-images.githubusercontent.com/101153516/178094772-f489ebe2-a5d2-4b5f-ad3f-d8cd928bfd58.jpg)

## Let's Have a Look At the Code!
I first set an array of tickers that identified the stock I was analyzing as seen in the image below. 

![Ticker_Array](https://user-images.githubusercontent.com/101153516/178094778-e336f6ae-44f3-4f1d-afe8-ed985ee1ea2f.jpg)

I then identified the number of rows I needed to loop through (thank you Stack Overflow for the inspiration!), set my ticker index to zero, and created three arrays for my code to populate. 

![Stack_Overflow_and_Zeros](https://user-images.githubusercontent.com/101153516/178094787-3a88b902-0900-4575-be76-fdd2ca393008.jpg)

To ensure my ticker for volume was reset for each loop, I created a loop to ensure that the ticker was, in fact, set to zero for every loop that runs. 

![Zero_Loop_Image](https://user-images.githubusercontent.com/101153516/178094796-fd43d581-85d5-4080-ad9d-2b489074d95a.jpg)

Next, I created several loops so the code analyzed the data and returned the data I was after. My first loop sets the rows I want to analyze, identifies the volume for the current ticker, and sets a starting price that I want to capture by ensuring the current row is the first row for that ticker. This loop also sets the ending price by ensuring the current row in the loop is the last row for the current ticker. The loop then does all of this for every ticker!

The last loop, before we get to formatting, outputs the data for the ticker, total daily volume, and the return.

![All_the_Loops](https://user-images.githubusercontent.com/101153516/178094808-2535fe23-44e9-42e7-a1d0-9cbcb1ccb4cb.jpg)

Formatting may be boring but it's super important! I then added a loop to automate some formatting to increase the visual appeal. I bolded my fonts, added some borders, and altered the number formatting. Check out the code below!

![All_the_Formatting](https://user-images.githubusercontent.com/101153516/178094814-c2b6b4a3-e5cf-46e6-9718-fa7afae9078b.jpg)


# Summary
## What Are the Advantages or Disadvantages of Refactoring Code?
There are some great advantages when it comes to refactoring code. As you can see, one of the advantages of refactoring code is making your code run quicker. The quicker the code runs, the quicker you can get to analyzing your results! Another advantage is refactoring makes your code easier to understand. You never know who will be reviewing your code after you write it so you may as well help them out! 

The biggest disadvantage of refactoring code is it can present new bugs that require time and effort to debug. This process can be frustrating but the end result is usually worth the effort! One other disadvantage of refactoring code is there isn't always a clear indicator that code needs to be updated. You can burn time better used for higher value add tasks if you're not careful. 

## How do These Pros and Cons Apply to Refactoring the Original VBA Script?
When refactoring my original code, I ran into SEVERAL bugs that I had to troubleshoot. It took quite a while to dial this code in to run, not only more efficiently, but run in general! Although it was not an easy task, I was able to shave off .5568771 seconds of run time for my 2017 analysis. I was also able to shave off .4913712 of run time for my 2018 analysis.

The advantage of my original code is that the analysis was compiled automatically which is much quicker than pulling the information together manually. The disadvantage of the original code is that it ran slower and was not as easy to read as the refactored code.


