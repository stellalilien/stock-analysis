# Stock Analysis

## Project Overview

This project was designed with the intention of analyzing a set of stocks and their performances in 2017 and 2018 for a client wanting to make informed recommendations based on past perfomances and trends. For the stocks in question Total Daily Volume was calculated as well as the stocks yearly return which was translated into a percentage. A button was added so the user can enter the year they would like to see and the code will populate the information needed.

Once this was developed the code was refactored to shorten analysis time.

## Results

### 2017 vs 2018

When comparing Total Daily Volume between 2017 and 2018 we can see that some of the stocks increased and some decreased but looking at the Returns we can see that the increases in Total Daily Volume did not necessarily correlate to an increase in return. Overall tickers ENPH and RUN both had significant increases in Volume and Return between the two years where all other tickers ended with decreases in Return in 2018 despite all but one increasing in 2017. Without more information or years to compare we can't supply definitive explanations for these performances. 

<img width="188" alt="2017 screenshot" src="https://user-images.githubusercontent.com/94754972/148664699-256a3a55-e2bb-444a-bf55-472cb3c38057.png">

<img width="186" alt="2018 screenshot" src="https://user-images.githubusercontent.com/94754972/148664702-9ac8bdd6-254b-4604-8a36-265df4bdd11f.png">

### Refactored Code

After the initial analysis the VBA code was refactored with the intention of reducing the overall run time. Using the original code each year was run and the times recorded and once the refactoring was completed the years were each run again so check for improvement. The runtimes were significantly decreased after the code was refactored:

#### 2017 Original Runtime
<img width="159" alt="original runtime 2017" src="https://user-images.githubusercontent.com/94754972/148664799-76099d8e-83a6-4241-a6f6-717833493867.png">

#### 2017 Refactored Runtime
<img width="164" alt="refactored runtime 2017" src="https://user-images.githubusercontent.com/94754972/148664807-936bc561-0475-4544-a1cd-d3a10ee3b87d.png">

#### 2018 Original Runtime
<img width="164" alt="original runtime 2018" src="https://user-images.githubusercontent.com/94754972/148664814-92610440-6112-4b4c-ab64-579623255715.png">

#### 2018 Refactored Runtime
<img width="166" alt="refactored runtime 2018" src="https://user-images.githubusercontent.com/94754972/148664819-e17633c8-9a34-41ec-b95d-dd924dab6b2f.png">

In the original code variables were initialized for starting and ending prices and the code looped through each ticker adding up the volume and then checking for starting and ending prices which then output results in the spreadsheet. When refactoring rather than simply looping through the tickers I added a ticker index and three output arrays.

<img width="278" alt="image" src="https://user-images.githubusercontent.com/94754972/148665108-bb40fcab-5de6-4629-9f56-8dad037472e6.png">

Once I set up the index and the arrays I set the ticker volumes to reset to zero for each loop.

<img width="398" alt="image" src="https://user-images.githubusercontent.com/94754972/148665292-fc66f76d-7e29-44ea-bbc5-5d61677e4497.png">

The next for loop goes through each row of the spreadsheet calculating volume and identifying the starting and ending prices. In the refactored code the ticker index is used, allowing all the necessary information to be collected while only looping through the data once.

<img width="660" alt="image" src="https://user-images.githubusercontent.com/94754972/148665519-be4a956c-f9e9-4e55-9259-d0b2539f45de.png">

The final adjustment was to add output arrays which loop through the ticker index to get the final calculations and insert them into the final worksheet.

<img width="501" alt="image" src="https://user-images.githubusercontent.com/94754972/148665564-91c9c72f-9d9c-4c88-9865-8ccfeba53121.png">


## Summary

### General Code Refactoring Advantages and Disadvantages

Generally speaking code refactoring is used to make code cleaner and more efficient. Ideally this allows programs to run with less lag time and also can make it simpler to address future changes in the data without having to rewrite large portions of code. In small datasets runtime lag may not be a significant issue however when handling larger sets of data maximizing efficiency can have a noticeable impact on user experience. There are situations where refactoring is not as advantageous however. There are cases where the outcomes can change unintentionally which makes it important to track what is being changed and to test careful. Another consideration is to be sure when refactoring code to update/add comments to accurately represent the intention of the refactored code in case others have questions or need to add to it later. 

In this specfic project the code refactoring had the advantage of decreasing the runtime due to organizing it in a way that it only had to loop through the dataset one time. Considering the amount of time it took and the difference in output times it could seem like a less efficient overall use of time however were additional data to be added at a later date that difference would make a much larger difference. 
