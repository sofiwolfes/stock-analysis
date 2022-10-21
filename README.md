# stock-analysis
Refactoring code to make it more efficient, by taking fewer steps and using less memory. Analyzing stock data to find the total daily volume and yearly return for each stock.

## Purpose and background:
By refactoring code, we are trying to make an analysis easier to read, and faster to run. 
In this case the difference in time might be less than seconds, but for bigger sets of data, where we have tu run loops over thousands of rows and lines of code, having an efficient way to perform the analysis could be a game changer.

### Advantages and disadvantages of refactoring code
We have listed some advantages below:
- Code is easier to read
- We review the code and add notes to explain what we are doing to facilitate its comprehension

On the other hand, some disadvantages are:
- If you are an inexperienced analyst, the code might be less complicated to understand despite being simpler and straight to the point.

### Advantages and disadvantages of the original and refactored VBA script
In this particular case, I think the first code we used was more steps by step, and the purpose was to learn basic concepts, which made the original non-refactored code extremely helpful.

The refactored code was a bit more complicated to understand, but once you got a hang of it, it made more sense to do it that way. Also, the benefits of having an efficient code that runs fast and smooth as mentioned above, to avoid long times of running code (especially if this would have been a bigger data set).

## Results:
When you compare the two exact same spreadsheets; the one done throughout the module, and the one done for the challenge, you get the same exact results with only two differences:
1. Different code (less lines of code, more efficient and straight to the point)
2. Faster running times:

See below the differences in running times for the analysis performed for the data of 2017:
![alt text]Resources/VBA_Challenge_2017.png

And below you can see here the differences in running times for the analysis performed for the data of 2018:



## Summary
As we predicted, the difference on time is not huge in this case given no-so-large data sets, but this same refactorization could mean a completely different story when we work with enormous data sets.






DELETE BELOW
The written analysis has the following:

Overview of Project
The purpose and background are well defined (2 pt).
Results
The analysis is well described with screenshots and code (4 pt).
Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
