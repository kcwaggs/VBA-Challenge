# README

## VBA-Challenge:

This is my submission for the Module 2 Challenge - VBA. 
Link to the files for this challenge are linked here so that you can take my CSV and run the script. 
  >>> https://static.bc-edx.com/data/dl-1-2/m2/lms/starter/Starter_Code.zip

## Notes: 

If you look at all of my commits, you will see in my first few attempts I was working on each individual portions of the code. Since this is a new subject for me it helped me better understand the rules and how the different lines of code interact with each other. Eventually all of my subs were merged to create one analysis script and one reset script for ease of use. 

One notable difference you will see here, is that I included color formatting for the summary tables and column headers. I also had this sectio of code include updated row/column sizing to accomodate for the size of data we were analyzing. It helped me when analyzing the data to differentiate the data points from titles/headers. You will see this in the ' WORKBOOK FORMATTING section. 

I saw in the instructions that we needed to include conditional formatting for the 'Yearly Change' column, but in the grading portion it asked to include conditional formatting for the 'Percent Change' as well. I ended up formatting both of these columns. 

In the retrieval of data portion of the grading notes, it states that we need to read and store data for the open and close price. We used these data points to calculate the Yearly Change/Percent Change fields. If you need to identify where I held this information, lines 82, 98, and 111 were used to store this information. 

Screenshot file is included in my repository with screenshots of... 
>>> My macro options.
>>> Running the 'AnalyzeStocks' script.
>>> Message Box stating completion of analysis.
>>> Confirmation that the script ran on all three tabs.
>>> Running the 'ResetWorkbook' script.
>>> Message Box stating completion of the reset (background showing formatting was withdrawn).

I made sure to label my steps so you can identify what I was aiming to accomplish, and it helped me stay organized. 

## Resources: 

I used this resource to assist in defining the LastRow variable. 
  >>> https://www.wallstreetmojo.com/vba-last-row/

For the following piece of code, my peer Ryan MacFarlane helped me write my loop script through all the worksheets. I had initially done it by defining each worksheet as i and looped that way, but it interfered with my code further down. 
  >>> For Each ws In Worksheets
        ws.Activate
      Next ws

I used this resource to better format my summary tables. I used it to understand how to format column width and row height. 
  >>> https://stackoverflow.com/questions/37972749/using-for-loop-to-set-ranges-of-column-widths-in-vba
