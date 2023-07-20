# Power-Automate-Extract-PBI-Data_into_Excel
We extract a table from PowerBI into a table in Excel.
Let's check the table format in excel first.
The table has one rows per date to extract the values. The date is the column key.

<img width="800" alt="Register  Table" src="https://github.com/calamardo33/Power-Automate-Extract-PBI-Data_into_Excel/assets/68512988/f418d442-f938-4905-8977-6d6a2235a581">

This flow extract the information for the previous day in a weekly basis.
Excel keeps the dates stored as numbers in the background. So we need to convert yesterday's date into a number. We will use it to update the correct excel row in the table afterwards.
See below formula used to convert the date into ticks.

<img width="550" alt="Register  Yesterday" src="https://github.com/calamardo33/Power-Automate-Extract-PBI-Data_into_Excel/assets/68512988/1c7501dc-3654-4cf6-8d92-fc43a8ccda72">


This time, the dax query is obtained from the Performance Analyzer so I am not including it in this repository.
I had to create several variables. Compose actions were not working for me this time.

<img width="450" alt="Register  1" src="https://github.com/calamardo33/Power-Automate-Extract-PBI-Data_into_Excel/assets/68512988/aed472ac-df35-43df-acee-d1613ce01817">

We convert the json format into an array.
We then loop every instance of the array.
With a Switch action, when the case matches what we are looking for, we set the value of the variable.
After every loop, we set the value of all the arrays.

<img width="600" alt="Register  2" src="https://github.com/calamardo33/Power-Automate-Extract-PBI-Data_into_Excel/assets/68512988/d1e5504a-38c4-45c4-b202-03105c330824">


We then update the row in excel.
the Key column in the table is the date. In order to identify the correct key, we use the compose output created previously.
The value for the different columns will be equal to the variables we have used.

<img width="450" alt="Register  3" src="https://github.com/calamardo33/Power-Automate-Extract-PBI-Data_into_Excel/assets/68512988/83400236-5e75-4b21-90a3-b11dd44021cf">
