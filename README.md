# PythonExcelPowerBIRefresh

Scenario: This "Example.xlsx" file data is queried in MS Power BI. When Power BI is refreshed, system gets latest data from excel into Power BI. The issue is, when we get the excel data updated using the OpenPyXl Python library and refresh power BI, the data doesn't get updated in Power BI. However, when we manually update and save the excel file, data gets updated in Power BI on refresh. In both the cases, data in excel file is correctly being saved.

Steps:
1.Open  example.pbix and refresh. (The data gets refreshed.)
2.Run  Python Code. (It will update example.xlsx file.)
3.Again, refresh example.pbix file, it wouldn’t get refreshed and will show blank values, however, excel file has been updated.
4.Now  open example.xlsx file, press ctrl + S and close the file.
5.Again, refresh example.pbix file. It will get updated.

In other words, When we update and save data in Excel using OpenPyXl Python library, some calculations are performed on it in excel, and refresh its powerbi connection, it shows blank table cells. When we again manually save the excel file and close it, the MSPowerBI connection works (It gets the calculated data from excel).
