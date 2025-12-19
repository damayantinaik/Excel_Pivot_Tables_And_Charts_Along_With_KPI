# Crafting Excel Pivot Tables and Pivot Charts with Slicers, Filters and KPI

This project consists of two parts: 
1. Creating pivot tables and pivot charts from a single data table
2. Creating pivot table and pivot chart from multiple tables connected in a Data model 
## Case: Pivot Table from one single data table
For this project, I’m going use data from a frozen desserts company. The dataset lists the  sales records over three years ; 2013, 2014, 2015. I need to prepare pivot tables & pivot charts summarizing the sales  under different circumstances: 
1.	Sales per year, month for each dessert in each region.
2.	 Sales by each salesperson during different years, and those values as % of grand total.
3.	Percent (%) difference in sales between months. 

Before creating a pivot table, I selected the data from ‘SalesData’ worksheet and converted it to a table called ‘YearlySales’. Upon creating the table, I proceeded to generate Pivot Tables & Pivot Charts.

From a Pivot table we can look into the details related  to a value. In PivotTableSalesPercentDifference if we click on a value, for example to see why there is sudden increase in the sales of 53.94% in April 2013 as compared to previous month.  To know further on the data,  if we double click the value 53.94%, it’ll list the values from April 2013 in a separate new sheet. We can name the sheet as ‘SalesApril2013’ and if required it can be sent to a person who is looking into it to derive business insights.

If pivot table data changes, pivot chart data also changes. So, if we update pivot table accordingly pivot chart also get updated.

I added multiple filters and slicers to the Pivot tables and charts. 

Find the workbook at:  https://github.com/damayantinaik/Excel_Pivot_Tables_And_Charts_Along_With_KPI/blob/main/Pivot_Tables_And_Pivot_Charts_With_Data_From_Single_Table.xlsx

## Case: Pivot Table from multiple data tables connected in a Data Model
To create the pivot table and pivot chart, I’ve a workbook with multiple worksheets, but I am going to work on two worksheets only named Customer Info and Order info. 

The Customer Info worksheet carries the records on customers whereas the Order Info worksheet lists order information records. 
I need to summarize the data from both worksheets in a pivot table & chart. 

To obtain the pivot table, first, I converted both worksheets into tables and then connected them by a common column 'Customer ID' in a data model using Power Pivot. Here the Customer Info Table is the parent table and Order Info is the child table. 

I created the Pivot Table showing a measure average freight for all countries and include the KPI icons for visually easy reading.   

The workbook is available at:  https://github.com/damayantinaik/Excel_Pivot_Tables_And_Charts_Along_With_KPI/blob/main/Pivot_Table_And_Chart_Showing_KPI_From_Multiple_Tables.xlsx

