# Coffee Dataset Analysis Dashboard using Mirosoft Excel-</br>
* Dataset-</br>https://github.com/tanjotveer-98/Advanced_Excel_Project/blob/777bb93a00551d3468a00b178cd74fb02481021b/coffeeDataset.xlsx
* Tools- MS Excel 2019. </br></br>
<p>The given dataset file contains three worksheets- </p>
 <ol><li> Customers- ID, name, email, phone number,address, city, country, postal code, loyalty Card.</li>
  <li>Products- ID, coffee type, roast type, size, unit prize, price per 100g, profit.</li>
  <li>Orders- ID, Date, Customer ID, Product ID, Quantity.</li></ol>
<p>In orders worksheet some fields were yet to be filled, for instance, Cusromer name, email, country, coffee type, roast type, size, unit price, sales.</p>

## Data Cleaning-</br>
* Check for duplicates.
* Check for empty values.</br>

## Filling up empty fields.- </br>
<ol>
<li>For Customer name, Email and  Country, Vlookup() can be used to import data from other sheets.</br>
For instance, =VLOOKUP(C95,customers!$A$2:$B$1001,2,FALSE) </li>
<li> For the remaining fields, INDEX() and MATCH() can be used to import all data at once dynamically.</br>
<b>For instance, =INDEX(products!$A$1:$G$49,MATCH(orders!$D95, products!$A$1:$A$49,0), MATCH(orders!L$1,products!$A$1:$G$1,0))</b></li>
<li> Sales field can be calculated using (Quantity*Unit Price) formula.</li>
</ol>

## Creating Pivot tables-
<p> Before creating any pivot tables, the data on 'orders' sheet needs to be converted into Table, so that in case any new data is added it can automatically be refreshed into pivot table created on top of that data.</p>
<ol>
 <li> The first pivot table is to analyze trend of coffee sales over the years and months. The line chart is used to display the trend over the period. Timeline is inserted to adjust the period dfor which we want to see the trend on the line graph. Slicers are added to filter the data based on roast type of coffee, loyaly card of thecustomer and product size. Following images shows the pivot table and the charts created from the aggregated data.</br>
<img src= "https://github.com/user-attachments/assets/8daa6c76-f3d1-4a8d-8d25-9eb1158411dc" height= 350 width= 350> 
<img src= "https://github.com/user-attachments/assets/fe86bd30-429e-4341-a2e5-491758295aae" height= 350 width= 350>
<img src= "https://github.com/user-attachments/assets/2b9966bf-aa33-4499-b1e1-4cf120fd4cb5" height= 350 width= 700>
</li>
 <li> The second pivot table is to analyze the sales of coffee products by country and visualize using bar chart. Following images shows the pivot table and the chart created from the aggregated data.</br>
 <img src= "https://github.com/user-attachments/assets/f905f219-ca0a-43dc-8cd0-d1490b5d7e8a" height= 350 width=600> </br> </br>
 </li>
 <li>
  The third pivot table is to analyze the sales of coffee by customers and display top 5 customers using bar chart. Following images shows the pivot table and the chart created from the aggregated data.</br>
 <img src="https://github.com/user-attachments/assets/07eb5928-6113-4553-b171-822da0742f13" height= 350 width= 600>
 </li>
</ol>

## Creating the final dashboard.
 <p> All the charts were copied to a new worksheets and all charts were connected using "Report Connections" feature to the timeline and the slicers. Following image shows the final result.</p>
 <img src= "https://github.com/user-attachments/assets/7e2a3033-a7db-43c5-a96e-a24f2e0377c5">


