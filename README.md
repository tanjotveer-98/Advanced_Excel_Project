# Coffee Dataset Analysis Dashboard using Mirosoft Excel-</br>
* Dataset-https://github.com/tanjotveer-98/Advanced_Excel_Project/blob/777bb93a00551d3468a00b178cd74fb02481021b/coffeeDataset.xlsx
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
<p> For Customername, Email and  Country, Vlookup() can be used to import data from other sheets.</p>
<p> For the remaining fields, INDEX() and MATCH() can be used to import all data at once dynamically.<br/>
For instance- =INDEX(products!$A$1:$G$49,MATCH(orders!$D95, products!$A$1:$A$49,0), MATCH(orders!L$1,products!$A$1:$G$1,0))</p>
<p> Sales field can be calculated using Quantity* Unit Price formula.</p>
