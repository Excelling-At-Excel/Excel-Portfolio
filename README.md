# Microsoft Excel Portfolio
## Showcasing my knowledge & understanding of Excel and it's many Functions.
> ###### For legal reasons, all numerical data within the Source Report has been changed using =RANDBETWEEN() and all names have been changed to a Customer Number (EX: Customer 37)

<p align="center">
<img src="https://www.versionmuseum.com/images/applications/microsoft-excel/microsoft-excel%5E2016%5Eexcel-logo-new.png" width="950" height="250" />
</p>
  
# [Project 1: Sales Report By Payment Card Type](https://github.com/Excelling-At-Excel/Excel-Portfolio/raw/main/Sales%20Report%20By%20Card%20Type.xlsx)
> ### With no available Unique Identifiers

## Created a workbook that utilizes a variety of formulas to pull data from a Source Report and Outputs it into a user-friendly summary

### Formulas/Functions included in this report are as follows:
> Note: All Full Formulas are included within the Linked Report, but some will not be shown here due to length.

* * *

### The following Formulas are housed in Sheet1 of the Linked Report.  (Where the Source Report is pasted into)

> * =IF(ISNUMBER(SEARCH("PAGE:",$G1))=TRUE,IF(ISNUMBER(VALUE(TRIM(RIGHT(G1,2))))=TRUE,TRIM(RIGHT(G1,2)),""),"")
  > > Searches a specified Cell for a predetermined Text-string.  If the Text-String is found, then pull the last 2 characters from the string.  (In this reports case, I needed to pull the Page Number from the specified cell).  If one of the 2 characters being pulled is a space, then remove the space. Lastly, this will convert the outcome into a Numerical Value instead of a Text Value.  (If I pulled from "Page: 53", the 53 would be considered a Text-String instead of a Numerical Value  

---------------------------------------

> * =IF(ISNUMBER(SEARCH("PAGE:",$G1))=TRUE,CONCAT($D1," - ",$G1),"")
  > > Searches a specified Cell for a predetermined Text-String.  If the Text-String is found, then Concatenate the Text-String with the another specified Cell, (The Numerical Value, from the above formula output) to make a Unique Identifier to be used with a future formula.

---------------------------------------

> * =IFERROR(IF(IF(COUNTIF($L$1:$L1,$L1)>1,$K1+100,$K1)="","",IF(COUNTIF($L$1:$L1,$L1)>1,$K1+100,$K1)),"")
  > > Using the data from the last two formulas, count how many instances of the exact Unique Identifier have been used in the cells above it.  (The reasons for this, is due to the fact that the Page Numbers in this report roll back to 1, after reaching "Page: 99" instead of carrying on to 100+.)  If this is the first instance of the Unique Identifier, then output The Page Number that we obtained from our first formula.  If the Unique Identifier has already been used, then check the value of the given to us from the Page Number output and add 100 to it. (EX: If "Page: 3" has been used in a cell above, set the output to "Page: 103")

---------------------------------------

> ### The following Formula is a 7 part, nested If-Statement (Shown in the Linked Report)
> > * =If(IsNumber(Search("abc",A1)),Concatenate(A1, "-",Left(B1,7))
> > 
> > > Check a specified Cell for One of Seven different criteria Text-Strings. If one of the seven Text-Strings is found within the cell, then Concatenate a predetermined Text-String with the Page Number that was given to use from our last formula.  The output will make the official Unique Identifier that will be used for the formulas housed in the Output Dashboard Sheet.  (If the specified Text-String is not found, then set Cell to Blank.)

---------------------------------------

### The following Formulas are housed in Sheet2 of the Linked Report.  (Where the Output is displayed in a Dashboard)

> * =COUNTIF('Source Report'!$M:$M," * abc * ")
> > Searches the specified range for any Cells that contain a predetermined Text-String.  Then, count how many instances of the Text-String are within the range and output the Numerical Value.  (This value will be used in a later formula)

---------------------------------------

> * =CONCATENATE($B$2," - ","Page:  ",$F$2)
> > Concatenates a predetermined Text-String with the Numerical Value from the above formula, to create a true Unique Identifier.  (This will be used in a later formula)

---------------------------------------

> * =ABS(MATCH($G$2,'Source Report'!$M:$M,0)-MATCH('Source Report'!$G$3,'Source Report'!$M:$M,0))-1
> > Checks a predetermined Text-String within a predetermined Range and outputs the absolute value of your criteria and then compares it to a second predetermined Text-String within the range.  Next, it will find the difference between the absolute values and will output the Numerical Value.  (This will be used in a later formula)

---------------------------------------

> * =IFNA(VLOOKUP($B3,OFFSET('Source Report'!$A$1,MATCH($G$2,'Source Report'!$M:$M,0),0,$J$3,3),3,0),0)
> > Uses a predetermined Cell as a Lookup-Value to be used with the VLookup.  Then, while using the Offset Function to start the Match function at the top row of the Source Report, Match a predetermined Cell to its Unique Identifier within a predetermined Range.  Then using the value obtained from the formula above, that will decide how many rows are needed for the Offset function, in order to guarantee no overlaps in data.  Lastly, finishing the VLookup, after finding the correct Match via using the Unique Identifier created in an earlier formula, output the data from the third column of the range as a Numerical Value.  (This formula is repeated one cell to the left and will pull from the second column instead of the third.)

<p align="center">  
<img src="https://i.imgur.com/LhQV3oz.png"/>
</p>

<p align="center">  
<img src="https://i.imgur.com/dI60Bnm.png"/>
</p>
  
---------------------------------------

# [Project 2: Summary of Funds](https://github.com/Excelling-At-Excel/Excel-Portfolio/raw/main/Summary%20of%20Funds.xlsx)
> ### Report with individual Summaries, per Program, and a Summary Wrap-Up for all Programs

## Created a workbook that utilizes a variety of formulas to pull data from a multiple Source Report and Outputs it into multiple Program Summary Dashboards and are then tied into an all encompassing Summary Dashboard.

### Formulas/Functions included in this report are as follows:
> Note: All Full Formulas are included within the Linked Report, but some will not be shown here due to length.

* * *

### The following Formula is housed in Sheet1 (Criteria) of the Linked Report.  (Where the first Source Report is housed)

> * =CONCATENATE("SII:  ",B2," - ",C2)
  > > Concatenates a predetermined Text-String with a specified Cell with another Text-String, to create a Unique Identifier. (Used in a future formula)

### The following Formulas are housed in Sheet2 (CM) of the Linked Report.  (Where the Output is displayed in a Summary Dashboard)

> * =MID(CELL("filename",A1),FIND("]",CELL("filename",A1))+1,255)
  > > Finds the name of the sheet that this formula is in and outputs the exact name of the sheet.  (Used in a future formula)

---------------------------------------

> * =VLOOKUP($F$1,Criteria!$B:$D,3,0)
  > > Utilizing VLookup, use a specified Cell as criteria to search in Sheet1 (Criteria) to then output the corresponding Text-String.  (If "CM" is the Criteria, then output "SII:  CM - Financial Mgmt"

---------------------------------------

> * =MID(A4,12,999)
  > > Obtain a Text-String by using a specified Cell.  Using the =Mid formula, select the Cell where the output of the above formula was placed.  Then, starting at the 12th character, pull the remainder of the Text-String and output as a Text-String.

---------------------------------------

> * =SUMIFS('Daily Funds Status'!K:K,'Daily Funds Status'!C:C,F1,'Daily Funds Status'!A:A,"52FA")/1000
  > > Using data from Sheet3 (Daily Funds Status), use column "K" as a Sum Range and Column "C" as the Criteria Range.  $F$1 is being used as the Unique Identifier that will be used to match against data Column "C" in Sheet3.  The Text-String of "52FA" will be used as the second Unique Identifier to be matched against Column "A" in sheet3.  Lastly, Divide by 1000 in order to set the output result to show in "Thousands" 
  > > (This formula is then repeated two times, while only changing the Criteria for the Sum Range)
  > > > 1. $E:$6 uses Column "P"
  > > > 1. $E:$11 uses Column "T"

---------------------------------------

> * =E15-SUM(E16:E18)
> > Takes the Numerical Value within the Cell and subtracts out the Sum of the following line items:
> > > 1. Total Support Costs posted to SABRS"
> > > 1. GPC Transactions On Log Not posted to SABRS
> > > 1. Other Transactions not yet in SABRS

---------------------------------------

### The following Formulas are housed in Sheet4 (Summary) of the Linked Report.  (Where the Output of all Individual Program Summaries are displayed in an all encompassing Summary Dashboard)

> * =CONCATENATE($D$4," 1200 OMN SABRS Balance  (K)")
> > Concatenates a predetermined Cell with a predetermined Text-String to output a Header for the Summary Dashboard.  (Pulls the date from the Source Report that is referenced within this sheet.)
> > > (This step is then repeated for the Cell directly to the right, with the only change being the predetermined Text-String)

---------------------------------------

> * =IF($B4=CM!$F$1,CM!$E$4,"ERROR - DATE")
> > Checks if the criteria in a predetermined Cell matches against the Criteria another predetermined Cell in the corresponding Sheet.  If the criteria both match 1:1, then pull data Cell $E$4 from the corresponding sheet and Output as a Numerical Value.  (If there is an error then it will inform you instead of just zeroing out)
> > > (This step is then repeated for the next four Cells directly to the right, with the only change being the predetermined Text-String and which Cell in the corresponding sheet is being referenced)
> > > > 1. Sheet3 (Summary) $E:$4 is referencing Sheet2 (CM) $E$7
> > > > 1. Sheet3 (Summary) $F:$4 is referencing Sheet2 (CM) $E$12
> > > > 1. Sheet3 (Summary) $G:$4 is referencing Sheet2 (CM) $E$9
> > > > 1. Sheet3 (Summary) $H:$4 is referencing Sheet2 (CM) $E$21

<p align="center">  
<img src="https://i.imgur.com/mubV3LG.png"/>
</p>

<p align="center">  
<img src="https://i.imgur.com/uPelMih.png"/>
</p>

---------------------------------------

# [Project 3: Credit Card Transaction Detail](https://github.com/Excelling-At-Excel/Excel-Portfolio/raw/main/Credit%20Card%20Transactions%20Detail.xls)
> ### Report of Credit Card transactions, broken out by card type and currency.

## Created a simple workbook that utilizes a basic excel functions, such as SumIfs and CountIfs that gathers data from a singular source report and disperses it into a multilayered Summary Dashboard.

### Formulas/Functions included in this report are as follows:
> Note: All Full Formulas are included within the Linked Report, but some will not be shown here due to length.

---------------------------------------

### The following Formula is housed in Sheet1 (Source Report) of the Linked Report.

> * =IF($F8="00000","3rd Party","")
  > > This is used in column "AJ" and is used as a helper-column that then lets the formulas in Sheet2 differentiate between Regular Transactions and then 3rd Party Transactions

---------------------------------------

### The following Formulas are housed in Sheet2 (Output Dashboard) of the Linked Report.

> * =SUMIFS('Source Report'!$N:$N,'Source Report'!$L:$L,$C$2,'Source Report'!$N:$N,"<0",'Source Report'!$AD:$AD,"<>eCommerce, Channel Encrypt",'Source Report'!$AJ:$AJ,"<>3rd Party")
  > > Uses the following as Criteria for the SumIfs:
  > > > Header of the Summary Section (Card Type) as Criteria 1
  > > > If the Value is less than zero (Credit)  or  more than zero (Debit)
  > > > Excludes any Transactions that contain the text *"eCommerce, Channel Encrypt"*
  > > > Excludes any Transactions that contain the text *"3rd Party"*
  > > This formula, if less than zero, will display what the total "Credit" value is, depending on where is falls with the above criteria.

---------------------------------------

> * =SUM(SUMIFS('Source Report'!$N:$N,'Source Report'!$F:$F,$C$25,'Source Report'!$N:$N,"<0",'Source Report'!$L:$L,{"VISA","MASTERCARD"}))
> > This formula will look at the "Source Report" and give the user the "Credit" value for the card types; "Visa" and "Mastercard", that have the criteria of *"3rd Party"* (Thanks to the Helper-Column) in Sheet1 (Source Report)
> > > Uses the following as Criteria for the Sum(SumIfs)
> > > > Header of the Summary Section (3rd Party) as Criteria 1
> > > > If the Value isless than zero (Credit)  or  more than zero (Debit)
> > > The following part of the formula is set as an *"Array"* in order to use multiple criteria in a singular formula
> > > > "Visa" and "Mastercard"
> > > This will only give the Value that meets all of the above criteria.  The user of this report needed to show the difference of 3rd Parties by "Visa" and "Mastercard" combined and then needed to show in a separate Summary field what the value for "Discover" was.

---------------------------------------

> * =SUMIFS('Source Report'!$N:$N,'Source Report'!$AD:$AD,"eCommerce, Channel Encrypt",'Source Report'!$N:$N,"<0")
> > This formula will look at the "Source Report" and give the user the "Credit" value for all card types that have the criteria of "eCommerce, Channel Encrypt" in Sheet1 (Source Report)
> > > Uses the following as Criteria for the SumIfs
> > > > If Column "AD" includes "eCommerce, Channel Encrypt" then sum

---------------------------------------

> ### These four Formulas are used together in order to verify that all data was pulled correctly and that there are no Variances.
> > * =ROUND(SUM('Source Report'!$N:$N),2)
> > * =ROUND(SUM(D3:D4,D7:D8,D11:D12,D15:D16,D26:D27,D34:D35,D30:D31),2)
> > * =I4-I5
> > * =I6=0

---------------------------------------

<p align="center">  
<img src="https://i.imgur.com/CYKEI4X.png"/>
</p>

<p align="center">  
<img src="https://i.imgur.com/v6fm0sn.png"/>
</p>

---------------------------------------
