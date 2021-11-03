# Microsoft Excel Portfolio
## Showcasing my knowledge & understanding of Excel and it's many Functions.
> ###### For legal reasons, all numerical data within the Source Report has been changed using =RANDBETWEEN() and all names have been changed to a Customer Number (EX: Customer 37)

<img src="https://www.versionmuseum.com/images/applications/microsoft-excel/microsoft-excel%5E2016%5Eexcel-logo-new.png" width="950" height="225" />


# [Project 1: Sales Report By Card Type]
> ### With no given Unique Identifiers

## Created a workbook that utilizes a variety of formulas to pull data from a Source Report and Outputs it into a user friendly dashboard

### Formulas/Functions included in this report are as follows:
> Note: Full Formulas are included within the Linked Report, but will not be shown here due to length.

### The following Formulas are housed in Sheet1 of the Linked Report.  (Where the Source Report is pasted into)

> * =IF(ISNUMBER(SEARCH("PAGE:",$G1))=TRUE,IF(ISNUMBER(VALUE(TRIM(RIGHT(G1,2))))=TRUE,TRIM(RIGHT(G1,2)),""),"")
  > > Searches a specified Cell for a pre-determined text-string.  If the Text-String is found, then pull the last 2 characters from the string.  (In this reports case, I needed to pull the Page Number from the specified cell).  If one of the 2 characters being pulled is a space, then remove the space. Lastly, this will convert the outcome into a Numerical Value instead of a Text Value.  (If I pulled from "Page: 53", the 53 would be considered a Text-String instead of a Numerical Value  

##

> * =IF(ISNUMBER(SEARCH("PAGE:",$G1))=TRUE,CONCAT($D1," - ",$G1),"")
  > > Searches a specified Cell for a pre-determined Text-String.  If the Text-String is found, then Concatenate the Text-String with the another specified Cell, (The Numerical Value, from the above formula output) to make a Unique Identifier to be used with a future formula.

##

> * =IFERROR(IF(IF(COUNTIF($L$1:$L1,$L1)>1,$K1+100,$K1)="","",IF(COUNTIF($L$1:$L1,$L1)>1,$K1+100,$K1)),"")
  > > Using the data from the last two formulas, count how many instances of the exact Unique Identifier have been used in the cells above it.  (The reasons for this, is due to the fact that the Page Numbers in this report roll back to 1, after reaching "Page: 99" instead of carrying on to 100+.)  If this is the first instance of the Unique Identifier, then output The Page Number that we obtained from our first formula.  If the Unique Identifier has already been used, then check the value of the given to us from the Page Number output and add 100 to it. (EX: If "Page: 3" has been used in a cell above, set the output to "Page: 103")

##

> ### The following Formula is a 7 part, nested If-Statement (Shown in the Linked Report)
> > * =If(IsNumber(Search("abc",A1)),Concatenate(A1, "-",Left(B1,7))
> > 
> > > Check a specified Cell for One of Seven different criteria Text-Strings. If one of the seven Text-Strings is found within the cell, then Concatenate a pre-determined Text-String with the Page Number that was given to use from our last formula.  The output will make the official Uniqe Identifier that will be used for the formulas housed in the Output Dashboard Sheet.  (If the specified Text-String is not found, then set Cell to Blank.)

##

### The following Formulas are housed in Sheet2 of the Linked Report.  (Where the Output is displayed in a Dashboard)
> * =IfNA(VLookup(Offset(Match)))
> * =CountIf('Source Report'!$M:$M,"* Partial Match * )
> * =ABS(Match(*Criteria*,'Source Report'!$M:$M,0)-Match('Output Dashboard'!$G$3,'Source Report'!$M$M,0))-1


* Engineered features from the text of each job description to quantify the value companies put on python, excel, aws, and spark. 
* Optimized Linear, Lasso, and Random Forest Regressors using GridsearchCV to reach the best model. 
* Built a client facing API using flask 
