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

> * =COUNTIF('Source Report'!$M:$M," * abc * ")
> > Searches the specified range for any Cells that contain a pre-determined Text-String.  Then, count how many instances of the Text-String are within the range and output the Numerical Value.  (This value will be used in a later formula)

##

> * =CONCATENATE($B$2," - ","Page:  ",$F$2)
> > Concatenates a pre-determined Text-String with the Numerical Value from the above formula, to create a true Uniqe Identifier.  (This will be used in a later formula)

##

> * =ABS(MATCH($G$2,'Source Report'!$M:$M,0)-MATCH('Source Report'!$G$3,'Source Report'!$M:$M,0))-1
> > Checks a pre-determined Text-String within a pre-determined Range and outputs the absolute value of your criteria and then compares it to a second pre-determined Text-String withing the range.  Next, it will find the difference between the absolute values and will output the Numerical Value.  (This will be used in a later formula)

##

> * =IFNA(VLOOKUP($B3,OFFSET('Source Report'!$A$1,MATCH($G$2,'Source Report'!$M:$M,0),0,$J$3,3),3,0),0)
> > Uses a pre-determined Cell as a Lookup-Value to be used with the VLookup.  Then, while using the Offset Function to start the Match function at the top row of the Source Report, Match a pre-determined Cell to it's Unique Identifier withing a pre-determined Range.  Then using the value obtained from the formula above, that will decide how many rows are needed for the Offset function, in order to guarantee no overlaps in data.  Lastly, finishing the VLookup, after finding the correct Match via using the Unique Identifier created in an earlier formula, output the data from the third column of the range as a Numerical Value.  (This formula is repeated one cell to the left and will pull from the second column instead of the third.)

##

