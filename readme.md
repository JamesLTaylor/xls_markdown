# Create Excel Workbooks from Markdown

## Overview

This library allows you to write plain text documents that can be converted to 
excel workbooks. The main advantages of this are:

 * you can modify the format of published sheets at anytime
 * the contents of workbooks can be properly managed in your source control system
 * the structure and format of all your workbooks is standardized
 
## Uses

This library has been written to allow [QuantSA](www.QuantSA.org) to replace its 
example workbooks with markdown versions that will lead to more standardized 
structures and better source control.  The examples below use functions from QuantSA
but this tool can easily be used for any other library.

## Example

As an example of how this works; if you want an example sheet that demonstrates
how to get an accrual fraction you would write a text file with:

```
# Dates

[B2]
## Accrual Fractions
valueDate,2018-07-27
settlementDate,2018-10-27

timeToExercise,=QSA.GetYearFraction(valueDate,settlementDate,"ACT365")[0.00]
```

and after generating the workbook you would obtain:

![example1](images/tiny_example.PNG)

### Explanation

|markdown|explanation|
|---|---|
|```# Dates```|Start a new tab with this name|
|||
|```[B2]```|The contents below this will be inserted from cell B2|
|```## Accrual Fractions```|Add a heading|
|```valueDate,2018-07-27```|Add a date, give a name to the date for use in formulas below |
|```settlementDate,2018-10-27```||
|||
|```timeToExercise,=QSA.GetYearFraction(valueDate,settlementDate,"ACT365")[0.00]```| Call a function using the valued defined above.|

## Bigger Example

![generated sheet](images/generated.PNG)