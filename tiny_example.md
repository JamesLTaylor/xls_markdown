# Dates

[B2]
## Accrual Fractions
convention,[ACT365,ACT360,_30360]ACT365
valueDate,2018-07-27
settlementDate,2018-10-27

timeToExercise,=QSA.GetYearFraction(valueDate,settlementDate,convention)[0.00]