# Description

[B2]
## Summary

[text]
The following sheet demonstrates creating a QuantSA zero curve and 
calling the Black formula

The sheet demonstrates:
 * QSA.CreateDatesAndRatesCurve
 * QSA.FormulaBlack
[endtext]

# BlackFormula

[B2]
valueDate,2018-07-27

## Caplet
callOrPut,[Call,Put]Call
strike,7.00%
notional,1000000
exerciseDate,2018-10-27
settlementDate,2018-10-27
accrueStartDate,2018-10-27
accrueEndDate,2019-01-27

## Black Formula
strike,=forward[0.00%]
timeToExercise,=QSA.GetYearFraction(valueDate,settlementDate,"ACT365")[0.00]
forward,=QSA.GetSimpleForward(zarCurve,accrueStartDate,accrueEndDate)[0.00%]
vol,15%
discountFactor,=QSA.GetDF(zarCurve,settlementDate)[0.00]
accrualFraction,=QSA.GetYearFraction(accrueStartDate,accrueEndDate)[0.00]
	
caplet value,=1000000*accrualFraction*QSA.FormulaBlack(callOrPut,strike,timeToExercise,forward,vol,discountFactor)[0.00]

[E2]
[table_c]
dates, continuousRates
2018-07-27,7.10%
2019-07-27,7.20%
2020-07-27,7.30%
[endtable_c]

[create_h]
zarCurve,=QSA.CreateDatesAndRatesCurve(**,dates,continuousRates)
