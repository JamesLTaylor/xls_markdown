# Description

[A2]
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
strike,=C17
timeToExercise,=QSA.GetYearFraction(valueDate,settlementDate,"ACT365")
forward,=QSA.GetSimpleForward(zarCurve,accrueStartDate,accrueEndDate)
vol,15%
discountFactor,=QSA.GetDF(zarCurve,settlementDate)
accrualFraction,=QSA.GetYearFraction(accrueStartDate,accrueEndDate)
	
caplet value,=1000000*accrualFraction*QSA.FormulaBlack(callOrPut,strike,timeToExercise,forward,vol,discountfactor)

[G2]
[table_c]
dates, continuousRates
2018-07-27,7.10%
2019-07-27,7.20%
2020-07-27,7.30%
[endtable_c]

[create_h]
zarCurve,=QSA.CreateDatesAndRatesCurve("name",dates,continuousRates)
