


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    
    If weight <= 0.23 Then
    	  cost = 0.00

    ElseIf weight <= 0.50 Then
    	  cost = 8.09

    ElseIf weight <= 2 Then
    	  cost = 10.97

    ElseIf weight <= 5 Then
        cost = 12.77

    ElseIf weight <= 7 Then
        cost = 13.79

    ElseIf weight <= 9 Then
        cost = 15.79

    ElseIf weight <= 11 Then
        cost = 17.04

    ElseIf weight <= 13 Then
        cost = 18.04

    ElseIf weight <= 43 Then
        cost = 18.10 + (weight - 13) / 2.0 + 1.00

    ElseIf weight <= 45 Then
        cost = 39.71

    ElseIf weight <= 47 Then
        cost = 40.51

    ElseIf weight <= 49 Then
        cost = 41.10

    ElseIf weight <= 51 Then
        cost = 42.37

    ElseIf weight <= 53 Then
        cost = 43.25

    ElseIf weight <= 55 Then
        cost = 43.57

    ElseIf weight <= 57 Then
        cost = 44.09

    ElseIf weight <= 59 Then
        cost = 45.19

    ElseIf weight <= 61 Then
        cost = 45.61

    ElseIf weight <= 63 Then
        cost = 46.05

    ElseIf weight <= 65 Then
        cost = 46.72

    ElseIf weight <= 67 Then
        cost = 47.32

    ElseIf weight <= 70 Then
        cost = 48.26

    ElseIf weight <= 75 Then
        cost = 57.56

    ElseIf weight <= 80 Then
        cost = 61.32

    ElseIf weight <= 85 Then
        cost = 62.99

    ElseIf weight <= 90 Then
        cost = 67.20

    ElseIf weight <= 95 Then
        cost = 70.12

    ElseIf weight <= 100 Then
        cost = 73.00

    ElseIf weight <= 110 Then
        cost = 78.30

    ElseIf weight <= 120 Then
        cost = 84.42

    ElseIf weight <= 130 Then
        cost = 91.75

    ElseIf weight <= 140 Then
        cost = 101.16

    ElseIf weight <= 150 Then
        cost = 107.88

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>