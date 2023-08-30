


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    
    If weight <= 0.37 Then
    	  cost = 0.00

    ElseIf weight <= 2 Then
    	  cost = 8.97

    ElseIf weight <= 5 Then
        cost = 10.98

    ElseIf weight <= 7 Then
        cost = 11.86

    ElseIf weight <= 9 Then
        cost = 13.58

    ElseIf weight <= 11 Then
        cost = 14.65

    ElseIf weight <= 13 Then
        cost = 15.51

    ElseIf weight <= 43 Then
        cost = 15.56 + (weight - 13) / 2.0 + 0.95

    ElseIf weight <= 45 Then
        cost = 34.14

    ElseIf weight <= 47 Then
        cost = 34.90

    ElseIf weight <= 49 Then
        cost = 35.33

    ElseIf weight <= 51 Then
        cost = 36.43

    ElseIf weight <= 53 Then
        cost = 37.19

    ElseIf weight <= 55 Then
        cost = 37.46

    ElseIf weight <= 57 Then
        cost = 37.90

    ElseIf weight <= 59 Then
        cost = 38.85

    ElseIf weight <= 61 Then
        cost = 39.27

    ElseIf weight <= 63 Then
        cost = 39.59

    ElseIf weight <= 65 Then
        cost = 40.16

    ElseIf weight <= 67 Then
        cost = 40.68

    ElseIf weight <= 70 Then
        cost = 41.39

    ElseIf weight <= 75 Then
        cost = 49.49

    ElseIf weight <= 80 Then
        cost = 52.71

    ElseIf weight <= 85 Then
        cost = 54.16

    ElseIf weight <= 90 Then
        cost = 57.77

    ElseIf weight <= 95 Then
        cost = 60.29

    ElseIf weight <= 100 Then
        cost = 62.76

    ElseIf weight <= 110 Then
        cost = 67.32

    ElseIf weight <= 120 Then
        cost = 72.58

    ElseIf weight <= 130 Then
        cost = 78.89

    ElseIf weight <= 140 Then
        cost = 86.97

    ElseIf weight <= 150 Then
        cost = 92.75

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>