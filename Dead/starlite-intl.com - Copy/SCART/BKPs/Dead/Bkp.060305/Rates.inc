


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    
    If weight <= 0.37 Then
    	  cost = 0.00

    ElseIf weight <= 2 Then
    	  cost = 7.87

    ElseIf weight <= 5 Then
        cost = 10.03

    ElseIf weight <= 7 Then
        cost = 11.11

    ElseIf weight <= 9 Then
        cost = 12.83

    ElseIf weight <= 11 Then
        cost = 13.91

    ElseIf weight <= 13 Then
        cost = 14.76

    ElseIf weight <= 43 Then
        cost = 14.76 + (weight - 13) / 2.0 + 0.95

    ElseIf weight <= 45 Then
        cost = 33.33

    ElseIf weight <= 47 Then
        cost = 34.10

    ElseIf weight <= 49 Then
        cost = 34.53

    ElseIf weight <= 51 Then
        cost = 35.63

    ElseIf weight <= 53 Then
        cost = 36.39

    ElseIf weight <= 55 Then
        cost = 36.66

    ElseIf weight <= 57 Then
        cost = 37.1

    ElseIf weight <= 59 Then
        cost = 38.05

    ElseIf weight <= 61 Then
        cost = 38.47

    ElseIf weight <= 63 Then
        cost = 38.79

    ElseIf weight <= 65 Then
        cost = 39.36

    ElseIf weight <= 67 Then
        cost = 39.88

    ElseIf weight <= 70 Then
        cost = 40.59

    ElseIf weight <= 75 Then
        cost = 48.69

    ElseIf weight <= 80 Then
        cost = 51.91

    ElseIf weight <= 85 Then
        cost = 53.36

    ElseIf weight <= 90 Then
        cost = 56.97

    ElseIf weight <= 95 Then
        cost = 59.49

    ElseIf weight <= 100 Then
        cost = 61.96

    ElseIf weight <= 110 Then
        cost = 66.52

    ElseIf weight <= 120 Then
        cost = 71.78

    ElseIf weight <= 130 Then
        cost = 78.02

    ElseIf weight <= 140 Then
        cost = 85.87

    ElseIf weight <= 150 Then
        cost = 91.05

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>