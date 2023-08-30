


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    
    If weight <= 0.23 Then
    	  cost = 0.00

    ElseIf weight <= 0.50 Then
    	  cost = 7.69

    ElseIf weight <= 2 Then
    	  cost = 9.89

    ElseIf weight <= 5 Then
        cost = 12.11

    ElseIf weight <= 7 Then
        cost = 13.07

    ElseIf weight <= 9 Then
        cost = 14.98

    ElseIf weight <= 11 Then
        cost = 16.15

    ElseIf weight <= 13 Then
        cost = 17.10

    ElseIf weight <= 43 Then
        cost = 17.16 + (weight - 13) / 2.0 + 1.00

    ElseIf weight <= 45 Then
        cost = 37.64

    ElseIf weight <= 47 Then
        cost = 38.48

    ElseIf weight <= 49 Then
        cost = 38.96

    ElseIf weight <= 51 Then
        cost = 40.16

    ElseIf weight <= 53 Then
        cost = 41.00

    ElseIf weight <= 55 Then
        cost = 41.30

    ElseIf weight <= 57 Then
        cost = 41.79

    ElseIf weight <= 59 Then
        cost = 42.83

    ElseIf weight <= 61 Then
        cost = 43.23

    ElseIf weight <= 63 Then
        cost = 43.65

    ElseIf weight <= 65 Then
        cost = 44.28

    ElseIf weight <= 67 Then
        cost = 44.85

    ElseIf weight <= 70 Then
        cost = 45.74

    ElseIf weight <= 75 Then
        cost = 54.56

    ElseIf weight <= 80 Then
        cost = 58.12

    ElseIf weight <= 85 Then
        cost = 59.71

    ElseIf weight <= 90 Then
        cost = 63.7

    ElseIf weight <= 95 Then
        cost = 66.47

    ElseIf weight <= 100 Then
        cost = 69.20

    ElseIf weight <= 110 Then
        cost = 74.22

    ElseIf weight <= 120 Then
        cost = 80.02

    ElseIf weight <= 130 Then
        cost = 86.97

    ElseIf weight <= 140 Then
        cost = 95.89

    ElseIf weight <= 150 Then
        cost = 102.26

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>