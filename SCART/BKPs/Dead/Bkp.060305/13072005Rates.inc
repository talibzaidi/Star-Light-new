


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    If weight <= 2 Then
    	  cost = 6.97

    ElseIf weight <= 5 Then
        cost = 8.97

    ElseIf weight <= 7 Then
        cost = 9.98

    ElseIf weight <= 9 Then
        cost = 11.48

    ElseIf weight <= 11 Then
        cost = 12.48

    ElseIf weight <= 13 Then
        cost = 13.27

    ElseIf weight <= 43 Then
        cost = 13.27 + (weight - 13) / 2.0 + 0.95

    ElseIf weight <= 45 Then
        cost = 30.46

    ElseIf weight <= 47 Then
        cost = 31.07

    ElseIf weight <= 49 Then
        cost = 31.47

    ElseIf weight <= 51 Then
        cost = 32.49

    ElseIf weight <= 53 Then
        cost = 33.19

    ElseIf weight <= 55 Then
        cost = 33.44

    ElseIf weight <= 57 Then
        cost = 33.80

    ElseIf weight <= 59 Then
        cost = 34.20

    ElseIf weight <= 61 Then
        cost = 34.62

    ElseIf weight <= 63 Then
        cost = 34.90

    ElseIf weight <= 65 Then
        cost = 35.32

    ElseIf weight <= 67 Then
        cost = 35.78

    ElseIf weight <= 70 Then
        cost = 36.48

    ElseIf weight <= 75 Then
        cost = 43.98

    ElseIf weight <= 80 Then
        cost = 46.96

    ElseIf weight <= 85 Then
        cost = 48.31

    ElseIf weight <= 90 Then
        cost = 51.65

    ElseIf weight <= 95 Then
        cost = 53.98

    ElseIf weight <= 100 Then
        cost = 56.27

    ElseIf weight <= 110 Then
        cost = 60.49

    ElseIf weight <= 120 Then
        cost = 65.36

    ElseIf weight <= 130 Then
        cost = 70.24

    ElseIf weight <= 140 Then
        cost = 74.50

    ElseIf weight <= 150 Then
        cost = 79.11

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>