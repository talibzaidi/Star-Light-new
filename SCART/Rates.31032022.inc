


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    
    If weight <= 0.23 Then
    	  cost = 0.00

    ElseIf weight <= 0.33 Then
    	  cost = 2.49

    ElseIf weight <= 0.56 Then
    	  cost = 9.42

    ElseIf weight <= 1.0 Then
    	  cost = 10.49

    ElseIf weight <= 1.5 Then
    	  cost = 14.67

    ElseIf weight <= 3.3 Then
        cost = 16.37

    ElseIf weight <= 6.5 Then
        cost = 19.87

    ElseIf weight <= 8.5 Then
        cost = 21.50

    ElseIf weight <= 11 Then
        cost = 22.02

    ElseIf weight <= 13 Then
        cost = 23.07

    ElseIf weight <= 43 Then
        cost = 20.89 + (weight - 13) / 2.0 + 1.00

    ElseIf weight <= 45 Then
        cost = 42.74

    ElseIf weight <= 47 Then
        cost = 43.17

    ElseIf weight <= 49 Then
        cost = 43.78

    ElseIf weight <= 51 Then
        cost = 45.10

    ElseIf weight <= 53 Then
        cost = 46.02

    ElseIf weight <= 55 Then
        cost = 47.39

    ElseIf weight <= 57 Then
        cost = 48.77

    ElseIf weight <= 59 Then
        cost = 49.07

    ElseIf weight <= 61 Then
        cost = 50.55

    ElseIf weight <= 63 Then
        cost = 51.01

    ElseIf weight <= 65 Then
        cost = 52.75

    ElseIf weight <= 67 Then
        cost = 53.37

    ElseIf weight <= 70 Then
        cost = 54.35

    ElseIf weight <= 75 Then
        cost = 59.86

    ElseIf weight <= 80 Then
        cost = 63.77

    ElseIf weight <= 85 Then
        cost = 65.51

    ElseIf weight <= 90 Then
        cost = 69.89

    ElseIf weight <= 95 Then
        cost = 72.92

    ElseIf weight <= 100 Then
        cost = 75.92

    ElseIf weight <= 110 Then
        cost = 81.43

    ElseIf weight <= 120 Then
        cost = 87.80

    ElseIf weight <= 130 Then
        cost = 95.42

    ElseIf weight <= 140 Then
        cost = 105.21

    ElseIf weight <= 150 Then
        cost = 112.20

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>