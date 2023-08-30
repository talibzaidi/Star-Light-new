


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    
    If weight <= 0.23 Then
    	  cost = 0.00

    ElseIf weight <= 0.50 Then
    	  cost = 7.32

    ElseIf weight <= 2 Then
    	  cost = 9.42

    ElseIf weight <= 5 Then
        cost = 11.53

    ElseIf weight <= 7 Then
        cost = 12.45

    ElseIf weight <= 9 Then
        cost = 14.27

    ElseIf weight <= 11 Then
        cost = 15.38

    ElseIf weight <= 13 Then
        cost = 16.29

    ElseIf weight <= 43 Then
        cost = 16.34 + (weight - 13) / 2.0 + 1.00

    ElseIf weight <= 45 Then
        cost = 35.85

    ElseIf weight <= 47 Then
        cost = 36.65

    ElseIf weight <= 49 Then
        cost = 37.10

    ElseIf weight <= 51 Then
        cost = 38.25

    ElseIf weight <= 53 Then
        cost = 39.05

    ElseIf weight <= 55 Then
        cost = 39.33

    ElseIf weight <= 57 Then
        cost = 39.80

    ElseIf weight <= 59 Then
        cost = 40.79

    ElseIf weight <= 61 Then
        cost = 41.23

    ElseIf weight <= 63 Then
        cost = 41.57

    ElseIf weight <= 65 Then
        cost = 42.17

    ElseIf weight <= 67 Then
        cost = 42.71

    ElseIf weight <= 70 Then
        cost = 43.56

    ElseIf weight <= 75 Then
        cost = 51.96

    ElseIf weight <= 80 Then
        cost = 55.35

    ElseIf weight <= 85 Then
        cost = 56.87

    ElseIf weight <= 90 Then
        cost = 60.67

    ElseIf weight <= 95 Then
        cost = 63.30

    ElseIf weight <= 100 Then
        cost = 65.90

    ElseIf weight <= 110 Then
        cost = 70.69

    ElseIf weight <= 120 Then
        cost = 76.21

    ElseIf weight <= 130 Then
        cost = 82.83

    ElseIf weight <= 140 Then
        cost = 91.32

    ElseIf weight <= 150 Then
        cost = 97.39

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>