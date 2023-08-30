


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    
    If weight <= 0.23 Then
    	  cost = 1.50

    ElseIf weight <= 0.33 Then
    	  cost = 3.13

    ElseIf weight <= 0.56 Then
    	  cost = 11.79

    ElseIf weight <= 1.0 Then
    	  cost = 13.28

    ElseIf weight <= 1.5 Then
    	  cost = 18.35

    ElseIf weight <= 3.3 Then
        cost = 20.47

    ElseIf weight <= 6.5 Then
        cost = 24.85

    ElseIf weight <= 8.5 Then
        cost = 26.89

    ElseIf weight <= 11 Then
        cost = 27.54

    ElseIf weight <= 13 Then
        cost = 28.85

    ElseIf weight <= 43 Then
        cost = 26.13 + (weight - 13) / 2.0 + 1.77

    ElseIf weight <= 45 Then
        cost = 53.45

    ElseIf weight <= 47 Then
        cost = 53.99

    ElseIf weight <= 49 Then
        cost = 54.76

    ElseIf weight <= 51 Then
        cost = 57.78

    ElseIf weight <= 53 Then
        cost = 57.56

    ElseIf weight <= 55 Then
        cost = 59.27

    ElseIf weight <= 57 Then
        cost = 61.00

    ElseIf weight <= 59 Then
        cost = 61.36

    ElseIf weight <= 61 Then
        cost = 63.22

    ElseIf weight <= 63 Then
        cost = 63.92

    ElseIf weight <= 65 Then
        cost = 62.25

    ElseIf weight <= 67 Then
        cost = 65.99

    ElseIf weight <= 70 Then
        cost = 67.98

    ElseIf weight <= 75 Then
        cost = 74.87

    ElseIf weight <= 80 Then
        cost = 79.77

    ElseIf weight <= 85 Then
        cost = 81.94

    ElseIf weight <= 90 Then
        cost = 87.41

    ElseIf weight <= 95 Then
        cost = 91.24

    ElseIf weight <= 100 Then
        cost = 94.97

    ElseIf weight <= 110 Then
        cost = 101.86

    ElseIf weight <= 120 Then
        cost = 109.82

    ElseIf weight <= 130 Then
        cost = 119.36

    ElseIf weight <= 140 Then
        cost = 131.59

    ElseIf weight <= 150 Then
        cost = 140.34

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>