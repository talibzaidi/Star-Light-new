


<% 
' [BN, 4/29/04] Sub to compute shipping and handling charge, based on total weight of shopping cart.
' ExchangeRate = 1 for USA and some number greater than 1 for Canada.

MaxSandH = 99.99  ' This is the amount for S&H when the weight is beyond the max in the rate table. 

Function SandH(weight, ExchangeRate)
 
    
    If weight <= 0.23 Then
    	  cost = 1.00

    ElseIf weight <= 0.33 Then
    	  cost = 2.65

    ElseIf weight <= 0.56 Then
    	  cost = 9.99

    ElseIf weight <= 1.0 Then
    	  cost = 11.25

    ElseIf weight <= 1.5 Then
    	  cost = 15.55

    ElseIf weight <= 3.3 Then
        cost = 17.35

    ElseIf weight <= 6.5 Then
        cost = 21.06

    ElseIf weight <= 8.5 Then
        cost = 22.79

    ElseIf weight <= 11 Then
        cost = 23.34

    ElseIf weight <= 13 Then
        cost = 24.45

    ElseIf weight <= 43 Then
        cost = 22.14 + (weight - 13) / 2.0 + 1.50

    ElseIf weight <= 45 Then
        cost = 45.30

    ElseIf weight <= 47 Then
        cost = 45.76

    ElseIf weight <= 49 Then
        cost = 46.41

    ElseIf weight <= 51 Then
        cost = 48.97

    ElseIf weight <= 53 Then
        cost = 48.78

    ElseIf weight <= 55 Then
        cost = 50.23

    ElseIf weight <= 57 Then
        cost = 51.70

    ElseIf weight <= 59 Then
        cost = 52.00

    ElseIf weight <= 61 Then
        cost = 53.58

    ElseIf weight <= 63 Then
        cost = 54.17

    ElseIf weight <= 65 Then
        cost = 52.75

    ElseIf weight <= 67 Then
        cost = 55.92

    ElseIf weight <= 70 Then
        cost = 57.61

    ElseIf weight <= 75 Then
        cost = 63.45

    ElseIf weight <= 80 Then
        cost = 67.60

    ElseIf weight <= 85 Then
        cost = 69.44

    ElseIf weight <= 90 Then
        cost = 74.08

    ElseIf weight <= 95 Then
        cost = 77.30

    ElseIf weight <= 100 Then
        cost = 80.48

    ElseIf weight <= 110 Then
        cost = 86.32

    ElseIf weight <= 120 Then
        cost = 93.07

    ElseIf weight <= 130 Then
        cost = 101.15

    ElseIf weight <= 140 Then
        cost = 111.52

    ElseIf weight <= 150 Then
        cost = 118.93

 Else cost = MaxSandH     ' Defined above.
    End If

SandH = cost * ExchangeRate

End Function

%>