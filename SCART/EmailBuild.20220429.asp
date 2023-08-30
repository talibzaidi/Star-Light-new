<%@ LANGUAGE = VBScript %>

<% response.buffer=true %>

<% ar = Request("Area") %>
<% Area = Request("Area") %>
<% sar = ReQuest("sar") %>
<% SID = ReQuest("SID") %>
<% spec = 2 %>
<% ar = Request("Area") %>
<% Area = Request("Area") %>
<% sar = ReQuest("sar") %>
<% SID = ReQuest("SID") %>
<% spec = 2 %>

<%
' 4/22/07, BN.
' 5/2/07, BN: No longer needed. Replaced by the Validator1(theForm) javascript function in Scart.inc.
Sub IncompleteData(FieldName)
    Response.Write "<br><br><br><br><br><br><center>" & _
					"<font face=Tahoma color=red size=4>You did not give <b>" & FieldName & "</b>. " & _
					"<br><br>Use your browser's Back button to return.</font>" & _
					"</center>"
	Response.End
End Sub    ' IncompleteData
%>


<% 
If False Then   ' 5/2/07, BN: No longer needed. Replaced by the Validator1(theForm) javascript function in Scart.inc.
   If ReQuest("Email") = "" Then ' "email@domain.com"
		IncompleteData("an Email Address")
		'response.Redirect "formcorrection.asp?" & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   ElseIf ReQuest("Address") = "" Then
   		IncompleteData("a Street Address")
		'response.Redirect "formcorrection.asp?" & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   ElseIf ReQuest("Phone") = "" Then
		IncompleteData("a Phone Number")
		'response.Redirect "formcorrection.asp?" & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   ElseIf ReQuest("StateProv") = "" Then
		IncompleteData("a State or Province")
		'response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   ElseIf ReQuest("Country") = "SELECT COUNTRY" Then
 		IncompleteData("a Country")
		'response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   ElseIf ReQuest("Name") = "" Then
   		IncompleteData("a Name")
		'response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   ElseIf ReQuest("Postal") = "" Then
     	IncompleteData("a Postal Code")
		' response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   ElseIf ReQuest("City") = "" Then
      	IncompleteData("a City")
		' response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   ElseIf ReQuest("Payment") = "Choose" Then
		IncompleteData("a Payment Method")
		'response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   End If
End If
   
LName   = ReQuest("LName")
Address = ReQuest("Address")
' To remove any possible apostrophes in LName and Address ...
LName   = Replace(LName, "'", "&apos;")
Address = Replace(Address, "'", "&apos;")

'LName = Replace(LName, "'", "&#8217;")
'LName = Replace(LName, "'", "'")

Response.Write("**LName = "    & LName)      & "<br>"

Session("CustomerName") 	= ReQuest("FName") & " " & ReQuest("LName")
Session("FName") 			= ReQuest("FName")
' Session("LName") 			= ReQuest("LName")
Session("LName") 			= LName
Session("Company")			= ReQuest("Company")
Session("Email")			= ReQuest("Email")
Session("Phone")			= ReQuest("Phone")
'Session("Address") 		= ReQuest("Address")
Session("Address") 			= Address
Session("City") 			= Request("City")
Session("StateProv")		= ReQuest("StateProv")
Session("Postal") 			= ReQuest("Postal")
Session("Country")			= ReQuest("Country")
Session("CustomerEmail") 	= ReQuest("Email")
Session("PaymentMeth")      = ReQuest("Payment")   ' e.g. Visa, Master Card, Discover, American Express, Check, Money Order, Pay Pal


If TRUE Then
    Response.Write("LName = "    & LName)      & "<br>"
    Response.Write("Address = "  & Address)    & "<br>"
End If

'Response.Write "<br>Session('Country') = " & Session("Country")  ' ***
%>


<%
' 11/10/15: Same as Function asGetItemParameters(iItemID) in Scart.inc.SRs.asp? So redundant here?
Function asGetItemParameters(iItemID)  
    Dim bParameters 
    iItemID = Trim(iItemID)     ' Added 11/10/15.

    'If Session("Country") = "USA" Then    
     
    ' 11/10/15: The code for the following 2 cases is the same except for the "bParameters =..." lines (?)
    ' So I could condense this If-Then-Else statement (?)
    If Session("Country") <> "Canada" Then     ' 1/15/09: Started treating all countries besides Canada the same as USA.
        'RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID +   "' ", "DSN=STAREC1" , 1, 4

        ' 11/10/15: Using the connection string method instead of the DSN method above.
        Set ConnUSA = Server.CreateObject("ADODB.Connection")
        ConnUSA.Open Session("ConnectionString")
        SQLstring = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID + "' "
        'RXS.Open SQLstring, ConnUSA, adOpenStatic, adLockOptimistic, adCmdText 
        RXS.Open SQLstring, ConnUSA, 3, 3, &H0001 

	    bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), 1 )
    Else   										' 1/15/09: i.e. Session("Country") = "Canada"
	    'RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID + "' ", "DSN=STAREC1" , 1, 4
	
        ' 11/10/15: Using the connection string method instead of the DSN method above.
        Set ConnCanada = Server.CreateObject("ADODB.Connection")
        ConnCanada.Open Session("ConnectionString")
        SQLstring = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID + "' "
        'RXS.Open SQLstring, ConnCanada, adOpenStatic, adLockOptimistic, adCmdText 
        RXS.Open SQLstring, ConnCanada, 3, 3, &H0001 
    
        bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("Pname") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Freight")*RXS("Exch")), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")*(1/(1-(RXS("GPM"))))),RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), RXS("Exch")  )
    End if
    ' Return array containing product info.
    asGetItemParameters = bParameters
End Function    ' asGetItemParameters
%>

<%
' Makes a row of the head (intro) part of the table for the Order Email.
Function MakeHeadRow(Label, Value, colSpan)
    MakeHeadRow = "<tr>" & _
	    "<td bgcolor=silver><font face=Tahoma><b>" & Label & "</b></font></td>" & _
	    "<td colspan=" & colSpan & " bgcolor=lightGrey><font face=Tahoma>" & Value & "</b></font></td>" & _
	    "</tr>"
End Function	' MakeHeadRow


' Make a generic row of the table for the Order Email.
Function MakeRow(Label, colSpan1, Value, colSpan2)
    MakeRow = "<tr>" & _
	    "<td colspan=" & colSpan1 & " bgcolor=silver align=right><font face=Tahoma><b>" & Label & "</b></font></td>" & _
	    "<td colspan=" & colSpan2 & " bgcolor=lightGrey><font face=Tahoma>" & Value & "</b></font></td>" & _
	    "</tr>"
End Function	' MakeRow


' Makes a row of the tail part of the table for the Order Email.
Function MakeTailRow(Label, Value)
    MakeTailRow = "<tr>" & _
	    "<td bgcolor=silver colspan=4 align=right><font face=Tahoma><b>" & Label & "</b></font></td>" & _
	    "<td bgcolor=lightGrey align=right><font face=Tahoma color=#b9000>" & Value & "</b></font></td>" & _
	    "</tr>"
End Function	' MakeTailRow


' Makes a row of the middle part of the table for the Order Email to be sent to Starlite.
Function MakeMiddleRowForStarlite(val1, val2, val3, val4, val5, val6)
    MakeMiddleRowForStarlite = "<tr>" & _
	    "<td bgcolor=lightGrey align=center><font face=Tahoma>" & val1 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=left><font face=Tahoma>" & val2 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=center><font face=Tahoma>" & val3 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=right><font face=Tahoma>" & val4 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=right><font face=Tahoma>" & val5 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=right><font face=Tahoma>" & val6 & "</font></td>" & _
	    "</tr>"
End Function	' MakeMiddleRowForStarlite

' Makes a row of the middle part of the table for the Order Email to be sent to customer.
Function MakeMiddleRowForCustomer(val1, val2, val3, val4, val5)
    MakeMiddleRowForCustomer = "<tr>" & _
	    "<td bgcolor=lightGrey align=center><font face=Tahoma>" & val1 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=left><font face=Tahoma>" & val2 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=center><font face=Tahoma>" & val3 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=right><font face=Tahoma>" & val4 & "</font></td>" & _
	    "<td bgcolor=lightGrey align=right><font face=Tahoma>" & val5 & "</font></td>" & _
	    "</tr>"
End Function	' MakeMiddleRowForCustomer
%>




<html>


<head>
<meta name="keywords" content="GPS,Navigation,Garmin,CB-Radios,Uniden,Cobra,Motorola,2-way radios,Hand-tools,Pyramid ">
<meta name="description" content="Online store for GPS Global Positioning Systems, Navigation equipment, CB Radios, FRS Radios, GMRS Radios, Antennas, Car Audio, Hand Tools.  Shopping on a secure SSL line. Accepting Visa,
Mastercard, Discover, American Express cards.">
<title>Starlite International LLC - Online Store</title>
<script language="Javascript">
<!--
	once = new MakeArray(6)
	over = new MakeArray(6)
	under = new MakeArray(6)
	standard = new MakeArray(1)
	once[0].src = "../Images/question1.gif"
	once[1].src = "../Images/scart1.gif"
	once[2].src = "../Images/home1.gif"
	once[3].src = "../Images/new1.gif"
                once[4].src = "../Images/cat1.gif"
	once[5].src = "../Images/ex1.gif"    
	over[0].src = "../Images/question2.gif"
	over[1].src = "../Images/scart2.gif"
	over[2].src = "../Images/home2.gif"
	over[3].src = "../Images/new2.gif"
	over[4].src = "../Images/cat2.gif"
	over[5].src = "../Images/ex2.gif"
	under[0].src = "../Images/helpnav.gif"
	under[1].src = "../Images/shoppingcartnav.gif"
	under[2].src = "../Images/homenav.gif"
	under[3].src = "../Images/newproductsnav.gif"
	under[4].src = "../Images/onlinecataloguenav.gif"
	under[5].src = "../Images/specialsnav.gif"
	standard[0].src = "../Images/emptynav.jpg"
function MakeArray(n) 

	{

	this.length = n

	for (var i = 1; i<=n; i++) 

		{

		this[i-1] = new Image()

		}

	return this

	}

function msover(inum,d_inum) 

	{

		if ((over[inum].src != "")) 

			{

			document.images[d_inum].src = over[inum].src
			document.images[7].src = under[inum].src
			}

	}


function msout(inum,d_inum) 

	{

		if ((once[inum].src != "")) 

			{

			document.images[d_inum].src = once[inum].src
			document.images[7].src = standard[0].src
			}

	}

// -->
</script>
</head>



<body>

<!--#include file="RATES.INC"-->


<%
'Address=Request("Address")   ' Already done above.
City=Request("City")
Payment=Request("Payment") 
expfrom=Request("expfrom") 
expto=Request("expto") 
Cardholder=Request("Cardholder")




If Err.number <> 0 Then
     response.redirect "error.asp"
end if
%>

<%
' Get a reference to the cart if it exists otherwise create it
If IsObject(Session("cart")) Then
	Set dictCart = Session("cart")
Else
	' We use a dictionary so we can name our keys to correspond to our
	' item numbers and then use their value to hold the quantity.  An
	' array would also work, but would be a little more complex and
	' probably not as easy for readers to follow.
	Set dictCart = Server.CreateObject("Scripting.Dictionary")
End If

'19 Get all the parameters passed to the script
set RXS = CreateObject("ADODB.Recordset")
	
%>

<%
' weightfactor = 0
if request("Payment") = "C.O.D." Then 
	shipcrg = 8.75
end if
Dim Key
Dim aParameters ' as Variant (Array)
Dim sTotal, sShipping

If Request("OptOutOfEmailing") = "on" Then
	OptedIn = "No"
	OptedIn2 = "False"
Else
	OptedIn = "Yes"
	OptedIn2 = "True"
End If

' 1/6/09: A customer gets our emails based on their value of "OptedIn2 AND WeEmail".
' This WeEmail = "True" had by accident been left out till now (1/6/09), and the WeEmail field of Orders db table was apparently defaulting to False!
WeEmail = "True"	' This is the default for WeEmail, i.e. We email based on truth value of "OptedIn2 AND True" = truth value of "OptedIn2".
					' Sani can later (using for e.g. Admin2/Orders.asp) set WeEmail = False in the Orders db table, to override customer's OptedIn2 value.

If Request("BigShip") = "on" Then
	BigShipment = "True"
Else
	BigShipment = "False"
End If

shpcrg = 0



	
sTotal = 0
weightTotal = 0		' [BN, 4/29/04] Added this.

Nw = Now()
'OrderNum = Right(Year(Nw),2) & "0" & Month(Nw) & "0" & Day(Nw) & "0" & Hour(Nw) & "0" & Minute(Nw)
OrderNum = Right(Year(Nw),2) & "-" & Month(Nw) & "-" & Day(Nw) & "-" & Hour(Nw) & "-" & Minute(Nw) & "-" & Second(Nw)
OrderDate = Nw 
Session("OrderNum")     = OrderNum



InvoiceHead =	MakeRow("Order Number:", 1, OrderNum, 4) & _
                MakeRow("Name:", 1, ReQuest("FName") & " " & LName, 4) & _
                MakeRow("Company:", 1, ReQuest("Company"), 4) & _
				MakeRow("Email:", 1, ReQuest("Email"), 4) & _
				MakeRow("Opt in to Emailings:", 1, OptedIn, 4) & _
				MakeRow("Phone:", 1, ReQuest("Phone"), 4) & _
				MakeRow("Address:", 1, Address, 4) & _
				MakeRow("City:", 1, ReQuest("City"), 4) & _
				MakeRow("State/Province:", 1, ReQuest("StateProv"), 4) & _
				MakeRow("Postal Code:", 1, ReQuest("Postal"), 4) & _
				MakeRow("Country:", 1, ReQuest("Country"), 4) & _
				MakeRow("Payment Method:", 1, ReQuest("Payment"), 4) & _
				MakeRow("Tax 1:", 1, ReQuest("Taxx1") & "%", 4) & _
				MakeRow("Tax 2:", 1, ReQuest("Taxx2") & "%", 4) & _
				MakeRow("&nbsp;", 1, "&nbsp;", 4)


' 4/18/07 BN: Continue creating body of the email to send to Starlite ...
InvoiceMiddleForStarlite = MakeMiddleRowForStarlite("<b>ID #</b>", "<b>Description</b>", "<b>Qty</b>", "<b>Price</b>", "<b>Totals</b>", "<b>Weight</b>")
InvoicelMiddleForCustomer = MakeMiddleRowForCustomer("<b>ID #</b>", "<b>Description</b>", "<b>Qty</b>", "<b>Price</b>", "<b>Totals</b>")
' BN: Loop over products in shopping cart, add one row per product to the email ...
For Each Key in dictCart
	aParameters = asGetItemParameters(Key)      
	InvoiceMiddleForStarlite = InvoiceMiddleForStarlite & _
		MakeMiddleRowForStarlite(CStr(aParameters(6)), CStr(aParameters(1)), CStr(dictCart(key)), FormatCurrency(CStr(aParameters(4))), FormatCurrency(dictCart(Key) * CSng(aParameters(4))), FormatNumber(dictCart(Key) * CSng(aParameters(8)),2))
	InvoicelMiddleForCustomer = InvoicelMiddleForCustomer & _
		MakeMiddleRowForCustomer(CStr(aParameters(6)), CStr(aParameters(1)), CStr(dictCart(key)), FormatCurrency(CStr(aParameters(4))), FormatCurrency(dictCart(Key) * CSng(aParameters(4))))
	If aParameters(7) = True Then	' Oversize.
		osize = 10.00
	Else							' Not oversize.
		osize = 0		
    End if
	' weightfactor = weightfactor + (dictCart(Key) * CDbl(aParameters(8)))
	sTotal = sTotal + (dictCart(Key) * CSng(aParameters(4)))
	' [BN, 4/29/04]: Added ...
	weightTotal = weightTotal + (dictCart(Key) * CSng(aParameters(8)))
	ExchangeRate = aParameters(9)    ' This is actually the same for each product in shopping cart, but what the heck.
	RXS.Close
Next		' BN: Next iteretion of looping over products in shopping cart.


'*******CALCULATION********************
ssTotal = sTotal
'Response.Write "<br>sTotal = " & sTotal
osizeTotal = osize * ExchangeRate
'Response.Write "<br>osize = " & osize
'Response.Write "<br>osizeTotal = " & osizeTotal
sTotal = sTotal + osizeTotal  ' 4/17/07, BN, Error, was using just osize, not osizeTotal         ' subtotal + oversize 
If Request("BigShip") = "on" Then
	' sfreight =((sTotal * 0.0375)+7.95)*3        ' new subtot * freight + insurance	
	fTotal = SandH(weightTotal, ExchangeRate) * 3
Else
	' sfreight = ((sTotal * 0.0375)+7.95)        ' new subtot * freight + insurance	
	fTotal = SandH(weightTotal, ExchangeRate)
End If
If Session("Country") = "Canada" Then
	ExtraShippingAmountForCanada = 4.0		' 4/18/07 BN: Sani wanted to charge more for shipping and handling to Canada.
	fTotal = fTotal + ExtraShippingAmountForCanada
End If 
'If weightTotal = 0 Then fTotal = 0.75   ' [BN] Not worth the trouble to distinguish between U.S. and Canada exchange rate.
If weightTotal = 0 Then fTotal = 0.0

'Response.Write "<br>fTotal = " & fTotal  ' ***

' sTotal = sTotal + sfreight
sTotal = sTotal + fTotal 
	
tax1 = sTotal * CDbl(Request("Taxx1")/100) 		' tax 1
tax2 = sTotal  * CDbl(Request("Taxx2")/100)		' tax 2
taxtoten = tax1 + tax2 

'Response.Write "<br>taxtoten = " & taxtoten   ' ***
grndTotal0 = sTotal + taxtoten    ' To create a version without a comma, so it doesn't choke the database when stored.
grndTotal = FormatNumber(sTotal + taxtoten, 2)
'Response.Write "<br>grndTotal = " & grndTotal ' ***
'Response.End   ' ***

'*******CALCULATION********************
	
' 4/18/07 BN: Continue creating body of the email to send to Starlite ...
	
InvoiceTail =	MakeTailRow("Sub Total:", FormatCurrency(ssTotal)) & _
				MakeTailRow("Over Sized Item Charge:", FormatCurrency(osizeTotal)) & _
				MakeTailRow("Shipping and Handling (within North America):", FormatCurrency(fTotal)) & _
				MakeTailRow("Tax:", FormatCurrency(taxtoten)) & _
				MakeTailRow("Grand Total:", FormatCurrency(grndTotal))
		
' [BN, 5/3/04]: Added...
If fTotal = MaxSandH  Then
	InvoiceTail = InvoiceTail & "<TR><TD COLSPAN=5 ALIGN='center' bgcolor='#DDDDDD'><font face=tahoma size=2 color=#b9000> <B>Shipping and Handling may need to be adjusted.<BR>We will notify you by email." & "</b></font></TD></TR>"
End If


' [BN, 5/2/04]: Added...
InvoiceTail2 = InvoiceTail & "<TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><font face=tahoma><B>Total Weight:</B></font></TD>" & _
	"<TD bgcolor=#BBBBBB></TD>" & _
	"<TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 ><b>" + FormatNumber(weightTotal ,2) & "</b></font></TD></TR>"

InvoiceForStarlite = "<table border=0 align=center cellpadding=4 bgcolor=white>" & _
			MakeRow("Date & Time:", 1, Now(), 4) & InvoiceHead & InvoiceMiddleForStarlite & InvoiceTail2 & _
			"</table>"
			
InvoiceForCustomerHeader =	"" & _
			"<tr><td align=center colspan=5><img src='https://www.starlite-intl.com/Images/logo50.jpg' width=50 align=absmiddle>&nbsp;&nbsp;&nbsp;&nbsp; " & _
				"<b><font face=Tahoma>Thank you very much for your order from Star Lite International, LLC</font></b></td>" & _
			"</tr>" & _
			"<tr><td align=center colspan=5><font face=Tahoma color=red>$$Disclaimer$$</font></td></tr>" & _
			"<tr><td align=center colspan=5><font face=Tahoma>" & _
					"<b>Email:</b> <a href='mailto:sales@starlite-intl.com'>sales@starlite-intl.com</a>,&nbsp;&nbsp;&nbsp; " & _
					"<b>Tel.:</b> 248-546-4489,&nbsp;&nbsp;&nbsp; " & _
			"<tr><td align=center colspan=5><font face=Tahoma>" & _
					"<b>Order Line:</b> 1-800-387-8535,&nbsp;&nbsp;&nbsp; " & _
					"<b>Fax:</b> 248-546-1462</font></td></tr>" & _
			"<tr><td align=center colspan=5><b><font face=Tahoma>Please <a href='https://www.starlite-intl.com'>Visit Us</a> Again<br><br></font></b></td></tr>" 


InvoiceForCustomer = "<table border=1 align=center><tr><td>" & _
			"<table border=0 align=left cellpadding=4 bgcolor=white width=700>" & _
			InvoiceForCustomerHeader & _
			InvoiceHead & InvoicelMiddleForCustomer & InvoiceTail & _
			"</table>" & _
			"</td></tr></table>" 
			

Session("InvoiceForCustomer") = InvoiceForCustomer			
Session("InvoiceForStarlite") = InvoiceForStarlite	


' PayPal users, that are non-US and non-Canadian.
If request.form("Payment") = "Pay Pal" AND NOT ((ReQuest("Country") = "Canada") OR (ReQuest("Country") = "USA")) Then	
	ReplacementText = 	"Please note that international shipping, insurance and handling costs are higher than shown below. " & _
					"We will notify you by email regarding the additional costs. " & _
					"If you decline the additional costs, your order will be cancelled and your money fully refunded. " & _
 					"Additional restrictions for international PayPal payment acceptance may apply."
	Session("Charge") = Cstr(grndTotal)
	PaymentMethod = "NonUSorCanadianCustomerPayPal"
	'Response.redirect "EmailSend.asp?PaymentMethod=NonUSorCanadianCustomer"

' Non-PayPal users, that are non-US and non-Canadian.
ElseIf request.form("Payment") <> "Pay Pal" AND NOT ((ReQuest("Country") = "Canada") OR (ReQuest("Country") = "USA")) Then	
	ReplacementText = 	"Please note that all our international customers must pay by <b>wire transfer</b>. " &_
					"Also, shipping, insurance and handling costs will be higher than shown below and you will be responsible for any taxes and or duties if applicable in your country. " & _
					"We ship insured after receiving confirmation from our bank that the money was transferred. " & _
					"If you agree to the above terms of sale, please let us know by return e-mail."
					
	PaymentMethod = "NonUSorCanadianCustomerNonPayPal"
	'Response.redirect "EmailSend.asp?PaymentMethod=NonUSorCanadianCustomer"
	
ElseIf (Request("Payment") = "Visa") OR (Request("Payment") = "Master Card") OR (Request("Payment") = "Discover") OR (Request("Payment") = "American Express") Then
	ReplacementText = 	"Your order will ship once your credit card payment is verified and approved."
	' US/Canadian customer using a credit card.
	Session("Charge") = Cstr(grndTotal)
	PaymentMethod = "CreditCard" 
	'response.redirect "cashout.asp"		' Webpage on the way to LinkPoint webpage.
	'Response.redirect "EmailSend.asp?PaymentMethod=CreditCard"
	%>
	<!-- <form method="POST" action="https://www.linkpointcentral.com/lpc/servlet/lppay" name="LinkPoint">
	<input type="hidden" name="mode" value="fullpay"> 
	<input type="hidden" name="chargetotal" value="<%=Session("Charge")%>">
	<input type="hidden" name="storename" value="330566"> 
	<input type="submit" value="Continue to Secure Payment Form">
	</form> -->
	
	<%	 
ElseIf request.form("Payment") = "Pay Pal" Then			' US/Canadian customer using PayPal.
	ReplacementText = 	"Your order will ship once your PayPal payment is verified and approved."
	Session("Charge") = Cstr(grndTotal)
	PaymentMethod = "PayPal"
	'Response.redirect "EmailSend.asp?PaymentMethod=PayPal"
	%>
	<!-- <form action="https://www.paypal.com/cgi-bin/webscr" method="post" name="paypal">
		<input type="hidden" name="cmd" value="_xclick">
		<input type="hidden" name="business" value="starlite@starlite-intl.com">
		<input type="hidden" name="item_name" value="OrderFor_<%=request.form("Name")%>">
		<input type="hidden" name="currency_code" value="<% If Session("Country") = "Canada" Then %>CAD
														 <% Else %>USD
														 <% End If %>">
		<input type="hidden" name="amount" value="<%=FormatNumber(grndTotal,2)%>">
		<input type="hidden" name="custom" value="<%=request.form("Email")%>">
		<input type="hidden" name="return" value="https://www.starlite-intl.com/scart/ThankYou.asp">
		<input type="hidden" name="cancel_return" name="https://www.starlite-intl.com/scart/paypalcancel.asp?name=<%=request.form("Name")%>">
	</form>
	<script language="javascript">
	document.paypal.submit();
	</script> -->
	
	<%
ElseIf request.form("Payment") = "Check" Then			' US/Canadian customer using a check.
	ReplacementText = 	"We are looking forward to receiving your Check.<br>" & _
						"Your order will ship <em>after</em> your Check has cleared."
	PaymentMethod = "Check"
	'Response.redirect "EmailSend.asp?PaymentMethod=Check"  
	
ElseIf request.form("Payment") = "Money Order" Then		' US/Canadian customer using a money order.
	ReplacementText = 	"We are looking forward to receiving your Money Order.<br>" & _
						"Your order will ship <em>after</em> your Money Order has cleared."
	PaymentMethod = "MoneyOrder"
	'Response.redirect "EmailSend.asp?PaymentMethod=MoneyOrder"  
End If


Session("InvoiceForCustomer") = Replace(Session("InvoiceForCustomer"), "$$Disclaimer$$", ReplacementText)


' Display emails on this webpage, for debugging purposes.
If True Then
	Response.Write "<table align=left border=0 width=800>" 
	Response.Write "<tr><td><b><font face=Tahoma color=red>Invoice for Customer ...</font></b></td></tr>"
	Response.Write "<tr><td>"
	Response.Write Session("InvoiceForCustomer")
	Response.Write "</td></tr>"
	Response.Write "<tr><td>&nbsp;</td></tr>"
	Response.Write "<tr><td><b><font face=Tahoma color=red>Invoice For Starlite ...</font></b></td></tr>"
	Response.Write "<tr><td>"
	Response.Write Session("InvoiceForStarlite")   ' InvoiceForStarlite
	Response.Write "</td></tr>"
	Response.Write "</table>"
End If   ' False


'*********************************************************************************
' [6/18/07, BN] Save order data to Orders table in database. 

Response.Write "<br>grndTotal0 = " & grndTotal0
'Response.End

' I have not been able to save the Date. Don't know why I am having this problem.
If TRUE Then
OrderSQL = "INSERT INTO Orders (OrderDate,OrderNumber,FName,LName,Email,OptInToEmailings,WeEmail,Phone," & _
	"Address,City,State,BigShipment,ZIP,Country,PaymentMethod,Tax1,Tax2,SubTotal,OversizeCharge,SandH,Tax,GrandTotal,TotalWeight) "
'Response.Write "<br><br>OrderSQL = " & OrderSQL
OrderSQL = OrderSQL & " VALUES(#" & OrderDate & "# ,'" & OrderNum & "' ,'" & _
	ReQuest("FName") & "' ,'" & LName & "' ,'" & ReQuest("Email") & "' ," & OptedIn2 & " ," & WeEmail & " ,'" & _
	ReQuest("Phone") & "' ,'" & Address & "' ,'" & ReQuest("City") & "' ,'" & _
	ReQuest("StateProv") & "' ," & BigShipment & " ,'" & ReQuest("Postal") & "' ,'" & ReQuest("Country") & "', '" & _
	ReQuest("Payment") & "', " & ReQuest("Taxx1") & ", " & ReQuest("Taxx2") & ", " & _
	ssTotal & ", " & osizeTotal & ", " & fTotal & ", " & _
	taxtoten & ", " & grndTotal0 & ", " & weightTotal & " );"
'Response.Write "<br><br>OrderSQL = " & OrderSQL
End If

'OrderSQL = "INSERT INTO Orders (FName,LName,Email) "
'OrderSQL = OrderSQL & " VALUES('" & ReQuest("FName") & "' ,'" & ReQuest("LName") & "' ,'" & ReQuest("Email") & "' );"
	
Response.Write "<br><br>OrderSQL = " & OrderSQL
'Response.End
'On Error Resume Next
Set Conn = Server.CreateObject("ADODB.Connection")
'Conn.Open Session("ConnectionString2")
Conn.Open Session("ConnectionString") 
Set obj = Conn.Execute(OrderSQL)
If Err.number <> 0 then
	Response.Write "<br><br><br><br><br><br><br><br><center>There was an error recording your order in our database. Please notify Starlite at sales@starlite-intl.com.</center>"
	Response.End
end if
'Response.End

Conn.Close

'*********************************************************************************

'Response.End

Response.Redirect "EmailSend.asp?PaymentMethod=" & PaymentMethod 
%>



</body>

</html>







