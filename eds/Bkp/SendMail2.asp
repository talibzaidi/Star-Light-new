<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<%
Dim myMail
Set myMail = Server.CreateObject("CDONTS.NewMail")
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
if request("Payment") = "C.O.D." then 
	shipcrg = 5
end if
Dim Key
Dim aParameters ' as Variant (Array)
Dim sTotal, sShipping
shpcrg = 0
	strBODY = "<html><head></head><body>" 
	strBODY = strBODY & "<TABLE Border=0 CellPadding=3 CellSpacing=2 width=450><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Name:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>"
	strBODY = strBODY & Request("Name")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Email:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Email")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Address:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Address")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>City:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("City")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>State/Province:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("StateProv")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Country:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Country")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Payment Method:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Payment")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Credit Card Number:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Cardnum")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Expires:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("expfrom") + "/" + Request("expto")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Cardholders Name:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Cardholder")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Shipper:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Shipper")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Tax 1:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Taxx1")
	strBODY = strBODY & "%</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Tax 2:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>"+ Request("Taxx2")
	strBODY = strBODY & "%</b></font><br><br></TD></tr><tr><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>ID #</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Description</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Qty.</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Price</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Totals</b></font></TD></tr>"
	'44*************************
	sTotal = 0
	For Each Key in dictCart
		aParameters = asGetItemParameters(Key)
        '48*************************
	strBODY = strBODY & "<TR><TD ALIGN=Center bgcolor=#DDDDDD>" + CStr(aParameters(6))
	strBODY = strBODY & "</TD><TD ALIGN=Left bgcolor=#DDDDDD>" + CStr(aParameters(1))
	strBODY = strBODY & "</TD><TD ALIGN=Center bgcolor=#DDDDDD>" + CStr(dictCart(key))
	strBODY = strBODY & "</TD><TD ALIGN=Right bgcolor=#DDDDDD>" + CStr(aParameters(4))
	strBODY = strBODY & "</TD><TD ALIGN=Right bgcolor=#DDDDDD>$" + (FormatNumber(dictCart(Key) * CSng(aParameters(4)),2))
	'*************************
	if aParameters(7) = true then
		      osize = 5
		      end if
	sTotal = sTotal + (dictCart(Key) * CSng(aParameters(4)))
	RXS.Close
	Next
	'*************************
	strBODY = strBODY & "</TD></TR><TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Sub Total:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>" + FormatNumber(sTotal,2) + "</b></font></TD></TR>"
	strBODY = strBODY & "<TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Tax:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>" + FormatNumber(CDbl(Request("Taxx1"))+CDbl(Request("Taxx2")),1)
	strBODY = strBODY & "<TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>C.O.D. Charge:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>"  + formatcurrency(shipcrg)
	strBODY = strBODY & "<TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Freight & Insurance:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>4" 
	strBODY = strBODY & "%</b></font></TD></TR><TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Over Sized Item Charge:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>$" + FormatNumber(osize,2) 
	'*******CALCULATION********************
	'*******CALCULATION********************

	sTotal = sTotal + osize + shipcrg	        ' subtotal + oversize + cod
	sTotal = sTotal + (sTotal * 0.04)		' new subtot * freight + insurance	
	tax1 = sTotal * CDbl(Request("Taxx1")) 		' tax 1
	tax2 = sTotal  * CDbl(Request("Taxx2"))		' tax 2
	grndTotal = sTotal + tax1 + tax2

	'*******CALCULATION********************
	'*******CALCULATION********************
	
	strBODY = strBODY & "</b></font></TD></TR><TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Total:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>$" + FormatNumber(grndTotal,2) 
	strBODY = strBODY & "</b></font></TD></TR></TABLE></body></html>" 

Function asGetItemParameters(iItemID)
	Dim bParameters 
        if Session("Country") = "USA" then       
        RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates  WHERE  ITEMID = " + iItemID +   " ", "DSN=STAREC1" , 1, 4
	bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")/RXS("GPM")), RXS("PID"),RXS("ITEMID"), RXS("OverSize") )
	else
	 RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates  WHERE  ITEMID = " + iItemID +   " ", "DSN=STAREC1" , 1, 4
	bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("Pname") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Freight")*RXS("Exch")), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")/RXS("GPM")),RXS("PID"),RXS("ITEMID"), RXS("OverSize") )
	end if
' Return array containing product info.
asGetItemParameters = bParameters

End Function
%>
<%
myMail.From = "Starlite@ECommerce.Com" 
myMail.To = "sanction@anyperson.com" 
myMail.Subject = "Starlite Order" 
myMail.BodyFormat = 0 
myMail.MailFormat = 0 
myMail.Body = strBODY
myMail.Send 
response.redirect "thankyou.asp"
%> 