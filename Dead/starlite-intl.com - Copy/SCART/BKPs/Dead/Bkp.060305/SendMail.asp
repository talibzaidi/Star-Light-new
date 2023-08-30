
<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
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
<% if ReQuest("Email") = "email@domain.com" then
   response.Redirect "formcorrection.asp?" & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("Address") = "" then
   response.Redirect "formcorrection.asp?" & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("Phone") = "" then
   response.Redirect "formcorrection.asp?" & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("StateProv") = "" then
   response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("Country") = "SELECT COUNTRY" then
   response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("Name") = "" then
   response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("Postal") = "" then
   response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("City") = "" then
   response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("Payment") = "Choose%20an%20Option" then
   response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")
   elseif ReQuest("Payment") = "Choose an Option" then
   response.Redirect "formcorrection.asp?"  & "Address=" & ReQuest("Address") & "&email=" & ReQuest("Email") & "&phone=" & ReQuest("Phone") & "&StateProv=" & ReQuest("StateProv") & "&Country=" & ReQuest("Country") & "&Taxx1=" & ReQuest("Taxx1") & "&Taxx2=" & ReQuest("Taxx2") & "&Payment=" & ReQuest("Payment") & "&Name=" & ReQuest("Name") & "&Postal=" & ReQuest("Postal") & "&City=" & ReQuest("City")

   end if
%>




<%
Dim myMail
Set myMail = Server.CreateObject("CDONTS.NewMail")
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



<body background="../Images/background.jpg" bgcolor="#FFFFFF" link="#000000" vlink="#000000" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<!--#include file="RATES.INC"-->

<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td background="../Images/topback.gif"><div align="left"><table border="0" cellpadding="0" cellspacing="0" width="575">
             <tr>
                <td> <!--#include file="NAV.INC"--><img src="../Images/toptitle.jpg" width="411" height="29"><br>
                </td>
            </tr>
            <tr>
                <td width="575"><img src="../Images/emptynav.jpg" width="164" height="14"><img src="../Images/bottitle.JPG" width="411" height="14"></td>
            </tr>
            <tr>
                <td><img src="../Images/leftbar.gif" width="176" height="23"><img src="../Images/blanka1.gif" WIDTH="399" HEIGHT="23"></td>
            </tr>
        </table>
        </div></td>
        <td width="100%" background="../Images/topback.gif">&nbsp;</td>
    </tr>
    <tr>
	<td width>&nbsp;
<table border="0">
<tr>
<td width="170">&nbsp;</td>
<td width="380">
<center>
<center> 
<%


Address=Request("Address") 
City=Request("City")
Payment=Request("Payment") 
expfrom=Request("expfrom") 
expto=Request("expto") 
Cardholder=Request("Cardholder")




If Err.number <> 0 then
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
if request("Payment") = "C.O.D." then 
	shipcrg = 8.75
end if
Dim Key
Dim aParameters ' as Variant (Array)
Dim sTotal, sShipping
If Request("OptIntoEmailing") = "on" Then
	OptedIn = "No"
Else
	OptedIn = "Yes"
End If
shpcrg = 0
	strBODY = "<html><head></head><body>" 
	strBODY = strBODY & "<TABLE Border=0 CellPadding=3 CellSpacing=2 width=350><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Name:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>"
	strBODY = strBODY & Request("Name")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Email:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Email")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Opt In to Emailings:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + OptedIn
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Phone:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Phone")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Address:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Address")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>City:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("City")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>State/Province:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("StateProv")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Postal Code:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Postal")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Country:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Country")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Payment Method:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Payment")
	strBODY = strBODY & "</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Tax 1:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>" + Request("Taxx1")
	strBODY = strBODY & "%</b></font></TD></tr><tr><TD colspan=1 bgcolor=#BBBBBB><font face=tahoma size=2><b>Tax 2:</b></font></TD><TD colspan=4 bgcolor=#DDDDDD><font face=tahoma size=2><b>"+ Request("Taxx2")
	strBODY = strBODY & "%</b></font><br><br></TD></tr><tr><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>ID #</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Description</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Qty.</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Price</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Totals</b></font></TD><TD bgcolor=#BBBBBB><font face=tahoma size=2><b>Weight</b></font></TD></tr>"
	'44*************************
	
	sTotal = 0
	weightTotal = 0	' [BN, 4/29/04] Added this.

	For Each Key in dictCart
		aParameters = asGetItemParameters(Key)

        '48*************************
	strBODY = strBODY & "<TR><TD ALIGN=Center bgcolor=#DDDDDD><font face=tahoma size=2>" + CStr(aParameters(6))
	strBODY = strBODY & "</font></TD><TD ALIGN=Left bgcolor=#DDDDDD><font face=tahoma size=2>" + CStr(aParameters(1))
	strBODY = strBODY & "</font></TD><TD ALIGN=Center bgcolor=#DDDDDD><font face=tahoma size=2>" + CStr(dictCart(key))
	strBODY = strBODY & "</font></TD><TD ALIGN=Right bgcolor=#DDDDDD><font face=tahoma size=2>" + CStr(aParameters(4))
	strBODY = strBODY & "</font></TD><TD ALIGN=Right bgcolor=#DDDDDD><font face=tahoma size=2>$" + (FormatNumber(dictCart(Key) * CSng(aParameters(4)),2))
	strBODY = strBODY & "</font></TD><TD ALIGN=Right bgcolor=#DDDDDD><font face=tahoma size=2>" + (FormatNumber(dictCart(Key) * CSng(aParameters(8)),2))
	'*************************
	if aParameters(7) = true then
		      osize = 8.75
        end if
        
        ' weightfactor = weightfactor + (dictCart(Key) * CDbl(aParameters(8)))
	sTotal = sTotal + (dictCart(Key) * CSng(aParameters(4)))

       	 ' [BN, 4/29/04]: Added ...
        	weightTotal = weightTotal + (dictCart(Key) * CSng(aParameters(8)))
	ExchangeRate = aParameters(9)    ' This is actually the same for each product in shopping cart, but what the heck.


	RXS.Close
	Next

	'*******CALCULATION********************
	'*******CALCULATION********************
	ssTotal = sTotal
	sTotal = sTotal + osize         ' subtotal + oversize 
	if Request("BigShip") = "on" then
	' sfreight =((sTotal * 0.0375)+7.95)*3        ' new subtot * freight + insurance	
	fTotal = SandH(weightTotal, ExchangeRate) * 3
	If weightTotal = 0 Then fTotal = 0.89   ' [BN] Not worth the trouble to distinguish between U.S. and Canada excgange rate.
	else
        
	' sfreight = ((sTotal * 0.0375)+7.95)        ' new subtot * freight + insurance	
	fTotal = SandH(weightTotal, ExchangeRate)
    If weightTotal = 0 Then fTotal = 0.89   ' [BN] Not worth the trouble to distinguish between U.S. and Canada excgange rate.
	end if

	' sTotal = sTotal + sfreight
	sTotal = sTotal + fTotal 
	
	tax1 = sTotal * CDbl(Request("Taxx1")/100) 		' tax 1
	tax2 = sTotal  * CDbl(Request("Taxx2")/100)		' tax 2
	taxtoten = tax1 + tax2 

        grndTotal = sTotal + taxtoten

	'*******CALCULATION********************
	'*******CALCULATION********************
	'*************************
	strBODY = strBODY & "</TD></TR><TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Sub Total:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>$" + FormatNumber(ssTotal,2) + "</b></font></TD></TR>"

	' strBODY = strBODY & "</b></font></TD></TR><TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Over Sized Item Charge:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>$" + FormatNumber(osize,2) 
	osizeTotal = osize * ExchangeRate
	strBODY = strBODY & "</b></font></TD></TR><TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Over Sized Item Charge:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>$" + FormatNumber( osizeTotal , 2) 

	' strBODY = strBODY & "<TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Shipping and Handling:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>" & FormatCurrency(sfreight)
	strBODY = strBODY & "<TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Shipping and Handling:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>" & FormatCurrency(fTotal)

	strBODY = strBODY & "<TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Tax:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>" + FormatCurrency(taxtoten)
	
	strBODY = strBODY & "</b></font></TD></TR><TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Grand Total:</B></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 color=#b9000><b>$" + FormatNumber(grndTotal,2) 
	
	' [BN, 5/3/04]: Added...
 If fTotal = MaxSandH  Then
strBODY = strBODY & "</b></font></TD></TR><TR><TD COLSPAN=5 ALIGN='center' bgcolor='#DDDDDD'><font face=tahoma size=2 color=#b9000> <B>Shipping and Handling may need to be adjusted.<BR>We will notify you by email."
End If


	' [BN, 5/2/04]: Added...
	strBODY = strBODY & "</b></font></TD></TR><TR><TD COLSPAN=4 ALIGN=Right bgcolor=#BBBBBB><B>Total Weight:</B></TD><TD bgcolor=#BBBBBB></TD><TD ALIGN=Right bgcolor=#BBBBBB><font face=tahoma size=2 ><b>" + FormatNumber( weightTotal ,2)    

	strBODY = strBODY & "</b></font></TD></TR></TABLE></body></html>" 


Function asGetItemParameters(iItemID)
	Dim bParameters 

if Session("Country") = "USA" then       
        	RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID +   "' ", "DSN=STAREC1" , 1, 4
	bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), 1 )
else
	RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID + "' ", "DSN=STAREC1" , 1, 4
	bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("Pname") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Freight")*RXS("Exch")), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")*(1/(1-(RXS("GPM"))))),RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), RXS("Exch")  )
end if

' Return array containing product info.
asGetItemParameters = bParameters

End Function


		myMail.From			= Request("Email")     ' "starlite@starlite-intl.com"  
		myMail.To			= "starlite@starlite-intl.com" 
		myMail.Subject		= "Electronic Order " 
		myMail.BodyFormat	= 0 
		myMail.MailFormat	= 0 
		myMail.Body			= StrBody
		myMail.Send 


if (Request("Payment") = "Visa") OR (Request("Payment") = "Master Card") OR (Request("Payment") = "Discover") OR (Request("Payment") = "American Express") Then
	if (ReQuest("Country") = "Canada") OR (ReQuest("Country") = "United States") then
		Session("Charge") = Cstr(grndTotal)
		response.redirect "cashout.asp"
	else
		response.redirect "thankyou.asp"
	end if 
else
	if request.form("Payment") = "Pay Pal" then
	%>
		<form action="https://www.paypal.com/cgi-bin/webscr" method="post" name="paypal">
		<input type="hidden" name="cmd" value="_xclick">
		<input type="hidden" name="business" value="starlite@starlite-intl.com">
		<input type="hidden" name="item_name" value="OrderFor_<%=request.form("Name")%>">
		<input type="hidden" name="currency_code" value="<%if Session("Country") = "Canada" then
		%>CAD
		<%else%>
		USD<%
		end if%>">
		<input type="hidden" name="amount" value="<%=FormatNumber(grndTotal,2)%>">
		<input type="hidden" name="custom" value="<%=request.form("Email")%>">
		<input type="hidden" name="return" value="https://www.starlite-intl.com/scart/thankyou.asp">
		<input type="hidden" name="cancel_return" name="https://www.starlite-intl.com/scart/paypalcancel.asp?name=<%=request.form("Name")%>">
		</form>
		<script language="javascript">
		document.paypal.submit();
		</script>
	<%else
		response.redirect "thankyou.asp"
	end if
end if


%>
<%=StrBody%>

<br><a href="thankyou.asp"><font face="tahoma" size="5" color="#000000"><b>Click here to proceed.</b></font></a>
</center>
</center>
</td>
</tr>
</table> </td>
                <td valign="top">


              
        </td>
     
    </tr>
    <tr>
        <td><img src="../Images/bottompage.GIF" width="575" height="52"></td>
        <td background="../Images/botback.gif">&nbsp;</td>
    </tr>
</table>
</body>
</html>







