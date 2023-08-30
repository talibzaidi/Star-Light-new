<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
		
		
<!--#INCLUDE FILE="PaymentAuthorizeNet/simlib.asp"-->
<!--#INCLUDE FILE="PaymentAuthorizeNet/simdata.asp" -->

		
<html>

<head>
    <TITLE> USAePay Payment Form via POST </TITLE>
</head>


<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">


<% 
'****************** 
' 7/5/17: The following is based on sample file PaymentUSAePay/USAePayTest/sample.asp
' (and PaymentAuthorizeNet/sim.asp)
'****************** 
%>

<!-- 
<h3>Final Order</h3>

Description: CC AUTH ONLY<br>
Total Amount : 19.99<br>
<br>
-->

<!--<FORM action="https://test.authorize.net/gateway/transact.dll">-->
<!-- Uncomment the line ABOVE for test accounts OR the line BELOW for LIVE accounts-->
<!-- <FORM action="https://secure.authorize.net/gateway/transact.dll"> -->
<%
Dim sequence
Dim amount
Dim ret


loginid		= "25aC8ErB"
txnkey      = "3e7T4u866T4MkZLn"    ' transactionKey	8/5/15.

' *** IF YOU WANT TO PASS CURRENCY CODE uncomment the next 2 lines **
' Dim currencycode
' Assign the transaction currency (from your shopping cart) to currencycode variable

' Trim $ dollar sign if it exists
' amount = Request("x_amount")

' Just a test so we use a hard-coded amount of 19.99
' NOTE: You must make absolutely sure that you are
' passing a number to the "amount" variable; otherwise,
' the fingerprint calculation will not work

' So, if you are gathering the amount value from a field
' or any other user input, make sure you collect it as or
' convert it to a number, using either VBScript, JScript or
' JavaScript.


'amount = 1.99
amount = Session("Charge")


If Mid(amount, 1,1) = "$" Then
	amount = Mid(amount,2)
End If

' Seed random number for more security and more randomness
Randomize
sequence = Int(1000 * Rnd)

'*** for testing only:   *********************************
'sequence = 931


' Now add the SIM related data, such as the fingerprint,
' to the HTML form.
' Response.Write ("<input type='text' name='x_description' value='" & Request("x_description") & "' style='width:300;'>" & vbCrLf)

'Response.Write ("<input type='text' name='x_description' value='CC AUTH ONLY' style='width:300;'>" & vbCrLf)
'Response.Write("<br>")

' Again, make sure all required values are properly declared
' in their respective places

'ret = InsertFP (loginid, txnkey, amount, sequence)


' *** IF YOU ARE PASSING CURRENCY CODE uncomment and use the following instead of the InsertFP function above  ***
' ret = InsertFP (loginid, txnkey, amount, sequence, currencycode)

'Response.Write ("<input type='text' name='x_login' value='" & loginid & "' style='width:300;'>" & vbCrLf)
'Response.Write("<br>")
'Response.Write ("<input type='text' name='x_amount' value='" & amount & "' style='width:300;'>" & vbCrLf)
'Response.Write("<br>")

' *** IF YOU ARE PASSING CURRENCY CODE uncomment the line below *****
' Response.Write ("<input type=""text"" name=""x_currency_code"" value=""" & currencycode & """>" & vbCrLf)

%>

<!--
<INPUT type="text" name="x_show_form" value="PAYMENT_FORM" style="width:300;"><br>
<INPUT type="text" name="x_color_background" value="#eeeeee" style="width:300;"><br>
<br>


<INPUT type="submit" value="Authorize.Net  ::  Accept Order">
</form>
<br>
<br>
<br>
-->


<br><br><br><br><br><br>
<table border="0" align=center>
	<tr>
		<td>
		<% ' My version (not Authorize.net's) of the form %>
		<form action="https://secure.authorize.net/gateway/transact.dll" method="POST" name="AuthorizeNet">
			<input type="text" name='x_description' value='<%=Date()%>' style="width:300;">x_description<br> 
			<% ret = InsertFP (loginid, txnkey, amount, sequence) %>
			<input type="text" name='x_login' value='<%=loginid%>' style="width:300;">x_login<br>
			<input type="text" name='x_amount' value='<%=amount%>' style="width:300;">x_amount<br>
			<input type="text" name="x_type" value='AUTH_ONLY' style="width:300;">x_type<br>	
			
			<input type="text" name="x_invoice_num" value=<%=Session("OrderNum")%> style="width:300;">x_invoice_num<br>	
			<input type="text" name="x_first_name" value='<%=Session("FName")%>' style="width:300;">x_first_name<br>	
			<input type="text" name="x_last_name" value='<%=Session("LName")%>' style="width:300;">x_last_name<br>	
			<input type="text" name="x_email" value='<%=Session("Email")%>' style="width:300;">x_email<br>	
			<input type="text" name="x_phone" value='<%=Session("Phone")%>' style="width:300;">x_phone<br>	
			<input type="text" name="x_address" value='<%=Session("Address")%>' style="width:300;">x_address<br>	
			<input type="text" name="x_city" value='<%=Session("City")%>' style="width:300;">x_city<br>
			<input type="text" name="x_state" value='<%=Session("StateProv")%>' style="width:300;">x_state<br>
			<input type="text" name="x_zip" value='<%=Session("Postal")%>' style="width:300;">x_zip<br>
			<input type="text" name="x_country" value='<%=Session("Country")%>' style="width:300;">x_country<br>	
								
			<input type="text" name="x_show_form" value="PAYMENT_FORM" style="width:300;">x_show_form<br>
			<input type="text" name="x_color_background" value="#eeeeee" style="width:300;">x_color_background<br>	
			<!-- <input type="submit" value="Authorize.Net  ::  Accept Order"> -->
		</form> 
		
		<script language="javascript">
		<!-- document.AuthorizeNet.submit(); -->
		</script>

		</td>
	</tr>
</table> 


<% ' ************************************************************************* %>

<%
' The parameters for the USAePay payment form etc. can be configured here.
' To post to the USAePay payment form in our SANDBOX account:
'URL         = "https://sandbox.usaepay.com/interface/epayform/"
' To post to the USAePay payment form in our PRODUCTION account:
URL         = "https://www.usaepay.com/interface/epayform/"

Description	= "Sample Transaction"
'UMkey       = "_GzuIx1hT7dOT4DGlVc45ac9iHLdpgR5"    ' My source key for USAePay SANDBOX useage. 
UMkey       = "v0tm18XWN8NKP222CtS55tu7467PFGHx"    ' My source key for USAePay PRODUCTION useage.
'UMcommand   = "sale"                               ' Sani does not want us to use "sale" in Production account, but only "authonly".
UMcommand   = "authonly"                            ' Authorize card only; only allowed in Production account.
UMamount    = "0.99"                                ' transaction amount. 
UMinvoice   = "Test-0004"                           ' invoice/order number.
    

' UMcustreceipt tells whether to send receipt to customer. But this line didn't work. 
' On 6/26/17 Jenny Silva (of USAePay Integration Support) hardcoded this line into the form at 
' her server, and it worked: 
' I was then able to receive receipt to my test customer email address "staff@intelligineering.com". -->
UMcustreceipt = "yes"                              

If FALSE Then   ' [BN, 1/30/18] Need to False this out in general else UMbillcompany will be set to "Intelligineering"
                ' for real orders if customer leaves that field blank, so that Session("Company") == "" below.

' Customer's Billing Address fields:
' See http://help.usaepay.com/merchant/epaymentform
UMbillfname     = "Bernard"
UMbilllname     = "Nadel"
UMbillcompany   = "Intelligineering"
UMbillstreet    = "14060 Balfour"
UMbillstreet2   = ""
UMbillcity      = "Oak Park"
UMbillstate     = "MI"
UMbillzip       = "48237"
UMbillcountry   = "USA"
UMbillphone     = "313-307-1874"
UMemail         = "staff@intelligineering.com"   ' Test customer email(s). 
UMwebsite       = "www.intelligineering.com"
%>

<!-- METADATA  (commenting out ASP a la: https://stackoverflow.com/questions/4431170/server-side-comments-whats-the-equivalent-of-in-asp-classic)
(Shipping info is to be input by customer on USAePay server paymemt form, at
URL         = "https://sandbox.usaepay.com/interface/epayform/"
or
URL         = "https://www.usaepay.com/interface/epayform/"
where (in Production account only?) the customer can check a box to set shipping info to equal billing info if appropriate. 
-->
<%
' Customer's Shipping Address fields:
' See http://help.usaepay.com/merchant/epaymentform
UMshipfname     = ""    ' "Bernard"
UMshiplname     = ""    ' "Nadel"
UMshipcompany   = ""    ' "Intelligineering"
UMshipstreet    = ""    ' "14060 Balfour"
UMshipstreet2   = ""
UMshipcity      = ""    ' "Oak Park"
UMshipstate     = ""    ' "MI"
UMshipzip       = ""    ' "48237"
UMshipcountry   = ""    ' "USA"
UMshipphone     = ""    ' "313-307-1874"

End If   ' TRUE / FALSE
%>

<%
If Session("Charge") <> "" Then
	UMamount = Session("Charge")
End If
If Session("OrderNum") <> "" Then
	UMinvoice = Session("OrderNum")
End If
If Session("FName") <> "" Then
	UMbillfname = Session("FName")
End If
If Session("LName") <> "" Then
	UMbilllname = Session("LName")
End If
If Session("Company") <> "" Then
	UMbillcompany = Session("Company")
End If
If Session("Email") <> "" Then
	UMemail = Session("Email")    ' Customer email.  
End If
If Session("Phone") <> "" Then
	UMbillphone = Session("Phone")      
End If
If Session("Address") <> "" Then
	UMbillstreet = Session("Address")      
End If
If Session("City") <> "" Then
	UMbillcity = Session("City")     
End If
If Session("StateProv") <> "" Then
	UMbillstate = Session("StateProv")     
End If
If Session("Postal") <> "" Then
	UMbillzip = Session("Postal")     
End If
If Session("Country") <> "" Then
	UMbillcountry = Session("Country")     
End If

SubmitbtnLabel  = "Continue to Payment Form"
%>


<%
' Print the variables and their value to the screen.
Response.Write("Description: " & Description & " <br />")
Response.Write("Amount: " & UMamount & " <br />")
Response.Write("Payment Method: " & Session("PaymentMeth") & " <br />")
Response.Write("Company: " & UMbillcompany & " <br />")
'Response.End
%>

<!--
Create the HTML form containing necessary POST values
Additional fields can be added here as outlined at
http://help.usaepay.com/merchant/epaymentform
-->

<%
Response.Write("<FORM method='POST' name='USAePayForm' action='" & URL & "' >")
' Additional fields can be added here as outlined at:
' http://help.usaepay.com/developer/transactionapi?s[]=transaction&s[]=api
Response.Write("	<INPUT type='hidden' name='UMkey' value='" & UMkey & "' />")
Response.Write("	<INPUT type='hidden' name='UMcommand' value='" & UMcommand & "' />")
Response.Write("	<INPUT type='hidden' name='UMamount' value='" & UMamount & "' />")
Response.Write("	<INPUT type='hidden' name='UMinvoice' value='" & UMinvoice & "' />")


' "Credit Card" fields:
' Leave the "Credit Card" fields to be filled out by customer on the USAePAY payment form.


' "Billing Information" fields:
' See http://help.usaepay.com/merchant/epaymentform
Response.Write("	<INPUT type='hidden' name='UMbillfname' value='" & UMbillfname & "' />")
Response.Write("	<INPUT type='hidden' name='UMbilllname' value='" & UMbilllname & "' />")
Response.Write("	<INPUT type='hidden' name='UMbillcompany' value='" & UMbillcompany & "' />")
Response.Write("	<INPUT type='hidden' name='UMbillstreet' value='" & UMbillstreet & "' />")
Response.Write("	<INPUT type='hidden' name='UMbillstreet2' value='" & UMbillstreet2 & "' />")
Response.Write("	<INPUT type='hidden' name='UMbillcity' value='" & UMbillcity & "' />")
Response.Write("	<INPUT type='hidden' name='UMbillstate' value='" & UMbillstate & "' />")
Response.Write("	<INPUT type='hidden' name='UMbillzip' value='" & UMbillzip & "' />")
Response.Write("	<INPUT type='hidden' name='UMbillcountry' value='" & UMbillcountry & "' />")
Response.Write("	<INPUT type='hidden' name='UMbillphone' value='" & UMbillphone & "' />")
Response.Write("	<INPUT type='hidden' name='UMemail' value='" & UMemail & "' />")  ' customer's email. 
Response.Write("	<INPUT type='hidden' name='UMwebsite' value='" & UMwebsite & "' />")


' "Shipping Information" fields:
' See http://help.usaepay.com/merchant/epaymentform
Response.Write("	<INPUT type='hidden' name='UMshipfname' value='" & UMshipfname & "' />")
Response.Write("	<INPUT type='hidden' name='UMshiplname' value='" & UMshiplname & "' />")
Response.Write("	<INPUT type='hidden' name='UMshipcompany' value='" & UMshipcompany & "' />")
Response.Write("	<INPUT type='hidden' name='UMshipstreet' value='" & UMshipstreet & "' />")
Response.Write("	<INPUT type='hidden' name='UMshipstreet2' value='" & UMshipstreet2 & "' />")
Response.Write("	<INPUT type='hidden' name='UMshipcity' value='" & UMshipcity & "' />")
Response.Write("	<INPUT type='hidden' name='UMshipstate' value='" & UMshipstate & "' />")
Response.Write("	<INPUT type='hidden' name='UMshipzip' value='" & UMshipzip & "' />")
Response.Write("	<INPUT type='hidden' name='UMshipcountry' value='" & UMshipcountry & "' />")
Response.Write("	<INPUT type='hidden' name='UMshipphone' value='" & UMshipphone & "' />")     

'Response.Write("	<input type='submit' value='" & SubmitbtnLabel & "' />")
Response.Write("</FORM>")

' document.USAePayForm.submit() below is to cause the form to be submitted automatically, 
' instead of requiring customer to manually press its submit button.
%>

<script language="javascript">
    document.USAePayForm.submit();
</script>


</body>

</html>