<!--
6/18/17: This file 
(URL: https://www.starlite-intl.com/SCART/PaymentUSAePay/USAePayTest/sample.asp)
is based on the analogous file: PaymentAuthorizeNet > AuthorizenetTest > sample.asp
(URL: https://www.starlite-intl.com/SCART/PaymentAuthorizeNet/AuthorizenetTest/sample.asp)
     
This file is designed to connect to USAePay.net using the method 
under heading "Payment Form via POST" in http://help.usaepay.com/merchant/epaymentform.
-->

<!--
This sample code is designed to connect to Authorize.net using the SIM method.
For API documentation or additional sample code, please visit:
http://developer.authorize.net

Most of this page can be modified using any standard html. The parts of the
page that cannot be modified are noted in the comments.  This file can be
renamed as long as the file extension remains .asp
-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" 
  "http://www.w3.org/TR/html4/loose.dtd">
<HTML lang='en'>
<HEAD>
	<TITLE> USAePay Payment Form via POST </TITLE>
</HEAD>
<BODY>

<!-- This section generates the "Submit Payment" button using ASP           -->
<!--# INCLUDE FILE="simlib.asp"-->

<!-- METADATA  (commenting out ASP a la: https://stackoverflow.com/questions/4431170/server-side-comments-whats-the-equivalent-of-in-asp-classic)
<%
' the parameters for the payment can be configured here
' the API Login ID and Transaction Key must be replaced with valid values
Dim loginID, transactionKey, amount, description, label, testMode, url
loginID			= "25aC8ErB"
transactionKey	= "3e7T4u866T4MkZLn"
amount			= "19.99"
description		= "Sample Transaction"
' The is the label on the 'submit' button
label			= "Submit Payment"
testMode		= "false"
' By default, this sample code is designed to post to our test server for
' developer accounts: https://test.authorize.net/gateway/transact.dll
' for real accounts (even in test mode), please make sure that you are
' posting to: https://secure.authorize.net/gateway/transact.dll
' url				= "https://test.authorize.net/gateway/transact.dll"
url             = "https://secure.authorize.net/gateway/transact.dll"

' If an amount or description were posted to this page, the defaults are overidden
If Request.Form("amount") <> "" Then
	amount = Request.Form("amount")
End If
If Request.Form("description") <> "" Then
	description = Request.Form("description")
End If

' also check to see if the amount or description were sent using the GET method
If Request.QueryString("amount") <> "" Then
	amount = Request.QueryString("amount")
End If
If Request.QueryString("description") <> "" Then
	description = Request.QueryString("description")
End If
%>
-->


<!-- METADATA  (commenting out ASP a la: https://stackoverflow.com/questions/4431170/server-side-comments-whats-the-equivalent-of-in-asp-classic)
<%
' an invoice is generated using the date and time
Dim invoice
invoice	= Year(Date) & Month(Date) &  Day(Date) & Hour(Now) & Minute(Now) & Second(Now)
' a sequence number is randomly generated
Dim sequence
Randomize
sequence	= Int(1000 * Rnd)
' a time stamp is generated (using a function from simlib.asp)
Dim timeStamp
'timeStamp = simTimeStamp()
' a fingerprint is generated using the functions from simlib.asp and md5.asp
Dim fingerprint
'fingerprint = HMAC (transactionKey, loginID & "^" & sequence & "^" & timeStamp & "^" & amount & "^")
%>
-->


<!-- METADATA  (commenting out ASP a la: https://stackoverflow.com/questions/4431170/server-side-comments-whats-the-equivalent-of-in-asp-classic)
<%
' Create the HTML form containing necessary SIM post values
Response.Write("<FORM method='post' action='" & url & "' >")
' Additional fields can be added here as outlined in the SIM integration guide
' at: http://developer.authorize.net
Response.Write("	<INPUT type='hidden' name='x_login' value='" & loginID & "' />")
Response.Write("	<INPUT type='hidden' name='x_amount' value='" & amount & "' />")
Response.Write("	<INPUT type='hidden' name='x_description' value='" & description & "' />")
Response.Write("	<INPUT type='hidden' name='x_invoice_num' value='" & invoice & "' />")
Response.Write("	<INPUT type='hidden' name='x_fp_sequence' value='" & sequence & "' />")
Response.Write("	<INPUT type='hidden' name='x_fp_timestamp' value='" & timeStamp & "' />")
Response.Write("	<INPUT type='hidden' name='x_fp_hash' value='" & fingerprint & "' />")
Response.Write("	<INPUT type='hidden' name='x_test_request' value='" & testMode & "' />")
Response.Write("	<INPUT type='hidden' name='x_show_form' value='PAYMENT_FORM' />")
Response.Write("	<input type='submit' value='" & label & "' />")
Response.Write("</FORM>")
%>
-->
<!-- This is the end of the code generating the "submit payment" button.    -->

<%
' The parameters for the USAePay payment form etc. can be configured here.
' To post to the USAePay payment form in our SANDBOX account:
'URL         = "https://sandbox.usaepay.com/interface/epayform/"
' To post to the USAePay payment form in our PRODUCTION account:
URL         = "https://www.usaepay.com/interface/epayform/"

Description	= "Sample Transaction"
'UMkey       = "_GzuIx1hT7dOT4DGlVc45ac9iHLdpgR5"    ' My source key for SANDBOX useage. 
UMkey       = "v0tm18XWN8NKP222CtS55tu7467PFGHx"    ' My source key for PRODUCTION useage.
'UMcommand   = "sale"
UMcommand   = "authonly"                            ' Authorize card only; only allowed in Production account?
UMamount    = "0.99"                                ' transaction amount. 
UMinvoice   = "Test-0003"                           ' invoice number.

' UMcustreceipt tells whether to send receipt to customer. But this line didn't work. 
' On 6/26/17 Jenny Silva (of USAePay Integration Support) hardcoded this line into the form at 
' her server, and it worked: 
' I was then able to receive receipt to my test customer email address "staff@intelligineering.com". -->
UMcustreceipt = "yes"                              

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

' Customer's Shipping Address fields:
' See http://help.usaepay.com/merchant/epaymentform
UMshipfname     = "Bernard"
UMshiplname     = "Nadel"
UMshipcompany   = "Intelligineering"
UMshipstreet    = "14060 Balfour"
UMshipstreet2   = ""
UMshipcity      = "Oak Park"
UMshipstate     = "MI"
UMshipzip       = "48237"
UMshipcountry   = "USA"
UMshipphone     = "313-307-1874"

SubmitbtnLabel  = "Continue to Payment Form"

%>


<%
' Print the variables and their value to the screen.
Response.Write("Description: " & Description & " <br />")
Response.Write("Amount: " & UMamount & " <br />")
%>

<!--
Create the HTML form containing necessary POST values
Additional fields can be added here as outlined at
http://help.usaepay.com/merchant/epaymentform
-->

<%
Response.Write("<FORM method='POST' action='" & URL & "' >")
' Additional fields can be added here as outlined at:
' http://help.usaepay.com/developer/transactionapi?s[]=transaction&s[]=api
Response.Write("	<INPUT type='hidden' name='UMkey' value='" & UMkey & "' />")
Response.Write("	<INPUT type='hidden' name='UMcommand' value='" & UMcommand & "' />")
Response.Write("	<INPUT type='hidden' name='UMamount' value='" & UMamount & "' />")
Response.Write("	<INPUT type='hidden' name='UMinvoice' value='" & UMinvoice & "' />")

' Customer's Billing Address fields:
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
Response.Write("	<INPUT type='hidden' name='UMemail' value='" & UMemail & "' />")  <!-- customer's email. -->
Response.Write("	<INPUT type='hidden' name='UMwebsite' value='" & UMwebsite & "' />")

' Customer's Shipping Address fields:
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
         

Response.Write("	<input type='submit' value='" & SubmitbtnLabel & "' />")
Response.Write("</FORM>")
%>

</BODY>
</HTML>