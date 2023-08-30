<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
		
		
<!--#INCLUDE FILE="PaymentAuthorizeNet/simlib.asp"-->
<!--#INCLUDE FILE="PaymentAuthorizeNet/simdata.asp" -->

		
<html>

<head>
</head>


<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">


<% '****************** The following is based on sample file PaymentAuthorizeNet/sim.asp ****************** %>

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
'Response.Write amount
'Response.End


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
			<input type="submit" value="Authorize.Net  ::  Accept Order"> 
		</form> 
		
		<script language="javascript">
		document.AuthorizeNet.submit(); 
		</script>

		</td>
	</tr>
</table> 


</body>

</html>