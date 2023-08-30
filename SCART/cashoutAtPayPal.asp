<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>

<html>

<head>
</head>


<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<%
If False Then
	Response.Write "<br>Session('CustomerName') = " & Session("CustomerName")
	Response.Write "<br>Session('Country') = " & Session("Country")
	Response.Write "<br>Session('Charge') = " & Session("Charge")
	Response.Write "<br>Session('CustomerEmail') = " & Session("CustomerEmail")
	Response.End
End If
%>

<br><br><br><br><br><br>
<table border="0" align=center>
	<tr>
		<td align=center>
		
		<form action="https://www.paypal.com/cgi-bin/webscr" method="post" name="PayPal">
			<input type="hidden" name="cmd" value="_xclick">
			<input type="hidden" name="business" value="starlite@starlite-intl.com">
			<input type="hidden" name="item_name" value="The order for <%=Session("CustomerName")%> of <%=Date()%>">
			<input type="hidden" name="currency_code" value="<% If Session("Country") = "Canada" Then %>CAD
															 <% Else %>USD
															 <% End If %>">
			<input type="hidden" name="amount" value="<%=Session("Charge")%>">
			<input type="hidden" name="custom" value="<%=Session("CustomerEmail")%>">
			<input type="hidden" name="return" value="https://www.starlite-intl.com/scart/thankyou.asp">
			<input type="hidden" name="cancel_return" name="https://www.starlite-intl.com/scart/paypalcancel.asp?name=<%=Session("CustomerName")%>">

			<% ' Customer Info %>
			<input type="hidden" name="first_name" value="<%=Session("FName")%>">
			<input type="hidden" name="last_name" value="<%=Session("LName")%>">
			<input type="hidden" name="address1" value="<%=Session("Address")%>">
			<input type="hidden" name="city" value="<%=Session("City")%>">
			<input type="hidden" name="state" value="<%=Session("StateProv")%>">
			<input type="hidden" name="country" value="<%=Session("Country")%>"> <% ' 4/16/08: Not passing properly %>
			<input type="hidden" name="zip" value="<%=Session("Postal")%>">
			<input type="hidden" name="night_phone_a" value="<%=Session("Phone")%>">
			<input type="hidden" name="email" value="<%=Session("Email")%>">

		</form>
		
		<script language="javascript">
		document.PayPal.submit();
		</script> 
		
		</td>
	</tr>
</table> 

</body>

</html>
