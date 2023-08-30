<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
		
		
<html>

<head>
</head>


<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<br><br><br><br><br><br>
<table border="0" align=center>
	<tr>
		<td align=center>
		
		<form action="https://www.linkpointcentral.com/lpc/servlet/lppay" method="POST" name="LinkPoint">
			<input type="hidden" name="mode" value="fullpay"> 
			<input type="hidden" name="chargetotal" value="<%=Session("Charge")%>"> 
			<input type="hidden" name="storename" value="330566"> 
			<input type="submit" value="Continue to Secure Payment Form"> 
		</form> 
		
		<script language="javascript">
		document.LinkPoint.submit();
		</script>

		</td>
	</tr>
</table> 

</body>

</html>