<%@ LANGUAGE = VBScript %>


<html>

<head>
<link rel="stylesheet" type="text/css" href="../../Misc/StyleSheet1.css">
<title>Starlite-intl.com - GPS, tracking GPS, cb radios, flash memory, audio/video, marine radios, fish finders, hand tools</title>
<meta name="keywords" content="GPS,Navigation,Garmin,CB-Radios,Uniden,Cobra,Motorola,2-way radios,Hand-tools,Pyramid ">
<meta name="description" content="Online store for GPS Global Positioning Systems, Navigation equipment, CB Radios, FRS Radios, GMRS Radios, Antennas, Car Audio, Hand Tools.  Shopping on a secure SSL line. Accepting Visa, Mastercard, American Express cards.">
<meta name="Keywords" content="Star, Lite, International, Sales, Electronics, Shop, Online, Commerce, Store, Buy, Sell, Free, Special, USA, Canada, Shipping, Business, Price, Best ">
<title>Starlite-intl.com - GPS, tracking GPS, cb radios, flash memory, audio/video, marine radios, fish finders, hand tools</title>
</head>



<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">


<!--#include virtual="Misc/Header.INC"-->


<br>
<br>

<%
SQL = "Select * from Company WHERE ID = 1"
Set conn = Server.CreateObject("ADODB.Connection")
Conn.Open Session("ConnectionString")
Set RS = Conn.Execute(SQL)
%>


<table cols="1" align='center' width='900' cellpadding='0' cellspacing='20' border=0>
<tr>
<td align='center'>

		<font size="3" face="Tahoma"><strong>
		<%=RS("Name")%><br>
		<%=RS("Addr")%><br>
		<%=RS("Postal")%><br>
		<%=RS("City")%><br>
		<%=RS("StPro")%><br><br>
		<%=RS("Phon1")%><br>
		<%=RS("Phon2")%><br>
		<%=RS("Fax")%><br>
		<%=RS("800num")%><br><br>
        Email: <a href="mailto: <%=RS("Email")%>"><%=RS("Emailname")%></a>
		</strong></font>
		
</td>
</tr>
</table>


</body>

</html>


