<%@ LANGUAGE = VBScript %>


<html>

<head>
<link rel="stylesheet" type="text/css" href="../../Misc/StyleSheet1.css">
<title>Star Lite Intl. - GPS, GPS Sensors, GPS accessories, CB radios, 2-way radios, Marine electronics, Flash memory, Scanners, MP3, audio/video, hand tools</title>
<meta name="keywords" content="GPS, GPS navigation, GPS sensors, OEM GPS, GPS accessories, CB, CB radio, CB radios, Garmin gps, global positioning, WalkyTalky, mobile tracking, fleet tracking, USglobasat gps, bluetooth, flash memory, gmrs, marine radios, navigation electronics, 2-way radios, radio scanners, marine radios, car audio, car stereos, power amplifiers, antennas, power supplies, regulated power supplies, DJ, accessories, hand tools, mechanics tools, Uniden, Cobra, Midland, MIT, Pyramid, Pyle, Solarcon">
<meta name="description" content="Large selection of - GPS, GPS sensors, GPS accessories, GPS OEM, PDA, tracking gps, bluetooth gps, fish finders, sounders, CB radios and Walky-talky, flash memory, MP3, radio scanners, digital cameras, car audio and car video, DJ, hand tools and mechanics tools">
</head>



<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<table border="0" bordercolor="green" align="center">		<% ' Start Table 1 %>
<tr><td>

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


</td>
</tr>
</table>

</body>

</html>


