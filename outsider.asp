<%@ LANGUAGE = VBScript %>

<%response.buffer=true%>

<% URL = ReQuest("WarpURL") %>


<head>
<meta name="keywords" content="GPS,Navigation,Garmin,CB-Radios,Uniden,Cobra,Motorola,2-way radios,Hand-tools,Pyramid ">
<meta name="description" content="Online store for GPS Global Positioning Systems, Navigation equipment, CB Radios, FRS Radios, GMRS Radios, Antennas, Car Audio, Hand Tools.  Shopping on a secure SSL line. Accepting Visa,
Mastercard, American Express cards.">
<!-- <meta name ="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds"> -->
<title>Starlite International LLC - Online Store</title>
</head>


<html>

<FRAMESET COLS="*" frameborder="0" scrolling="no" topmargin="2" leftmargin="2" MARGINWIDTH="2" MARGINHEIGHT="2" border=0 NORESIZE>
		
	<FRAMESET ROWS="44,*" frameborder="0" NORESIZE scrolling="auto" topmargin="2" leftmargin="2" MARGINWIDTH="2" MARGINHEIGHT="2" border=0>
		<FRAME SRC="top.asp?warp=<%ReQuest("Warp")%>"  NAME="top" NORESIZE  scrolling="no" topmargin="0" leftmargin="0" MARGINWIDTH="0" MARGINHEIGHT="0" frameborder="0" border=0>
		<FRAME SRC="<%=URL%>" NAME="main"  scrolling="auto" topmargin="0" leftmargin="0" MARGINWIDTH="0" MARGINHEIGHT="0"  frameborder="0" border=0>
	</FRAMESET>
		
</FRAMESET>	

</HTML>
