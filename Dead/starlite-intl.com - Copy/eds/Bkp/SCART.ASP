<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<% ar = Request("Area") %>
<% Area = Request("Area") %>
<% if Area="iii" then
   Area = Request("Manufat")
   ar = Request("Manufat")

   end if
   if Area = "New Products" then
     response.redirect "./scart.asp?pid=0&sid=11&area=New!&sar=New%20Products"
   end if	

   if Area="Choose a catalog area." then
   response.redirect "http://www.starlite-intl.com/index.asp"
   end if
%>
<% sar = ReQuest("sar") %>
<% 'sar = Replace( sar, " ", "%20") %>
<% SID = ReQuest("SID") %>
<html>

<head>
<meta name="keywords" content="GPS,Navigation,Garmin,CB-Radios,Uniden,Cobra,Motorola,2-way radios,Hand-tools,Pyramid ">
<meta name="description" content="Online store for GPS Global Positioning Systems, Navigation equipment, CB Radios, FRS Radios, GMRS Radios, Antennas, Car Audio, Hand Tools.  Shopping on a secure SSL line. Accepting Visa,
Mastercard, American Express cards.">
<meta name ="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds">

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

<body background="../Images/background.jpg" bgcolor="#FFFFFF"
link="#000000" vlink="#000000" topmargin="0" leftmargin="0"
marginwidth="0" marginheight="0">


<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td><div align="left"><table border="0" cellpadding="0"
        cellspacing="0" width="575">
             <tr>
                <td> <!--#include file="NAV.INC"--><img
                src="../Images/toptitle.jpg" width="411" height="29"><br>
                </td>
            </tr>
            <tr>
                <td width="575"><img src="../Images/emptynav.jpg"
                width="164" height="14"><img
                src="../Images/bottitle.JPG" width="411" height="14"></td>
            </tr>
            <tr>
                <td><img src="../Images/leftbar.gif" width="176"
                height="23"><img src="../Images/blanka1.gif"></td>
            </tr>
        </table>
        </div></td>
        <td width="100%"
        background="../Images/topback.gif">&nbsp;</td>
    </tr>
    <tr>
        <td><table border="0" cellpadding="5" cellspacing="0">
<tr><td align="center" valign="bottom">
<% If Session("Country") = "Canada" Then%>
<img src="../Images/can1.gif" border="1">
<%elseif Session("Country") = "USA" Then %>
<img src="../Images/USA1.gif" border="1">
<%end if%>
</td></tr> 
           <tr>
                <td align="left" valign="top" width="166">


&nbsp;&nbsp;<font face="tahoma" size="3"><b><u>Catalog Sections:</u></b>
<br>
<%
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	dim sdsqstring
	sdsqstring = "select AreaName from Area51 WHERE  AreaVisible = yes ORDER BY AreaName Asc"
        	Set RSS = Conn.Execute(sdsqstring)
%>

<table>
<%do while not rsS.eof %>


<% if RSS("AreaName")= "New!" then%>
<tr><td><% LINKER = Replace( RSS("AreaName") , " ", "%20") %>
<img src="../images/arrow2.gif"><font face="tahoma" size="2"><a href="scart.asp?pid=0&sid=11&area=New!&sar=New%20Products"><b>
 <%="New Products"%>
<%elseif RSS("AreaName") = "Manufa" then%>


<%else%>

<tr><td><% LINKER = Replace( RSS("AreaName") , " ", "%20") %>
<img src="../images/arrow2.gif"><font face="tahoma" size="2"><a href="./scart.asp?area=<%=LINKER%>&amp;sid=0"><b>


 <%=RSS("AreaNAme")%>
 <%end if%>
</b></a><br>
</td></tr>
<% rsS.movenext
      loop
      rsS.close
      conn.close
%>
<tr><td>
<img src="../images/arrow2.gif"><font face="tahoma" size="2"><a href="Service.asp?sid=0"><b>Terms And Conditions</b></a>

</td></tr>




</table>
</td>
                <td valign="top" width="382"><!--#include file="BANNER.INC"-->
                <!--#include file="SCART.INC"--><br><br>

                <div align="center"><center><table border="1"
                cellpadding="3" cellspacing="0" width="95%"
                bgcolor="#0000FF" bordercolor="#000000">
                    <tr>
                        <td align="center"><a href="#top"><font color="#FFFFFF"
                        size="4"><strong>RETURN TO TOP OF PAGE</strong></font></a></td>
                    </tr>
                </table>
                </center></div></td>
            </tr>
        </table>
        </td>
        <td>&nbsp;</td>
    </tr>
    <tr>
        <td><img src="../Images/bottompage.GIF" width="575"
        height="52"></td>
        <td
        background="../Images/botback.gif">&nbsp;</td>
    </tr>
</table>
</body>
</html>
