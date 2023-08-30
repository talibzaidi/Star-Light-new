<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<% spec = 2 %>
<% locale = "Classified"%>

<html>

<head>
<title>Starlite-Intl.</title>
<meta name="keywords"
content="GPS,Navigation,Garmin,CB-Radios,Uniden,Cobra,Motorola,2-way
radios,Hand-tools,Pyramid ">
<meta name="description" content="Online store for GPS Global Positioning
Systems, Navigation equipment, CB Radios, FRS Radios, GMRS Radios, Antennas,
Car Audio, Hand Tools.  Shopping on a secure SSL line. Accepting Visa,
Mastercard, American Express cards.">
<title>Starlite International LLC - Online Store</title>
<script language="Javascript">
<!--

	once = new MakeArray(6)
	over = new MakeArray(6)
	under = new MakeArray(6)
	standard = new MakeArray(1)
 
	once[0].src = "Images/question1.gif"
	once[1].src = "Images/scart1.gif"
	once[2].src = "Images/home1.gif"
	once[3].src = "Images/new1.gif"
                once[4].src = "Images/cat1.gif"
	once[5].src = "Images/ex1.gif"    
		
	over[0].src = "Images/question2.gif"
	over[1].src = "Images/scart2.gif"
	over[2].src = "Images/home2.gif"
	over[3].src = "Images/new2.gif"
	over[4].src = "Images/cat2.gif"
	over[5].src = "Images/ex2.gif"

	under[0].src = "Images/helpnav.gif"
	under[1].src = "Images/shoppingcartnav.gif"
	under[2].src = "Images/homenav.gif"
	under[3].src = "Images/newproductsnav.gif"
	under[4].src = "Images/onlinecataloguenav.gif"
	under[5].src = "Images/specialsnav.gif"

	standard[0].src = "Images/emptynav.jpg"

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

<body background="../../Images/background.gif" bgcolor="#FFFFFF"
link="#000000" vlink="#000000" topmargin="0" leftmargin="0"
marginwidth="0" marginheight="0">


<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td><div align="left"><table border="0" cellpadding="0"
        cellspacing="0" width="575">
            <tr>
                <td> <!--#include file="NAV.INC"--><img
                src="../../Images/toptitle.jpg" width="411" height="29"><br>
                </td>
            </tr>
            <tr>
                <td width="575"><img src="../../Images/emptynav.jpg"
                width="164" height="14"><img
                src="../../Images/bottitle.JPG" width="411" height="14"></td>
            </tr>
            <tr>
                <td><img src="../../Images/leftbar.gif" width="176"
                height="23"><a href="../../class.asp"><img
                src="../../Images/classifieds.gif" alt="Classifieds"
                border="0" width="115" height="23"></a><a
                href="../../links.asp"><img src="../../Images/links.gif"
                alt="Links" border="0" width="91" height="23"></a><a
                href="../../contact.asp"><img
                src="../../Images/ContactUS.gif" alt="Contact Us"
                border="0" width="126" height="23"></a><a
                href="../../help.asp"><img src="../../Images/Help.gif"
                alt="Help" border="0" width="67" height="23"></a></td>
            </tr>
        </table>
        </div></td>
        <td width="100%"
        background="../../Images/topback.gif">&nbsp;</td>
    </tr>
    <tr>
        <td><table border="0" cellpadding="5" cellspacing="0">
            <tr>
                <td align="center" valign="top" width="166"><img
                src="../../Images/logo.gif" width="140" height="145">
<!--#include file="SPECIAL.INC"-->
<br>

</td>
                <td valign="top" width="382"><div align="center"><center>
                </center></div>

<br>
           <form action="../../GOTIT.asp" method="post">
    <input type="hidden" name="Date" value="<%=Date%>"><table
    border="0" width="100%">
       
        <tr  background="bg.gif">
            <td><div align="center"><center><table border="0" cellpadding="5"
            width="99%" bgcolor="#0080C0">
                             <tr>
                    <td><font color="#000000" face="Arial"><strong><u>Post
                    in</u>: </strong></font><table border="0"
                    width="100%">
                        <tr>
                            <td><font face="Arial"><input
                            type="radio" checked name="R1"
                            value="Buy"><b>Buy</b></font></td>
                            <td><font face="Arial"><input
                            type="radio" name="R1" value="Sell"><b>Sell</b></font></td>
                        </tr>
                    
                    </table>
                    <p><font face="Arial"><b>Your contact name: &nbsp;</b><input
                    type="text" size="20" name="Author"></font></p>
		    <p><font face="Arial"><b>Your contact phone: </b><input
                    type="text" size="20" name="Contact"></font></p>
                    <p><font face="Arial"><b>Your ad:</b>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input
                    type="text" size="20" maxlength="200"
                    name="Message" height="4"></font></p>
                    <p><font face="Arial"><b>Your ad cannot exceed
                    200 characters.</b></font></p>
                    <p><font face="Arial"><b>Your email:</b>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input
                    type="text" size="20" name="Email"></font></p>
                    <p><font face="Arial"><b>Date: <%=Date%></b></font></p>
                    <p >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit"
                    name="B1" value="Submit"></p><br>
                    </td>
                </tr>
            </table>
            </center></div></td>
        </tr>
      
    </table>
</form>

</div align="center"></center>
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
        <td><img src="../../Images/bottompage.GIF" width="575"
        height="52"></td>
        <td
        background="../../Images/botback.gif">&nbsp;</td>
    </tr>
</table>
</body>
</html>
