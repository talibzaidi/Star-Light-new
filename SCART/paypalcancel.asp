<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<% ar = Request("Area") %>
<% Area = Request("Area") %>
<% sar = ReQuest("sar") %>
<% SID = ReQuest("SID") %>
<% spec = 6 
msgbody="<html><head></head><body>The Pay Pal order placed by " & request.querystring("name") & " has been cancelled.</body></html>"
Set myMail = Server.CreateObject("CDONTS.NewMail")
myMail.From = "OrderCancel@StarliteEcommerce" 
		myMail.To = "starlite@starlite-intl.com" 
		myMail.Subject = "PayPal Order Cancel" 
		myMail.BodyFormat = 0 
		myMail.MailFormat = 0 
		myMail.Body = msgbody
		myMail.Send 
set myMail = nothing%>
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
        <td background="../Images/topback.gif"><div align="left"><table border="0" cellpadding="0"
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
	<td width=>&nbsp;
<table border="0">
<tr>
<td width="170">&nbsp;</td>
<td width="380">
<center>
<br>
									<font face="tahoma" size="6"><b>Your order has been cancelled. </b></font><br>
									<br>
<font face="tahoma" size="3"><b>Please shop again. <br>
<a href="http://www.starlite-intl.com">Click here to return to starlite.</a></b></font>
                </center></div>
                <br><br>  <br><br> <br><br><br><br>
<CENTER>
</td>
</tr>
</table> </td>
                <td valign="top" >


              
        </td>
     
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
