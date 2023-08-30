<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<% ar = Request("Area") %>
<% Area = Request("Area") %>
<% sar = ReQuest("sar") %>
<% SID = ReQuest("SID") %>
<% spec = 6 %>




<html>


<head>
<meta name="keywords" content="GPS,Navigation,Garmin,CB-Radios,Uniden,Cobra,Motorola,2-way radios,Hand-tools,Pyramid ">
<meta name="description" content="Online store for GPS Global Positioning Systems, Navigation equipment, CB Radios, FRS Radios, GMRS Radios, Antennas, Car Audio, Hand Tools.  Shopping on a secure SSL line. Accepting Visa,
Mastercard, American Express cards.">
<meta name="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds">
<title>Email Send</title>

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


<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<!--#include virtual="Misc/Header.INC"-->

<br>

<%'=Session("InvoiceForStarlite")%>
<%'=Session("InvoiceForCustomer")%>


<%
' BN: Send the email to Starlite ...
If True Then
Dim myMail
Set myMail			= Server.CreateObject("CDONTS.NewMail")
myMail.From			= "sales@starlite-intl.com"   
myMail.To			= "sales@starlite-intl.com" 
myMail.BCC			= "bn@intelligineering.com"	
myMail.Subject		= "Electronic Order # " & Session("OrderNum")
myMail.BodyFormat	= 0 
myMail.MailFormat	= 0 
myMail.Body			= Session("InvoiceForStarlite") 
myMail.Send 
End If   ' True / False


' BN: Send the email to customer ...
If True Then
Dim customerMail
Set customerMail		= Server.CreateObject("CDONTS.NewMail")
customerMail.From		= "sales@starlite-intl.com"    
customerMail.To			= Session("CustomerEmail")
customerMail.BCC		= "sales@starlite-intl.com, staff@mitmazel.com"	
customerMail.Subject	= "Your Order # " & Session("OrderNum") & " from Star Lite" 
customerMail.BodyFormat	= 0 
customerMail.MailFormat	= 0 
customerMail.Body		= Session("InvoiceForCustomer")
customerMail.Send 
End If   ' True / False


PaymentMethod = Request.QueryString("PaymentMethod") 
Response.Redirect "EmailShow.asp?PaymentMethod=" & PaymentMethod 
%>


</body>

</html>
