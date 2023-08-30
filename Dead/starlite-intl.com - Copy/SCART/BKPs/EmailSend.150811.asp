<%@ LANGUAGE = VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->




<%response.buffer=true%>
<% ar = Request("Area") %>
<% Area = Request("Area") %>
<% sar = ReQuest("sar") %>
<% SID = ReQuest("SID") %>
<% spec = 6 %>




<html>


<head>
<link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->

<meta name="keywords" content="GPS,Navigation,Garmin,GPS sensors,OEM GPS,USGlobalSat,Pharos,CB-Radios,Uniden,Cobra,2-way radios ">
<meta name="description" content="Source for GPS - Global Positioning Systems, Navigation equipment, CB Radios, GMRS Radios, Scanners, Antennas, Digital equipment and Hand Tools.">
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
If FALSE Then   ' 4/29/10: Tech support said CDONTS was unreliable, so I FALSE-d this out and started using ASPemail below.
	Dim myMail
	Set myMail			= Server.CreateObject("CDONTS.NewMail")
	myMail.From			= "sales@starlite-intl.com"   
	myMail.To			= "sales@starlite-intl.com" 
	' myMail.BCC		= "bn@intelligineering.com"	
	myMail.Subject		= "Electronic Order # " & Session("OrderNum")
	myMail.BodyFormat	= 0 
	myMail.MailFormat	= 0 
	myMail.Body			= Session("InvoiceForStarlite") 
	myMail.Send 
	Response.Write "<br>" & Session("InvoiceForStarlite")
End If   ' True / False


' BN: Send the email to Starlite ...  4/29/10: Using ASPemail instead of CDONTS.
' Based on an example at Sani's hosting company, at http://www.hosting.com/support/programming/aspmail/
If TRUE Then		
	Set Mailer 			= Server.CreateObject("SMTPsvg.Mailer")
	Mailer.FromName 	= "www.starlite-int.com"
	Mailer.FromAddress	= "starlite@starlite-intl.com"
	Mailer.RemoteHost 	= "mail.starlite-intl.com"
	Mailer.AddRecipient   "Sani", "starlite@starlite-intl.com"
	Mailer.AddExtraHeader "X-MimeOLE:Produced starlite-intl.com"
	Mailer.Subject 		= "Electronic Order # " & Session("OrderNum")
	Mailer.BodyText 	= Session("InvoiceForStarlite")
	Mailer.ContentType = "text/html"
	If Mailer.SendMail then
	  Response.Write "<br>The email was successfully submitted ..."
	Else
	  Response.Write "<br><br><br><br><br><br>The email send was not successful. The error was " & Mailer.Response
	  Response.End
	End If
	Set Mailer = Nothing
End If   ' True / False



' BN: Send the email to customer ...
If FALSE Then    ' 4/29/10: Tech support said CDONTS was unreliable, so I FALSE-d this out and started using ASPemail below.
Dim customerMail
Set customerMail		= Server.CreateObject("CDONTS.NewMail")
customerMail.From		= "sales@starlite-intl.com"    
customerMail.To			= Session("CustomerEmail")
customerMail.BCC		= "sales@starlite-intl.com"   		' "sales@starlite-intl.com, staff@mitmazel.com"	
customerMail.Subject	= "Your Order # " & Session("OrderNum") & " from Star Lite" 
customerMail.BodyFormat	= 0 
customerMail.MailFormat	= 0 
customerMail.Body		= Session("InvoiceForCustomer")
customerMail.Send 
Response.Write "<br><br><br><br>" & Session("InvoiceForCustomer")
End If   ' True / False


' BN: Send the email to customer ...  4/29/10: Using ASPemail instead of CDONTS.
' Based on an example at Sani's hosting company, at http://www.hosting.com/support/programming/aspmail/
If TRUE Then		
	Set Mailer 			= Server.CreateObject("SMTPsvg.Mailer")
	Mailer.FromName 	= "www.starlite-int.com"
	Mailer.FromAddress	= "starlite@starlite-intl.com"
	Mailer.RemoteHost 	= "mail.starlite-intl.com"
	Mailer.AddRecipient   "", Session("CustomerEmail")
	Mailer.AddBCC 		  "Sani", "starlite@starlite-intl.com"
	Mailer.AddExtraHeader "X-MimeOLE:Produced starlite-intl.com"
	Mailer.Subject 		= "Electronic Order # " & Session("OrderNum")
	Mailer.BodyText 	= Session("InvoiceForCustomer")
	Mailer.ContentType = "text/html"
	If Mailer.SendMail then
	  Response.Write "<br>The email was successfully submitted ..."
	Else	  
	  Response.Write "<br><br><br><br><br><br>The email send was not successful. The error was " & Mailer.Response
	  Response.End
	End If
	Set Mailer = Nothing
End If   ' True / False



PaymentMethod = Request.QueryString("PaymentMethod") 
'Response.Write "<br><br><br><br>PaymentMethod = " & PaymentMethod 
'Response.End
Response.Redirect "EmailShow.asp?PaymentMethod=" & PaymentMethod 


If PaymentMethod = "CreditCard" Then
	Response.Redirect "cashoutAuthorizeNet.asp"
ElseIf PaymentMethod = "PayPal" OR PaymentMethod = "NonUSorCanadianCustomerPayPal" Then
	Response.Redirect "cashoutAtPayPal.asp"
Else
	Response.Redirect "EmailShow.asp?PaymentMethod=" & PaymentMethod 
End If
%>


</body>

</html>









