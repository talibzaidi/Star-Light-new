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




<%
If False Then
SubmitGoToLinkPoint = Request.QueryString("SubmitGoToLinkPoint") 
'Response.Write "SubmitGoToLinkPoint = " & SubmitGoToLinkPoint
'Response.End
If SubmitGoToLinkPoint <> "" Then   ' User has pressed this form's Submit button to proceed to LinkPoint payment page.
	'Response.redirect "cashout.asp"		' cashout.asp is a webpage on the way to LinkPoint webpage.
	Response.redirect "cashoutAuthorizeNet.asp"		' cashoutAuthorizeNet.asp is a webpage on the way to Authorize.net webpage.
	' 4/22/07, BN: Apparently Linkpoint knows that we will be coming there from file cashout.asp. 
	' I apparenttly can NOT change the name of this file without letting LinkPoint know. I can NOT even add 
	' Querystring parameters to the call e.g. cashout.asp?GoTo=LinkPoint, so that I could then also do 
	' cashout.asp?GoTo=PayPal.
End If
End If    ' False
%>


<%
If False Then
SubmitGoToPayPal = Request.QueryString("SubmitGoToPayPal") 
'Response.Write "SubmitGoToPayPal = " & SubmitGoToPayPal
'Response.End
If SubmitGoToPayPal <> "" Then   ' User has pressed the Submit button on this form, to proceed to PayPal payment processor.
	Response.redirect "cashoutAtPayPal.asp"		' cashoutAtPayPal.asp is a webpage on the way to the PayPal payment page.
End If
End If    ' False
%>


<html>


<head>
<link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->

<meta name="keywords" content="GPS,Navigation,Garmin,CB-Radios,Uniden,Cobra,Motorola,2-way radios,Hand-tools,Pyramid ">
<meta name="description" content="Online store for GPS Global Positioning Systems, Navigation equipment, CB Radios, FRS Radios, GMRS Radios, Antennas, Car Audio, Hand Tools.  Shopping on a secure SSL line. Accepting Visa,
Mastercard, American Express cards.">
<meta name="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds">
<title>Thank You</title>

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
<% 
'Response.Write Session("InvoiceForCustomer")

PaymentMethod = Request.QueryString("PaymentMethod")
'Response.Write "<br>PaymentMethod = "& PaymentMethod

' 8/3/08: This page is now being skipped for cases PaymentMethod = "CreditCard" and PaymentMethod = "PayPal", because in the 
' "CreditCard" case, I could not get it to submit invisibly - although I could in the "PayPal" case for some weird reason.
' So we now just jump diretly from EmailSend.asp to cashoutAuthorizeNet.asp or to cashoutAtPayPal.asp in these two cases, and
' just bypass this page completely.
' However, in the other cases ("NonUSorCanadianCustomer", "Check", "MoneyOrder") we still come to this page from EmailSend.asp.
Select Case PaymentMethod
Case "NonUSorCanadianCustomerNonPayPal"
	Response.Write Session("InvoiceForCustomer")
Case "Check" 
	Response.Write Session("InvoiceForCustomer")
Case "MoneyOrder"
	Response.Write Session("InvoiceForCustomer")
Case "CreditCard"
	Response.Write "<table align=center border=0 cellpadding=5><tr><td>"
	Response.Write Session("InvoiceForCustomer")
	Response.Write "</td><td valign=top align=center width=250>"
%>	
	<!-- <form id=form1 name=form1> -->
	<form action="cashoutAuthorizeNet.asp" method="POST" id=form1 name="form1">
	<font face=Tahoma color=navy><b>To continue to our secure credit card payment page ...</b></font><br><br>
	<input type="submit" value="Go" name="SubmitGoToLinkPoint"> 
	</form> 
	
	<script language="javascript">
	document.form1.submit();  // Automatic submission of this form, to avoid need for user to see it and click the Submit button. Works!
	</script>
<%
	Response.Write "</td></tr></table>"
Case "PayPal", "NonUSorCanadianCustomerPayPal"
	Response.Write "<table align=center border=0 cellpadding=5><tr><td>"
	Response.Write Session("InvoiceForCustomer")
	Response.Write "</td><td valign=top align=center width=250>"
%>	
	<!-- <form id=form2 name=form2> -->
	<form action="cashoutAtPayPal.asp" method="POST" id=form2 name="form2">
	<font face=Tahoma color=navy><b>To continue to our secure PayPal payment page ...</b></font><br><br>
	<input type="submit" value="Go to PayPal" name="SubmitGoToPayPal"> 
	</form> 
	
	<script language="javascript">
	document.form2.submit();  // Automatic submission of this form, to avoid need for user to see it and click the Submit button. Works!
	</script>
<% 
	Response.Write "</td></tr></table>"
End Select 
%>




<% If PaymentMethod <> "CreditCard" AND PaymentMethod <> "PayPal" Then %>
<br>
<center>      
<font face="tahoma" size="3"><b><a href="http://www.starlite-intl.com">Click here</a> to continue shopping.</b></font>
</center>
<br>
<% Else %>
<br><br>
<% End If %>

</body>


</html>
