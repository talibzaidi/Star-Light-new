<%@ LANGUAGE = VBScript %>

<!-- 
11/15/17
This file is based on (is a simplified version of) Sani's file EmailSend.asp 
from SCART folder, that uses CDOSys to send email programmatically; specifically, that 
uses CDOSys to send invoices to Sani's customers and to Sani.

This file is POST-ed to by the form in Sani's Contact Us page of file Misc2 > contact.asp
-->

<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->

<html>
<!--  foneFrame v1.0.1 Copyright 2011 Azalea Software, Inc. www.QRdvark.com/foneFrame/ 31aug11  -->


<head>
    <link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/mobile1/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
    <meta name="keywords" content="GPS,Navigation,Garmin,GPS sensors,OEM GPS,USGlobalSat,Pharos,CB-Radios,Uniden,Cobra,2-way radios ">
    <meta name="description" content="Source for GPS - Global Positioning Systems, Navigation equipment, CB Radios, GMRS Radios, Scanners, Antennas, Digital equipment and Hand Tools.">
    <meta name="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds">
    <title>Email Send</title>

	<meta charset=utf-8>
	
	<meta name="viewport" content="width=device-width; initial-scale=1.0">
	<!-- foneFrame.css is the stylesheet with comments, so it is readable.
	     foneFrame-min.css is the minimized version; it is smaller and loads faster. -->
	<link href="https://www.starlite-intl.com/mobile1/foneFrame.css" rel="stylesheet" type="text/css">
	<!-- The following 2 lines are not strict HTML5. -->
	<meta name="HandheldFriendly" content="true"/>
	<meta name="MobileOptimized" content="320"/>

	<!-- You can use different style sheets for mobile vs. computer browsers: -->
	<!--  <link href="style-mobile.css" rel="stylesheet" type="text/css" media="handheld"> -->
	<!--  <link href="style-computer.css" rel="stylesheet" type="text/css" media="screen"> -->
	<!-- The favicon & iOS home screen icon are both 57x57 PNG's. Use a full URL file path for Android devices.  -->
	<!--  <link rel="apple-touch-icon-precomposed" href="http://yoursite.com/apple-touch-icon.png">  -->
	<!--  <link rel="icon" type="image/vnd.microsoft.icon" href="http://yoursite.com/favicon.png">  -->
	<!-- Site maps help search spiders where to find your pages.  www.xml-sitemaps.com  -->
	<!--  <link rel="alternate" type="application/rss+xml" title="ROR" href="ror.xml"> -->
	<!-- Your Google Analytics code goes here, just before the </head> tag. -->

    <!-- 11/10/13: For the accordion menu from menucool.com, where its HTML is in a separate file, and does not have to be repeated in each webpage that has the menu. -->
    <link href="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.css" rel="stylesheet" type="text/css" />
    <script src="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.js" type="text/javascript"></script>

	<script type="text/javascript">amenu.close(true);</script>
</head>


<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<table style='width:100%;'>
<tr><td>
<!-- #include virtual="mobile1/Misc/Header.INC" -->
</td></tr>
</table>


<br>

<% 
'11/15/17:

UserFullName         = Request.Form("UserFullName")
UserEmailAddress     = Request.Form("UserEmailAddress")
UserAddress          = Request.Form("UserAddress")
UserPhone            = Request.Form("UserPhone")
EmailSubject         = Request.Form("EmailSubject")
UserEmailEnquiryToSL = Request.Form("UserEmailEnquiryToSL")

If FALSE Then
    Response.Write "<br>UserFullName = "         & UserFullName
    Response.Write "<br>UserEmailAddress = "     & UserEmailAddress
    Response.Write "<br>UserAddress = "          & UserAddress
    Response.Write "<br>UserPhone = "            & UserPhone
    Response.Write "<br>EmailSubject = "         & EmailSubject
    Response.Write "<br>UserEmailEnquiryToSL = " & UserEmailEnquiryToSL & "<br><br>"  
End If 

'Response.End
%>



<% ' *********************************************************************************** 
   ' [8/11/15, BN] Send Contact Us email to Starlite ...
%>

<%
    ' [8/11/15, BN]: Send the email to Starlite ...  8/11/15: Using CDOSYS instead of ASPemail or CDONTS.
    ' CDONTS is now obsolete and ASPemail is apparently not on the NEW server at HostMySite.com.
    ' Based on Knowledgebase article at
    ' https://solutions.hostmysite.com/index.php?/Knowledgebase/Article/View/8596/0/Using-CDOSys-to-create-an-ASP-Mail-form-that-uses-Authentication
    ' (See my version in test files: contactform.asp and send_form_email.asp in folder MyTests/EmailOnNewServer/.)
    ' See also http://www.w3schools.com/asp/asp_send_email.asp, where it says:
    ' "CDOSYS is a built-in component in ASP. This component is used to send e-mails with ASP." It replaces CDONTS.
    ' See also http://www.itgeared.com/articles/1433-asp-sending-email-cdosys/ 
    ' See also https://www.codeproject.com/Tips/754136/Send-Email-using-Classic-ASP 
%>

<%
If TRUE Then	
    Dim email_to, email_subject, host, username, password, reply_to, port, from_address
    Dim first_name, last_name, home_address, email_from, telephone, comments, error_message
    Dim ObjSendMail, email_message

    email_to =  "starlite@starlite-intl.com"    ' "bn2@intelligineering.com"   'Enter the email address you want to send to.
    email_subject = EmailSubject                
    host = "mail.starlite-intl.com"             'The mail server name. (Commonly mail.yourdomain.xyz if your mail is hosted with HostMySite).
    username = "starlite@starlite-intl.com"     'A valid email address you have setup.
    from_address = "starlite@starlite-intl.com" 'If your mail is hosted with HostMySite this has to match the email address [in line?] above.
    password = "S6a2n2I6"                       'Password for the above email address.
    reply_to = UserEmailAddress                 'Enter the email address you want recipient to reply to. UserEmailAddress is the email address of the user who submitted the Contact Us email.
    port = "25"                                 'This is the default port. Try port 50 if this port gives you issues and your mail is hosted with HostMySite.


    Sub Died(errors)
        Response.Write "errors = " & errors
        'Your error code can go here 
        Response.write("<br />There were error(s) found with the form you submitted. These errors appear below.<br /><br />")
        Response.write(errors & "<br />")
        Response.write("Please go back and fix these errors.<br />")
        Response.write("<font color='red'>And / or you may also need to make sure that cookies are enabled in your browser.</font><br /><br />")
        Response.End
    End Sub

    ' 8/13/15: Session variables, set in EmailBuild.asp, are empty below.
    ' This turned out to be because cookies were not enabled in the browser that Sani was using on his cell phone. 
    ' Cookies need to be enabled for Session variables to work!
    If FALSE Then 
        Response.Write "Session('FName') = " & Session("FName") & "<br />"
        Response.Write "Session('LName') = " & Session("LName") & "<br />"
        Response.Write "Session('Address') = " & Session("Address") & "<br />"
        Response.Write "Session('CustomerEmail') = " & Session("CustomerEmail") & "<br />"
        Response.Write "Session('InvoiceForStarlite') = " & Session("InvoiceForStarlite") & "<br />"
    End If

    first_name = Session("FName")  ' Request.Form("first_name")  'required 
    last_name = Session("LName")   ' Request.Form("last_name")  'required 
    home_address = Session("Address")  ' Request.Form("home_address")  'not required
    'email_from = Session("CustomerEmail") ' Request.Form("email")  'required   
    ' 11/15/17:
    email_from = UserEmailAddress ' Request.Form("email")  'required   
    telephone = Request.Form("telephone")  'not required
    comments = Session("InvoiceForStarlite") ' Request.Form("comments")  'required 


    'Validate expected data exists
    'If Request.Form("first_name") = "" Or Request.Form("last_name") = ""  Or Request.Form("email") = ""  Or Request.Form("comments") = "" Then
    'If first_name = "" Or last_name = ""  Or email_from = ""  Or comments = "" Then
        'Call Died("Required field/s have not been entered")
    'End If

    IF TRUE Then
        errors = ""
        If first_name = "" Then
            errors = errors & "The first name field has not been entered" & "<br />"
        End If
        If last_name = "" Then
            errors = errors & "The last name field has not been entered" & "<br />"
        End If
        If Not(errors = "") Then
            'Call Died(errors)
        End If
    End If

    
    errors = ""

    Dim rg
    Set rg = New RegExp
    rg.Global = True
    rg.Pattern = "^[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}$"
    'If Not rg.Test(Request.Form("email")) Then 
    If Not rg.Test(email_from) Then 
        errors = errors & "The Email Address you entered does not appear to be valid.<br />"
    End If
    rg.Pattern = "^[A-Za-z .'-]+$"
    'If Not rg.Test(Request.Form("first_name")) Then 
    If Not rg.Test(first_name) Then 
        errors = errors & "The First Name you entered does not appear to be valid.<br />"
    End If
    'If Not rg.Test(Request.Form("last_name")) Then 
    'Response.Write "last_name = " & last_name
    If FALSE AND Not rg.Test(last_name) Then   ' I FALSE-d this out because it didn't seem to be working properly on Last Names with apostrophes in them.
        errors = errors & "The Last Name you entered does not appear to be valid.<br />"
    End If
    'If Len(comments) < 2 Then
    '    errors = errors & "The Comments you entered do not appear to be valid.<br />"
    'End If
    If Not errors = "" Then
        'Call Died(errors)
    End If


    Function CleanString(str)
        Dim bad(5)
        bad(0) = "content-type"
        bad(1) = "bcc:"
        bad(2) = "to:"
        bad(3) = "cc:"
        bad(4) = "href"
        For Each Item in bad
            str = Replace(str, Item, "")
        Next
        CleanString = str
    End Function


    If FALSE Then
        email_message = "Invoice for Star Lite ...<br /><br />"
        email_message = email_message & "First Name: " & CleanString(first_name) & "<br />"
        email_message = email_message & "Last Name: " & CleanString(last_name) & "<br />"
        email_message = email_message & "Home Address: " & CleanString(home_address) & "<br />"
        email_message = email_message & "Email: " & CleanString(email_from) & "<br />"
        email_message = email_message & "Telephone: " & CleanString(telephone) & "<br />"
        email_message = email_message & "Comments: " & CleanString(comments) & "<br />"
    End If   ' True / False

    If FALSE Then
    email_message = ""
    email_message = email_message & "<strong>Customer's Name:</strong>&nbsp;"
    email_message = email_message & UserFullName & "<br /><br />"
    email_message = email_message & "<strong>Customer's Address:</strong>&nbsp;"
    email_message = email_message & UserAddress & "<br /><br />"
    email_message = email_message & "<strong>Customer's Phone:</strong>&nbsp;"
    email_message = email_message & UserPhone & "<br /><br />"
    email_message = email_message & "<strong>Enquiry:</strong><br />"
    email_message = email_message & UserEmailEnquiryToSL & "<br /><br />"
    End If

    email_message = ""
    email_message = email_message & "<strong>Customer:</strong>&nbsp;&nbsp;"
    email_message = email_message & UserFullName & ", &nbsp;"
        email_message = email_message & UserPhone & ", &nbsp;"
    email_message = email_message & UserAddress & "<br /><br />"

    email_message = email_message & "<strong>Enquiry:</strong><br />"
    email_message = email_message & UserEmailEnquiryToSL & "<br /><br />"

    Set ObjSendMail = CreateObject("CDO.Message")
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = host
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
    ObjSendMail.Configuration.Fields.Update

    ' REMEMBER, THERE ARE TWO VERSIONS OF THIS FILE, ONE FOR THE MAIN SITE AND ONE FOR THE MOBILE SITE.
    ' KEEP THE FOLLOWING LINES IN SYNC IN BOTH THOSE FILES.

    ObjSendMail.From        = "starlite@starlite-intl.com"  ' from_address
    'ObjSendMail.Sender     = "starlite@starlite-intl.com"  ' Didn't work: Test to see if I can avoid the need for a From feild appearing to recipient.

    ObjSendMail.To          = "starlite@starlite-intl.com"  ' So starlite gets the Contact Us email.
                                                            ' ObjSendMail.To is not needed to just RECEIVE THE CONTACT US EMAIL, as long as I use .Bcc file (or .Cc field) below. 
                                                            ' But ObjSendMail.To IS NEEDED to allow THAT recipient TO REPLY. This makes sense, because of course a BCC recipient is NOT allowed to reply to the Contact Us form sender, since the latter was not even writing to the BCC recipient!
    'ObjSendMail.To          = "bn2@intelligineering.com"    ' So I get the Contact Us email.

    ' NOTE: Of course a BCC recipient (who is not also a TO recipient) is NOT allowed to reply to the Contact Us form sender, since the latter was not even writing to the BCC recipient!
    ObjSendMail.Bcc         = "bn2@intelligineering.com"    ' So I get a BCC copy of the Contact Us email.
    'ObjSendMail.Bcc         = "starlite@starlite-intl.com"
    'ObjSendMail.Bcc         = "bn2@intelligineering.com, starlite@starlite-intl.com"    ' So starlite and I both get a BCC copy of the Contact Us email.

    ObjSendMail.ReplyTo     = reply_to
    ObjSendMail.Subject     = email_subject
    ObjSendMail.HTMLBody    = email_message

    'This section sends the email
    On Error Resume Next
    ObjSendMail.Send

    If err.number <> 0 Then
        'Include your own failure message html here
        Response.write("<center><font color='blue' size='+1'>Unfortunately, the message could not be sent at this time. Please try again later.</font></center>")
    
        'Uncomment the line below to see errors with sending the message
        'Response.write("<br />Error: " & err.description)
    Else
        'Include your own success message html here
        Response.write("<center><font color='blue' size='+1'>Thank you for contacting us.<br><br>We will be in touch with you as soon as possible.</font></center>")
    End If

    set ObjSendMail = Nothing

End If   ' True / False

Response.End
%>


<% ' *********************************************************************************** 
   ' [8/11/15, BN] Send invoice email to the customer ...
   ' [11/17/17, BN] No longer used in this file.
%>

<!-- COLLAPSE THIS. IT'S NOT USED ... -->
<%
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
If FALSE Then		
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
%>  

<% 
If TRUE Then
    email_to = Session("CustomerEmail")     '  "bn2@intelligineering.com"    'Enter the email address you want to send to.
    comments = Session("InvoiceForCustomer") 

    If FALSE Then
        email_message = "Invoice for Customer ...<br /><br />"
        email_message = email_message & "First Name: " & CleanString(first_name) & "<br />"
        email_message = email_message & "Last Name: " & CleanString(last_name) & "<br />"
        email_message = email_message & "Home Address: " & CleanString(home_address) & "<br />"
        email_message = email_message & "Email: " & CleanString(email_from) & "<br />"
        email_message = email_message & "Telephone: " & CleanString(telephone) & "<br />"
        email_message = email_message & "Comments: " & CleanString(comments) & "<br />"
    End If   ' True / False


    email_message = comments & "<br />"


    Set ObjSendMail = CreateObject("CDO.Message")
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = host
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
    ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
    ObjSendMail.Configuration.Fields.Update

    ObjSendMail.From        = from_address
    ObjSendMail.To          = email_to                              ' So Customer gets the invoice for Customer.
    ObjSendMail.Bcc         = "bn2@intelligineering.com; starlite@starlite-intl.com"        ' So starlite and I both get a BCC copy of the invoice for Customer.
    ObjSendMail.Subject     = email_subject
    ObjSendMail.HTMLBody    = email_message


    'This section sends the email
    On Error Resume Next
    ObjSendMail.Send

    If err.number <> 0 Then
        'Include your own failure message html here
        Response.write("<br />Unfortunately, the message could not be sent at this time. Please try again later.")
    
        'Uncomment the line below to see errors with sending the message
        Response.write("<br />Error: " & err.description)
    Else
        'Include your own success message html here
        Response.write("<br />Thank you for contacting us. We will be in touch with you very soon.")
    End If
    'Response.End

    set ObjSendMail = Nothing

End If
%>


<%
PaymentMethod = Request.QueryString("PaymentMethod") 
'Response.Write "<br><br><br><br>PaymentMethod = " & PaymentMethod 
'Response.End
'Response.Redirect "EmailShow.asp?PaymentMethod=" & PaymentMethod 


If PaymentMethod = "CreditCard" Then
	'Response.Redirect "cashoutAuthorizeNet.asp"                  ' 7/16/17: OLD gateway.
    Response.Redirect "cashoutUSAePay.asp"   ' 7/5/2017 = 170705  ' 7/16/17: NEW gateway.
ElseIf PaymentMethod = "PayPal" OR PaymentMethod = "NonUSorCanadianCustomerPayPal" Then
	Response.Redirect "cashoutAtPayPal.asp"
Else
	Response.Redirect "EmailShow.asp?PaymentMethod=" & PaymentMethod 
End If
%>


</body>

</html>