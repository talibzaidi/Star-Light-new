<!DOCTYPE html>

<!-- 
    [BN, 10/23/15] This file works with sibling file contactform.asp

    [BN, 8/11/15] Test of using new email system on NEW server at HostMySite.com, based on Knowledgebase article at
    https://solutions.hostmysite.com/index.php?/Knowledgebase/Article/View/8596/0/Using-CDOSys-to-create-an-ASP-Mail-form-that-uses-Authentication
    See also http://www.w3schools.com/asp/asp_send_email.asp where it says:
    "CDOSYS is a built-in component in ASP. This component is used to send e-mails with ASP." It replaces CDONTs.
-->

<html xmlns="http://www.w3.org/1999/xhtml">

<head>
    <title></title>
</head>

<body>

    <%
    Dim email_to, email_subject, host, username, password, reply_to, port, from_address
    Dim first_name, last_name, home_address, email_from, telephone, comments, error_message
    Dim ObjSendMail, email_message

    'email_to = "name@yourdomain.xyz"           'Enter the email you want to send the form to
    email_to = "bn2@intelligineering.com"       'Enter the email you want to send the form to
    email_subject = "This is a test"            'You can put whatever subject here
    'host = "mail.yourdomain.xyz"               'The mail server name. (Commonly mail.yourdomain.xyz if your mail is hosted with HostMySite) 
    'host = "mail.starlite-intl.com" 
     host = "win-mail13.hostmanagement.net"            'The mail server name. (Commonly mail.yourdomain.xyz if your mail is hosted with HostMySite) 
    'username = "name@yourdomain.xyz"           'A valid email address you have setup 
    username = "starlite@starlite-intl.com"     'A valid email address you have setup 
    'from_address = "name@yourdomain.xyz"       'If your mail is hosted with HostMySite this has to match the email address above 
    from_address = "starlite@starlite-intl.com" 'If your mail is hosted with HostMySite this has to match the email address above 
    'password = "your_password"                 'Password for the above email address
    'password = "S6a2n2I6"
    'password = "$621zBn7!"                     'Password for the above email address
    'password = "x718Bnr#?_s"                   'Changed Jan.26 2023
     password = "T6#_S?ByzP"
    'reply_to = "name@yourdomain.xyz"           'Enter the email you want customers to reply to
    reply_to = "starlite@starlite-intl.com"     'Enter the email you want customers to reply to
    port = "25"                                 'This is the default port. Try port 50 if this port gives you issues and your mail is hosted with HostMySite

    Sub Died(errors)
        'Your error code can go here 
        Response.write("We are very sorry, but there were error(s) found with the form you submitted. These errors appear below.<br /><br />")
        Response.write(errors & "<br /><br />")
        Response.write("Please go back and fix these errors.<br /><br />")
        Response.End
    End Sub

    'Validate expected data exists
    If Request.Form("first_name") = "" Or Request.Form("last_name") = ""  Or Request.Form("email") = ""  Or Request.Form("comments") = "" Then
        Call Died("Required field/s have not been entered")
    End If

    first_name = Request.Form("first_name")  'required 
    last_name = Request.Form("last_name")  'required 
    home_address = Request.Form("home_address")  'not required
    email_from = Request.Form("email")  'required 
    telephone = Request.Form("telephone")  'not required
    comments = Request.Form("comments")  'required 

    Dim rg
    Set rg = New RegExp
    rg.Global = True

    rg.Pattern = "^[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}$"
    If Not rg.Test(Request.Form("email")) Then 
        error_message = error_message & "The Email Address you entered does not appear to be valid.<br />"
    End If

    rg.Pattern = "^[A-Za-z .'-]+$"
    If Not rg.Test(Request.Form("first_name")) Then 
        error_message = error_message & "The First Name you entered does not appear to be valid.<br />"
    End If

    If Not rg.Test(Request.Form("last_name")) Then 
        error_message = error_message & "The Last Name you entered does not appear to be valid.<br />"
    End If

    If Len(comments) < 2 Then
        error_message = error_message & "The Comments you entered do not appear to be valid.<br />"
    End If

    If Not error_message = "" Then
        Call Died(error_message)
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

    email_message = "Form details below.<br /><br />"
    email_message = email_message & "First Name: " & CleanString(first_name) & "<br />"
    email_message = email_message & "Last Name: " & CleanString(last_name) & "<br />"
    email_message = email_message & "Home Address: " & CleanString(home_address) & "<br />"
    email_message = email_message & "Email: " & CleanString(email_from) & "<br />"
    email_message = email_message & "Telephone: " & CleanString(telephone) & "<br />"
    email_message = email_message & "Comments: " & CleanString(comments) & "<br />"

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

    ObjSendMail.To = email_to
    ObjSendMail.Subject = email_subject
    ObjSendMail.From = from_address

    ObjSendMail.HTMLBody = email_message

    'This section sends the email
    On Error Resume Next
    ObjSendMail.Send

    If err.number <> 0 Then
        'Include your own failure message html here
        Response.write("Unfortunately, the message could not be sent at this time. Please try again later.")
    
        'Uncomment the line below to see errors with sending the message
        'Response.write("<br />Error: " & err.description)
    Else
        'Include your own success message html here
        Response.write("Thank you for contacting us. We will be in touch with you very soon.")
    End If

    set ObjSendMail = Nothing

    %>

    </body>
</html>