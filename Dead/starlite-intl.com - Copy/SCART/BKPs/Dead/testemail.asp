<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>

<head>
<title>Sans Titre</title>
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
<meta name="generator" content="HAPedit 3.0">
</head>

<body bgcolor="#FFFFFF">
<%
Dim myMail
Set myMail = Server.CreateObject("CDONTS.NewMail")
myMail.From = "NewOrder@StarliteEcommerce" 
        myMail.To = "starlite@starlite-intl.com"
         'myMail.To = "sani@wwnet.com"
        'myMail.To = "marketing@mitmazel.com"
        ' myMail.To = "bn@intelligineering.com"
        myMail.Subject = "this is a test"
        myMail.BodyFormat = 0 
        myMail.MailFormat = 0 
        myMail.Body = "This email is being sent as a test from the server.  Please let us know if you got it. IAC"
        myMail.Send 
%>
done
</body>

</html>