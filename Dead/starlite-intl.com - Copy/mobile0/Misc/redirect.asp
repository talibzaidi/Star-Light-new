<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">


<% 
target = request.querystring("target")
Response.Write "target = " & target

If target = "Home" Then
	Response.redirect "http://starlite-intl.com/mobile"
ElseIf target = "Terms" Then
	Response.redirect "https://www.starlite-intl.com/mobile/Misc2/Terms_and_Conditions.asp"
ElseIf target = "Contact Us" Then
	Response.redirect "https://www.starlite-intl.com/mobile/Misc2/contact.asp"
End If

Response.End
%>


<head>
    <title>Redirect</title>
</head>


<body>
</body>


</html>
