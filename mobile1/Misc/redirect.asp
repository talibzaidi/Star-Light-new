<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">


<% 
target = request.querystring("target")
Response.Write "target = " & target

If target = "Home" Then
	Response.redirect "http://www.starlite-intl.com/mobile1/index.asp"
ElseIf target = "Terms" Then
	Response.redirect "http://starlite-intl.com/mobile1/Misc2/Terms_and_Conditions.asp"
ElseIf target = "Contact Us" Then
	Response.redirect "http://www.starlite-intl.com/mobile1/Misc2/contact.asp"
ElseIf target = "OEM GPS Sensors" Then
	Response.redirect "http://www.starlite-intl.com/mobile1/Search/SearchSummary.asp?AID=45&SID=173"
ElseIf target = "Night Vision Optics" Then
	Response.redirect "http://www.starlite-intl.com/mobile1/Search/SearchSummary.asp?AID=53&SID=297"
	ElseIf target = "Communications" Then
	Response.redirect "http://www.starlite-intl.com/mobile1/Search/SearchSummary.asp?AID=44"
ElseIf target = "Shopping Cart" Then
	Response.redirect "http://www.starlite-intl.com/mobile1/scart/scart.asp?action=viewcart&pid=0&sid=11"
End If

Response.End
%>


<head>
    <title>Redirect</title>
</head>


<body>
</body>


</html>
