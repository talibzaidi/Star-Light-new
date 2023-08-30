<%@ Language=VBScript %>

<%
Pwd = Request("Pwd")
btnSubmit = Request("btnSubmit")
'Response.Write "<br>Pwd = " & Pwd
'Response.Write "<br>btnSubmit = " & btnSubmit
If btnSubmit = "Submit" AND Pwd = "zeevi" Then   ' Successful login
	Session("LoggedIn") = "Yes"
	Response.Redirect "http://www.starlite-intl.com/Admin2/DisplayProducts.asp"
Else
	Session("LoggedIn") = "No"
End If
%>

<HTML>


<% 
Session("LoggedIn") = "No"
Response.Redirect "http://www.starlite-intl.com" 
%>

