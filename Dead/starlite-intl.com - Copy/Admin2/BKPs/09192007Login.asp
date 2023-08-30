<%@ Language=VBScript %>

<%
Pwd = Request("Pwd")
btnSubmit = Request("btnSubmit")
'Response.Write "<br>Pwd = " & Pwd
'Response.Write "<br>btnSubmit = " & btnSubmit
If btnSubmit = "Submit" AND Pwd = "zeevi" Then   ' Successful login
	Session("LoggedIn") = "Yes"
	Response.Redirect "http://www.starlite-intl.com/Admin2/Orders.asp"
ElseIf btnSubmit = "Submit" Then 
	Session("LoggedIn") = "No"
Else  ' Got to this page without pressing Submit button to submit a password.
	' Leave Session("LoggedIn") as it was.
End If
%>

<HTML>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="/Misc/StyleSheet1.css" type="text/css">
</HEAD>


<BODY topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<% InSection = "Login" %>

<!--#include virtual="Admin2/Misc/AdminHeader.asp"-->
<!--#include virtual="Misc/Header.INC"-->

<br><br><br><br>
<form>
<table align=center cellpadding=10>
<tr>
	<td>
	<b>Login Password:</b>
	</td>
	<td>
	<input type=txt name=pwd></input>
	</td>
	<td>
	<input type=submit name="btnSubmit" value="Submit"></input>
	</td>
</tr>
</table>
</form>

<% If btnSubmit = "Submit" Then %>
<center><font size=4 face=Tahoma color=red>Invalid Password. Try Again.</font></center>
<% End If %>

</BODY>

</HTML>
