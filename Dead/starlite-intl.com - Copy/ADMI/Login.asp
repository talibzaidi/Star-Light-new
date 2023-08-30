<%@ LANGUAGE = VBScript %> 
<!--#include file="ADOVBS.INC"-->
<%
msg=""

Action = Left(UCase(Request("Action")),6)
mSubmitted = date & " " & time


If Action = "SUBMIT" Then
	If Request("CoName") = "" OR _
		Request("Password") = "" Then
		msg="Invalid."
	End If

	If msg = "" Then
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	SQL = "SELECT Access99 from ADMI WHERE  User = '"& Request("CoName") &"' AND Pass = '"& Request("Password") &"' " 
	Set RS = Conn.Execute(SQL)
	Session("Access") = RS("Access99")
	If Session("Access") > 0 Then
		Response.Redirect "sanction.asp"
        End If

End If  'msg = ""
End If  'Action = "Submit"
%>
<html>

<head>

<title>Sanction - Version (Orange) - 4HRSMN</title>
</head>

<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0" link="#FFFFFF" vlink="#FFFFFF">

<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td bgcolor="#FFBD00"><font face="Arial"><img
        src="Simages/sanction.gif"
        width="330" height="82"></font></td>
        <td bgcolor="#FFBD00"><font face="Arial"></font>&nbsp;</td>
        <td align="right" bgcolor="#FFBD00"><font face="Arial" color="#FFBD00">No back doors, sorry. =(</font></td>
    </tr>
    <tr>
        <td><font face="Arial"><img
        src="Simages/btcurve.gif"
        width="330" height="82"></font></td>
        
        <td colspan="2"><font size="2" face="Arial">Welcome to Sanction: The
        Active web tool for mission critical administration.</font></td>
    </tr>
    <tr>
        <td><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td align="center"><font face="Arial"></font><form method="POST" action = "login.asp">
                    <p><input type="text" size="20" name="CoName" value="User<%=Request("CoName")%>"></p>
                    <p><input type="Password" size="20" name="Password" value="<%=Request("Password")%>"></p>
                    <p><INPUT TYPE=SUBMIT NAME="Action" VALUE="Submit"></p><font
                color="#FFFFFF"><%=msg%></font>
                </form></td>
        <td><font face="Arial"></font>&nbsp;</td>
    </tr>
</table>
</body>
</html>
