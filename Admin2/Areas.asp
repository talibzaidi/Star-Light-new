<%@ Language=VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<!--#include file="../ADOVBS.INC"-->


<HTML>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="/Misc/StyleSheet1.css" type="text/css">
</HEAD>


<BODY topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<% InSection = "Areas" %>

<!--#include virtual="Admin2/Misc/AdminHeader.asp"-->
<!--#include virtual="Misc/Header.INC"-->
<!--#include virtual="Admin2/Misc/LoginCheck.asp"-->


<%
sortField = Request("sortField")
'Response.Write "<br>You just chose sortField = " & sortField 
sortDirection = Request("sortDirection")
'Response.Write "<br>You just chose sortDirection = " & sortDirection 

If sortField = "" Then
	OrderBySplice = "ORDER BY AID "
Else
	OrderBySplice = "ORDER BY " & sortField & " "
End If

%>

<center>
<form method='Post'>
Sort by: &nbsp;&nbsp;
<select name='sortField'>
	<option>AID
	<option>AreaName
</select>

<select name='sortDirection'>
	<option value='ASC'>Ascending
	<option value='DESC'>Descending
</select>
&nbsp;&nbsp;
<input type='submit' value='sort'>
</form>
</center>

<%
Dim Conn, rsAreas, AreasSQL
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Session("ConnectionString")
' The WHERE Not IsNull(AID) is needed in case, for some reason, a record "is deleted" (?!)
AreasSQL = "SELECT * FROM Area51 WHERE Not IsNull(AID) " & OrderBySplice & sortDirection
'Response.Write "AreasSQL = " & AreasSQL & "<br>"
'Response.End
Set rsAreas = Server.CreateObject("ADODB.Recordset")
rsAreas.Open AreasSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 


rsAreas.Movefirst 
Response.Write "<table align=center border=1 width='90%' >"
Response.Write "<tr bgcolor=lightblue><td colspan=7 align='center'><b>Areas Table</b></td></tr>"
Response.Write "<tr bgcolor=lightblue><td>i</td><td>AID</td><td width=150>AreaName</td><td>AreaVisible</td>" & _
				"<td>AreaDesc</td></tr>"
row = 0
Do While NOT rsAreas.EOF
	row = row + 1
	AID = rsAreas("AID")
	AreaName = rsAreas("AreaName")
	AreaVisible = rsAreas("AreaVisible")
	AreaDesc = rsAreas("AreaDesc")
	If IsNull(AreaDesc) OR AreaDesc = "" Then
		AreaDesc = "&nbsp;"
	End If
	
	Response.Write "<tr>"
	Response.Write "<td valign=top>" & row & "</td>"
	Response.Write "<td valign=top>" & AID & "</td>"
	Response.Write "<td valign=top>" & AreaName & "</td>"
	Response.Write "<td valign=top>" & AreaVisible & "</td>"
	Response.Write "<td valign=top>" & AreaDesc & "</td>"
	Response.Write "</tr>"
	rsAreas.moveNext
Loop
Response.Write  "</table>"

rsAreas.Close
Set rsAreas = Nothing
Set conn = Nothing
%>

<br><br>

</BODY>

</HTML>
