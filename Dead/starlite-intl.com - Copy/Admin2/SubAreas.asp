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

<% InSection = "SubAreas" %>

<!--#include virtual="Admin2/Misc/AdminHeader.asp"-->
<!--#include virtual="Misc/Header.INC"-->
<!--#include virtual="Admin2/Misc/LoginCheck.asp"-->


<%
sortField = Request("sortField")
'Response.Write "<br>You just chose sortField = " & sortField 
sortDirection = Request("sortDirection")
'Response.Write "<br>You just chose sortDirection = " & sortDirection 

If sortField = "" Then
	OrderBySplice = "ORDER BY SID "
Else
	OrderBySplice = "ORDER BY " & sortField & " "
End If

%>

<center>
<form method='Post'>
Sort by: &nbsp;&nbsp;
<select name='sortField'>
	<option>SID
	<option>AID
	<option>SubName
	<option>Warranties
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
Dim Conn, rsSubAreas, SubAreasSQL
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Session("ConnectionString")
' The WHERE Not IsNull(SID) is needed in case, for some reason, a record "is deleted" (?!)
SubAreasSQL = "SELECT * FROM SubArea  WHERE Not IsNull(SID) " & OrderBySplice & sortDirection
'Response.Write "SubAreasSQL = " & SubAreasSQL & "<br>"
'Response.End
Set rsSubAreas = Server.CreateObject("ADODB.Recordset")
rsSubAreas.Open SubAreasSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 

rsSubAreas.Movefirst 
Response.Write "<table align=center border=1 width='90%' >"
Response.Write "<tr bgcolor=lightblue><td colspan=7 align='center'><b>SubAreas Table</b></td></tr>"
Response.Write "<tr bgcolor=lightblue><td>i</td><td>SID</td><td>AID</td><td width=300>SubName</td>" & _
				"<td>SubVisible</td><td>SubDesc</td><td width=120>Warranties</td></tr>"
row = 0
Do While NOT rsSubAreas.EOF
	row = row + 1
	SID = rsSubAreas("SID")
	AID = rsSubAreas("AID")
	SubName = rsSubAreas("SubName")
	SubVisible = rsSubAreas("SubVisible")
	SubDesc = rsSubAreas("SubDesc")
	If IsNull(SubDesc) OR SubDesc = "" Then
		SubDesc = "&nbsp;"
	End If
	Warranties = rsSubAreas("Warranties")
	If IsNull(Warranties) OR Warranties = "" Then
		Warranties = "&nbsp;"
	End If
	
	Response.Write "<tr>"
	Response.Write "<td valign=top>" & row & "</td>"
	Response.Write "<td valign=top>" & SID & "</td>"
	Response.Write "<td valign=top>" & AID & "</td>"
	Response.Write "<td valign=top>" & SubName & "</td>"
	Response.Write "<td valign=top>" & SubVisible & "</td>"
	Response.Write "<td valign=top>" & SubDesc & "</td>"
	Response.Write "<td valign=top>" & Warranties & "</td>"
	Response.Write "</tr>"
	rsSubAreas.moveNext
Loop
Response.Write  "</table>"

rsSubAreas.Close
Set rsSubAreas = Nothing
Set conn = Nothing
%>

<br><br>

</BODY>

</HTML>
