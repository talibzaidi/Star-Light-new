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

<% InSection = "Products" %>

<!--#include virtual="Admin2/Misc/AdminHeader.asp"-->
<!--#include virtual="Misc/Header.INC"-->
<!--#include virtual="Admin2/Misc/LoginCheck.asp"-->


<%
sortField = Request("sortField")
'Response.Write "<br>You just chose sortField = " & sortField 
sortDirection = Request("sortDirection")
'Response.Write "<br>You just chose sortDirection = " & sortDirection 

If sortField = "" Then
	OrderBySplice = "ORDER BY PID "
Else
	OrderBySplice = "ORDER BY " & sortField & " "
End If

%>

<center>
<form method='Post'>
Sort by: &nbsp;&nbsp;
<select name='sortField'>
	<option>PID
	<option>ItemID
	<option>PName
	<option>MSL
	<option>Cost
	<option>Duty
	<option>GPM
	<option value='Manufa'>Manufacturer
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
Dim Conn, rsProduct, ProductSQL
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Session("ConnectionString")
' The WHERE Not IsNull(PID) is needed in case, for some reason, a record "is deleted" (?!)
ProductSQL = "SELECT * FROM Product WHERE Not IsNull(PID) " & OrderBySplice & sortDirection
'Response.Write "ProductSQL = " & ProductSQL & "<br>"
'Response.End
Set rsProduct = Server.CreateObject("ADODB.Recordset")
rsProduct.Open ProductSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 

rsProduct.moveFirst
Response.Write "<table align=center border=1>"
Response.Write "<tr bgcolor=lightblue><td colspan=10 align='center'><b>Products Table</b></td></tr>"
Response.Write "<tr bgcolor=lightblue><td>i</td><td>PID</td><td>ItemID</td><td>PName</td><td>Old MSL</td><td>MSL</td><td>Cost</td><td>Duty</td><td>GPM</td><td>Manufacturer</td></tr>"
row = 0
Do While NOT rsProduct.EOF
	row = row + 1
	PID = rsProduct("PID")
	ItemID = rsProduct("ItemID")
	ProductName = rsProduct("PName")
	Manufacturer = rsProduct("Manufa")
	Description = rsProduct("Descr")
	MSLOld = rsProduct("MSLOld")
	'MSLNew = MSLOld * 1.13
	MSL = rsProduct("MSL")
	Cost = rsProduct("Cost")
	Duty = rsProduct("Duty")
	GPM = rsProduct("GPM")
	Manufa = rsProduct("Manufa")
	
	Response.Write "<tr>"
	Response.Write "<td valign=top>" & row & "</td>"
	Response.Write "<td valign=top>" & PID & "</td>"
	Response.Write "<td valign=top>" & ItemID & "</td>"
	Response.Write "<td valign=top>" & ProductName & "</td>"
	Response.Write "<td valign=top>" & MSLOld & "</td>"
	Response.Write "<td valign=top>" & MSL & "</td>"
	Response.Write "<td valign=top>" & Cost & "</td>"
	Response.Write "<td valign=top>" & Duty & "</td>"
	Response.Write "<td valign=top>" & GPM & "</td>"
	Response.Write "<td valign=top>" & Manufa & "</td>"
	Response.Write "</tr>"
	rsProduct.moveNext
Loop
Response.Write  "</table>"

rsProduct.Close
Set rsProduct = Nothing
Set conn = Nothing
%>

</BODY>

</HTML>
