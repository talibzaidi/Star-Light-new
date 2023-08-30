<%@ Language=VBScript %>



<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<!--#include file="../ADOVBS.INC"-->


<%
Btn 				= Trim(Request.QueryString("Btn"))
OrderID 			= Trim(Request.QueryString("OrderID"))
FName 				= Trim(Request.QueryString("FName"))
Email 				= Trim(Request.QueryString("Email"))
OptInToEmailings 	= Trim(Request.QueryString("OptInToEmailings"))
WeEmail 			= Trim(Request.QueryString("WeEmail"))
Phone 				= Trim(Request.QueryString("Phone"))
Address 			= Trim(Request.QueryString("Address"))
City 				= Trim(Request.QueryString("City"))
State 				= Trim(Request.QueryString("State"))
ZIP 				= Trim(Request.QueryString("ZIP"))

If FALSE Then
	Response.Write "<br>Btn: " 		& Btn
	Response.Write "<br>OrderID: " 	& OrderID 
	Response.Write "<br>FName: " 	& FName 
	Response.Write "<br>Email: " 	& Email 
	Response.Write "<br>OptIn: " 	& OptInToEmailings
	Response.Write "<br>ON: " 		& WeEmail
	Response.Write "<br>Phone: " 	& Phone
	Response.Write "<br>Address : " & Address 
	Response.Write "<br>City: " 	& City
	Response.Write "<br>State: " 	& State 
	Response.Write "<br>ZIP: " 		& ZIP 
End If

If Btn = "Delete" Then
	SQL = "DELETE FROM Orders WHERE OrderID = " & OrderID
	Response.Write "<br>SQL = " & SQL
	
	Set Conn = Server.CreateObject("ADODB.Connection")
  	Conn.Open Session("ConnectionString")
   	Conn.Execute(SQL)
	Response.Redirect "Orders.asp" 
	
ElseIf Btn = "Edit" Then	
	OrderSQL = "SELECT * FROM Orders WHERE OrderID = " & OrderID 
	'Response.Write "<br>OrderSQL = " & OrderSQL & "<br>"
	
	Set Conn = Server.CreateObject("ADODB.Connection")
  	Conn.Open Session("ConnectionString")
	Set rsOrder = Server.CreateObject("ADODB.Recordset")
	rsOrder.Open OrderSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 
	'Response.End
	
	' Update record
	rsOrder("FName") 			= FName
	rsOrder("Email") 			= Email
	rsOrder("OptInToEmailings") = OptInToEmailings
	rsOrder("WeEmail") 			= WeEmail
	rsOrder("Phone") 			= Phone
	rsOrder("Address") 			= Address
	rsOrder("City") 			= City
	rsOrder("State") 			= State
	rsOrder("ZIP") 				= ZIP
			
	rsOrder.Update
	rsOrder.Close
	Conn.Close
	 
End If
%>


<HTML>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 6.0">
<link rel="stylesheet" href="/Misc/StyleSheet1.css" type="text/css">
</HEAD>


<BODY topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<% InSection = "Orders" %>

<!--#include virtual="Admin2/Misc/AdminHeader.asp"-->
<!--#include virtual="Misc/Header.INC"-->
<!--#include virtual="Admin2/Misc/LoginCheck.asp"-->


<%
sortField = Request("sortField")
'Response.Write "<br>You just chose sortField = " & sortField 
sortDirection = Request("sortDirection")
'Response.Write "<br>You just chose sortDirection = " & sortDirection 

If sortField = "" Then
	OrderBySplice = "ORDER BY OrderDate DESC "
Else
	OrderBySplice = "ORDER BY " & sortField & " " & sortDirection 
End If
%>

<br>
<center>
<form method='Post'>
Sort by: &nbsp;&nbsp;
<select name='sortField'>
	<option>OrderID
	<option>OrderDate
	<option value='LName'>Last Name
	<option value='FName'>First Name
	<option>Email
	<option>State
	<option>Country
	<option>SubTotal
	<option>GrandTotal
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
Dim Conn, rsOrders, OrdersSQL
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Session("ConnectionString")
' The WHERE Not IsNull(OrderID) is needed because, for some reason, the "first" record of Orders "is deleted" (?!)
OrdersSQL = "SELECT * FROM Orders WHERE Not IsNull(OrderID) " & OrderBySplice ' & sortDirection
'Response.Write "OrdersSQL = " & OrdersSQL & "<br>"
'Response.End
Set rsOrders = Server.CreateObject("ADODB.Recordset")
rsOrders.Open OrdersSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 

rsOrders.Movefirst ' move 1,1		' [BN, 6/27/07] This moves to the second record. Needed (for now) because first record of Orders "is deleted" (!?)
Response.Write "<table align=center border=1 width='98%' >"
Response.Write "<tr bgcolor=lightblue><td colspan=22 align='left'><b>Orders Table</b></td></tr>"
Response.Write "<tr bgcolor=lightblue><td></td><td></td><td>i</td><td>OrderID</td><td>OrderDate</td><td>FName</td><td>LName</td>" & _
				"<td>Email</td><td>OptIn</td><td>ON</td><td>Phone</td><td>Address</td><td>City</td><td>State</td>" & _
				"<td>Far</td><td>ZIP</td><td>Country</td><td>PaymentMethod</td><td>SubTotal</td>" & _
				"<td>S&H</td><td>Tax</td><td>GrandTotal</td></tr>"
row = 0
Do While NOT rsOrders.EOF
	row = row + 1
	OrderID = rsOrders("OrderID")
	OrderDate = rsOrders("OrderDate")
	FName = rsOrders("FName")
	LName = rsOrders("LName")
	Email = rsOrders("Email")
	OptInToEmailings = rsOrders("OptInToEmailings")
	WeEmail = rsOrders("WeEmail")
	Phone = rsOrders("Phone")
	Address = rsOrders("Address")
	City = rsOrders("City")
	State = rsOrders("State")
	BigShipment = rsOrders("BigShipment")
	ZIP = rsOrders("ZIP")
	Country = rsOrders("Country")
	PaymentMethod = rsOrders("PaymentMethod")
	SubTotal = rsOrders("SubTotal")
	SandH = rsOrders("SandH")
	Tax = rsOrders("Tax")
	GrandTotal = rsOrders("GrandTotal")
	
	Response.Write "<tr>"
	
	Response.Write "<td valign=top bgcolor=pink><form><input type='Submit' name='Btn' value='Delete'><input type=hidden name=OrderID value=" & OrderID & "></form></td>"
	
	Response.Write "<form>"
	Response.Write "<td valign=top>" & "<input type='Submit' name='Btn' value='Edit'>" & "</td>"
	Response.Write "<td valign=top>" & row & "</td>"
	Response.Write "<td valign=top>" & OrderID & "</td>"
	Response.Write "<td valign=top>" & OrderDate & "</td>"
	'Response.Write "<td valign=top>" & FName & "</td>"
	Response.Write "<td valign=top><input type=text name=FName value='" & FName & "' size=10 maxlength=50></input></td>"
	Response.Write "<td valign=top>" & LName & "</td>"
	Response.Write "<td valign=top><input type=text name=Email value='" & Email & "' size=30 maxlength=50></input></td>"
	'Response.Write "<td valign=top>" & OptInToEmailings & "</td>"
	Response.Write "<td valign=top><input type=text name=OptInToEmailings value='" & OptInToEmailings & "' size=2 maxlength=5></input></td>"
	Response.Write "<td valign=top><input type=text name=WeEmail value='" & WeEmail & "' size=2 maxlength=5></input></td>"
	Response.Write "<td valign=top><input type=text name=Phone value='" & Phone & "' size=20 maxlength=50></input></td>"
	'Response.Write "<td valign=top>" & Address & "</td>"
	Response.Write "<td valign=top><input type=text name=Address value='" & Address & "' size=30 maxlength=50></input></td>"
	'Response.Write "<td valign=top>" & City & "</td>"
	Response.Write "<td valign=top><input type=text name=City value='" & City & "' size=15 maxlength=50></input></td>"
	Response.Write "<td valign=top><input type=text name=State value='" & State & "' size=10 maxlength=50></input></td>"
	Response.Write "<td valign=top>" & BigShipment & "</td>"
	Response.Write "<td valign=top><input type=text name=ZIP value='" & ZIP & "' size=7 maxlength=15></input></td>"
	Response.Write "<td valign=top>" & Country & "</td>"
	Response.Write "<td valign=top>" & PaymentMethod & "</td>"
	Response.Write "<td valign=top>" & SubTotal & "</td>"
	Response.Write "<td valign=top>" & SandH & "</td>"
	Response.Write "<td valign=top>" & Tax & "</td>"
	Response.Write "<td valign=top>" & GrandTotal & "</td>"
	Response.Write ""
	Response.Write "<input type=hidden name=OrderID value='" & OrderID & "'>"
	Response.Write "</form>"
	
	Response.Write "</tr>"
	
	rsOrders.moveNext
Loop
Response.Write  "</table>"

Response.Write  ""

rsOrders.Close
Set rsOrders = Nothing
Set conn = Nothing
%>

<br><br>


</BODY>

</HTML>

