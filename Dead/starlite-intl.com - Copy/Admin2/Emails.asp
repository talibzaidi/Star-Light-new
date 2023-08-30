<%@ Language=VBScript %>  


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<!--#include file="../ADOVBS.INC"-->


<%
Btn 				= Trim(Request.Form("Btn"))
EmailID 			= Trim(Request.Form("EmailID"))
EmailSubject 		= Trim(Request.Form("EmailSubject"))
EmailDescription	= Trim(Request.Form("EmailDescription"))
EmailHeader			= Trim(Request.Form("EmailHeader"))
EmailBody 			= Trim(Request.Form("EmailBody"))
EmailFooter 		= Trim(Request.Form("EmailFooter"))
DateAdded 			= Trim(Request.Form("DateAdded"))

If FALSE Then
	Response.Write "<br>Btn: " 				& Btn
	Response.Write "<br>EmailID: " 			& EmailID
	Response.Write "<br>EmailSubject: " 	& EmailSubject 
	Response.Write "<br>EmailDescription: " & EmailDescription
	Response.Write "<br>EmailHeader: " 		& EmailHeader
	Response.Write "<br>EmailBody: " 		& EmailBody 
	Response.Write "<br>EmailFooter: " 		& EmailFooter
	Response.Write "<br>DateAdded: " 		& DateAdded 
End If

If Btn = "Delete" Then
	SQL = "DELETE FROM Emails WHERE EmailID = " & EmailID
	Response.Write "<br><br>SQL = " & SQL
	
	Set Conn = Server.CreateObject("ADODB.Connection")
  	Conn.Open Session("ConnectionString")
   	Conn.Execute(SQL)
	Response.Redirect "Emails.asp" 
	
ElseIf Btn = "Insert New Email" Then
	EmailSQL = "SELECT * FROM Emails ORDER BY EmailID DESC" 
	'Response.Write "<br><br>EmailSQL = " & EmailSQL & "<br>"
	
	Set Conn = Server.CreateObject("ADODB.Connection")
  	Conn.Open Session("ConnectionString")
	Set rsEmail = Server.CreateObject("ADODB.Recordset")
	rsEmail.Open EmailSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 
	
	MaxEmailID = rsEmail("EmailID")    ' This is the max value of EmailID because we sorted recordset rsEmail by EmailID DESC
	NewEmailID = MaxEmailID + 1
	'Response.Write "<br>MaxEmailID = " & MaxEmailID
	'Response.End
	
	rsEmail.AddNew()
	' Update / initialize new blank record
	rsEmail("EmailID") 				= NewEmailID
	rsEmail("EmailSubject") 		= "Edit this Email Subject"
	rsEmail("EmailDescription") 	= "Edit this Email Description"
	rsEmail("EmailHeader") 			= "Edit this Email Header"
	rsEmail("EmailBody") 			= "Edit this Email Body" 
	rsEmail("EmailFooter") 			= "Edit this Email Footer"
	rsEmail("DateAdded") 			= Now()
			
	rsEmail.Update
	rsEmail.Close
	Conn.Close
	
	'rsEmail.AddNew 
	Response.Redirect "Emails.asp" 
	
ElseIf Btn = "Edit" Then	
	EmailSQL = "SELECT * FROM Emails WHERE EmailID = " & EmailID
	'Response.Write "<br><br>EmailSQL = " & EmailSQL & "<br>"
	
	Set Conn = Server.CreateObject("ADODB.Connection")
  	Conn.Open Session("ConnectionString")
	Set rsEmail = Server.CreateObject("ADODB.Recordset")
	rsEmail.Open EmailSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 
	'Response.End
	
	' Update record
	rsEmail("EmailSubject") 		= EmailSubject
	rsEmail("EmailDescription") 	= EmailDescription
	rsEmail("EmailHeader") 			= EmailHeader
	rsEmail("EmailBody") 			= EmailBody 
	rsEmail("EmailFooter") 			= EmailFooter
	rsEmail("DateAdded") 			= DateAdded
			
	rsEmail.Update
	rsEmail.Close
	Conn.Close
	 
End If
%>


<HTML>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<link rel="stylesheet" href="/Misc/StyleSheet1.css" type="text/css">
</HEAD>


<BODY topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<% InSection = "Emails" %>

<!--#include virtual="Admin2/Misc/AdminHeader.asp"-->
<!--#include virtual="Misc/Header.INC"-->
<!--#include virtual="Admin2/Misc/LoginCheck.asp"-->


<%
sortField = Request("sortField")
'Response.Write "<br>You just chose sortField = " & sortField 
sortDirection = Request("sortDirection")
'Response.Write "<br>You just chose sortDirection = " & sortDirection 
If sortDirection = "" Then
	sortDirection = "DESC "
End If

If sortField = "" Then
	OrderBySplice = "ORDER BY EmailID "
Else
	OrderBySplice = "ORDER BY " & sortField & " "
End If

%>

<form method='Post'>
<table align=center border=0 width=100%>
<tr>
	<td width=150>
	<input type='Submit' name='Btn' value='Insert New Email'>
	</td>
	
	<td align=center>
Sort by: &nbsp;&nbsp;
<select name='sortField'>
	<option>EmailID
	<option>SortOrder
	<option>EmailType
	<option>EmailSubject
	<option>EmailDescription
	<option>BCC
	<option>DateAdded
</select>

<select name='sortDirection'>
	<option value='DESC'>Descending
	<option value='ASC'>Ascending
</select>
&nbsp;&nbsp;
<input type='submit' value='sort'>
	</td>

	<td width=150>
	</td>	
</tr>
</table>
</form>


<% 
Dim Conn, rsEmails, EmailsSQL 
Set Conn = Server.CreateObject("ADODB.Connection")
Response.Write "Session('ConnectionString') = " & Session("ConnectionString") & "<br><br>"
'Conn.Open Session("ConnectionString")  ' Opted-In Customer DB.
EmailsSQL = "SELECT * FROM Emails WHERE Not IsNull(EmailID) " & OrderBySplice & sortDirection

Response.Write "Session('ConnectionString2') = " & Session("ConnectionString2")
Conn.Open Session("ConnectionString2")  ' Purchased DB.
' The WHERE Not IsNull(EmailID) is needed in case, for some reason, a record "is deleted" (?!)
EmailsSQL = "SELECT * FROM Emails"
'Response.Write "EmailsSQL = " & EmailsSQL & "<br>"
'Response.End

Set rsEmails = Server.CreateObject("ADODB.Recordset")
rsEmails.Open EmailsSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 


rsEmails.Movefirst 
Response.Write "<table align=center border=1 width='98%' >"
Response.Write "<tr bgcolor=lightblue><td colspan=17 align='left'><b>Emails Table</b></td></tr>"
Response.Write "<tr bgcolor=lightblue><td></td><td></td><td>i</td><td>EmailID</td><td>SortOrder</td><td>EmailType</td>" & _
				"<td>EmailSubject</td><td>EmailDescription</td><td>EmailHeader HTML</td><td>EmailHeader</td><td>EmailBody HTML</td><td>EmailBody</td>" & _
				"<td>BCC</td><td>PicToEmbed</td><td>EmailFooter HTML</td><td>DateAdded</td></tr>"
row = 0
Do While NOT rsEmails.EOF
	row = row + 1
	EmailID				= rsEmails("EmailID")
	SortOrder			= rsEmails("SortOrder")
	EmailType			= rsEmails("EmailType")
	EmailSubject		= rsEmails("EmailSubject")
	EmailDescription	= rsEmails("EmailDescription")
	EmailHeader			= rsEmails("EmailHeader")
	EmailBody			= rsEmails("EmailBody")
	BCC					= rsEmails("BCC")
	PicToEmbed			= rsEmails("PicToEmbed")
	EmailFooter			= rsEmails("EmailFooter")
	DateAdded			= rsEmails("DateAdded")
	
	Response.Write "<tr>"
	
	Response.Write "<form method=Post>"
	Response.Write "<td valign=top bgcolor=pink><input type='Submit' name='Btn' value='Delete'><input type=hidden name=EmailID value=" & EmailID & "></td>"
	Response.Write "</form>"
	
	Response.Write "<form method=Post>"
	Response.Write "<td valign=top>" & "<input type='Submit' name='Btn' value='Edit'>" & "</td>"
	
	Response.Write "<td valign=top>" & row & "</td>"
	Response.Write "<td valign=top>" & EmailID & "</td>"
	Response.Write "<td valign=top>" & SortOrder & "</td>"
	Response.Write "<td valign=top>" & EmailType & "</td>"

	Response.Write "<td valign=top><input type=text name=EmailSubject value='" & EmailSubject & "' size=20 maxlength=255></input></td>"
	Response.Write "<td valign=top><input type=text name=EmailDescription value='" & EmailDescription & "' size=20 maxlength=50></input></td>"
	Response.Write "<td valign=top><textarea name=EmailHeader rows=10 cols=20>" & EmailHeader & "</textarea></td>"
	Response.Write "<td valign=top>" & EmailHeader & "</td>"
	Response.Write "<td valign=top><textarea name=EmailBody rows=20 cols=20>" & EmailBody & "</textarea></td>"
	Response.Write "<td valign=top>" & EmailBody & "</td>"	
	Response.Write "<td valign=top>" & BCC & "</td>"
	Response.Write "<td valign=top>" & PicToEmbed & "</td>"
	Response.Write "<td valign=top><textarea name=EmailFooter rows=20 cols=20>" & EmailFooter & "</textarea></td>"
	'Response.Write "<td valign=top>" & EmailFooter & "</td>"

	Response.Write "<td valign=top><input type=text name=DateAdded value='" & DateAdded & "' size=15 maxlength=15></input></td>"
	'Response.Write "<td valign=top>" & DateAdded & "</td>"

	Response.Write ""
	Response.Write "<input type=hidden name=EmailID value='" & EmailID & "'>"
	Response.Write "</form>"
	
	Response.Write "</tr>"
	
	rsEmails.moveNext
Loop
Response.Write  "</table>"

rsEmails.Close
Set rsEmails = Nothing
Set conn = Nothing
%>

<br><br>

</BODY>

</HTML>
