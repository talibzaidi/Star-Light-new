<%@ Language=VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->

<%  
' 10/23/15: On the "NEW server" at HostMySite, the included file "Admin2/Emailing/EmailMergeINC.asp"
' needs to be updated approximately a` la my file /MyTests/EmailOnNewServer/send_form_email.asp   
%>


<!--#include file="../ADOVBS.INC"-->

<%
Dim Conn, rsEmails, EmailsSQL
'Set Conn = Server.CreateObject("ADODB.Connection")			' 12/18/11: I had to copy these two lines from here to below, 
'Conn.Open Session("ConnectionString")						' else Conn was not recognized as an open object. Don't know why.
'Response.Write "<br>Session('ConnectionString') = " & Session("ConnectionString")
'Response.Write "<br>Conn = " & Conn
%>


<%
SendEmail		= Trim(Request.QueryString("SendEmail"))
SendEmailTest	= Trim(Request.QueryString("SendEmailTest"))
'Response.Write "<br>SendEmail = " & SendEmail
'Response.Write "<br>SendEmailTest = " & SendEmailTest
%>



<%
' A non-recursive version of Substitute. 
' Recursion can be avoided because (a) Replace() replaces ALL occurrences of replaced string by replacing string,
' and (b) by assumption that can use an explicit listing of all the different $$ metavariables that may be needed (one Replace() call for each).
Function Substitute(SourceString)
	NewSourceString = SourceString
	If FName <> "" Then 
		NewSourceString = Replace(NewSourceString, "$$FirstName$$", FName)		' Replaces ALL occurrences of "$$FirstName$$" in SourceString with FName.
	End If
	If LName <> "" Then 
		NewSourceString = Replace(NewSourceString, "$$LastName$$", LName)		' Replaces ALL occurrences of "$$LastName$$" in SourceString with LName.
	End If

	Substitute = NewSourceString
End Function	' Substitute
%>


<%
' Tests if Email address is valid.
Function ValidEmail(Email)
	AtPos 	= Instr(2, Email, "@")
	DotPos 	= Instr(4, Email, ".")
	If AtPos = 0 OR DotPos = 0 Then		' One of @ or . is missing
		ValidEmail = False
	ElseIf DotPos - AtPos < 2 Then		' The . must be at least 2 characters AFTER the @
		ValidEmail = False
	Else 
		ValidEmail = True
	End If
End Function	' ValidEmail
%>



<html>


<head>
	<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">
	<link rel="stylesheet" href="/Misc/StyleSheet1.css" type="text/css">
</head>


<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<% InSection = "EmailSend" %>

<!--#include virtual="Admin2/Misc/AdminHeader.asp"-->
<!--#include virtual="Misc/Header.INC"-->
<!--#include virtual="Admin2/Misc/LoginCheck.asp"-->
<br />

<%
'-------------------------------------------------------------------------------------------------------------
' Choose DB to Use and Email to Send ...

DBtoUse     = Request("DBToUse")
EmailToSend = Request("EmailToSend")
Response.Write "<br>DBtoUse = " & DBtoUse
Response.Write "<br>EmailToSend = " & EmailToSend

'If EmailToSend = "" Then
'	EmailToSend = 5
'End If
'Response.Write "<br>EmailToSend (2) = " & EmailToSend
%>


<%
'Response.Write "<br><br>DBtoUse = " & DBtoUse 
If (DBtoUse = 1) OR (DBtoUse = "") Then
    Database = "OptedInMailingList.mdb"
    If EmailToSend = "" Then
        EmailToSend = 10
    End If
ElseIf DBtoUse = 2 Then
    Database = "PurchasedMailingList1.mdb"
    If EmailToSend = "" Then
        EmailToSend = 10
    End If
Else
    'Response.Write "<br><br><font color='red'><b>Please choose a database from the drop-down menu.</b></font>" 
End If
Response.Write "<br><br>Database = "    & Database 
Response.Write "<br>EmailToSend = " & EmailToSend 


Set Conn = Server.CreateObject("ADODB.Connection")  ' 12/18/11: I had to copy this line to here from above, else Conn was not recognized as an open object. Don't know why.

' 11/4/15: In added the following to allow choosing the appropriate database, now that we have two of them.
If Database = "OptedInMailingList.mdb" Then

    Response.Write "<br><br>Using the original (first) database, <font color='red'><b>OptedInMailingList.mdb</b></font>, of opted-in Starlite cutomers (which used to be called ec-star-001.mdb)"
    Response.Write "<br><br>Session('ConnectionString') = " & Session("ConnectionString")
    Conn.Open Session("ConnectionString")			    ' For original (first) database, PurchasedMailingList1.mdb, of company mailing addresses purchased from INFO-USA in Oct. 2015.	' 
    ' The WHERE Not IsNull(EmailID) is needed in case, for some reason, a record "is deleted" (?!)
    EmailsSQL = "SELECT * FROM Emails WHERE Not IsNull(EmailID) " & "ORDER BY EmailID "

    RecipientsSQL = "SELECT Last(OrderID), FName, LName, Email, OptInToEmailings, WeEmail FROM Orders " & _
                    "WHERE Not IsNull(OrderID) AND OptInToEmailings = True AND WeEmail = True " & _
	                "GROUP BY FName, LName, Email, OptInToEmailings, WeEmail ORDER BY LName"

ElseIf Database = "PurchasedMailingList1.mdb" Then
    
    Response.Write "<br><br>Using the new (second) database, <font color='red'><b>PurchasedMailingList1.mdb</b></font>, of company mailing addresses purchased from INFO-USA in Oct. 2015."
    Response.Write "<br><br>Session('ConnectionString2') = " & Session("ConnectionString2")
    Conn.Open Session("ConnectionString2")  
    EmailsSQL = "SELECT * FROM Emails"          
    
    RecipientsSQL = "SELECT ID, [Executive Email] AS Email, State, [Last Name] AS LName, [First Name] AS FName FROM CompanyProspects " & _
	                "ORDER BY ID"  
End If 

Response.Write "<br><br>EmailsSQL = " & EmailsSQL 
Response.Write "<br><br>RecipientsSQL = " & RecipientsSQL & "<br><br>"

Set rsEmails = Server.CreateObject("ADODB.Recordset")
'Response.Write "<br>IsObject(Conn) = "		& IsObject(Conn)
'Response.Write "<br>IsObject(rsEmails) = "	& IsObject(rsEmails)
'Response.Write "<br>adLockOptimistic = "	& adLockOptimistic

'Response.Write "<br><br>Conn = " & Conn
rsEmails.Open EmailsSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 
'Set rsEmails = Conn.Execute(EmailsSQL)
%>


<center>
<FORM name="ChooseDBandEmailForm" method="Post">
	<b>Database to Use:</b> &nbsp;&nbsp;
	<select name="DBToUse">
			<option value=1 <% If DBtoUse = 1 Then Response.Write "selected" End If %> >OptedInMailingList.mdb</option>
            <option value=2 <% If DBtoUse = 2 Then Response.Write "selected" End If %> >PurchasedMailingList1.mdb</option>
	</select>
	&nbsp;&nbsp;
    <input type="hidden" name="EmailToSend" value=<%=EmailToSend%> >
	<input type="submit" value="Choose">
</FORM>
</center>
<br />


<center>
<FORM name="Form1" method="Post">
	<b>Email to Send:</b> &nbsp;&nbsp;
	<select name="EmailToSend">
		<%
		rsEmails.Movefirst 
		Do While NOT rsEmails.EOF
            If EmailToSend <> rsEmails("EmailID") Then
	            Response.Write "<option value=" & rsEmails("EmailID") & ">" & rsEmails("EmailID") & ": " & rsEmails("EmailSubject")
            ElseIf EmailToSend = rsEmails("EmailID") Then
            %>
                <option <%Response.Write "selected" %> value=<%=rsEmails("EmailID")%> ><%=rsEmails("EmailID")%>:<%=rsEmails("EmailSubject")%></option>
            <% ' The above insertion of "selected" does not seem to work.
            End If
			rsEmails.moveNext
		Loop
		%>
	</select>
	&nbsp;&nbsp;
    <input type="hidden" name="DBToUse" value=<%=DBtoUse%> >
	<input type="submit" value="Choose">
</FORM>
</center>
<br />

<%
'-------------------------------------------------------------------------------------------------------------
' Display Email chosen for sending ...
%>


<%
If EmailToSend <> "" Then
	EmailToSendSQL = "SELECT * FROM Emails WHERE EmailID=" & EmailToSend
	'Response.Write "<br>EmailToSendSQL = " & EmailToSendSQL & "<br>"
	'Response.End
	Set rsEmailToSend = Server.CreateObject("ADODB.Recordset")
	rsEmailToSend.Open EmailToSendSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 
	
	EmailID				= Trim(rsEmailToSend("EmailID"))
	EmailDescription	= Trim(rsEmailToSend("EmailDescription"))
	EmailBCC			= Trim(rsEmailToSend("BCC"))
	EmailSubject		= Trim(rsEmailToSend("EmailSubject"))
	'EmailFrom			= "sales@starlite-intl.com"
    EmailFrom			= "starlite@starlite-intl.com"
	'EmailReplyTo		= "sales@starlite-intl.com"
	EmailReplyTo		= "starlite@starlite-intl.com"
	EmailHeader			= Trim(rsEmailToSend("EmailHeader"))
	EmailBody			= Trim(rsEmailToSend("EmailBody"))
	EmailFooter			= Trim(rsEmailToSend("EmailFooter"))

	FullEmailBody		= EmailBody
	' If there is a header Then pre-pend it.
	If EmailHeader <> "" AND Not IsNull(EmailHeader) Then
		FullEmailBody	= EmailHeader & "<br>" & FullEmailBody
	End If
	
	' If there is a picture to embed Then embed it.
	If rsEmailToSend("PicToEmbed") <> "" AND Not IsNull(rsEmailToSend("PicToEmbed")) Then
		PathToPic = "http://www.starlite-intl.com/Admin2/PicsToEmbedInEmail/" & rsEmailToSend("PicToEmbed") 
	FullEmailBody = FullEmailBody & "<center><img src='" & PathToPic & "'></center>"
	End If
	
	' If there is a footer Then append it.
	If EmailFooter <> "" AND Not IsNull(EmailFooter) Then
		FullEmailBody	= FullEmailBody & EmailFooter
	End If
%>


<table align="center" border="1" cellpadding="5" width="800" bgcolor="pink">
<tr>
	<td colspan="3">
		<font face="Tahoma">
		<b>Email ID:</b>&nbsp;<%=EmailID%><br>
		<b>Email Description:</b>&nbsp;<%=EmailDescription%><br>
		<b>BCC:</b>&nbsp;<%=EmailBCC%><br><br>
		
		<b>Subject:</b>&nbsp;<%=EmailSubject%><br>
		<b>From:</b>&nbsp;<%=EmailFrom%><br><br>
		<hr><br>
		<%=FullEmailBody%>
		</font>
	</td>
</tr>
</table>

<% End If	' EmailToSend <> "" %>


<%
'-------------------------------------------------------------------------------------------------------------
' List recipients for email being sent 
' and mail to them if the "Send the email" button was pressed ...
%>

<br><br>

<table align=center border=1 cellpadding=5>
<FORM name="Form2">
<tr>
	<td bgcolor=red>
	<input type="Submit" name="SendEmail" value="Send the email above to the people below">
	</td>
	<td>
	<input type="Submit" name="SendEmailTest" value="Send the email above to Sani and Berel">
	</td>
</tr>
	<input type=hidden name=EmailToSend value=<%=EmailToSend%> >
</FORM>
</table>

<br>


<%
' The WHERE Not IsNull(OrderID) is needed in case, for some reason, a record "is deleted" (?!)
If SendEmailTest <> "" Then
	'RecipientsSQL = "SELECT * FROM Orders WHERE (LName = 'Zeevi') OR (LName = 'Shmerel') " & "ORDER BY LName"
    'RecipientsSQL = "SELECT * FROM CompanyProspects WHERE (LName = 'Zeevi') OR (LName = 'Shmerel') " & "ORDER BY LName   
    'RecipientsSQL = "SELECT ID, [Executive Email] AS Email, State, [Last Name] AS LName, [First Name] AS FName FROM CompanyProspects " & _
    '                "WHERE ([Last Name] LIKE 'Zeevi') OR ([Last Name] LIKE 'Shmerel') " & _
	'                "ORDER BY ID" 
Else
	'RecipientsSQL = "SELECT * FROM Orders WHERE Not IsNull(OrderID) AND OptInToEmailings = True ORDER BY LName"
	'RecipientsSQL = "SELECT FName, LName, Email, OptInToEmailings, WeEmail FROM Orders GROUP BY Email, LName, FName, OptInToEmailings, WeEmail ORDER BY LName"
	'RecipientsSQL = "SELECT Last(OrderID), FName, LName, Email, OptInToEmailings, WeEmail FROM Orders WHERE Not IsNull(OrderID) AND OptInToEmailings = True GROUP BY Email, LName, FName, OptInToEmailings, WeEmail ORDER BY LName"
	'RecipientsSQL = "SELECT Last(OrderID), FName, LName, Email, OptInToEmailings, WeEmail FROM Orders WHERE Not IsNull(OrderID) AND OptInToEmailings = True AND WeEmail = True " & _
	'               "GROUP BY FName, LName, Email, OptInToEmailings, WeEmail ORDER BY LName"
End If

Response.Write "<table width='1000' align='center'><tr><td>"
Response.Write "RecipientsSQL = " & RecipientsSQL
Response.Write "</td></tr></table><br />"

Set rsRecipients = Server.CreateObject("ADODB.Recordset")

' 11/4/15: The following 2 lines were added by HostMySite tech support. But I was going to add something similar myself anyway.
' Tech support is still missing the point: I want to know if and how I can define connection string ConnectionString 
' to connect to PurchasedMailingList1.accdb, a .accdb version of PurchasedMailingList1.mdb.
'Set Conn = Server.CreateObject("ADODB.Connection") 
'Conn.Open Session("ConnectionString2")

If TRUE Then
    Response.Write "<br><br>Session('ConnectionString2) = " & Session("ConnectionString2")
    Response.Write "<br><br>RecipientsSQL = "       & RecipientsSQL
    Response.Write "<br><br>Conn = "                & Conn
    Response.Write "<br><br>adOpenStatic = "        & adOpenStatic
    Response.Write "<br><br>adLockOptimistic = "    & adLockOptimistic
    Response.Write "<br><br>adCmdText = "           & adCmdText
End If

rsRecipients.Open RecipientsSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 

Response.Write "<table align=center border=1 cellpadding=10 XXXwidth='800'>"
Response.Write "<tr>"
Response.Write "<td>"


If Database = "OptedInMailingList.mdb" Then

    rsRecipients.Movefirst 
    Response.Write "<table align=center cellpadding=3 border=0>"
    Response.Write "<tr bgcolor='lightblue'>"
    Response.Write "<td align=center><b>Row</b></td> <td align=center><b>Record ID</b></td> <td align=center><b>Email</b></td>  " & _
            "<td align=center><b>FName</b></td> <td align=center><b>LName</b></td>" & _
            "<td align=center><b>Opted In</b></td> <td align=center><b>ON</b></td>"
    Response.Write "</tr>"
    row = 0
    ' Bad = 0
    Do While (NOT rsRecipients.EOF)
	
        'ID 				    = rsRecipients("OrderID")
	    FName 				= rsRecipients("FName")
	    LName 				= rsRecipients("LName")
	    Email 				= rsRecipients("Email")
	    OptInToEmailings	= rsRecipients("OptInToEmailings")
	    WeEmail				= rsRecipients("WeEmail")
	
        If ValidEmail(Email) Then 
	        row = row + 1
	        SubstitutedEmailSubject = Substitute(EmailSubject)
	        SubstitutedEmailBody = Substitute(FullEmailBody)

	        Response.Write "<tr>"
		        Response.Write "<td align=center>" & row & "</td>"
                Response.Write "<td>" & ID & "</td>"
                Response.Write "<td>" & Email & "</td>"
		        Response.Write "<td>" & FName & "</td>"
		        Response.Write "<td>" & LName & "</td>"
		        'If Not ValidEmail(Email) Then 
		        '	Bad = Bad + 1
		        '	Response.Write " BAD: " & Bad
		        'End If
		        Response.Write "<td>" & OptInToEmailings & "</td>"
		        Response.Write "<td>" & WeEmail & "</td>"
		        Response.Write "<td>"
		
		        EmailTo = Email   ' "bn@intelligineering.com" 
		        ' The following INCLUDE-d code actually sends the email ...
		        'If False Then		' Can set to False during debugging and further development.
		        If (SendEmail <> "") OR (SendEmailTest <> "") Then
        %>
			        <!--#include virtual = "Admin2/Emailing/EmailMergeINC.asp" -->
        <%	
			        Response.Write "<font color=blue>Sent email " & EmailToSend & " at " & Now() & "</font>"
		        Else
			        Response.Write "&nbsp;"
		        End If
		        Response.Write "</td>"
	        Response.Write "</tr>"
	
        End If  ' ValidEmail(Email)

	    rsRecipients.moveNext
    Loop
    Response.Write "</table>"

ElseIf Database = "PurchasedMailingList1.mdb" Then

    rsRecipients.Movefirst 
    Response.Write "<table align=center cellpadding=3 border=0>"
    Response.Write "<tr bgcolor='lightblue'>"    
    Response.Write "<td align=center><b>Row</b></td> <td align=center><b>Record ID</b></td> <td align=center><b>State</b></td>" & _
       "<td align=center><b>Email</b></td> <td align=center><b>FName</b></td> <td align=center><b>LName</b></td>" & _ 
       "<td align=center><b>Sent</b></td> <td align=center><b>Opted Out</b></td>"

    
    Response.Write "</tr>"
    row = 0
    ' Bad = 0
    Do While (NOT rsRecipients.EOF)
	
        ID 				    = rsRecipients("ID")
        State 		        = rsRecipients("State")
        'Email               = "bn2@intelligineering.com"
        Email               = rsRecipients("Email")            ' "bn2@intelligineering.com"
        LName 		        = rsRecipients("LName")
        FName 		        = rsRecipients("FName")

	
        If ValidEmail(Email) Then 
	        row = row + 1
	        SubstitutedEmailSubject = Substitute(EmailSubject)
	        SubstitutedEmailBody = Substitute(FullEmailBody)

	        Response.Write "<tr>"
		        Response.Write "<td align=center>" & row & "</td>"
		        Response.Write "<td>" & ID & "</td>"
		        Response.Write "<td>" & State & "</td>"
		        Response.Write "<td>" & Email & "</td>"
                Response.Write "<td>" & FName & "</td>"
                Response.Write "<td>" & LName & "</td>"
		        'If Not ValidEmail(Email) Then 
		        '	Bad = Bad + 1
		        '	Response.Write " BAD: " & Bad
		        'End If
		        Response.Write "<td>"
		
		        EmailTo = Email   ' "bn@intelligineering.com" 
		        ' The following INCLUDE-d code actually sends the email ...
		        'If False Then		' Can set to False during debugging and further development.
		        If (SendEmail <> "") OR (SendEmailTest <> "") Then
        %>
			        <!-- #include virtual = "Admin2/Emailing/EmailMergeINC.asp" -->
        <%	
			        Response.Write "<font color=blue>Sent email " & EmailToSend & " at " & Now() & "</font>"
		        Else
			        Response.Write "&nbsp;"
		        End If
		        Response.Write "</td>"
	        Response.Write "</tr>"
	
        End If  ' ValidEmail(Email)

	    rsRecipients.moveNext
    Loop
    Response.Write "</table>"

End If


Response.Write "</td>"
Response.Write "</tr>"
Response.Write "</table>"
%>


<%
rsEmails.Close
rsEmailToSend.Close
Conn.Close
Set rsEmails		= Nothing
Set rsEmailToSend	= Nothing
Set Conn			= Nothing
%>

<br><br>

</body>

</html>
