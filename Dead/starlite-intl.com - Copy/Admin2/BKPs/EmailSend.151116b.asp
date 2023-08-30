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
    'Response.Write "<br>Substitute: FName = " & FName & " LName = " & LName
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

DBtoUse     = Request("DBtoUse")
PrevDBUsed  = Request("PrevDBUsed")
EmailToSend = Request("EmailToSend")
Response.Write "<br>DBtoUse = "     & DBtoUse
Response.Write "<br>PrevDBUsed = "  & PrevDBUsed
Response.Write "<br>EmailToSend = " & EmailToSend

Action      = Request("Action")
Response.Write "<br>Action = " & Action     ' Can be "" or "DisplayRecipients" or "SendToRecipients"

If Action = "SendToRecipients" Then     ' Specify set of recipients to send to.
    ' 11/16/15: Send the email to all records having record ID between the following minimum and maximum record IDs inclusive:
    MinID       = Request("MinID")
    MaxID       = Request("MaxID")
    Response.Write "<br>MinID = " & MinID
    Response.Write "<br>MaxID = " & MaxID
    If (MinID = "") OR (MaxID = "") Then
        Response.Write "<br>Error, empty string in invalid for MinID or MaxID."
        Response.End
    End If
End If
%>


<%
'Response.Write "<br><br>DBtoUse = " & DBtoUse 
If (DBtoUse = 1) Then
    Database = "OptedInMailingList.mdb"
    If (EmailToSend = "") OR (DBtoUse <> PrevDBUsed) Then
        EmailToSend = 5
    End If
ElseIf (DBtoUse = 2) Then
    Database = "PurchasedMailingList1.Tiny.mdb"
    If (EmailToSend = "") OR (DBtoUse <> PrevDBUsed) Then
        EmailToSend = 10
    End If
Else
%>
    <center>
    <FORM name="ChooseDBForm" XXmethod="Post">
	    <b>Choose a Database to Use:</b> &nbsp;&nbsp;
	    <select name="DBtoUse">
			    <option value=1 <% If DBtoUse = 1 Then Response.Write "selected" End If %> >OptedInMailingList.mdb</option>
                <option value=2 <% If DBtoUse = 2 Then Response.Write "selected" End If %> >PurchasedMailingList1.mdb</option>
	    </select>
        <input type=hidden name=Action value='DisplayRecipients' >
	    <input type="submit" value="Choose">
    </FORM>
    </center>
    <br />
<%
End If


If Database <> "" Then   ' i.e. if a database has been chosen.

    Set Conn = Server.CreateObject("ADODB.Connection")  ' 12/18/11: I had to copy this line to here from above, else Conn was not recognized as an open object. Don't know why.

    ' 11/4/15: In added the following to allow choosing the appropriate database, now that we have two of them.
    If Database = "OptedInMailingList.mdb" Then

        Response.Write "<br><br>Using the original (first) database, <font color='red'><b>OptedInMailingList.mdb</b></font>, of opted-in Starlite cutomers (which used to be called ec-star-001.mdb)"
        Response.Write "<br><br>Session('ConnectionString') = " & Session("ConnectionString")
        Conn.Open Session("ConnectionString")			    ' For original (first) database, PurchasedMailingList1.mdb, of company mailing addresses purchased from INFO-USA in Oct. 2015.	' 
        ' The WHERE Not IsNull(EmailID) is needed in case, for some reason, a record "is deleted" (?!)
        EmailsSQL = "SELECT * FROM Emails WHERE Not IsNull(EmailID) " & "ORDER BY EmailID "

        ' Display ~ all recipient records:
        DisplayRecipientsSQL = "SELECT Last(OrderID), FName, LName, Email, OptInToEmailings, WeEmail FROM Orders " & _
                        "WHERE Not IsNull(OrderID) AND OptInToEmailings = True AND WeEmail = True " & _
	                    "GROUP BY FName, LName, Email, OptInToEmailings, WeEmail ORDER BY LName"

        ' Send email to this subset of recipients = ~ all records:
        SendToRecipientsSQL = "SELECT Last(OrderID), FName, LName, Email, OptInToEmailings, WeEmail FROM Orders " & _
                        "WHERE Not IsNull(OrderID) AND OptInToEmailings = True AND WeEmail = True " & _
	                    "GROUP BY FName, LName, Email, OptInToEmailings, WeEmail ORDER BY LName"

    ElseIf Database = "PurchasedMailingList1.Tiny.mdb" Then
    
        Response.Write "<br><br>Using the new (second) database, <font color='red'><b>PurchasedMailingList1.mdb</b></font>, of company mailing addresses purchased from INFO-USA in Oct. 2015."
        Response.Write "<br><br>Session('ConnectionString2') = " & Session("ConnectionString2")
        Conn.Open Session("ConnectionString2")  
        EmailsSQL = "SELECT * FROM Emails"          
    
        ' Display ~ all records:
        DisplayRecipientsSQL    = "SELECT ID, [Executive Email] AS Email, State, [Last Name] AS LName, [First Name] AS FName FROM CompanyProspects " & _
                        "ORDER BY ID"  

        ' Send email to this subset of recipients:
        If (Action = "SendToRecipients") Then
        SendToRecipientsSQL = "SELECT ID, [Executive Email] AS Email, State, [Last Name] AS LName, [First Name] AS FName FROM CompanyProspects " & _
	                    "WHERE (ID >= " & MinID & ") AND (ID <= " & MaxID & ") " & _
                        "ORDER BY ID"  
        End If

    End If 

    Response.Write "<br><br>EmailsSQL = "       & EmailsSQL 
    Response.Write "<br><br>DisplayRecipientsSQL = "    & DisplayRecipientsSQL
    Response.Write "<br><br>SendToRecipientsSQL = "     & SendToRecipientsSQL & "<br><br>"

    Set rsEmails = Server.CreateObject("ADODB.Recordset")
    'Response.Write "<br>IsObject(Conn) = "		& IsObject(Conn)
    'Response.Write "<br>IsObject(rsEmails) = "	& IsObject(rsEmails)
    'Response.Write "<br>adLockOptimistic = "	& adLockOptimistic

    'Response.Write "<br><br>Conn = " & Conn
    rsEmails.Open EmailsSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 
    'Set rsEmails = Conn.Execute(EmailsSQL)

%>


<table align="center">
    <tr>
        <td>
        <% PrevDBUsed = DBtoUse %>
        <FORM name="ChooseDBForm" XXmethod="Post">
	        <b style="font-size:12px;">Choose a Database to Use:</b> &nbsp;&nbsp; 
	        <select name="DBtoUse">
			        <option value=1 <% If DBtoUse = 1 Then Response.Write "selected" End If %> >OptedInMailingList.mdb</option>
                    <option value=2 <% If DBtoUse = 2 Then Response.Write "selected" End If %> >PurchasedMailingList1.mdb</option>
	        </select>
	        &nbsp;&nbsp;
            <!-- <input type="hidden" name="PrevEmailToSend" value=<%=EmailToSend%> > -->
            <input type="hidden" name="PrevDBUsed" value=<%=PrevDBUsed%> >
            <input type=hidden name=Action value='DisplayRecipients' >
	        <input type="submit" value="Choose">
        </FORM> 

        <center>
        <FORM name="ChooseEmailForm" XXmethod="Post">
	        <b style="font-size:12px;">Choose an Email to Send:</b> &nbsp;&nbsp;
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
            <!-- Changing the Email does not require changing the db. -->
            <input type="hidden" name="DBtoUse" value=<%=DBtoUse%> >
            <input type="hidden" name="PrevDBUsed" value=<%=DBtoUse%> >
            <input type=hidden name=Action value='DisplayRecipients' >
            <input type="submit" value="Choose">
        </FORM>

        </td>
    </tr>
</table>
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
    <FORM name="MinAndMaxIDForm">
    <tr />
        <td>
        <%  If (MinID = "") AND (MaxID = "") Then
                Response.Write "<span style='color:red;'>Please choose a Min and Max Record ID</span>"
            End If
        %>
        </td>
        <td>
        </td>
    </tr>
    <tr>
	    <td bgcolor='pink'>
	    <input type="Submit" name="SendEmail" value="Send the email above to the people below">
        <br /><font size='2'>from min. Record ID:</font> <input name="MinID" />
        <br /><font size='2'>to max. Record ID:</font> <input name="MaxID" />
	    </td>
	    <td>
	    <input type="Submit" name="SendEmailTest" value="Send the email above to Sani and Berel">
	    </td>
    </tr>
	    <input type=hidden name=EmailToSend value=<%=EmailToSend%> >
        <input type=hidden name=DBtoUse value=<%=DBtoUse%> >
        <input type=hidden name=PrevDBUsed value=<%=DBtoUse%> >
        <input type=hidden name=Action value='SendToRecipients' >
    </FORM>
</table>

<br>


<%

' ****************************************************************************************************


Set rsDisplayRecords = Server.CreateObject("ADODB.Recordset")

' 11/4/15: The following 2 lines were added by HostMySite tech support. But I was going to add something similar myself anyway.
' Tech support is still missing the point: I want to know if and how I can define connection string ConnectionString 
' to connect to PurchasedMailingList1.accdb, a .accdb version of PurchasedMailingList1.mdb.
'Set Conn = Server.CreateObject("ADODB.Connection") 
'Conn.Open Session("ConnectionString2")

If TRUE Then
    Response.Write "<br><br>Action = "              & Action            ' Can be "" or "DisplayRecipients" or "SendToRecipients"

    If (DBtoUse = 1) Then
        Response.Write "<br><br>Session('ConnectionString) = " & Session("ConnectionString")        ' For db1.
    ElseIf (DBtoUse = 2) Then
        Response.Write "<br><br>Session('ConnectionString2) = " & Session("ConnectionString2")      ' For db2.
    ElseIf (DBtoUse = "") Then
        Response.Write "<br><br>No database has been chosen yet."        ' This case cannot occur in this location of the code.
    Else
        Response.Write "<br><br>Invalid database has been chosen."
    End If 

    Response.Write "<br><br>DisplayRecipientsSQL = " & DisplayRecipientsSQL
    Response.Write "<br><br>SendToRecipientsSQL = "  & SendToRecipientsSQL 
    'Response.Write "<br><br>Conn = "                & Conn
    'Response.Write "<br><br>adOpenStatic = "        & adOpenStatic
    'Response.Write "<br>adLockOptimistic = "        & adLockOptimistic
    'Response.Write "<br>adCmdText = "               & adCmdText
    Response.Write "<br><br>"
End If


'rsRecipients.Open RecipientsSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText 
rsDisplayRecords.Open DisplayRecipientsSQL, Conn, adOpenStatic, adLockOptimistic, adCmdText

Response.Write "<table align=center border=1 cellpadding=10 XXXwidth='800'>"
Response.Write "<tr>"
Response.Write "<td>"


If Database = "OptedInMailingList.mdb" Then

    rsDisplayRecords.Movefirst 
    Response.Write "<table align=center cellpadding=3 border=0>"
    Response.Write "<tr bgcolor='lightblue'>"
    Response.Write "<td align=center><b>Row</b></td> <td align=center><b>Record ID</b></td> <td align=center><b>Email</b></td>  " & _
            "<td align=center><b>FName</b></td> <td align=center><b>LName</b></td>" & _
            "<td align=center><b>Opted In</b></td> <td align=center><b>ON</b></td>"
    Response.Write "</tr>"
    row = 0
    ' Bad = 0
    Do While (NOT rsDisplayRecords.EOF)
	
        'ID 				= rsRecipients("OrderID")
	    FName 				= rsDisplayRecords("FName")
	    LName 				= rsDisplayRecords("LName")
	    Email 				= rsDisplayRecords("Email")
	    OptInToEmailings	= rsDisplayRecords("OptInToEmailings")
	    WeEmail				= rsDisplayRecords("WeEmail")
	
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
			        <!--# include virtual = "Admin2/Emailing/EmailMergeINC.asp" -->
        <%	
			        Response.Write "<font color=blue>Sent email " & EmailToSend & " at " & Now() & "</font>"
		        Else
			        Response.Write "&nbsp;"
		        End If
		        Response.Write "</td>"
	        Response.Write "</tr>"
	
        End If  ' ValidEmail(Email)

	    rsDisplayRecords.moveNext
    Loop
    Response.Write "</table>"

ElseIf Database = "PurchasedMailingList1.Tiny.mdb" Then

    rsDisplayRecords.Movefirst 
    Response.Write "<table align=center cellpadding=3 border=0>"
    Response.Write "<tr bgcolor='lightblue'>"    
    Response.Write "<td align=center><b>Row</b></td> <td align=center><b>Record ID</b></td> <td align=center><b>State</b></td>" & _
       "<td align=center><b>Email</b></td> <td align=center><b>FName</b></td> <td align=center><b>LName</b></td>" & _ 
       "<td align=center><b>Sent</b></td> <td align=center><b>Opted Out</b></td>"

    
    Response.Write "</tr>"
    row = 0
    ' Bad = 0
    Do While (NOT rsDisplayRecords.EOF)
	
        ID 				    = rsDisplayRecords("ID")
        State 		        = rsDisplayRecords("State")
        'Email               = "bn2@intelligineering.com"
        Email               = rsDisplayRecords("Email")            ' "bn2@intelligineering.com"
        LName 		        = rsDisplayRecords("LName")
        FName 		        = rsDisplayRecords("FName")

	
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
			        <!-- # include virtual = "Admin2/Emailing/EmailMergeINC.asp" -->
        <%	
			        Response.Write "<font color=blue>Sent email " & EmailToSend & " at " & Now() & "</font>"
		        Else
			        Response.Write "&nbsp;"
		        End If
		        Response.Write "</td>"
	        Response.Write "</tr>"
	
        End If  ' ValidEmail(Email)

	    rsDisplayRecords.moveNext
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

End If   '  Database <> ""
%>

<br><br>

</body>

</html>
