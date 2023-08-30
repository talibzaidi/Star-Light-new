<%
' 11/8/15: Adjusted for HostMySite, a` la their article https://solutions.hostmysite.com/index.php?/Knowledgebase/Article/View/8596/0/using-cdosys-to-create-an-asp-mail-form-that-uses-smtp-authentication

' 6/29/07: Copied from file EmailMergeINC.asp of Mit Mazel.

' 10/6/03: New method of sending email (See MS KB article 810702 at http://support.microsoft.com/?kbid=810702) ...
' Mike of JAM DATA had written me the following email on 10/3/03:

' Bernie,
' I loaded Microsoft Exchange service pack 3 a couple of days ago.  I believe this is what is causing your problems.  Microsoft apparently is
' discontinuing support for CDONTS.  I have attached some changes that you can make that will make the program work.  The problem deals with the CDO object
' and changes that were made to it during service pack 3.  For more information see: http://aspfaq.com/show.asp?id=2026

'Dim myCDOMail, sch, cdoConfig

sch = "http://schemas.microsoft.com/cdo/configuration/" 

' [11/8/15] For HostMySite, a` la their article https://solutions.hostmysite.com/index.php?/Knowledgebase/Article/View/8596/0/using-cdosys-to-create-an-asp-mail-form-that-uses-smtp-authentication
host         = "mail.starlite-intl.com"     'The mail server name. (Commonly mail.yourdomain.xyz if your mail is hosted with HostMySite)
username     = "starlite@starlite-intl.com" 'A valid email address you have setup 
from_address = "starlite@starlite-intl.com" 'If your mail is hosted with HostMySite this has to match the email address above 
password     = "K3ez6N2h"                   'Password for the above email address
reply_to     = "starlite@starlite-intl.com" 'The email you want customers to reply to
port         = "25"                         'This is the default port. Try port 50 if this port gives you issues and your mail is hosted with HostMySite.


Set cdoConfig = Server.CreateObject("CDO.Configuration") 
With cdoConfig.Fields 
    .Item(sch & "sendusing")                = 2		    ' i.e. cdoSendUsingPort 
    .Item(sch & "smtpserver")               = host	
    .Item(sch & "smtpserverport")           = port	
    .Item(sch & "smtpusessl")               = False     ' smtpusessl means: SMTP use SSL ?
    .Item(sch & "smtpconnectiontimeout")    = 60	
    .Item(sch & "smtpauthenticate")         = 1	
    .Item(sch & "sendusername")             = username	
    .Item(sch & "sendpassword")             = password				

    ' From http://www.c-amie.co.uk/technical/cdo-error-80040211/ 
    '.Item(sch & cdoSMTPAuthenticate) = cdoBasic
    '.Item(sch & cdoSendUserName) = username
    '.Item(sch & cdoSendPassword) = password

    .Update 
End With 
      
If TRUE Then
    Response.Write "<br><br>EmailFrom = "               & EmailFrom
    Response.Write "<br><br>EmailTo = "                 & EmailTo
    Response.Write "<br><br>SubstitutedEmailSubject = " & SubstitutedEmailSubject
    Response.Write "<br><br>EmailReplyTo = "            & EmailReplyTo
    Response.Write "<br><br>SubstitutedEmailBody = "    & SubstitutedEmailBody
End If
		
Set myCDOMail = Server.CreateObject("CDO.Message")          
With myCDOMail 
Set .Configuration = cdoConfig 
	.From		= EmailFrom 
	.To			= EmailTo   ' "sales@starlite-intl.com" 
    '.To			= "bn2@intelligineering.com"   '  "sales@starlite-intl.com" 
	'.BCC		= EmailBCC
	.Subject	= "subject"  ' SubstitutedEmailSubject
'	.ReplyTo	= EmailReplyTo
	.HTMLBody	= "body"  'SubstitutedEmailBody    ' To send an HTML message; otherwise use .TextBody = <string> to send a plain text message.     '     .Bcc = "staff@mitmazel.com"		
	.Send
End With 

Set myCDOMail  = Nothing     
Set cdoConfig = Nothing
%>