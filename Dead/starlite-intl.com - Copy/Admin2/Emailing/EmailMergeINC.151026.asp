<%

' 6/29/07: Copied from file EmailMergeINC.asp of Mit Mazel.

' 10/6/03: New method of sending email (See MS KB article 810702 at http://support.microsoft.com/?kbid=810702) ...
' Mike of JAM DATA had written me the following email on 10/3/03:

' Bernie,
' I loaded Microsoft Exchange service pack 3 a couple of days ago.  I believe this is what is causing your problems.  Microsoft apparently is
' discontinuing support for CDONTS.  I have attached some changes that you can make that will make the program work.  The problem deals with the CDO object
' and changes that were made to it during service pack 3.  For more information see: http://aspfaq.com/show.asp?id=2026

'Dim myCDOMail, sch, cdoConfig

sch = "http://schemas.microsoft.com/cdo/configuration/" 
    
Set cdoConfig = Server.CreateObject("CDO.Configuration") 
With cdoConfig.Fields 
.Item(sch & "sendusing") = 2						' i.e. cdoSendUsingPort 
'.Item(sch & "smtpserver") = "mail.jamdata.net"		' [11/21/04] This used to work when Mit Mazel was hosted at JAM Data.
.Item(sch & "smtpserver") = "localhost"				' [11/21/04] Apparently I need to use this now that Mit Mazel is hosted at Interland.
.Update 
End With 
      
		
Set myCDOMail = Server.CreateObject("CDO.Message")          
With myCDOMail 
Set .Configuration = cdoConfig 
	.From		= EmailFrom 
	.To			= EmailTo   ' "sales@starlite-intl.com" 
	'.BCC		= EmailBCC
	.Subject	= SubstitutedEmailSubject
	.ReplyTo	= EmailReplyTo
	.HTMLBody	= SubstitutedEmailBody    ' To send an HTML message; otherwise use .TextBody = <string> to send a plain text message.     '     .Bcc = "staff@mitmazel.com"		
	.Send
End With 

Set myCDOMail  = Nothing     
Set cdoConfig = Nothing
%>