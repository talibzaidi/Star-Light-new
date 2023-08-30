<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<%

		'**************************************************************
		'* COPYRIGHT 2000 WWW.ANYPERSON.COM			      *
		'* ANYMAILER - ASP FORM EMAIL				      *
		'* RESPONDS TO ANY FORM USING THE 'GET' METHOD                *
		'* THIS CODE IS FREE TO USE AS LONG AS IT RETAINS THIS HEADER *
		'* USAGE : <form action=AnyMail.asp method=GET>		      *
		'**************************************************************

		'****** INITIALIZE CDONTS MAIL OBJECT FOR USE *****************
		'****** INSTALLS WITH NT OPTION PACK 4.0 UNDER IIS
		Dim myMail
		Set myMail = Server.CreateObject("CDONTS.NewMail")

		'****** ALWAYS GOOD TO REDIRECT TO AN ERROR PAGE IF SOMETHING GOES WRONG
		'If Err.number <> 0 then
		'    response.redirect "error.asp"
		'end if
	
		'******* THIS SCRIPT WILL PARSE A FORM INTO ITS CONSTITUENT PARTS ****
		'******* CONVERT THEM INTO A TABULAR HTML FORMAT AND *****************
		'******* MAIL THEM TO WHOMEVER YOU CHOOSE ****************************
		

		GigaParse = Request.QueryString
		dim count
		count = 0
		
		if GigaParse <> "" then
		mailer = GigaParse
		loobol = len(mailer)
		loobol = loobol + 1
		testvar = loobol
	
		'********* INITIATE HTML TABLE LAYOUT **********************************
		mailDeamon = "<br>"
		mailDaemon = MailDaemon & "<center><font size=7><b>This will be emailed to you.</b><br><Table border=0  width=500 cellpadding=2 cellspacing=1><tr>"
		mailDaemon = MailDaemon & "<font face=arial size=3>"
		mailDaemon = MailDaemon & "<td colspan=2 bgcolor=#0000cc><font face=arial size=3 color=#ffffff><b>&nbsp;&nbsp;INCOMING E-MAIL FROM WEBSITE&nbsp;&nbsp;</FONT></TD></TR><TR>"
		mailDaemon = MailDaemon & "<td width=100 valign=center bgcolor=#cccccc align=left><font face=arial size=2><b>&nbsp;"		


		'********* PARSE THE CAPTURED STRING CHAR BY CHAR **********************
		do Until count = loobol 
				
		
		'************ DEAL WITH AMPERSANDS *************************************
		charz = left(mailer,count)
		char = right(charz,1)
 		 if (char = "&")  then
		 mailDaemon = MailDaemon & ""
		 
		end if
		
		'*********** DEAL WITH EQUALS SIGNS, PLUS SIGNS AND AMPERSAND DISPLAY **	
		if (char = "&") then
		char = right(charz,0)
 		mailDaemon = MailDaemon & "&nbsp;<tr><td width=100 valign=center align=left bgcolor=#cccccc><font face=arial size=2><b>&nbsp;"
 		end if

		if (char = "+") then
		char = right(charz,0)
 		mailDaemon = MailDaemon & "&nbsp;"
 		end if

		if (char = "=") then
		char = right(charz,1)
 		mailDaemon = MailDaemon & "&nbsp;</font></td><td width=100 valign=center bgcolor=#eeeeee align=left><font face=arial size=2>&nbsp;"
 		mailDaemon = MailDaemon & "&nbsp;&nbsp;"
 		'response.write("<b>•</b>")
		'response.write("&nbsp;&nbsp;")
 		else


		'********* DEAL WITH ESCAPE KEYS =,%,(,),$,^, etc ********************
		'********* ADD MORE ESCAPE SEQUENCES WITH A SIMPLE CUT AND PASTE

		if char = "%" then
		count = count + 2
		charz = left(mailer,count)
		char = right(charz,0)
		ParseCode = right(charz,2)
			'******* AFTER CLAIMING TWO DIGIT ESCAPE CODE ****************
			'******* DEAL WITH ESCAPE KEYS ON CASE BY CASE BASIS *********
			'*************************************************************
			If ParseCode = "26" then
			mailDaemon = MailDaemon & "&"
			elseif ParseCode = "28" then
			mailDaemon = MailDaemon & "("
			elseif ParseCode = "29" then
			mailDaemon = MailDaemon & ")"
			elseif ParseCode = "3D" then
			mailDaemon = MailDaemon & "="
			elseif ParseCode = "2C" then
			mailDaemon = MailDaemon & ","
			elseif ParseCode = "21" then
			mailDaemon = MailDaemon & "!"
			elseif ParseCode = "7E" then
			mailDaemon = MailDaemon & "~"
			elseif ParseCode = "60" then
			mailDaemon = MailDaemon & "`"
			elseif ParseCode = "23" then
			mailDaemon = MailDaemon & "#"
			elseif ParseCode = "24" then
			mailDaemon = MailDaemon & "$"
			elseif ParseCode = "5E" then
			mailDaemon = MailDaemon & "^"
			elseif ParseCode = "5C" then
			mailDaemon = MailDaemon & "\"
			elseif ParseCode = "7C" then
			mailDaemon = MailDaemon & "|"
			elseif ParseCode = "5B" then
			mailDaemon = MailDaemon & "["
			elseif ParseCode = "5D" then
			mailDaemon = MailDaemon & "]"
			elseif ParseCode = "7B" then
			mailDaemon = MailDaemon & "{"
			elseif ParseCode = "7D" then
			mailDaemon = MailDaemon & "}"
			elseif ParseCode = "2B" then
			mailDaemon = MailDaemon & "+"
			elseif ParseCode = "2F" then
			mailDaemon = MailDaemon & "/"
			elseif ParseCode = "3C" then
			mailDaemon = MailDaemon & "&lt;"
			elseif ParseCode = "3E" then
			mailDaemon = MailDaemon & "&gt;"
			elseif ParseCode = "3F" then
			mailDaemon = MailDaemon & "?"
			elseif ParseCode = "3B" then
			mailDaemon = MailDaemon & ";"
			elseif ParseCode = "3A" then
			mailDaemon = MailDaemon & ":"
			elseif ParseCode = "27" then
			mailDaemon = MailDaemon & "'"
			elseif ParseCode = "22" then
			mailDaemon = MailDaemon & "&quot;"
			end if
		'******* END OF DEAL WITH ESCAPE KEYS ON CASE BY CASE BASIS *********
		'********************************************************************
		end if	

		
		

		'********* WHEN CONDITIONS ARE CLEAR, TYPE **************************
		mailDaemon = MailDaemon & ""
		mailDaemon = MailDaemon + char
 		end if
		count = (count+1)
		loop
		

		end if

		
		'********* TIDY UP THE HTML TABLE ************************************
		mailDaemon = MailDaemon & "<TR><td colspan=2 bgcolor=#0000cc><font face=arial size=3 color=#ffffff><b>&nbsp;&nbsp;&nbsp;</FONT></TD></TR>"
		mailDaemon = MailDaemon & "</Table></center>"



	        StrBODY = mailDaemon 
	

		'******* MAIL SENDING ROUTINE ****************************************
		'TO CUSTOMIZE THIS SECTION CHANGE THE FROM, TO AND SUBJECT
		'COMMENT OUT ARE RECOMMENDED HIDDEN FORM FIELDS

		myMail.From = Session("UName") & "@Anyperson.com"   	'Request("MailFrom")
		myMail.To = "your@website.com" 			'Request("MailTo")
		myMail.Subject = "Message from " & Session("UName") & "@Anyperson.com" 			'Request("MailSubject")
		myMail.BodyFormat = 0 
		myMail.MailFormat = 0 
		myMail.Importance = 0
		myMail.Body = StrBody
		myMail.Send 


		'******* CLEAN UP SOME VARIABLES AND REDIRECT THE USER ***************
		SET mymail = nothing
		SET count = nothing
		SET strBODY = nothing
		SET mailDaemon = nothing

		'response.redirect "index.asp"		'*** WHAT TO DO WHEN YOU ARE DONE

%>
<%

'********** ONLY DEBUGGING CODE BEYOND THIS POINT! ************************************
'********** FEEL FREE TO DELETE YOU WILL NOTICE THE myMail.Send ABOVE IS CURRENTLY ****
'********** COMMENTED OUT *************************************************************
%>

<%=strBody%>




