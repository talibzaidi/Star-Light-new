<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<%
Dim myMail
Set myMail = Server.CreateObject("CDONTS.NewMail")



		myMail.From = "TESURE@anyperson.com" 
		myMail.To = "sanction@anyperson.com" 
		myMail.Subject = "Fuckup" 
		myMail.BodyFormat = 0 
		myMail.MailFormat = 0 
		myMail.Body = "boo"
		myMail.Send 

%>
