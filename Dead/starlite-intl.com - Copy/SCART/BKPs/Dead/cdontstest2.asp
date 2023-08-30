<%
Dim myMail
Set myMail = CreateObject("CDONTS.NewMail")

HTML = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & vbCrLf
HTML = HTML & "<html>"
HTML = HTML & "<head>"
HTML = HTML & "<meta http-equiv=""Content-Type"""
HTML = HTML & "content=""text/html; charset=iso-8859-1"">"
HTML = HTML & "<title>Sample NewMail</title>"
HTML = HTML & "</head>"
HTML = HTML & "<body>"
HTML = HTML & "This is a sample message being sent using HTML. <BR></body>"
HTML = HTML & "</html>"

myMail.From = "astopani@hostmysite.com"
myMail.To = "astopani@hostmysite.com"
myMail.Subject = "Sample Message"
myMail.BodyFormat = 0
myMail.MailFormat = 0
myMail.Body = HTML
myMail.Send
Set myMail = Nothing
%>
Done!
