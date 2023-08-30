<html>

<head>

<title>ASPMail Test</title>
</head>

<body>

<%

if request.form("domain")<>"" then

	'set variables

	varDomain=request.form("domain")
	varServer="mail."& vardomain 
	varFromAdd="root@"& vardomain 
	varName=request.form("name")
	varEmail=request.form("email")
	varSubject=request.form("subject")
	varMessage=request.form("message")
	



	'CDONTS script
	Set MyCDONTSMail = CreateObject("CDONTS.NewMail")
	MyCDONTSMail.From = varFromAdd
	MyCDONTSMail.To = varEmail
	MyCDONTSMail.Subject = varSubject
	MyCDONTSMail.Body = varMessage
	'MyCDONTSMail.MailFormat = 0
	'MyCDOTNSMail.BodyFormat = 0
	MyCDONTSMail.Send
    	set MyCDONTSMail=nothing
	
	response.write("<BR>")
	response.write("From Address: "&varFromAdd)
	response.write("<BR>")
	response.write("To Address: "&varEmail)
	response.write("<BR>")
	
end if
%>

<form method="POST" action="cdontstest.asp">

  This is a test CDONTS script that will test the function of CDONTS on the 
  server.&nbsp; </p>
  <p>For example, if the domain was thedomain.xyz, this will send the email from root@thedomain.xyz</p>
  <p>DOMAIN NAME: <input type="text" name="domain" size="20"></p>
  <p>TO NAME: <input type="text" name="name" size="20"></p>
  <p>TO EMAIL ADDRESS: <input type="text" name="email" size="20"></p>
  <p>SUBJECT: <input type="text" name="subject" size="20"></p>
  <p>MESSAGE: <textarea rows="2" name="message" cols="20"></textarea> </p>
  <p><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>

</body>

</html>