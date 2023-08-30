
<% @ Language=VBScript EnableSessionState=False %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->


<!-- #XXXINCLUDE VIRTUAL = "/Misc/CheckLogin.inc.asp" ' 5/1/10: Commented out. No longer appropriate because this file was just made available to ALL users. -->



<% ' [7/26/03] This page was used as a model for eMailToRequestDeletionForm.asp. %>

<% 

Sub thisPage_onenter()
' The following test does not avoid a crash if any Recordset DTC is being automatically opened on this page
' and needs a Session variable as an input parameter. In that case, I should re-write the Recordset(s) to not
' be opened automatically.
'If Session("UserID") = "" Then
'	Response.Redirect "../SessionProblem.asp"
'End If
End Sub


'Sub cmdSendMsgToMM_onclick()		
'   Response.Cookies("MMCookie")("MsgToMitMazel") = Request.Cookies("MMCookie")("MsgToMM")
'End Sub
%>

<html>

<head>
	<link rel="stylesheet" type="text/css" href="../Misc/StyleSheet1.css">
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<title>Email Us</title>

	<link rel="shortcut icon" href="http://www.mitmazel.com/MM/favicon.ico" type="image/x-icon"> 
	<link rel="icon" href="http://www.mitmazel.com/MM/favicon.ico" type="image/x-icon">
</head>


<body>



<!-- # INCLUDE VIRTUAL = "MM/header.htm" -->
<% 
InArea 		= "Questions"				' Parameter of NavBars1.inc file 
InSubArea 	= "ContactUs"		
%>
<!-- #INCLUDE VIRTUAL = "Misc/Header.INC" -->

<!--# include virtual="Misc/Header.INC"-->

<table class="PageHeadingTable" >
<tr>
	<td>
	<IMG hspace=0 vspace=0 src='../images/navimages/Q30.transp.png' align="absmiddle" border=0>
	</td>
	
	<td>
        <font color=lavender style="">
	Email Us
        </font>
	</td>
</tr>
</table>



<%  
' This failure cases do not result in a jump to eMailMsgToMM.asp, nor in any emails being sent to Mit Mazel.
If CBool(Request.Cookies("MMCookie")("ValidMemberLogin")) = True AND CBool(Instr(Request.Cookies("MMCookie")("UserEmail"), "invalid")) Then 
%>
	<br><br><br>
	<center>
	<font color=red>
	You cannot use this page because the email address that you originally gave us is no longer valid. 
	So Mit Mazel will not be able to reply to you.
	<br><br>
	If you have a new email address, you need to first  
	<a href="http://www.mitmazel.com/MM/singles/ProfileEdit.asp#EmailAddress">update it</a> in your profile and then login to Mit Mazel again. 
	<br><br>
	If necessary, you can email us at <a href="mailto:staff@mitmazel.com">staff@mitmazel.com</a>
	</font>
	</center>
<% 
	Response.End
End If 
%>


<% ' Due to the Response.End above, we only do the following if the If-Then condition above is False. %>

<form action="emailQtoMM.asp" METHOD="POST">

<div style="text-align:center; margin:0px 0px 20px 0px; border:0px solid red; line-height:1.5">
Use this page to send us your requests, questions, comments or suggestions. 
</div>

<table align="center" border=0 cellPadding="0" cellSpacing="0" width="65%" bgcolor="lavender">

	<tr>
		<td width="13"><img src="../images/RoundCorners/B0C0D0-TL.png" width="13" height="13"></td>
		<!-- In IE, bgcolor of #A8B9CB looks more like #B0C0D0 of the corner files, than does #B0C0D0 itself!! -->
		<!-- Not so in Firefox and Chrome, but looking good in IE is more important. -->
		<td bgColor="#A8B9CB"  bgColorXXX="#B0C0D0"><!-- Blank top section --></td>
		<td width="13"><img src="../images/RoundCorners/B0C0D0-TR.png" width="13" height="13"></td>
	</tr>

	<tr>
		<td bgColor="#A8B9CB"><!-- Blank left section --></td>
		<td height=10></td>
		<td bgColor="#A8B9CB"><!-- Blank right section --></td>
	</tr>
	
	<tr>
		<td bgColor="#A8B9CB"><!-- Blank left section --></td>
	
  		<td align="left"><font size="2">	 

		<table align="center" border="0" cellPadding="10" cellSpacing="1" width="100%">
	
		<% If CBool(Request.Cookies("MMCookie")("ValidMemberLogin")) = FALSE Then %>
		<tr>
			<td align="right">
			Your Email Address:
			</td>

			<td align="left">
			<input name="UserEmailAddress" size="50" maxlength="50">
			</td>

			<td valign="top">
			</td>
		</tr>
		<% End If %>

		<tr>
			<td align="right" valign="top">
				Your Message:
			</td>

			<td align="left" valign="top">
				<textarea name="UserMsgToMM" rows="6" cols="80"></textarea>
			</td>

			<td valign="top">
			</td>
		</tr>
	
		</table>
	
		<td bgColor="#A8B9CB"><!-- Blank right section --></td>
	
		</td>
	</tr>

	<tr>
		<td bgColor="#A8B9CB"><!-- Blank left section --></td>
		<td height=10></td>
		<td bgColor="#A8B9CB"><!-- Blank right section --></td>
	</tr>

	<tr>
		<td width="13"><img src="../images/RoundCorners/B0C0D0-BL.png" width="13" height="13"></td>
		<td bgColor="#A8B9CB"><!-- Blank bottom section --></td>
		<td width="13"><img src="../images/RoundCorners/B0C0D0-BR.png" width="13" height="13"></td>
	</tr>

</table>


<br>
<table align=center border=0>
<tr>
	<td>
		<table align=center border=0>
		<tr>
			<td></td>
		</tr>
		<tr>
			<td align="middle">
			<input type=Submit value="SEND MESSAGE">
			&nbsp;&nbsp; (Click once only - be patient)
			</td>
		</tr>
		</table>
				
	<br>
		
	<% If CBool(Request.Cookies("MMCookie")("ValidMemberLogin")) = True Then %>		
		<center style="font-size: smaller">
		You will receive a reply by email to 
		<font color=black style="BACKGROUND-COLOR: lavender">&nbsp;<% =Request.Cookies("MMCookie")("UserEmail") %>&nbsp;</font>.
		<br>If this is not your current email address you should first 
		<a href="http://www.mitmazel.com/MM/singles/ProfileEdit.asp#EmailAddress">update it</a> in your profile and then login again.
		</center>
	<% End If %>
				
	</td>
</tr>
</table>
      
</form>

<br><br>

            
</body>

</html>