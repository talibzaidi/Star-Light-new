<%@ LANGUAGE = VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->


<html>

<head>
    <link rel="stylesheet" type="text/css" href="../Misc/StyleSheet1.css"> <% ' 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. %>
    <title>Contact Us</title>
    <meta name="description" content="Large selection of GPS, GPS sensors, GPS accessories, GPS OEM, tracking gps, bluetooth gps, fish finders, sounders, CB radios and Walky-talky, flash memory, radio scanners, digital cameras, car audio, dash cam, night vision optics">
    <!-- [11/3/20, BN:] The following <script> line is re Google's reCaptcha version 2, from https://developers.google.com/recaptcha/docs/display#auto_render -->
    <!-- [11/3/20, BN:] I decided to comment out that <script> line because I reached a dead end with the <form> below, just after the <body> tag.
    <script src="https://www.google.com/recaptcha/api.js" async defer></script>
    -->
</head>


<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0"> 


<!-- [11/3/20, BN:] The following <form> is re Google's reCaptcha version 2, from https://developers.google.com/recaptcha/docs/display#auto_render -->
<!-- The value for attribute data-sitekey below was obtained by registering (creating an account for) our site starlite-intl.com at https://www.google.com/recaptcha/admin/create -->
<!-- As part of doing that at https://www.google.com/recaptcha/admin/create, I created a label "starlite-intl.com label". I'm not sure what that's for. -->
<!-- [11/3/20, BN:] I decided to comment out the following form because I could not see what to do next. Youtube was useless.
     I need to Google how to proceed using classic ASP. This URL looks promising:
     https://developers.google.com/recaptcha/old/docs/asp = Using reCAPTCHA with Classic ASP.
     See also my "Starlite Work Log" Excel file, under the References tab.
<form action="?" method="POST">
  <div class="g-recaptcha" data-sitekey="6LeLbt4ZAAAAAEyuyokgIVPqKlbOlDJ7uQzECciE"></div>
  <br/>
  <input type="submit" value="Submit">
</form>
-->


<table bordercolor="green" border="0" align="center">

<tbody><tr><td>
<% InArea = "ContactUs" %>
<!--#include virtual="Misc/Header.INC"-->

<table width="398" height="349" cellspacing="20" cellpadding="0" border="0" align="center">
<tbody><tr><td><center style="color: rgb(0, 0, 0); font-size: medium; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; text-indent: 0px; text-transform: none; white-space: normal; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><font size="+1" color="blue" face="Tahoma">Contact us by mail, phone, fax or email</font></center>
<table style="color: rgb(0, 0, 0); font-size: medium; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;" width="322" height="268" bordercolor="green" border="0" align="center">
<tbody><tr><td xxxalign="center"><font face="Tahoma"><b>Star Lite International LLC</b><span>&nbsp;</span><br>
P.O. Box 965<span>&nbsp;</span><br>
Southfield, MI 48037- 0965<span>&nbsp;</span><br>
USA</font>
<p><font face="Tahoma"><b>S.L.I. Corporation</b><span>&nbsp;</span><br>
P.O. Box 23030, Devonshire Mall<span>&nbsp;</span><br>
Windsor, Ontario N8X 5B5<span>&nbsp;</span><br>
Canada</font></p>
<p><font face="Tahoma">Tel. 1-800-387-8535 (Order line)<span>&nbsp;</span><br>
Tel. 248-546-4489<span>&nbsp;</span><br>
Fax. 248-546-1462</font></p>
</td></tr></tbody>
</table>


<center style="color: rgb(0, 0, 0); font-size: medium; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; text-indent: 0px; text-transform: none; white-space: normal; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;">
    <br /><font size="+1" color="blue" face="Tahoma">Submit this form to email us</font>
</center>

</td></tr></tbody>
</table>
</td>
</tr>
</tbody>
</table>

<form name="ContactUsForm" method="post" action="../AskStarlite/EmailSend.asp" onsubmit="javascript:return WebForm_OnSubmit();"><table width="407" height="379" cellspacing="2" cellpadding="3" border="0" align="center"><tbody><tr><td align="right"><font face="Tahoma">Full Name<span style="color:Red;">*</span></font></td>
<td><font face="Tahoma"><input name="UserFullName" type="text"></font></td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="right"><font face="Tahoma">Email Address<span style="color:Red;">*</span></font></td>
<td><font face="Tahoma"><input name="UserEmailAddress" type="text"></font></td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="right"><font face="Tahoma">Full Address&nbsp;&nbsp;</font></td>
<td><font face="Tahoma"><input name="UserAddress" type="text"></font></td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="right"><font face="Tahoma">Phone Number&nbsp;&nbsp;</font></td>
<td><font face="Tahoma"><input name="UserPhone" type="text"></font></td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="right"><font face="Tahoma">Subject&nbsp;&nbsp;</font></td>
<td><font face="Tahoma"><input name="EmailSubject" type="text"></font></td>
<td>&nbsp;</td>
</tr>
<tr> 
<td valign="top" align="right"><font face="Tahoma">Enquiry<span style="color:Red;">*</span></font></td>
<td><font face="Tahoma"><textarea name="UserEmailEnquiryToSL" rows="4" cols="30"></textarea></font></td>
<td>&nbsp;</td>
</tr>
<tr>
<td><font face="Tahoma"><span style="color:Red;">*</span> required field</font></td>
<td align="right">
<font face="Tahoma"><input value="Reset" type="reset">
<input value="Submit" type="submit"></font>
</td><td>&nbsp;</td>
</tr>
</tbody>
</table>

</form>

<br><br>


<br>



<br>
</body></html>


