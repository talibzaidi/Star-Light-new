<% @ LANGUAGE = VBScript %>

<!doctype html> 

<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<html>

<head>  
	<link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/mobile1/Misc/StyleSheet1.css"> <% ' 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. %> 
	<title>Contact</title>
	<meta http-equiv="content-type" content="text/html; charset=ISO-8859-1">

	<meta name="viewport" content="width=device-width; initial-scale=1.0">
    <!-- foneFrame.css is the stylesheet with comments, so it is readable.
	     foneFrame-min.css is the minimized version; it is smaller and loads faster. -->
	<link href="https://www.starlite-intl.com/mobile1/foneFrame.css" rel="stylesheet" type="text/css">
	<!-- The following 2 lines are not strict HTML5. -->
	<meta name="HandheldFriendly" content="true"/>
	<meta name="MobileOptimized" content="320"/>

    <!-- 11/10/13: For the accordion menu from menucool.com, where its HTML is in a separate file, and does not have to be repeated in each webpage that has the menu. -->
    <link href="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.css" rel="stylesheet" type="text/css" />
    <script src="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.js" type="text/javascript"></script>
	<script type="text/javascript">amenu.close(true);</script>
</head>




<body bgcolor="white">

<% XXXInArea = "Terms" %>

<!-- [BN, 11/23/17] The width="100%" in the next line is necessary to get the header of this Contact Us page 
     to expand to the full width of the smart phone screen, even though width="100%" in not needed in the other 
     pages of this mobile site. I do not know why it isn't needed in at least some of those other pages too,
     because the content on some of them, like the Shopping Cart page, is not sufficient to induce the header to 
     expand to the full width of the smartphone screen.
     -->
<table style="border:0px solid green;" width="100%" bgcolor="" align='center'>		<% ' Start Table 1 %>
<tr><td><!-- #include virtual="mobile1/Misc/Header.INC" -->

<center><font color='blue' size='+1'><br />Contact us by mail, phone, fax or email</font></center>


<table align="center" border="0" bordercolor="green">

<tr>
<td XXXalign="center">

<% 
'Response.Write "<table width=" & PageWidth & " XXXbgcolor=lightblue align=left cellpadding=10>"
Response.Write "<table style='width:100%;' XXXbgcolor=lightblue align=left cellpadding=10>"
%>


<tr>
<td>
	<br />
	<b>Star Lite International LLC</b> 
	<br>P.O. Box 965
	<br />Southfield, MI 48037- 0965
    <br />USA

	<!-- <p><b>S.L.I. Corporation</b>
    <br />P.O. Box 23030, Devonshire Mall
	<br>Windsor, Ontario N8X 5B5
    <br />Canada
	</p> -->

	<p>
	Tel. 1-800-387-8535 (Order line)
	<br>
	Tel. 248-546-4489
	<br>
	Fax. 248-546-1462
	</p>
	
	<p>
	Email us at <a href="mailto:starlite@starlite-intl.com">starlite@starlite-intl.com</a>
	</p>
</td>
</tr>
</table>


</td>
</tr>

<tr>
<td>

</td>
</tr>
</table>


<center><font color='blue' size='+1'>Submit this form to email us</font></center>

<form name="ContactUsForm" method="post" 
    action="../AskStarlite/EmailSend.asp" onsubmit="javascript:return WebForm_OnSubmit();" >
<!-- This table is based on that in the Contact Us page at http://usglobalsat.com/ContactUs.aspx -->
<table border="0" cellspacing="2" cellpadding="3" align='center'>
                <tr>
                    <td>&nbsp;</td><td colspan="2"></td>
                </tr>
                 <tr>
                    <td align="right">Full Name<span style="color:Red;">*</span></td>
                    <td><input name="UserFullName" type="text" /></td>
                    <td>&nbsp;</td>
                 </tr>
                 <tr>
                    <td align="right">Email Address<span style="color:Red;">*</span></td>
                    <td><input name="UserEmailAddress" type="text" /></td>
                    <td>&nbsp;</td>
                 </tr>
                 <tr>
                    <td align="right">Full Address&nbsp;&nbsp;</td>
                    <td><input name="UserAddress" type="text" /></td>
                    <td>&nbsp;</td>
                 </tr>
                 <tr>
                    <td align="right">Phone Number&nbsp;&nbsp;</td>
                    <td><input name="UserPhone" type="text" /></td>
                    <td>&nbsp;</td>
                 </tr>
                 <tr>
                    <td align="right">Subject&nbsp;&nbsp;</td>
                    <td><input name="EmailSubject" type="text" /></td>
                    <td>&nbsp;</td>
                 </tr>
                 <tr>
                    <td align="right" valign="top">Enquiry<span style="color:Red;">*</span></td>
                    <td><textarea name="UserEmailEnquiryToSL" rows="4" cols="30"></textarea>
                        
                    </td>
                    <td>&nbsp;</td>
                 </tr>
                 <tr>
                    <td><span style="color:Red;">*</span><small>&nbsp;required field</small></td>
                    <td align="right">
                        <input type="reset" value="Reset">
                        <input type="submit" value="Submit">
                     <td>&nbsp;</td>
                 </tr>
                 </table>
</form>

<br><br>


</body>


</html>

