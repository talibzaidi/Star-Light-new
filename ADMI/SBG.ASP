<%@ LANGUAGE = VBScript %>

<% 
    If (Session("Access") < "1") Then 
	Response.Redirect "login.asp"
    End If
%>

<html>

<head>

<title>Sanction - Version (Orange)</title>
</head>

<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0" link="#FFffff" vlink="#FFffff">




<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td bgcolor="#FFBD00"><font face="Arial"><img
        src="Simages/sanction.gif"
        width="330" height="82"></font></td>
        <td bgcolor="#FFBD00"><font face="Arial"></font>&nbsp;</td>
        <td align="right" bgcolor="#FFBD00"><font face="Arial"><a href="sanction.asp"><img
        src="Simages/homegif.GIF"
        width="84" height="82" border="0"></a></font></td>
    </tr>
    <tr>
        <td><font face="Arial"><img
        src="Simages/btcurve.gif"
        width="330" height="82"></font></td>
       
        <td colspan="2"><font size="2" face="Arial" color="#FFBD00"> CHANGE YOUR BANNER INFORMATION HERE: Then use Image Upload on the main menu to upload your banner. Banner Size [ 100 * (up to 200) pixels. ]  </font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2">



<font size="2" face="Arial" color="#FFBD00">Click <a href="maint.asp"><b>here</b></a> if you want to <b>edit an existing</b> banner. <br>
Or click <a href="newad.asp"><b>here</b></a> if you want to <b>add</b> a banner.   <br>
Or click <a href="listad.asp"><b>here</b></a> if you want to <b>list</b> banners. 
<br>
Or click <a href="delad.asp"><b>here</b></a> if you want to <b>delete</b> banners.   </font>
 </td>


			       
		 
		<tr>
                     <td valign="top" >
		    </td>
                     <td valign="top" colspan="2">
			
			


		   
                </tr>
		
		
 </table>


</center>
</font>&nbsp;</td>
    
	    
    </tr>
</table>
</body>
</html>
