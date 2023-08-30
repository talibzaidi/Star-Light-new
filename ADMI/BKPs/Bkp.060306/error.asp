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

<body bgcolor="#000000" text="#FFFFFF" link="#FFBD00" vlink="#FFBD00" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">

<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td bgcolor="#FFBD00"><font face="Arial"><img
        src="Simages/sanction.gif"
        width="330" height="82"></font></td>
        <td bgcolor="#FFBD00"><font face="Arial"></font>&nbsp;</td>
        <td align="right" bgcolor="#FFBD00"><font face="Arial" color="#FFBD00">=)</font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/btcurve.gif"
        width="330" height="82"></font></td>
        <td><font face="Arial"></font>&nbsp;</td>
        <td><font size="2" face="Arial">AN ERROR HAS OCCURRED A RESULT OF INCORRECT INPUT . </font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2" valign="top" align="left"><br><font face="Arial" size="2" color="#FFBD00"></font><p align="left"><br>
                
		
                
		
	

		<font color="#FFBD00"
                face="Arial" size="2"><b>ERROR:  YOU CANNOT HAVE TWO PRODUCTS WITH THE SAME PRODUCT ID. 
                PLEASE HIT YOUR BACK BUTTON TO CORRECT THE ERROR. 
                MAKE SURE THAT YOU HAVE NOT ENTERED ANY INVALID CHARACTERS INTO 
                A PRICE FIELD OR PRODUCT ID FIELD <br>(i.e. ~!@#$%^&* etc. )</b> </font></p></td>
    
    </tr>
</table>
</body>
</html>