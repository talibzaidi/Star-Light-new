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

<table border="0" cellpadding="5" cellspacing="0" width="100%">

    <tr>
        <td bgcolor="#FFBD00">
        <font face="Arial"><img src="Simages/sanction.gif" width="330" height="82"></font>
        </td>
    </tr>
    
    <tr>
        <td>
        <font size="2" face="Arial">Welcome to Sanction: The Active web tool for mission critical administration,
        developed for use by <a href="http://www.ontbiz.com/iac">I.A.C.</a> 
        by <a href="http://www.anyperson.com/tds">AnyPerson Dot Com</a>. All access to this site is logged.
        </font>
        </td>
    </tr>
    
    <tr> 
		<td><br>
		<table align=center>
		<tr>
        <td valign="top" align="left">
        
        <a href="sanctionarea.asp">
        <font color="#FFBD00" face="Arial" size="2" color="#FFBD00">I want to create or edit an <b>Area</b></font>
        </a>
                
        <br>        
		<a href="sanctionsubarea.asp">
		<font color="#FFBD00" face="Arial" size="2" color="#FFBD00">I want to create a <b>Sub-Area</b></font>
		</a>
		
		<br>	
        <a href="sanctionproduct.asp">
        <font color="#FFBD00" face="Arial" size="2" color="#FFBD00">I want to edit a <b>Product</b></font>
        </a>
        <br>	
        
		<a href="sanctionglobal.asp">
		<font color="#FFBD00" face="Arial" size="2" color="#FFBD00">I want to effect a <b>Global Price Change</b></font>
		</a>
		<br>	
	
		<a href="sanctionglobalduty.asp">
		<font color="#FFBD00" face="Arial" size="2" color="#FFBD00">I want to effect a <b>Global Duty Change</b></font>
		</a>
		<br>	

		<a href="sanctionglobalgpm.asp">
		<font color="#FFBD00" face="Arial" size="2" color="#FFBD00">I want to effect a <b>Global GPM Change</b></font>
		</a>
		<br>	

		<a href="supload.asp">
		<font color="#FFBD00" face="Arial" size="2">I want to upload an <b>Image</b>.</font>
		</a>
		<br>
		<br>
		 
		<a href="sclass.asp">
		<font color="#FFBD00" face="Arial" size="2">I want to approve <b>classified</b> ads. </font>
		</a>
		<br>
		
		<a href="sanctionrates.asp">
		<font color="#FFBD00" face="Arial" size="2" color="#FFBD00">I want to edit <b>Rates</b></font>
		</a>
		<br>	 
      
		<a href="sbg.asp"><font color="#FFBD00" face="Arial" size="2">I want to edit <b>banners</b>. </font>
		</a>
		<br>
		
		<a href="sanctionlinks.asp">
		<font color="#FFBD00" face="Arial" size="2">I want to edit <b>links</b>. </font>
		</a>
		<br>
		<br>
		   
		<a href="sanctionadmin.asp">
		<font color="#FFBD00" face="Arial" size="2">I want to edit the <b>Administrator</b>.</font>
		</a>
		<br>
		<br>
		
		<a href="selp/sanction.asp" target="_new">
		<font color="#FFBD00" face="Arial" size="2">Access <b>help</b>.</font>
		</a>
		<br><br>
		
		<a href="slog.asp">
		<font color="#FFBD00" face="Arial" size="2"> <b>Logout</b></font>
		</a>
		
		</td>
		</tr>
		</table>
		</td>
       
    </tr>
</table>

</body>

</html>
