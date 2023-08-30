<%@ LANGUAGE = VBScript %>

<% 
    If (Session("Access") < "1") Then  
		Response.Redirect "login.asp"
    End If
%>

<% Msg = Request.QueryString("Msg") %>


<html>

<head>
<title>Sanction - Version (Orange)</title>
</head>


<body bgcolor="#000000" text="#FFFFFF" link="#FFBD00" vlink="#FFBD00" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">


<% If Msg <> "" Then %>
	<center>
		<font size='4' face="Tahoma" color='white'><%=Msg%></font>
	</center>
<% End If %>


<table border="0" cellpadding="5" cellspacing="0" width="100%">

    <tr>
        <td bgcolor="#FFBD00">
			<a href="Sanction.asp">
			<font face="Arial"><img src="Simages/sanction.gif" width="330" height="82"></font>
			</a>
        </td>
    </tr>
    
    <tr>
        <td>
        <font size="2" face="Arial">Welcome to Sanction: The database administration tool.
        </font>
        </td>
    </tr>
    
    <tr> 
		<td><br>
		<table border="0" align="center" cellpadding="5">
		<tr>
			<td valign="top" align="left">
			<font face="Arial" size="2" color="#FFBD00">
			<b>Area:</b> 
			</font>
			</td>
			<td>
			<font face="Arial" size="2" color="#FFBD00">
			<a href="sanctionarea.asp?ToDo=CreateArea">
			Create 
			</a>, 
			<a href="sanctionarea.asp?ToDo=EditArea">
			Edit
			</a>, 
			<a href="sanctionarea.asp?ToDo=DeleteArea">
			Delete
			</a>
			</font>
			</td>
		</tr>
        
        <tr>
			<td>        
			<font face="Arial" size="2" color="#FFBD00"><b>Sub-Area:</b></font>
			</td>
			<td>
			<a href="sanctionsubarea.asp">
			<font face="Arial" size="2" color="#FFBD00">Create</font>
			</a>
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Arial" size="2" color="#FFBD00"><b>Product:</b></font>
			</td>

			<td>
				<font face="Arial" size="2" color="#FFBD00">
					<a href="SanctionProduct.asp?Action=CreateNewProduct">Create</a>,&nbsp;
					<a href="SanctionProduct.asp?Action=EditOldProduct">Edit</a>,&nbsp;
					<a href="SanctionProduct.asp?Action=DeleteOldProduct">Delete</a>,&nbsp;
					<a href="http://www.starlite-intl.com/Admin2/login.asp?Pwd=787szd&btnSubmit=Submit">Extra Admin</a>
				</font>
			</td>
        </tr>
        
        <tr>
			<td>
			<a href="sanctionglobal.asp">
			<font face="Arial" size="2" color="#FFBD00">I want to effect a <b>Global Price Change</b></font>
			</a>
			<br />
	
			<a href="sanctionglobalduty.asp">
			<font face="Arial" size="2" color="#FFBD00">I want to effect a <b>Global Duty Change</b></font>
			</a>
			<br>	

			<a href="sanctionglobalgpm.asp">
			<font face="Arial" size="2" color="#FFBD00">I want to effect a <b>Global GPM Change</b></font>
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
			<font face="Arial" size="2" color="#FFBD00">I want to edit <b>Rates</b></font>
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
