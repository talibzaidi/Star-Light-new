<%@ LANGUAGE = VBScript %>

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
        <td align="right" bgcolor="#FFBD00"><font face="Arial"><a href="sanction.asp"><img
        src="Simages/homegif.GIF"
        width="84" height="82" border="0"></a></font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/btcurve.gif"
        width="330" height="82"></font></td>
        <td><font face="Arial"></font>&nbsp;</td>
        <td><font size="2" face="Arial"> ENTER YOUR BANNER INFORMATION HERE: Then use Image Upload off of main menu to upload your banner. Banner Size [ 350 *60 pix. ] Or click <a href="maint.asp"><b>here</b></a> if you want to edit an existing banner. Banners in shopping cart
ie: Area/Subarea must have <b>https://</b> instead of <i>http://</i> </font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2" valign="top" align="left">

<br>
<font face="Arial" size="2" color="#FFBD00"></font>

<form method="POST" action="savead.asp">
<%
		
		 SQL = "Select Distinct SubName, SID from SubArea ORDER BY Subname ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		
					
		  
  <p>
  <font face="Arial" size="2" color="#FFBD00">SubArea: 
  <% If (Session("Access") = 1) Then%>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<select name="Location" size="1" >
			<%Do While Not RS.EOF%>
			

	<option value="Home">Home</option>
	<option value="Classified">Classified</option>
	<option value="Links">Links</option>
<option value="<%=RS("Subname")%><%=Request("Location")%>">
			<font face ="arial" size="1"><%=RS("Subname")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			
			</select>
			<%End If%>
			
<br>
 


  <font face="Arial" size="2" color="#FFBD00">Ad Visible? &nbsp;&nbsp;</font><input type="checkbox" name="Show" value="T" checked><br>
   <font face="Arial" size="2" color="#FFBD00">Ad Name: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><input type="text" name="AName" size="40" maxlength="128"><br>
  <font face="Arial" size="2" color="#FFBD00">Ad URL: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><input type="text" name="URL" size="40" maxlength="128"><br>
  <font face="Arial" size="2" color="#FFBD00">Image URL: &nbsp;&nbsp;</font><input type="text" name="Image" size="40" maxlength="128" value="http://www.starlite-intl.com/imi/***.gif"><br>
  <br>
  <input type="submit" value="Add"> <input type="reset" value="Reset"></p>
</form>


</td>
       
    </tr>
</table>
</body>
</html>
