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
        <td align="right" bgcolor="#FFBD00"><font face="Arial" color="#FFBD00"><a href="sanction.asp"><img
        src="Simages/homegif.GIF"
        width="84" height="82" border="0"></a></font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/btcurve.gif"
        width="330" height="82"></font></td>
        <td><font face="Arial"></font>&nbsp;</td>
        <td><font size="2" face="Arial"> ENTER YOUR BANNER INFORMATION HERE: Then use Image Upload off of main menu to upload your banner. Banner Size ( 100 * 58 pix. ) Or click <a href="editad.asp"><b>here</b></a> if you want to edit an existing banner.  </font></td>
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

  adtemp = Request("AdNum")
  set RS = CreateObject("ADODB.Recordset")
  RS.Open "SELECT * FROM Banner WHERE Advertisement=" & adtemp, "DSN=STAREC1"
%>	
					
		  
  <p><font face="Arial" size="2" color="#FFBD00">Ad Name: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><input type="text" name="AName" size="40" maxlength="40" value="<%=RS("AName")%>"><br>
  
<br>
  <font face="Arial" size="2" color="#FFBD00">Area: &nbsp;&nbsp;
<% If (Session("Access") = 1) Then%>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<select name="Location" size="1" width="25" >
 			<option value="<%=RS("ALocation")%>"><%=RS("ALocation")%></option>
			<%
		 SQL = "Select Subname from Subarea ORDER BY Subname ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			<%Do While Not RS.EOF%>
			<option value="Home">Home</option>
	<option value="Classified">Classified</option>
	<option value="Links">Links</option>
			<option value="<%=RS("Subname")%><%=Request("Subname")%>">
			<font face ="arial" size="1"><%=RS("Subname")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
			<%End If%>

			

  </font>

<%
  set RS = CreateObject("ADODB.Recordset")
  RS.Open "SELECT * FROM Banner WHERE Advertisement=" & adtemp, "DSN=STAREC1"
%>

  <br>
  <font face="Arial" size="2" color="#FFBD00">Ad Visible?&nbsp;&nbsp; <input type="checkbox" name="Show" value="T" <% if RS("AShow") then Response.Write(" checked") %>><br>
  <font face="Arial" size="2" color="#FFBD00">Ad URL: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><input type="text" name="URL" size="40" maxlength="128" value="<%=RS("AURL")%>"><br>
  <font face="Arial" size="2" color="#FFBD00">Image URL: &nbsp;&nbsp;</font><input type="text" name="Image" size="40" maxlength="128" value="<%=RS("AImage")%>"><br>
  <br>
  <input type="submit" value="Add"> <input type="reset" value="Reset"></p>
  <input type="hidden" name="AdNum" value="<% = Request("AdNum") %>">
</form>


</td>
       
    </tr>
</table>
</body>
</html>
