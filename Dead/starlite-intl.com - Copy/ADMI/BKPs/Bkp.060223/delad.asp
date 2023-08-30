<%@ LANGUAGE = VBScript %>

<% 
    If (Session("Access") < "1") Then 
	Response.Redirect "login.asp"
    End If


Action = Left(UCase(Request("Action")),6)
mSubmitted = date & " " & time


If Action = "DELETE" Then
	              
		msg=""
	If msg = "" Then
'74 74 74 74 74 74 74 74 74 74 74 74 
	
	 SQL = "DELETE * FROM Banner WHERE Advertisement" + "= (" + ReQuest("AdNum") +")" 
	 Set conn = Server.CreateObject("ADODB.Connection")
    	 Conn.Open Session("ConnectionString")
    	 Conn.Execute(SQL)
	 Response.Redirect "sbg.asp" 
               
	 End If  'msg = ""
End If  'Action = "Submit"


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
        <td><font face="Arial"></font>&nbsp;</td>
        <td><font size="2" face="Arial" color="#FFBD00"> CHANGE YOUR BANNER INFORMATION HERE: Then use Image Upload off of main menu to upload your banner. Banner Size ( 100 * 58 pix. ) </font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2">



<p align="center"><font size="2" face="Arial" color="#FFBD00"><strong><big>Banner Maintenance</big></strong></font></p>

<p><font size="2" face="Arial" color="#FFBD00"><a href="newad.asp">Add a banner.</a></font></p>

<%If Session("Access") = 1 then %>
<form method="POST" action="delad.asp">
  <p>Edit an advertisement:<br>
  <select name="AdNum" size="1">
<%
  set rstAdvertisement = CreateObject("ADODB.Recordset")
  rstAdvertisement.Open "SELECT DISTINCT Advertisement, AName FROM Banner ORDER BY AName ASC", "DSN=STAREC1"

  do while not rstAdvertisement.EOF
    Response.Write("<option value=""" & rstAdvertisement("Advertisement") & """>" & rstAdvertisement("AName") & "&nbsp;</option>")
    rstAdvertisement.MoveNext
  loop

  rstAdvertisement.Close
  set rstAdvertisement = nothing
%>  </select><input name="Action" type="submit" value="Delete"></p>
</form>
<% End If %>




<p><font size="2" face="Arial" color="#FFBD00"><a href="listad.asp">List banners.</a></font></p>




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
