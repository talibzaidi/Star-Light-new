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

<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0" link="#FF0000" vlink="#FF0000">




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

<p><font size="1" face="Arial" color="#FFBD00"><a href="newad.asp">Add a banner.</a></font></p>


<table border="2" width=350>
  <tr>
    <th bgcolor="#000000"><font size="2" face="Arial" color="#FFBD00"><B>BANNER LIST:<B/></FONT></th>
    
  </tr>
<%

If Session("Access") = 1 then


  SQL =  "SELECT * FROM Banner "
  Set conn = Server.CreateObject("ADODB.Connection")
  Conn.Open Session("ConnectionString")
     Set RS = Conn.Execute(SQL)
  
             
  Dim StrBuf 
  StrBuf= ""
 

  LastAd = 0
  
  do while not RS.EOF
    if RS("Advertisement") <> LastAd then
      LastAd = RS("Advertisement")
      Response.Write("<tr><td valign=""top"">")
      StrBuf = "Not Visible"
      if RS("AShow") then
       StrBuf = "Visible"
      end if
     
      Response.Write("<font face=arial size=1 color=ffffff><b>" & RS("ALocation") & "</b></font><br>")
     
      Response.Write("<font face=arial size=1 color=ffbd00><img src=" & RS("AImage") & " " & "></font><br>")
      Response.Write("<font face=arial size=1 color=ffffff>URL: &nbsp;&nbsp;&nbsp;&nbsp;" & RS("AURL") & "</font><br>")
      Response.Write("<font face=arial size=1 color=ff0000>" & "<a href=""editad.asp?AdNum=" & LastAd & """>Edit Banner</a></font>")
      Response.Write("</td>")
      
      Response.Write("</tr>")
      RS.MoveNext
    end if
  loop

  RS.Close
  set RS = nothing


End If

%>
</table>




<p><font size="1" face="Arial" color="#FFBD00"><a href="listad.asp">List bannerss</a></font></p>




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
