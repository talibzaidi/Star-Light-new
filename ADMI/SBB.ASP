<%@ LANGUAGE = VBScript %>
<!--#include file="ADOVBS.INC"-->
<% 
    If (Session("Access") < "1") Then 
	Response.Redirect "login.asp"
    End If
%>
<%
msg=""

Action = Left(UCase(Request("Action")),6)



If Action = "SUBMIT" Then
	If Request("City") = "" OR _
		Request("Title") = "" Then
		msg="All fields must have non-empty responses."
	End If

	If msg = "" Then
	
              
		 SQL = "INSERT INTO COCbullet ( City, Area, Name, BannerURL, CoName, Title, Text1, Image1 ) "
                                 SQL = SQL & " VALUES( '" + Request("City") + "','" + Request("Area") + "','" + Request("Name")+ "','" + Request("BannerURL")+ "','" + Request("CoName")+ "','" + Request("Title") + "','" & Request("Text1") & "','" & Request("Image1") & " ' " & " )"
			
		 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
 		
    		 Conn.Execute(SQL)
		 Response.Redirect "sanction.asp" 
	End If
End If  'Action = "Submit"
%>
<html>

<head>

<title>Sanction - Version (Orange)</title>
</head>

<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0" link="#FFffff" vlink="#FFffff">



<FORM ACTION="sbb.asp" METHOD="POST">
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
        <td><font size="2" face="Arial" color="#FFBD00"> ENTER YOUR BULLETIN INFORMATION HERE:  Or click <a href="sbbedit.asp"><b>here</b></a> if you want to edit an existing bulletin. Click <a href="sbbdel.asp"><b>here</b></a> if you want to delete a bulletin. The Bulletin Title is what you see when the bulletin is listed!  </font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2"><font face="Arial"><center>


<INPUT type="submit"  align="bottom"  NAME="Action" VALUE="Submit" > <font face="arial" size="2" color="#FFBD00"><strong><%=msg%></font>
        
		<table border="0" cellpadding="3" cellspacing="0" width="80%" ALIGN="CENTER"><br>

			<%
		 If (Session("Access") = 3) Then
		 SQL = "Select Distinct CITYname from COCcity ORDER BY CITYname ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		 End If
%>
			       
		<tr>
                     <td valign="top" ><font size="2" face="Arial" color="#FFBD00">City :</font>
		    </td>
                     <td valign="top" width="100">
			<% If (Session("Access") = 3) Then%>
			<select name="City" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("CITYname")%><%=Request("City")%>">
			<font face ="arial" size="1"><%=RS("CITYname")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			<option value="Ontario<%=Request("City")%>">
			<font face ="arial" size="1">Ontario</font>
			</option>
			</select>
			<%End If%>
			<% If (Session("Access") = 2) Then%>
			<b><%=Session("ccCITY")%></b><INPUT TYPE=hidden NAME="City" VALUE="<%=Session("ccCITY")%>">
			<%End If%>
			
		    </td> <td valign="top" ></td>
                </tr>
		<%
		 If (Session("Access") = 3) Then
		 SQL = "Select Distinct Name from CoC_Main Where (isChamber = yes or isMember = yes)ORDER BY Name ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		 End If
%>
			       
		<tr>
                     <td valign="top" ><font size="2" face="Arial" color="#FFBD00">Company :</font>
		    </td>
                     <td valign="top" width="100">
			<% If (Session("Access") = 3) Then%>
			<select name="Name" size="1" width="25" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("Name")%><%=Request("Name")%>">
			<font face ="arial" size="1"><%=RS("Name")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
			<%End If%>
			<% If (Session("Access") = 2) Then
			   SQL = "Select Distinct Name from CoC_Main WHERE CITY LIKE '" + Session("ccCITY") + "' and (isChamber = yes or isMember =yes) ORDER BY Name ASC"
                	   Set conn = Server.CreateObject("ADODB.Connection")
    			   Conn.Open Session("ConnectionString")
    			   Set RS = Conn.Execute(SQL)	%>
			<select name="Name" size="1" width="25">
			<%Do While Not RS.EOF%>
			<option value="<%=RS("Name")%><%=Request("Name")%>">
			<font face ="arial" size="1"><%=RS("Name")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
			<%End If%>
			
		    </td> <td valign="top" ></td>
                </tr>
               			       
		<tr>
                     <td valign="top" ><font size="2" face="Arial" color="#FFBD00">Area :</font>
		    </td>
                     <td valign="top" width="100">
			<% If (Session("Access") = 3) Then%>
			<select name="Area" size="1" width="25" >
			
			<option value="Event">Event</option>
			<option value="Message">Message</option>
			<option value="Link">Link</option>
			
			
			</select>
			<%End If%>

			<% If (Session("Access") = 2) Then  %>
			<select name="Area" size="1" width="25" >
			
			<option value="Event">Event</option>
			<option value="Message">Message</option>
			<option value="Link">Link</option>
			
			
			
			
			</select>
			<%End If%>
			
		    </td> <td valign="top" ></td>
                </tr>
		<tr>
                     <td valign="top"><font size="2" face="Arial" color="#FFBD00">Bulletin or Link  URL (if no text is specified!):</font>
	            </td>
                     <td valign="top" ><input type="text" size="35" name="BannerURL" value="">
		    </td> <td valign="top" ></td>
                </tr>
		<tr>
                     <td valign="top"><font size="2" face="Arial" color="#FFBD00">Title:</font>
	            </td>
                     <td valign="top" ><input type="text" size="35" name="Title" value="same as client add<%=Request("Title")%>">
		    </td> <td valign="top" ></td>
                </tr>
	<tr>
                     <td valign="top"><font size="2" face="Arial" color="#FFBD00">Image Only:</font>
	            </td>
                     <td valign="top" ><input type="text" size="35" name="Image1" value="<%=Request("Image1")%>">
		    </td> <td valign="top" ></td>
                </tr>
	 <tr>
                     <td valign="top" valign="top"><font size="2" face="Arial" color="#FFBD00">Text Option:</font>
		    </td>
                     <td valign="top" width="100">
                    <textarea name="Text1" rows="5" cols="38" wrap="virtual" value="<%=Request("Text1")%>"></textarea>
		    </td>
                </tr>
		<tr>
                     <td valign="top"><font size="2" face="Arial" color="#FFBD00">Edited By:</font>
	            </td>
                     <td valign="top" ><b><%=Session("User")%></b><input type="hidden" size="35" name="Name" value="<%=Session("User")%>">
		    </td> <td valign="top" ></td>
                </tr>
		
 </table>

</form>
</center>
</font>&nbsp;</td>
    
	    
    </tr>
</table>
</body>
</html>
