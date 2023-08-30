<%@ LANGUAGE = VBScript %>
<!--#include file="ADOVBS.INC"-->
<% 
    If (Session("Access") <> "1") Then 
	Response.Redirect "login.asp"
    End If
%>
<%
msg=""

Action = Left(UCase(Request("Action")),6)
mSubmitted = date & " " & time


If Action = "SUBMIT" Then
	msg=""
	If msg = "" Then     
		SQL = "INSERT INTO SubArea ( Subname, AID, Subvisible) "
       'SQL = SQL & " VALUES('"& Request("SubArea") & "' ,'"  & Request("AID") &   "' )"		
       SQL = SQL & " VALUES('"& Request("SubArea") & "' ,'"  & Request("AID") &   "' ,1)"		
		Set conn = Server.CreateObject("ADODB.Connection")
		Response.Write "Session('ConnectionString') = " & Session("ConnectionString")
    	Conn.Open Session("ConnectionString")
    	Conn.Execute(SQL)
		Response.Redirect "sanctionproduct.asp" 
	End If  'msg = ""
	
Elseif  Action = "DELETE" Then
	msg=""
	If msg = "" Then
	 SQL = "DELETE * FROM SubArea WHERE SID" + "= (" + ReQuest("SID") +")" 
	 Set conn = Server.CreateObject("ADODB.Connection")
    	 Conn.Open Session("ConnectionString")
    	 Conn.Execute(SQL)
	 Response.Redirect "sanctionproduct.asp" 
	 End If  'msg = ""
	 
	 
Elseif  Action = "UPDATE" Then
	 msg=""
	 If msg = "" Then

	 Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT * FROM SubArea " & _
	 "WHERE SID =" & Request("SAID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 
 	 ' Update record
	 rst("Subname") = Request("EArea")
 	
	
	
                rst.Update

	rst.Close
	Conn.Close



	 Response.Redirect "sanctionsubarea.asp" 
               
	 End If  'msg = ""
Elseif  Action = "DOTEXT" Then

	 msg=""
	 If msg = "" Then

	 Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT * FROM SubArea " & _
	 "WHERE SID =" & Request("DAID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 


 	 ' Update record
	 rst("SubDesc") = Request("DArea")
 	
	
	
                rst.Update

	rst.Close
	Conn.Close



	 Response.Redirect "sanctionsubarea.asp" 
               
	 End If  'msg = ""


End If  'Action = "Submit"
%>


<html>




<head>
<title>Sanction - Version (Orange)</title>
</head>




<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">

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
        <td><font color="#FFBD00" size="2" face="Arial">Add a <b>Sub Area</b> to the online catalogue. <u>Changes are immediate</u>. Please create at least one  product after creation. Selecting a <b>Sub-Area</b> from the drop down box and clicking the 'delete' button will permanently remove the <b>Area</b>.  <u>Remember to select an Area to put the Sub-Area into.</u></font></td>
    </tr>
    <tr>
        <td valign=top><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2'><font face="Arial">





<FORM ACTION="sanctionsubarea.asp" METHOD=POST>

<table border="0" cellpadding="3" cellspacing="0"  ALIGN="CENTER">
               
		<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Sub-Area (Spell correctly!):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="SubArea" value="<%=Request("SubArea")%>">
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Submit" ></td>
                </tr>
  <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Area Select:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select AreaName, AID from Area51 ORDER BY AreaName ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="AID" size="1" >
			<option>
			<%Do While Not RS.EOF%>
			<option value="<%=RS("AID")%><%=Request("AID")%>">
			<font face ="arial" size="1"><%=RS("AreaName")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td ></td>
                </tr>
<tr><td colspan=2><hr></td></tr>
                <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Sub-Area Delete:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select Subname, SID from SubArea ORDER BY Subname ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="SID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("SID")%><%=Request("SID")%>">
			<font face ="arial" size="1"><%=RS("Subname")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Delete" ></td>
                </tr>
      <tr><td colspan=2><hr></td></tr>         
 <tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Sub-Area Change (Spell correctly!):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="EArea" value="<%=Request("Area")%>">
		    </td><td ></td>
                </tr>
              <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Sub-Area Edit:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select Distinct Subname, SID from SubArea ORDER BY Subname ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="SAID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("SID")%><%=Request("SAID")%>">
			<font face ="arial" size="1"><%=RS("Subname")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Update" ></td>
                </tr>
		  <tr><td colspan=2><hr></td></tr>           
		 <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>SubArea Description:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select Distinct Subname, SID from SubArea ORDER BY Subname ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="DAID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("SID")%><%=Request("DAID")%>">
			<font face ="arial" size="1"><%=RS("Subname")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="DoText" ></td>
                </tr>
<tr>
      
              <td ><font size="2" face="Arial" color="#FFBD00">Sub Area Text (Spell Correctly!):</font>
		    </td>

                    <td width="100"><font size="2" face="Arial"></font><textarea cols="25" rows="5" name="DArea" value="<%=Request("DArea")%>"></textarea>
		    </td><td ></td>
                </tr>
            


</table>
    </td><td ></td>
    </tr>
</table>

</FORM>


</font>&nbsp;</td>
        
    </tr>
</table>
</body>
</html>
























