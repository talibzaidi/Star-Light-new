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
                 AreaName = Replace( ReQuest("AreaName") , " ", "%20") 
		 Area = Replace( ReQuest("Area") , " ", "%20") 
		 SQL = "INSERT INTO Area51 ( AreaName ) "
                 SQL = SQL & " VALUES('"& Request("Area") &"' )"
		
		
		 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Conn.Execute(SQL)
		 Response.Redirect "sanctionsubarea.asp" 
	End If  'msg = ""
Elseif  Action = "DELETE" Then
	msg=""
	If msg = "" Then

	 SQL = "DELETE * FROM Area51 WHERE AID" + "= (" + ReQuest("AID") +")" 
	 Set conn = Server.CreateObject("ADODB.Connection")
    	 Conn.Open Session("ConnectionString")
    	 Conn.Execute(SQL)
	 Response.Redirect "sanctionsubarea.asp" 
               
	 End If  'msg = ""
Elseif  Action = "UPDATE" Then

	 msg=""
	 If msg = "" Then

	 Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT * FROM Area51 " & _
	 "WHERE AID =" & Request("EAID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 


 	 ' Update record
	 rst("AreaName") = Request("EArea")
 	
	
	
                rst.Update

	rst.Close
	Conn.Close



	 Response.Redirect "sanctionarea.asp" 
               
	 End If  'msg = ""

Elseif  Action = "DOTEXT" Then

	 msg=""
	 If msg = "" Then

	 Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT * FROM Area51 " & _
	 "WHERE AID =" & Request("DAID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 


 	 ' Update record
	 rst("AreaDesc") = Request("DArea")
 	
	
	
                rst.Update

	rst.Close
	Conn.Close



	 Response.Redirect "sanctionarea.asp" 
               
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
        <td><font color="#FFBD00" size="2" face="Arial">Add an <b> Area</b> to the online catalogue. <u>Changes are immediate</u>. Please create at least one <b>Sub-Area</b>  and product after creation. Selecting an <b>Area</b> from the drop down box and clicking the 'delete' button will permanently remove the <b>Area</b> . You can verify that an <b>Area</b> exists by its presence in the deletion drop-down box, all <b>Areas</b> are sorted alphabetically.  </font></td>
    </tr>
    <tr>
        <td valign=top><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2"><font face="Arial">



<FORM ACTION="sanctionarea.asp" METHOD=POST>



<table border="0" cellpadding="0" cellspacing="0" width="100%">
    
    <tr>
        <td align="center">
        
              
        
        </td>
    </tr>

            	
	    	
		</table>
		<table border="0" cellpadding="3" cellspacing="0"  ALIGN="CENTER">
               
		<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Area (Spell correctly!):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Area" value="<%=Request("Area")%>">
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Submit" ></td>
                </tr><tr><td colspan=2><hr></td></tr>
                <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Area Delete:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select Distinct AreaName, AID from Area51 ORDER BY AreaName ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="AID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("AID")%><%=Request("AID")%>">
			<font face ="arial" size="1"><%=RS("AreaName")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Delete" ></td>
                </tr>
               
<tr><td colspan=2><hr></td></tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Area Change (Spell correctly!):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="EArea" value="<%=Request("Area")%>">
		    </td><td ></td>
                </tr>
              <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Area Edit:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select Distinct AreaName, AID from Area51 ORDER BY AreaName ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="EAID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("AID")%><%=Request("EAID")%>">
			<font face ="arial" size="1"><%=RS("AreaName")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Update" ></td>
                </tr>
	<tr><td colspan=2><hr></td></tr>	
            <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Area Description:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select Distinct AreaName, AID from Area51 ORDER BY AreaName ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="DAID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("AID")%><%=Request("DAID")%>">
			<font face ="arial" size="1"><%=RS("AreaName")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="DoText" ></td>
                </tr>
<tr>
      
              <td ><font size="2" face="Arial" color="#FFBD00">Area Text (Spell Correctly!):</font>
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
