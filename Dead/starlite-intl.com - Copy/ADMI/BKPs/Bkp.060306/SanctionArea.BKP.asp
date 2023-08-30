<%@ LANGUAGE = VBScript %>

<!--#include file="ADOVBS.INC"-->


<% 
    If (Session("Access") <> "1") Then 
	Response.Redirect "login.asp"
    End If
%>


<%

ChoosenAID = Request("ChoosenAID")   
Response.Write "<br>ChoosenAID = " & ChoosenAID
If ChoosenAID = "" Then 
	'ChoosenAID = 66			' Default Area AID.
	DisplayAreaName = "None"
Else
	SQL = "Select * from Area51 WHERE AID = " & ChoosenAID 
	Response.Write "<br>SQL = " & SQL
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	Set RSDisplay = Conn.Execute(SQL)
	If NOT RSDisplay.EOF Then  ' In case have just deleted this Area.
	DisplayAreaName = RSDisplay("AreaName")
	End If
	Response.Write "<br>DisplayAreaName = " & DisplayAreaName	
End If

	
	
msg=""

Action = Request("Action")
'Response.Write "<br>Action = " & Action 
mSubmitted = date & " " & time


If Action = "Add Area" Then  ' BN: Submit new Area.
	msg=""
	If msg = "" Then
	Area = Replace( ReQuest("Area") , " ", "%20") 
	SQL = "INSERT INTO Area51 ( AreaName ) "
    SQL = SQL & " VALUES('"& Request("Area") &"' )"
	Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open Session("ConnectionString")
	Conn.Execute(SQL)
	End If  'msg = ""
	
	
Elseif  Action = "Delete Area" Then  ' BN: Delete and existing Area (given by ChoosenAID).
	msg=""
	If msg = "" Then
	SQL = "DELETE * FROM Area51 WHERE AID" + "= (" + ReQuest("ChoosenAID") +")" 
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	Conn.Execute(SQL) 
	End If  'msg = ""
	 
	 
Elseif  Action = "Edit Area Name" Then  ' Update to Request("EArea") the AreaName field of an existing Area given by AID = Request("EAID").
	msg=""
	If msg = "" Then
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT * FROM Area51 WHERE AID = " & ChoosenAID   ' Request("EAID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, adLockOptimistic, adCmdText 
 	' Update record
	rst("AreaName") = Request("EArea")
	rst.Update
	rst.Close
	Conn.Close
	End If  'msg = ""


Elseif  Action = "Edit Area Description" Then   ' Update to Request("DArea") the AreaDesc field of an existing Area given by AID = Request("DAID").
	msg=""
	If msg = "" Then
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT * FROM Area51 WHERE AID =" & ChoosenAID  ' Request("DAID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, adLockOptimistic, adCmdText 
 	' Update record
	rst("AreaDesc") = Request("DArea")
	rst.Update
	rst.Close
	Conn.Close     
	End If  'msg = ""
	
	
Else
	'Response.Write "<br>GOT HERE"

End If  'Action = "Submit"
%>


<html>



<head>
<title>Sanction - Version (Orange)</title>
</head>



<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">



<table border="0" cellpadding="8" cellspacing="0" width="100%">
    <tr>
        <td bgcolor="#FFBD00"><font face="Arial"><img src="Simages/sanction.gif" width="330" height="82"></font></td>
        <td align="right" bgcolor="#FFBD00"><font face="Arial"><a href="sanction.asp"><img src="Simages/homegif.GIF" width="84" height="82" border="0"></a></font></td>
    </tr>
    
    <tr>
        <td colspan=2><font color="#FFBD00" size="2" face="Arial">Add an <b> Area</b> to the online catalogue. <u>Changes are immediate</u>. Please create at least one <b>Sub-Area</b>  and product after creation. Selecting an <b>Area</b> from the drop down box and clicking the 'delete' button will permanently remove the <b>Area</b> . You can verify that an <b>Area</b> exists by its presence in the deletion drop-down box, all <b>Areas</b> are sorted alphabetically.  </font></td>
    </tr>
    
    <tr>
        <td colspan="2">


		<FORM ACTION="sanctionarea.asp" METHOD=GET>

		<table border="0" cellpadding="3" cellspacing="0"  ALIGN="CENTER" width=95%>
               
		<tr>
			<td>
			<b><font size="2" face="Arial" color="#FFBD00">Add Name for a New Area:</font></b>
			<br><b><font size="1" face="Arial" color="#FFBD00">Spell correctly!</font></b>
		    </td>
			<td width="200"><input type="text" size="50" name="Area">
		    </td>
		    <td>
		    <INPUT type="submit"  NAME="Action" VALUE="Add Area"></td>
		</tr>
		
		<tr>
			<td colspan=3><hr></td>
		</tr>
		
		
		<tr>
			<td><font size="2" face="Arial" color="#FFBD00"><b>Choose An Existing Area to Edit:<b></font>
		    </td>
			<td>
			<%
			SQL = "Select Distinct AreaName, AID from Area51 ORDER BY AreaName ASC"
			Set conn = Server.CreateObject("ADODB.Connection")
    		Conn.Open Session("ConnectionString")
    		Set RS = Conn.Execute(SQL)
  			'Response.Write "<br>DisplayAID = " & DisplayAID 
			%>
			<select name="ChoosenAID" size="1" >
			<!-- <option value=""> -->
			<% Do While Not RS.EOF
			'If RS("AID") <> CInt(ChoosenAID) Then %>
			<option value="<%=RS("AID")%>">
			<% 'Else %>
			<!-- <option selected value="<%=RS("AID")%>">	-->	
			<% 'End If %>
			<font face ="arial" size="1"><%=RS("AID")%>&nbsp;&nbsp;&nbsp;<%=RS("AreaName")%></font>
			<%
			RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td>
		    <td>
		    <INPUT type="submit" NAME="Action" VALUE="Choose Area">
		    </td>
		</tr>
		
		<% If (ChoosenAID <> "") AND (NOT RSDisplay.EOF) Then %>
		<tr>
			<td colspan=3><hr></td>
		</tr>
		
		<tr>
			<td>
			<font size="2" face="Arial" color="#FFBD00"><b>Chosen Area:<b>
			</font>
		    </td>
			<td>
			<font size="2" face="Arial" color="#FFBD00">
			<%=DisplayAreaName%>
			</font>
		    </td>
		    <td>
		    <INPUT type="submit" NAME="Action" VALUE="Delete Area" >
		    </td>
		</tr>
		
		<tr>
			<td>
			<b><font size="2" face="Arial" color="#FFBD00">New Area Name:</font></b>
			<br><b><font size="1" face="Arial" color="#FFBD00">Spell correctly!</font></b>
		    </td>
			<td>
			<input type="text" size="50" name="EArea">
		    </td>
		    <td>
		    <INPUT type="submit"  NAME="Action" VALUE="Edit Area Name">
		    </td>
		</tr>
	
		<tr>
			<td valign=top>
			<b><font size="2" face="Arial" color="#FFBD00">New Area Description:</font></b>
		    </td>                    
		    <td>
			<%
			SQL = "Select AreaName, AreaDesc from Area51 WHERE AID = " & ChoosenAID & " ORDER BY AreaName ASC"
			'SQL = "Select AreaName, AreaDesc from Area51 ORDER BY AreaName ASC"
			Set conn = Server.CreateObject("ADODB.Connection")
    		Conn.Open Session("ConnectionString")
    		Set RS = Conn.Execute(SQL)
    		'AreaName = RS("AreaName")
    		AreaDesc = RS("AreaDesc")
			%>
		    <textarea cols="90" rows="10" name="DArea"><%=AreaDesc%></textarea>
		    </td>
		    <td valign='top'>
		    <INPUT type="submit" NAME="Action" VALUE="Edit Area Description" >
		    </td>
		    <% RS.Close %>
		</tr>
		<% End If ' ChoosenAID <> "" %>
	
		</table>

		</FORM>

		</td>
 
    </tr>
</table>

</body>

</html>
