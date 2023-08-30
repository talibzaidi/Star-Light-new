<%@ LANGUAGE = VBScript %>

<!--#include file="ADOVBS.INC"-->


<% 
    If (Session("Access") <> "1") Then 
	Response.Redirect "login.asp"
    End If
%>

<%
ToDo = Request("ToDo")   ' This is set in preceeding file Sanction.asp.
'Response.Write "<br>ToDo = " & ToDo
%>


<%
If False Then
	ChoosenAID = Request("ChoosenAID")   
	'Response.Write "<br>ChoosenAID = " & ChoosenAID
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
		'Response.Write "<br>DisplayAreaName = " & DisplayAreaName	
	End If
End If ' False

	
Button = Request("Button")  ' Denotes button that was just pressed (if any) on this form.
'Response.Write "<br>Button = " & Button 
mSubmitted = date & " " & time

AIDChoosen = ""

If Button = "Create Area" Then  ' BN: User has just submitted a new Area.
	SQL = "INSERT INTO Area51 ( AreaName, AreaDesc ) VALUES('" & Request("AreaName") & "', '" & Request("AreaDescription") & "' )"
	Response.Write "<br>SQL = " & SQL
	Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open Session("ConnectionString")
	Conn.Execute(SQL)
	Response.Redirect "Sanction.asp"
	
Elseif Button = "Choose Area" Then  ' BN: User has just chosen an existing area Area to Delete or Edit.
	AIDChoosen = Request("ChoosenAID")
	'Response.Write "<br>AIDChoosen = " & AIDChoosen
	SQL = "Select AreaName, AreaDesc from Area51 WHERE AID = " & AIDChoosen
	Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open Session("ConnectionString")
    Set RS = Conn.Execute(SQL)
    ChosenAreaName = RS("AreaName")
    'Response.Write "<br>ChosenAreaName = " & ChosenAreaName
    ChosenAreaDescription = RS("AreaDesc")
    'Response.Write "<br>ChosenAreaDescription = " & ChosenAreaDescription
	
Elseif  Button = "Delete Area" Then  ' BN: Delete an existing Area (given by AIDChoosen).
	SQL = "DELETE * FROM Area51 WHERE AID" + "= (" + ReQuest("AIDChoosen") +")" 
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	Conn.Execute(SQL) 
	Conn.Close
	Response.Redirect "Sanction.asp"
	  
Elseif  Button = "Edit Area" Then  ' Update to an existing Area (given by AIDChoosen).
	strSQL = "SELECT * FROM Area51 WHERE AID = " & ReQuest("AIDChoosen")
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, adLockOptimistic, adCmdText 
 	' Update record
	rst("AreaName") = Request("AreaName")
	rst("AreaDesc") = Request("AreaDescription")
	rst.Update
	rst.Close
	Conn.Close    
	Response.Redirect "Sanction.asp"
	
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
        <td colspan=2>
        <font color="#FFBD00" size="2" face="Arial">
        Add an <b>Area</b> to the online catalogue.
        Changes are immediate. Please create at least one <b>Sub-Area</b> and <b>Product</b> after creation. 
        Selecting an <b>Area</b> from the drop down box and clicking the 'Delete' button will permanently remove the <b>Area</b>. 
        You can verify that an <b>Area</b> exists by its presence in the deletion drop-down box. 
        All <b>Areas</b> are sorted alphabetically.
        </font></td>
    </tr>
    
    <tr>
        <td colspan="2">


		<FORM ACTION="sanctionarea.asp" METHOD=GET>

		<table border="0" cellpadding="3" cellspacing="0"  ALIGN="CENTER" width=1000>
		
		<% If ((ToDo = "DeleteArea") OR (ToDo = "EditArea")) AND (AIDChoosen = "") Then ' i.e. Want to Delete or Edit an Area but have not yet chosen which one. %>
		<tr>
			<td width=150>
			<b><font size="2" face="Arial" color="#FFBD00">Area Name:</font></b>
		    </td>
			<td >
			<%
			SQL = "Select Distinct AreaName, AID from Area51 ORDER BY AreaName ASC"
			Set conn = Server.CreateObject("ADODB.Connection")
    		Conn.Open Session("ConnectionString")
    		Set RS = Conn.Execute(SQL)
			%>
			<select name="ChoosenAID" size="1" >
			<% Do While Not RS.EOF %>
			<option value="<%=RS("AID")%>">
			<font face ="arial" size="1"><%=RS("AID")%>&nbsp;&nbsp;&nbsp;<%=RS("AreaName")%></font>
			<%
			RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td>
		    <td width=150>
		    <INPUT type="submit"  NAME="Button" VALUE="Choose Area">
		    <input type=hidden name='ToDo' value='<%=ToDo%>'>
		    </td>
		</tr>
		<% End If ' ((ToDo = "DeleteArea") OR (ToDo = "EditArea")) AND (AIDChoosen = "") %>
		
	
        <% If (ToDo = "DeleteArea") AND (AIDChoosen <> "") Then %>
		<tr>
			<td width=150>
			<b><font size="2" face="Arial" color="#FFBD00">Area ID (AID):</font></b>
		    </td>
			<td >
			<%=AIDChoosen%>
		    </td>
		    <td width=150>
		    <INPUT type="submit"  NAME="Button" VALUE="Delete Area">
		    <input type=hidden name='ToDo' value='<%=ToDo%>'>
		    <input type=hidden name='AIDChoosen' value='<%=AIDChoosen%>'>
		    </td>
		</tr>
		
		<tr>
			<td>
			<b><font size="2" face="Arial" color="#FFBD00">Area Name:</font></b>
		    </td>
			<td >
			<%=ChosenAreaName%>
		    </td>
		    <td>
			&nbsp;
		    </td>
		</tr>
		
		<tr>
			<td valign=top>
			<b><font size="2" face="Arial" color="#FFBD00">Area Description:</font></b>
		    </td>
			<td >
			<%=ChosenAreaDescription%>
		    </td>
		    <td>
		    &nbsp;
		    </td>
		</tr>
		<% End If ' (ToDo = "DeleteArea") AND (AIDChoosen <> "") %>
		
		
        <% If (ToDo = "EditArea") AND (AIDChoosen <> "") Then %>
		<tr>
			<td width=150>
			<b><font size="2" face="Arial" color="#FFBD00">Area ID (AID):</font></b>
		    </td>
			<td >
			<%=AIDChoosen%>
		    </td>
		    <td width=150>
		    <INPUT type="submit"  NAME="Button" VALUE="Edit Area">
		    <input type=hidden name='ToDo' value='<%=ToDo%>'>
		    <input type=hidden name='AIDChoosen' value='<%=AIDChoosen%>'>
		    </td>
		</tr>
		
		<tr>
			<td>
			<b><font size="2" face="Arial" color="#FFBD00">Area Name:</font></b>
		    </td>
			<td >
		    <textarea cols="40" rows="1" name="AreaName"><%=ChosenAreaName%></textarea>
		    </td>
		    <td>
			&nbsp;
		    </td>
		</tr>
		
		<tr>
			<td valign=top>
			<b><font size="2" face="Arial" color="#FFBD00">Area Description:</font></b>
		    </td>
			<td >
		    <textarea cols="90" rows="10" name="AreaDescription"><%=ChosenAreaDescription%></textarea>
		    </td>
		    <td>
		    &nbsp;
		    </td>
		</tr>
		<% End If ' (ToDo = "EditArea") AND (AIDChoosen <> "") %>
               
               
        <% If (ToDo = "CreateArea") Then %>
		<tr>
			<td>
			<b><font size="2" face="Arial" color="#FFBD00">Area Name:</font></b>
			<br><b><font size="1" face="Arial" color="#FFBD00">Spell correctly!</font></b>
		    </td>
			<td >
			<input type="text" size="50" name="AreaName">
		    </td>
		    <td>
		    <% If ToDo = "CreateArea" Then %>
		    <INPUT type="submit"  NAME="Button" VALUE="Create Area">
		    <% Else %>
		    &nbsp;
		    <% End If %>
		    </td>
		</tr>
		
		<tr>
			<td valign=top>
			<b><font size="2" face="Arial" color="#FFBD00">Area Description:</font></b>
		    </td>
			<td >
			<textarea cols="90" rows="10" name="AreaDescription"></textarea>
		    </td>
		    <td>
		    </td>
		</tr>
		<% End If ' (ToDo = "CreateArea") %>
		
	
		</table>

		</FORM>

		</td>
 
    </tr>
</table>

</body>

</html>
