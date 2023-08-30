<%@ LANGUAGE = VBScript %>
<!--#include file="ADOVBS.INC"-->
<% 
    If (Session("Access") <> "1") Then 
	Response.Redirect "login.asp"
    End If
%>
<%
If Err.number <> 0 then
     response.redirect "error.asp"
end if

ID = Request("ID")
NAME = Replace( ReQuest("Name") , "'", "''") 
BANNERURL = Replace( ReQuest("BannerURL") , "'", "''") 
TITLE = Replace( ReQuest("Title") , "'", "''") 
IMAGE1 = Replace( ReQuest("Image1") , "'", "''") 
COLINK = Replace( ReQuest("CoLink") , "'", "''")
msg=""

Action = Left(UCase(Request("Action")),6)
mSubmitted = date & " " & time


If Action = "SUBMIT" Then
	msg=""
	If msg = "" Then
              
		On Error Resume Next
	Dim conn
	Dim rst
	Dim strSQL
	
	
	
	
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT * FROM Links " & _
	 "WHERE ID =" & Request.QueryString("ID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 


 	 ' Update record
 	 rst("CoLink") = COLINK
     rst.Update

	rst.Close
	Conn.Close
		
		
'If Err.number <> 0 then
'     response.redirect "./error.asp"
'end if
		 Response.Redirect "sanctioneditlinks.asp?ID=" + ID
	End If  'msg = ""
Elseif  Action = "DELETE" Then
	msg=""
	If msg = "" Then
'74 74 74 74 74 74 74 74 74 74 74 74 
	
	 SQL = "DELETE * FROM Links WHERE ID" + "= (" + ReQuest("ID") +")" 
	 Set conn = Server.CreateObject("ADODB.Connection")
    	 Conn.Open Session("ConnectionString")
    	 Conn.Execute(SQL)
	 Response.Redirect "sanctionlinks.asp" 
               
	 End If  'msg = ""
Elseif  Action = "EDIT" Then
	msg=""
	If msg = "" Then

	 
	 Response.Redirect "sanctioneditlinks.asp?ID=" & ReQuest("PID")
               
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
        <td><font color="#FFBD00" size="2" face="Arial">Edit links with the form below. Choose an existing link from the top to edit, or from the bottom to delete.<br></td>
    </tr>
    <tr>
        <td valign=top><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2'><font face="Arial">








<table border="0" cellpadding="0" cellspacing="0" width="100%">
    
    <tr>
        <td align="center">
        
              
        
        </td>
    </tr>

            	
	    	
		</table>
		<table border="0" cellpadding="3" cellspacing="0"  ALIGN="CENTER">
                <tr><FORM ACTION="sanctioneditlinks.asp?ID=<%=ID%>" METHOD=POST>
                    <td  bgcolor="#bbbbbb"><font size="2" face="Arial" color="#000000"><b>LINK EDIT:<b></font>
		    </td>
                    <td width="100"  bgcolor="#bbbbbb"><%
		 SQL = "Select * from Links ORDER BY ID ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="PID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("ID")%><%=Request("PID")%>">
			<font face ="arial" size="1"><%=RS("ID")%>&nbsp;&nbsp;<%=RS("Name")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td  bgcolor="#bbbbbb"><INPUT type="submit"  NAME="Action" VALUE="Edit" ></td>
                </tr>
            </form>

<%
		 SQL = "Select * from Links WHERE ID = " + ID + ""
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RVS = Conn.Execute(SQL)
		 	
%>

<FORM ACTION="sanctioneditlinks.asp?ID=<%=ID%>" METHOD=POST>
		<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Link Name:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Name" value="<%=RVS("Name")%><%=Request("Name")%>">
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Submit" ></td>
                </tr>
 
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Link HTML:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><TEXTAREA COLS="12" ROWS="7" NAME="CoLink" STYLE="width: 403px;"><%=RVS("CoLink")%><%=Request("CoLink")%></TEXTAREA>
		    </td><td ></td>
                </tr>
<tr>
<%  RVS.Close
        %>
                    <td align=center bgcolor=bbbbbb colspan=3> <INPUT type="submit"  NAME="Action" VALUE="Submit This Link" >
		    </td>
                </tr>
<tr><td colspan="3"> <br><br>
</td></tr>
         </form><FORM ACTION="sanctioneditlinks.asp" METHOD=POST>       <tr>
                    <td bgcolor="#790000"><font size="2" face="Arial" color="#FFFFFF"><b>LINK DELETE:<b></font>
		    </td>
                    <td width="100" bgcolor="#790000"><font size="2" face="Arial"></font><%
		 SQL = "Select* from Links ORDER BY ID ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="ID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("ID")%><%=Request("ID")%>">
			<font face ="arial" size="1"><%=RS("ID")%>&nbsp;&nbsp;<%=RS("Name")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td  bgcolor="#790000"><INPUT type="submit"  NAME="Action" VALUE="DELETE"></td>
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

