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

NAME = Request("Name")
ADDR = Replace( ReQuest("Addr") , "'", "''") 
POSTAL = Replace( ReQuest("Postal") , "'", "''") 
CITY = Replace( ReQuest("City") , "'", "''") 
STPRO = Replace( ReQuest("StPro") , "'", "''") 
EMAIL = Replace( ReQuest("Email") , "'", "''") 
EMAILNAME = Replace( ReQuest("Emailname") , "'", "''") 
PHON1 = Replace( ReQuest("Phon1") , "'", "''") 
PHON2 = Replace( ReQuest("Phon2") , "'", "''") 
FAX = Replace( ReQuest("Fax") , "'", "''") 
i800NUM = Replace( ReQuest("800num") , "'", "''") 
TEXT1 = Replace( ReQuest("Text1") , "'", "''") 
TEXT2 = Replace( ReQuest("Text2") , "'", "''") 
PASS = Replace( ReQuest("Pass") , "'", "''") 
USER = Replace( ReQuest("User") , "'", "''") 
              


msg=""

Action = Left(UCase(Request("Action")),6)
mSubmitted = date & " " & time


If Action = "SUBMIT" Then
    msg=""
    If msg = "" Then
              
        'On Error Resume Next
    Dim conn
    Dim rst
    Dim strSQL
    
    
    
    
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open Session("ConnectionString")
    strSQL = "SELECT * FROM Company"
    Set rst = Server.CreateObject("ADODB.Recordset")
    rst.Open strSQL, Conn, 3,3


      ' Update record
      rst("Name") = NAME
      rst("Addr") = ADDR
      rst("Postal") = POSTAL 
     rst("City") = CITY 
      rst("StPro") = STPRO 
      rst("Email") = EMAIL 
      rst("Phon1") = PHON1 
       rst("Phon2") =  PHON2
     rst("Text1") = TEXT1 
     rst("Text2") = TEXT2 
     rst("Fax") = FAX 
     rst("800num") = i800NUM
     rst("Emailname") = EMAILNAME

    ' rst("ADMI.Pass") = PASS 
    ' rst("ADMI.User") = USER 
        
                rst.Update

    rst.Close
    strSQL = "SELECT * FROM ADMI WHERE ID = 1"
    rst.Open strSQL, Conn, 3,3

      ' Update record
      rst("Pass") = PASS
     rst("User") = USER
        
                rst.Update

    rst.Close
    Conn.Close
        
If Err.number <> 0 then
     response.redirect "./error.asp"
end if
         Response.Redirect "sanctionadmin.asp"
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
        <td bgcolor="#FFBD00">
        <font face="Arial"><img src="Simages/sanction.gif" ></font>
        </td>
        <td align="right" bgcolor="#FFBD00">
        <font face="Arial"><a href="sanction.asp"><img src="Simages/homegif.GIF" border="0"></a></font>
        </td>
    </tr>
    
    <tr>
        <td colspan='2'><font face="Arial">

						<table border="0" cellpadding="3" cellspacing="0"  width='98%' ALIGN="CENTER">
						<%
						         SQL = "Select *,* from Company, ADMI"
						                 Set conn = Server.CreateObject("ADODB.Connection")
						             Conn.Open Session("ConnectionString")
						             Set RVS = Conn.Execute(SQL)
						             
						%>

						<FORM ACTION="sanctionadmin.asp" METHOD=POST>
						<tr>
									<td >
									<br>
									<font color="#FFBD00" size="2" face="Arial">This form changes your company profile.</font><br>
									</td>
						            <td width='50'>
						            <INPUT type="submit"  NAME="Action" VALUE="Submit" >
						            </td>
						</tr>
						<tr>
									<td >
									<font size="2" face="Arial" color="#FFBD00">Company Name:</font>
						            </td>
						            <td width="100"><font size="2" face="Arial"></font>
						            <input type="text" size="30" name="Name" value="<%=RVS("Name")%><%=Request("Name")%>">
						            </td>
						</tr>
						 
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Address:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Addr" value="<%=RVS("Addr")%><%=Request("Addr")%>">
						            </td>
						                </tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Postal Code:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Postal" value="<%=RVS("Postal")%><%=Request("Postal")%>">
						            </td>
						                </tr>

						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">City:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="City" value="<%=RVS("City")%><%=Request("City")%>">
						            </td>
						                </tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">State / Province:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="StPro" value="<%=RVS("StPro")%><%=Request("StPro")%>">
						            </td>
						                </tr>
						<tr>
						            <td >
						            <font size="2" face="Arial" color="#FFBD00">Text 1: (homepage)</font>
						            </td>
						            <td width="100">
						            <font size="2" face="Arial"></font>
						            <textarea cols="130" name="Text1" rows="15" wrap="virtual"  value="<%=Request("Text1")%>"><%=RVS("Text1")%></textarea>
						            </td>
						</tr>
						<tr>
						            <td >
						            <font size="2" face="Arial" color="#FFBD00">Text 2: (homepage)</font>
						            </td>
						            <td width="100">
						            <font size="2" face="Arial"></font>
						            <textarea cols="130" rows="15" name="Text2" wrap="virtual" value="<%=Request("Text2")%>"><%=RVS("Text2")%></textarea>
						            </td>
						</tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Email (250):</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Email" value="<%=RVS("Email")%><%=Request("Email")%>">
						            </td>
						                </tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Email Name (250):</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Emailname" value="<%=RVS("Emailname")%><%=Request("Emailname")%>">
						            </td>
						                </tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Phone 1:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Phon1" value="<%=RVS("Phon1")%><%=Request("Phon1")%>">
						            </td>
						                </tr>
						  <tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Phone 2:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Phon2" value="<%=RVS("Phon2")%><%=Request("Phon2")%>">
						            </td>
						                </tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Fax :</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Fax" value="<%=RVS("Fax")%><%=Request("Fax")%>">
						            </td>
						                </tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">800num:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="800num" value="<%=RVS("800num")%><%=Request("800num")%>">
						            </td>
						                </tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Password:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Pass" value="<%=RVS("Pass")%><%=Request("Pass")%>">
						            </td>
						                </tr>
						<tr>
						                    <td ><font size="2" face="Arial" color="#FFBD00">Username:</font>
						            </td>
						                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="User" value="<%=RVS("User")%><%=Request("User")%>">
						            </td>
						                </tr>
						<tr>
						<%  RVS.Close
						        %>
						                    <td align=center bgcolor=bbbbbb colspan=2> 
						                    <INPUT type="submit"  NAME="Action" VALUE="Submit This Information" >
						            </td>
						                </tr>
						<tr>
									<td colspan="2">
									<br>
									</td>
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