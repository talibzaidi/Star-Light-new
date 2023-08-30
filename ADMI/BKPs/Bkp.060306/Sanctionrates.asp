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
	strSQL = "SELECT * FROM Rates " & _
	 "WHERE ID = 1" 
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 


 	 ' Update record
 	 rst("Tax1Rate") = Request("Tax1Rate")
 	 rst("Tax2Rate") = Request("Tax2Rate")
 	 rst("ExchangeRate1") = Request("ExchangeRate1")
 	 rst("Duty") = 1
 	 rst("GPM") = 1
 	 rst("Freight") = Request("Freight")
  	 rst("Insurance") = Request("Insurance")
		
                rst.Update

	rst.Close
	Conn.Close
		
		
'If Err.number <> 0 then
'     response.redirect "./error.asp"
'end if
		 Response.Redirect "sanctionrates.asp"
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
        <td><font color="#FFBD00" size="2" face="Arial">Rates are set in decimal form (i.e. 0.5, 2.2, etc. )</font><br></td>
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
               

<%
		 SQL = "Select * from Rates "
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RVS = Conn.Execute(SQL)
		 	
%>

<FORM ACTION="sanctionrates.asp" METHOD=POST>
		<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Tax Rate 1:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Tax1Rate" value="<%=RVS("Tax1Rate")%><%=Request("Tax1Rate")%>">
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Submit" ></td>
                </tr>

<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Tax Rate 2:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Tax2Rate" value="<%=RVS("Tax2Rate")%><%=Request("Tax2Rate")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Exchange Rate ( Can ):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="ExchangeRate1" value="<%=RVS("ExchangeRate1")%><%=Request("ExchangeRate1")%>">
		    </td><td ></td>
                </tr>



<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Freight:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Freight" value="<%=RVS("Freight")%><%=Request("Freight")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Insurance:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Insurance" value="<%=RVS("Insurance")%><%=Request("Insurance")%>">
		    </td><td ></td>
                </tr>
 
<tr>
<%  RVS.Close
        %>
                    <td align=center bgcolor=bbbbbb colspan=3> <INPUT type="submit"  NAME="Action" VALUE="Submit" >
		    </td>
                </tr>
<tr><td colspan="3"> <br><br>
</td></tr>
         </form>
               
              
		


</table>
    </td><td ></td>
    </tr>
</table>




</font>&nbsp;</td>
        
    </tr>
</table>
</body>
</html>
