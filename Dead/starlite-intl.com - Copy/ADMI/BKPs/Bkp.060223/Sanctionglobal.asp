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
	 set RUS = CreateObject("ADODB.Recordset")
	 set RS = CreateObject("ADODB.Recordset")
  RS.Open  "Select ITEMID, Cost from Product WHERE SID=" & Request("SID") , "DSN=STAREC1" , 1, 4
            	
	
                 do while not rs.eof
	
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT Cost, ITEMID FROM PRODUCT " & _
	 "WHERE ITEMID =" & RS("ITEMID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 

                rst("Cost") = (rst("Cost")*Request("Prix"))

	rst.Update

	rst.Close
	Conn.Close

                RS.MoveNext
	loop
	              
       RS.Close
      
      
          
        

       Response.Redirect "Sanctionglobal.asp"

	
elseIf Action = "CHANGE" Then
	 set RUS = CreateObject("ADODB.Recordset")
	 set RS = CreateObject("ADODB.Recordset")
  RS.Open  "Select ITEMID, Cost from Product WHERE SID=" & Request("SSID") , "DSN=STAREC1" , 1, 4
            	
	
                 do while not rs.eof
	
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	strSQL = "SELECT Cost, ITEMID FROM PRODUCT " & _
	 "WHERE ITEMID =" & RS("ITEMID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 

                rst("Cost") = Request("Prixx")

	rst.Update

	rst.Close
	Conn.Close

                RS.MoveNext
	loop
	              
       RS.Close
      
      
          
        

       Response.Redirect "Sanctionglobal.asp"

	
End If  'Action = "Change"
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
        <td><font color="#FFBD00" size="2" face="Arial">Set your global price change in decimal form. [ 0.5 = 50% ]</u> Choose by Sub-Area. 1200 product/sub-area max.</font></td>
    </tr>
    <tr>
        <td><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2'><font face="Arial">

<FORM ACTION="sanctionglobal.asp" METHOD="POST">



<table border="0" cellpadding="0" cellspacing="0" width="100%">
    
    <tr>
        <td align="center">
        
              
        
        </td>
    </tr>

            	
	    	
		</table>
		<table border="0" cellpadding="3" cellspacing="0"  ALIGN="CENTER">
               
		<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Price Multiplier:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"><input type="text" size="3" name="prix" value="<%=Request("prix")%>"> <b> %<b></font>
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Submit" ></td>
                </tr>
 
                <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Sub-Area:<b></font>
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
		    </td><td ></td>
                </tr>
                 <tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">New COST:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"><input type="text" size="3" name="prixx" value="<%=Request("prixx")%>"> <b> <b></font>
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Change" ></td>
                </tr>
 
                <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Sub-Area:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select Subname, SID from SubArea ORDER BY Subname ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="SSID" size="1" >
 
			<%Do While Not RS.EOF%>
		              
			<option value="<%=RS("SID")%><%=Request("SSID")%>">
			<font face ="arial" size="1"><%=RS("Subname")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
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
