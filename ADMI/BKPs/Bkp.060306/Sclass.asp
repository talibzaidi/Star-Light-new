 <%@ LANGUAGE = VBScript %>
<!--#include file="ADOVBS.INC"-->
<% 
    If (Session("Access") < "1") Then 
	Response.Redirect "login.asp"
    End If
%>
<%
		
		 SQL = "Select Distinct Area from CLASSFD ORDER BY Area ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		 	
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
        <td><font face="Arial" ></font>&nbsp;</td>
        <td><font size="2" face="Arial" color="#FFBD00">SANCTION CLASSIFIED: User Level Admin. You will be able to delete any one classified ad submitted .  </font></td>
    </tr>
    <tr>
        <td><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2"><font face="Arial">

<FORM ACTION="sclass1.asp" METHOD=get>



<table border="0" cellpadding="0" cellspacing="0" width="100%">
 
		<table border="0" cellpadding="3" cellspacing="0" width="80%" ALIGN="CENTER">
                
		<tr>
                    <td align="center">

			<% If (Session("Access") = 1) Then %> 
			<select name="ID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("Area")%><%=Request("Area")%>">
			<font face ="arial" size="1"><%=RS("Area")%>&nbsp;&nbsp;&nbsp;</font>
			 </option>
			<% RS.MoveNext
			Loop
			RS.Close %>

			</select>
			<% End If%>
			
		    </td>
                   
                </tr>
               <tr>
        <td align="center">
        
           <br>  <INPUT type="submit" NAME="Action" VALUE="Submit" >
       
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
