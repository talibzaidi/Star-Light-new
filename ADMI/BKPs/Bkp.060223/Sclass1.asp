<%@ LANGUAGE = VBScript %>
<!--#include file="ADOVBS.INC"-->
<% 
    If (Session("Access") < "1") Then 
	Response.Redirect "login.asp"
    End If
   ID = ReQuest("ID")

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
        <td><font size="2" face="Arial" color="#FFBD00">Choose an ad to delete. Submit form. Repeat. </font></td>
    </tr>
    <tr>
        <td valign="top"><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2"><font face="Arial">

<FORM ACTION="sclass2.asp" METHOD=get>



<table border="0" cellpadding="4" cellspacing="0" width="100%">
    
    <tr>
        <td align="center">

		<INPUT type="submit"  NAME="Action" VALUE="Submit" >
        
        </td>
    </tr>

            	
	    
		
		<%
		
		 SQL = "Select * from CLASSFD WHERE Area Like '"& ID &"' "
                                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			<tr><td colspan=4><font color="#FFFFFF" size="4" face="Arial"><b> You must delete one at a time <b> </font><br><hr></td></tr>
		<%Do While Not RS.EOF%><tr>
                    <td ><font size="2" face="Arial"  color="#FFFFFF"><strong>DELETE :</strong></font><input type="radio" name="Classified" value="<%=RS("Index")%>"> </td>
                    
		    <td ><font size="2" face="Arial"  color="#FFFFFF">&nbsp;<strong>INDEX :</strong></font><font face ="arial" size="2" color="#FFFFFF"><%=RS("Index")%></font> </td>
		      <td ><font size="2" face="Arial"  color="#FFFFFF"><strong>AREA :</strong></font><font face ="arial" size="2" color="#FFFFFF"><%=RS("Area")%></font> </td>
			</tr>
	<tr>
		    <td ><font size="2" face="Arial"  color="#FFFFFF"><strong>MESSAGE :</strong></font><font face ="arial" size="2" color="#FFFFFF"><%=RS("Message")%></font> </td>
		   
		     <td ><font size="2" face="Arial"  color="#FFFFFF"><strong>AUTHOR :</strong></font><font face ="arial" size="2" color="#FFFFFF"><%=RS("Author")%></font> </td>
			  <td ><font size="2" face="Arial"  color="#FFFFFF"><strong>DATE :</strong></font><font face ="arial" size="2" color="#FFFFFF"><%=RS("Datet")%></font> </td>			
		  <td ></td>
                </tr><tr><td colspan=4><br><hr></td></tr>
               <% RS.MoveNext
			Loop
			RS.Close %>
            </table>

    </td>
    </tr>
</table>

</FORM>

</font>&nbsp;</td>
        
    </tr>
</table>
</body>
</html>
