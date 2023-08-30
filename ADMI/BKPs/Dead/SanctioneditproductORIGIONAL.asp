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

PID = Request("PID")
PNAME = Replace( ReQuest("PName") , "'", "''") 
ITEMID = Replace( ReQuest("ITEMID") , "'", "''") 
UPC = Replace( ReQuest("UPC") , "'", "''") 
SERIALNUMBER = Replace( ReQuest("SerialNumber") , "'", "''") 
DESCR = Replace( ReQuest("Descr") , "'", "'") 

TEXT1 = Replace( ReQuest("Text1") , "'", "'") 
TEXT2 = Replace( ReQuest("Text2") , "'", "'") 
PIC1 = Replace( ReQuest("Pic1") , "'", "''") 
PIC2 = Replace( ReQuest("Pic2") , "'", "''") 
MANUFA = Replace( ReQuest("Manufa") , "'", "''") 
MANURL = Replace( ReQuest("ManURL") , "'", "''") 
COST = Replace( ReQuest("Cost") , "'", "''") 
MSL = Replace( ReQuest("MSL") , "'", "''") 
GPM = Replace( ReQuest("GPM") , "'", "''") 
Duty = Replace( ReQuest("Duty") , "'", "''") 
WEIGHT = ReQuest("Weight")
                If Request("Special")= "True" Then
	SPECIAL = true
	else
	SPECIAL = false
	end if
	If Request("NewProduct")= "True" Then
	NEWPRODUCT = true
	else
	NEWPRODUCT = false
	end if
	If Request("OverSize")= "True" Then
	OVERSIZED = true
	else
	OVERSIZED = false
	end if


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
	strSQL = "SELECT * FROM PRODUCT " & _
	 "WHERE PID =" & Request.QueryString("PID")
	Set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, conn, adOpenStatic, _
 	adLockOptimistic, adCmdText 


 	 ' Update record
	 if ReQuest("SID") = "" Then
	 xSIDx = 11
	 else
	 xSIDx = ReQuest("SID")
	 end if
	 rst("SID") = xSIDx
 	 rst("PName") = PNAME
 	 rst("ITEMID") = ITEMID
 	 rst("UPC") = UPC 
 	 rst("SerialNumber") = SERIALNUMBER 
 	 rst("Descr") = DESCR 
 	 rst("Text1") = TEXT1 
  	 rst("Text2") =  TEXT2
	 rst("Pic1") = PIC1 
	 rst("Pic2") = PIC2 
	 rst("Manufa") = MANUFA 
	 rst("ManURL") = MANURL 
	 rst("Cost") = COST
	 rst("MSL") = MSL 
	 rst("GPM") = GPM
	 rst("Duty") = Duty
	 rst("Weight") = WEIGHT 
	 rst("Special") = SPECIAL 
	 rst("NewProduct") = NEWPRODUCT 
	 rst("OverSize") = OVERSIZED
	
	
                rst.Update

	rst.Close
	Conn.Close
		
		
'If Err.number <> 0 then
'     response.redirect "./error.asp"
'end if
		 Response.Redirect "sanction.asp" 
	End If  'msg = ""
Elseif  Action = "DELETE" Then
	msg=""
	If msg = "" Then
'74 74 74 74 74 74 74 74 74 74 74 74 
	'Response.Redirect "sanctioneditproduct.asp?PID=" & ReQuest("PID") 'DEBUG CODE
	 SQL = "DELETE * FROM Product WHERE PID" + "= (" + ReQuest("PID") +")" 
	 Set conn = Server.CreateObject("ADODB.Connection")
    	 Conn.Open Session("ConnectionString")
    	 Conn.Execute(SQL)
	 Response.Redirect "sanctioneditproduct.asp" 
               
	 End If  'msg = ""
Elseif  Action = "EDIT" Then
	msg=""
	If msg = "" Then

	 
	 Response.Redirect "sanctioneditproduct.asp?PID=" & ReQuest("PID")
               
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
        src="../../Simages/sanction.gif"
        width="330" height="82"></font></td>
        <td bgcolor="#FFBD00"><font face="Arial"></font>&nbsp;</td>
        <td align="right" bgcolor="#FFBD00"><font face="Arial"><a href="../../sanction.asp"><img
        src="../../Simages/homegif.GIF"
        width="84" height="82" border="0"></a></font></td>
    </tr>
    <tr>
        <td><font face="Arial"><img
        src="../../Simages/btcurve.gif"
        width="330" height="82"></font></td>
        <td><font face="Arial"></font>&nbsp;</td>
        <td><font color="#FFBD00" size="2" face="Arial">Edit products with the form below. Choose an existing product from the top to edit, or from the bottom to delete. New products are entered using the form. Every product needs a numeric Product ID, no two Product ID's can be the same.<br></td>
    </tr>
    <tr>
        <td valign=top><font face="Arial"><img
        src="../../Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2'><font face="Arial">








<table border="0" cellpadding="0" cellspacing="0" width="100%">
    
    <tr>
        <td align="center">
        
              
        
        </td>
    </tr>

            	
	    	
		</table>
		<table border="0" cellpadding="3" cellspacing="0"  ALIGN="CENTER">
                <tr><FORM ACTION="sanctioneditproduct.asp?PID=<%=PIDD%>" METHOD=POST>
                    <td  bgcolor="#bbbbbb"><font size="2" face="Arial" color="#000000"><b>PRODUCT EDIT:<b></font>
		    </td>
                    <td width="100"  bgcolor="#bbbbbb"><%
		 SQL = "Select PName, PID, ITEMID from PRODUCT ORDER BY ITEMID ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		 On error resume next
%>
			       
		                
			
			<select name="PIDD" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("PID")%><%=Request("PID")%>">
			<font face ="arial" size="1"><%=RS("ITEMID")%>&nbsp;&nbsp;<%=RS("PName")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td  bgcolor="#bbbbbb"><INPUT type="submit"  NAME="Action" VALUE="Edit" ></td>
                </tr>
            </form>

<%
		 SQL = "Select * from PRODUCT WHERE PID = " + PID + ""
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RVS = Conn.Execute(SQL)
		 baz =  Cint(RVS("SID"))		 	


		 SQQL = "Select * from SubArea WHERE SID = " & baz & ""
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RQVS = Conn.Execute(SQQL)
%>

<FORM ACTION="sanctioneditproduct.asp?PID=<%=PID%>" METHOD=POST>
		<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Product Name:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="PName" value="<%=RVS("PName")%><%=Request("PName")%>">
		    </td><td ><INPUT type="submit"  NAME="Action" VALUE="Submit" ></td>
                </tr>
  <tr>
                    <td><font size="2" face="Arial" color="#FFBD00"><b>Sub-Area Select:<b></font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><%
		
		 SQL = "Select SubName, SID from SubArea ORDER BY Subname ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="SID" size="1" >
			<option value="<%=RQVS("SID")%><%=Request("SID")%>"><%=RQVS("Subname")%></option>
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
                    <td ><font size="2" face="Arial" color="#FFBD00">PRODUCT ID:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="ITEMID" value="<%=RVS("ITEMID")%><%=Request("ITEMID")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">UPC CODE:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="UPC" value="<%=RVS("UPC")%><%=Request("UPC")%>">
		    </td><td ></td>
                </tr>

<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">SERIAL #:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="SerialNumber" value="<%=RVS("SerialNumber")%><%=Request("SerialNumber")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Description (250):</font>
		    </td>


                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Descr" value="<%=RVS("Descr")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">
Text: (65535 characters max.)</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><textarea cols="45" name="Text1" rows="5" wrap="virtual"  value="<%=Request("Text1")%>"><%=RVS("Text1")%></textarea>
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Bullets (65535 characters max. use the tilde <b>~</b> to add a bullet ):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><textarea cols="45" rows="5" name="Text2" wrap="virtual" value="<%=Request("Text2")%>"><%=RVS("Text2")%></textarea>
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Thumbnail (250):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Pic1" value="<%=RVS("Pic1")%><%=Request("Pic1")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Image (250):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Pic2" value="<%=RVS("Pic2")%><%=Request("Pic2")%>">
		    </td><td ></td>
                </tr>
  <tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Manufacturers Name(250):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Manufa" value="<%=RVS("Manufa")%><%=Request("Manufa")%>">
		    </td><td ></td>
                </tr>
  <tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Manufacturers Link (250):</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="ManURL" value="<%=RVS("ManURL")%><%=Request("ManURL")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Cost :</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Cost" value="<%=RVS("Cost")%><%=Request("Cost")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">MSL:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="MSL" value="<%=RVS("MSL")%><%=Request("MSL")%>">
		    </td><td ></td>
                </tr>
	<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">GPM:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="GPM" value="<%=RVS("GPM")%><%=Request("GPM")%>">
		    </td><td ></td>
                </tr>
	<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Duty:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Duty" value="<%=RVS("Duty")%><%=Request("Duty")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Gross Weight:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Weight" value="<%=RVS("Weight")%><%=Request("Weight")%>">
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Special:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="checkbox" size="30" name="Special" value="<%=True%>" <% if RVS("Special") then Response.Write(" checked") %>>
		    </td><td ></td>
                </tr>
<tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">New Product:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="checkbox" size="30" name="NewProduct" value="<%=True%>" <% if RVS("NewProduct") then Response.Write(" checked") %>>
		    </td><td ></td>
                </tr><tr>
                    <td ><font size="2" face="Arial" color="#FFBD00">Over Sized:</font>
		    </td>
                    <td width="100"><font size="2" face="Arial"></font><input type="checkbox" size="30" name="OverSize" value="<%=True%>" <% if RVS("OverSize") then Response.Write(" checked") %>>
		    </td><td ></td>
                </tr>
<tr>
<%  RVS.Close
        %>
                    <td align=center bgcolor=bbbbbb colspan=3> <INPUT type="submit"  NAME="Action" VALUE="Submit This Product" >
		    </td>
                </tr>
<tr><td colspan="3"> <br><br>
</td></tr>
         </form><FORM ACTION="../../sanctionproduct.asp" METHOD=POST>       <tr>
                    <td bgcolor="#790000"><font size="2" face="Arial" color="#FFFFFF"><b>PRODUCT DELETE:<b></font>
		    </td>
                    <td width="100" bgcolor="#790000"><font size="2" face="Arial"></font><%
		 SQL = "Select PName, PID, ITEMID from PRODUCT ORDER BY ITEMID ASC"
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Set RS = Conn.Execute(SQL)
		
%>
			       
		                
			
			<select name="PID" size="1" >
			<%Do While Not RS.EOF%>
			<option value="<%=RS("PID")%><%=Request("PID")%>">
			<font face ="arial" size="1"><%=RS("ITEMID")%>&nbsp;&nbsp;<%=RS("PName")%></font>
			</option>
			<% RS.MoveNext
			Loop
			RS.Close %>
			</select>
		    </td><td  bgcolor="#790000"><INPUT type="submit"  NAME="Action" VALUE="Delete" ></td>
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
