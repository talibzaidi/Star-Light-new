<%@ LANGUAGE = VBScript %>

<!--#include file="ADOVBS.INC"-->

<% 
' 2/22/06: This file is almost identical to SanctionProductInsert.asp.
' 12/16/11: SanctionProductInsert.asp is now obsolute because I unified it with this file. 

If (Session("Access") <> "1") Then 
	Response.Redirect "login.asp"
End If
%>


<%
If Err.number <> 0 then
     response.redirect "error.asp"
End If


Action	= Trim(Request.QueryString("Action"))

If (Action = "CreateNewProduct") OR (Action = "EditOldProduct") OR (Action = "DeleteOldProduct") Then
	PID = ""
	SID = ""
ElseIf (Action <> "SubmitNewProduct") Then
	PID_SID	= Request.QueryString("PID_SID")	' Returned by FORM1 in querystring.
	'Response.Write "<br>*** PID_SID = " & PID_SID
	If PID_SID <> "" Then							' Extract PID from before the underscore, and extract SID from after the underscore.
		p = Instr(PID_SID, "_")
		PID		= Mid(PID_SID, 1, p-1)				' All chars before the underscore.
		SID		= Mid(PID_SID, p+1)					' All chars after the underscore.
	Else
		'PID		= Request.QueryString("PID")
		'SID		= Request.QueryString("SID")
		PID		= Request("PID")
		SID		= Request("SID")
	End If
ElseIf Action = "SubmitNewProduct" Then
	SID	= Request("SID")						' Returned by FORM2, not in querystring.
End If
%>


<%
If TRUE Then
	Response.Write "<br>Action = "	& Action
	Response.Write "<br>PID_SID = "	& PID_SID
	Response.Write "<br>PID = "		& PID
	Response.Write "<br>SID = "		& SID
End If


' Next Action control, for when Submit buttons for Form1 or Form2 are clicked.
' Form1 stuff ...
If Action = "EditOldProduct" Then 
	NextAction			= "DisplayProductToEdit" 
	UsageMsg = "First choose an existing product to edit."
	strLabel			= "CHOOSE PRODUCT TO EDIT" 
ElseIf Action = "DeleteOldProduct" Then
	NextAction			= "DisplayProductToDelete" 
	UsageMsg = "First choose an existing product to delete. This will just display that product, not delete it."
	strLabel = "CHOOSE PRODUCT TO PHYSICALLY DELETE<br>(Logical deletion is an alternative option, by editing that field for a product)" 

' Form2 stuff ...
ElseIf Action = "DisplayProductToEdit" Then
	NextAction			= "SubmitProductChanges" 
	UsageMsg = "Edit this product using the form below. Then submit your changes using the button at top right or at bottom. Every product needs a numeric Item ID (containing no spaces). No two Item IDs should be the same."
ElseIf Action = "DisplayProductToDelete" Then 
	NextAction			= "DeleteThisProduct" 
	UsageMsg = "Delete this product using the button at top right or at bottom."
ElseIf Action = "CreateNewProduct" Then
	NextAction			= "SubmitNewProduct"
	UsageMsg = "Enter a new product using the form below. Then submit your data using the button at top right or at bottom. Every product needs a numeric Item ID (containing no spaces). No two Item IDs should be the same."
End If
Response.Write "<br>NextAction = "			& NextAction
	


' 12/15/11: I'm not sure, or I've forgotten, what exactly these Replacements are for. 
' However these Replacements ARE needed, or the corresponding boxes are empty when a product is selected for Editing in Sanction.
' Note some replace ' with itself. Is that an error? This does NOT occur in file SanctionProductInsert.asp !?
' 12/16/11: I made these Replacements all "1-single-quote" to "2-single-quote", like in file SanctionProductInsert.asp.
If TRUE Then
	PNAME           = Replace( ReQuest("PName") , "'", "''") 
	ITEMID          = Replace( ReQuest("ITEMID") , "'", "''") 
	UPC             = Replace( ReQuest("UPC") , "'", "''") 
	SERIALNUMBER    = Replace( ReQuest("SerialNumber") , "'", "''") 
	DESCR           = Replace( ReQuest("Descr") , "'", "''") 
	TEXT1           = Replace( ReQuest("Text1") , "'", "''") 
	'TEXT1          = Replace( TEXT1 , vbCrLf, "<br>") 
	TEXT2           = Replace( ReQuest("Text2") , "'", "''") 
	PIC1            = Replace( ReQuest("Pic1") , "'", "''") 
	PIC2            = Replace( ReQuest("Pic2") , "'", "''") 
	MANUFA          = Replace( ReQuest("Manufa") , "'", "''") 
	MANURL          = Replace( ReQuest("ManURL") , "'", "''") 
	COST            = Replace( ReQuest("Cost") , "'", "''") 
	VENDORS         = Replace( ReQuest("Vendors") , "'", "''") 
	MSL             = Replace( ReQuest("MSL") , "'", "''") 
	GPM             = Replace( ReQuest("GPM") , "'", "''") 
	Duty            = Replace( ReQuest("Duty") , "'", "''") 
	WEIGHT          = Replace( ReQuest("Weight") , "'", "''")
	REBATEDESCR		= Replace( Request("RebateDescr") , "'", "''") 
	HASACCESSORIES  = Replace( ReQuest("HasAccessories") , "'", "''")
	ISACCESSORYOF   = Replace( ReQuest("IsAccessoryOf") , "'", "''")
End If		' TRUE/FALSE

  
If Request("ShowPrice") = "True" Then
	SHOWPRICE = true
Else
	SHOWPRICE = false
End If
              
If Request("Special") = "True" Then
	SPECIAL = true
Else
	SPECIAL = false
End If

If Request("NewProduct") = "True" Then
	NEWPRODUCT = true
Else
	NEWPRODUCT = false
End If

If Request("Rebated") = "True" Then
	REBATED = true
Else
	REBATED = false
End If

If Request("OverSize") = "True" Then
	OVERSIZED = true
Else
	OVERSIZED = false
End If

If Request("Deleted") = "True" Then
	DELETED = true
Else
	DELETED = false
End If



mSubmitted = date & " " & time


' 11/12/16: The following case is for SUBMIT-ting an edited version of a pre-existing product / record (unlike the next case below). 
If Action = "SubmitProductChanges" Then	

	If TRUE Then
		Response.Write "<br>"
		Response.Write "<br>PID = "				& PID
		Response.Write "<br>PNAME = "			& PNAME
		Response.Write "<br>SID = "				& SID
		Response.Write "<br>ITEMID = "			& ITEMID
		Response.Write "<br>UPC = "				& UPC
		Response.Write "<br>SERIALNUMBER = "	& SERIALNUMBER
		Response.Write "<br>DESCR = "			& DESCR
		Response.Write "<br>TEXT1 = "			& TEXT1
		Response.Write "<br>TEXT2 = "			& TEXT2
		Response.Write "<br>PIC1 = "			& PIC1
		Response.Write "<br>PIC2 = "			& PIC2
		Response.Write "<br>MANUFA = "			& MANUFA
		Response.Write "<br>MANURL = "			& MANURL
		Response.Write "<br>COST = "			& COST
		Response.Write "<br>VENDORS = "			& VENDORS
		Response.Write "<br>MSL = "				& MSL
		Response.Write "<br>GPM = "				& GPM
		Response.Write "<br>Duty = "			& Duty
		Response.Write "<br>WEIGHT = "			& WEIGHT
		Response.Write "<br>REBATEDESCR = "		& REBATEDESCR
		Response.Write "<br>HASACCESSORIES = "	& HASACCESSORIES
		Response.Write "<br>ISACCESSORYOF = "	& ISACCESSORYOF
	End If


	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	'SQL = "SELECT * FROM PRODUCT WHERE PID =" & Request.QueryString("PID")
	SQL = "SELECT * FROM PRODUCT WHERE PID =" & PID
    Response.Write "<br>SQL = " & SQL
	Set rst = Server.CreateObject("ADODB.Recordset")
	'Response.End
	rst.Open SQL, conn, adOpenStatic, adLockOptimistic, adCmdText 

	' Update record
	'If Request("SID") = "" Then
	'	xSIDx = 11
	'Else
	'	xSIDx = Request("SID")
	'End If

	If SID = "" Then
		xSIDx = 11
	Else
		xSIDx = SID
	End If
	
	rst("SID")				= xSIDx
	rst("PName")			= PNAME
	rst("ITEMID")			= ITEMID
	rst("UPC")				= UPC 
	rst("SerialNumber")		= SERIALNUMBER 
	rst("Descr")			= DESCR 
	rst("Text1")			= TEXT1 
	rst("Text2")			= TEXT2
	rst("Pic1")				= PIC1 
	rst("Pic2")				= PIC2 
	rst("Manufa")			= MANUFA 
	rst("ManURL")			= MANURL 
	rst("Cost")				= COST
	rst("Vendors")			= VENDORS
	rst("MSL")				= MSL 
	rst("GPM")				= GPM
	rst("Duty")				= Duty
	rst("Weight")			= WEIGHT 
	rst("ShowPrice")		= SHOWPRICE 
	rst("Special")			= SPECIAL 
	rst("NewProduct")		= NEWPRODUCT 
	rst("Rebated")			= REBATED
	rst("RebateDescr")		= REBATEDESCR
	rst("OverSize")			= OVERSIZED
	rst("Deleted")			= DELETED
	rst("HasAccessories")	= HASACCESSORIES
	rst("IsAccessoryOf")	= ISACCESSORYOF
			
	rst.Update
	rst.Close
	Conn.Close

	' The following (hopefully) takes user back to the display of the product just edited, so he can check his changes.
	Response.Redirect "SanctionProduct.asp?Action=DisplayProductToEdit" & "&PID=" & PID & "&SID=" & SID & "&Msg=Your EDIT was successful"


' 11/12/16: The following case is for SUBMIT-ting a brand new product / record (unlike the previous case above). 
' The code was copied from the SUBMIT case of file SanctionProductInsert.111216.asp. I am trying to do away with the need for SanctionProductInsert.asp, which was almost identical to this SanctionProductEdit.asp file.
ElseIf Action = "SubmitNewProduct" Then	
    SQL = "INSERT INTO Product(SID, PName, ITEMID, UPC, SerialNumber, Descr, Text1, Text2, Pic1, Pic2, Manufa, ManURL, Cost, Vendors, MSL, GPM, Duty, Weight, ShowPrice, Special, NewProduct, OverSize, Deleted, HasAccessories, IsAccessoryOf) "
    SQL = SQL & " VALUES(" & Request("SID") & " ,'"  & PNAME &  "' ,'"  & ITEMID & "' ,'"  & UPC & "' ,'"  & SERIALNUMBER & "' ,'"  & DESCR & "' ,'"  & TEXT1 & "' ,'"  &  TEXT2 & "' ,'"  &  PIC1 & "' ,'"  &  PIC2 & "' ,'"  &  MANUFA & "' ,'"  &  MANURL & "' ,"  &  COST & " , '"  &  VENDORS & "' , " &  MSL & " , "  &  GPM & " , " &  Duty & " , " &  WEIGHT & " , " &  SHOWPRICE & " , " &  SPECIAL & " , "  &  NEWPRODUCT & " , "  &  OVERSIZED  &  " , "  &  DELETED &  " , '"  &   HASACCESSORIES &   "' , '"  &  ISACCESSORYOF & "' );"
    Response.Write "<br>SQL = " & SQL
	'Response.End 
	'On Error Resume Next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	Conn.Execute(SQL)
	If Err.number <> 0 then
    	Response.Redirect "error.asp"
	end if

	' The following (hopefully) takes user back to the display of the product just edited, so he can check his changes.
	' NO. Cannot use it because PID is not defined here. PID is created automatically, as an AutoNumber field, by above INSERT. How can I retreive that PID value?
	'Response.Redirect "SanctionProduct.asp?PID=" & PID & "&Action=DisplayProductToEdit"
	'Response.Redirect "SanctionProduct.asp?Action=CreateNewProduct"		' Anyway, Sani wanted this next, instead of using above line. 
	Response.Redirect "sanction.asp?Msg=Your CREATE was successful"      ' Above redirect is not a good idea. Hard to tell if anything happened.

Elseif  Action = "DeleteThisProduct" Then
	SQL = "DELETE * FROM Product WHERE PID = " & Request("PID") 
    Response.Write "<br>SQL = " & SQL
	'Response.End
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	Conn.Execute(SQL)								' This causes the actual deletion!

	Response.Redirect "sanction.asp?Msg=Your DELETE was successful"		' Remember: the product just deleted is no longer in db. So cannot here go back to display it.

End If  'Action = "SubmitProductChanges"
%>


<% Msg = Request.QueryString("Msg") %>


<html>


<head>
	<title>SanctionProduct.asp</title>
</head>



<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">


<% If Msg <> "" Then %>
	<center>
		<font size='4' face="Tahoma" color='white'><%=Msg%></font>
	</center>
<% End If %>


<table align="center" border="0" cellpadding="5" cellspacing="0" width="100%">
<tr>
    <td bgcolor="#FFBD00">
		<a href="sanction.asp"><img src="Simages/sanction.gif" align="middle" border='0' width="330" height="82"></a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<!-- <a href="sanction.asp"><img src="Simages/homegif.GIF" width="84" height="82" border="0"></a> -->
		<font face="Arial" size="4" color='black'>
			<b>Product:&nbsp;&nbsp;&nbsp; </b>
			<a href="SanctionProduct.asp?Action=CreateNewProduct">Create</a>,&nbsp;
			<a href="SanctionProduct.asp?Action=EditOldProduct">Edit</a>,&nbsp;
			<a href="SanctionProduct.asp?Action=DeleteOldProduct">Delete</a>,&nbsp;
			<a href="http://www.starlite-intl.com/Admin2/login.asp?pwd=787szd&btnSubmit=Submit">Extra Admin</a>
		</font>
	</td>
</tr>
    
<tr>
    <td>
    <font color="#FFBD00" size="2" face="Arial">
	<%=UsageMsg%>
	</font>
	</td>
</tr>
    
<tr>
    <td colspan="2"><font face="Arial">

	<table border="0" cellpadding="3" cellspacing="0"  ALIGN="CENTER" width='100%'>

	<% If Action = "EditOldProduct" OR Action = "DeleteOldProduct" Then %>
	<tr>
		<!-- <td  bgcolor="#bbbbbb"><font size="2" face="Arial" color="#000000"><b>PRODUCT EDIT:<b></font></td>  -->
		<FORM Name="FORM1" ACTION="SanctionProduct.asp?Action=<%=NextAction%>&PID_SID=<%=PID_SID%>" METHOD="GET" >

			<input type="hidden" name="PID" value="<%=PID%>" />
			
			<td  bgcolor="#bbbbbb"><font size="2" face="Arial" color="#000000"><b><%=strLabel%>:<b></font></td>
            
			<td width="100"  bgcolor="#bbbbbb">
				<%	SQL = "Select PName, PID, SID, ITEMID from PRODUCT ORDER BY ITEMID ASC"
					Set conn = Server.CreateObject("ADODB.Connection")
    				Conn.Open Session("ConnectionString")
    				Set RS = Conn.Execute(SQL)
					On error resume next
				%>
			       
				<select name="PID_SID" size="1" >
					<% 
					Do While Not RS.EOF
						PID = RS("PID") : SID = RS("SID") %>
						<option value="<%=PID%>_<%=SID%>" >
							<%=RS("ITEMID")%>&nbsp;&nbsp;<%=RS("PName")%>&nbsp;&nbsp;<%=PID%>
						</option>
						<% 
						RS.MoveNext
					Loop
					%>
				</select>
			</td>
		
			<td  bgcolor="#bbbbbb"><INPUT type="submit" NAME="Action" VALUE="<%=NextAction%>"></td>

		</FORM>
		<% RS.Close %>
	</tr>
	<% End If   '  Action = "EditOldProduct" OR Action = "DeleteOldProduct" %>


	<%
	If PID <> "" Then	' When editing a pre-existing product / record with a given PID value.
		SQL = "Select * from PRODUCT WHERE PID = " + PID + ""
		Set conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("ConnectionString")
		Set RVS = Conn.Execute(SQL)
		baz =  Cint(RVS("SID"))		 	

		SQQL = "Select * from SubArea WHERE SID = " & baz & ""
		Set conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("ConnectionString")
		Set RQVS = Conn.Execute(SQQL)
	Else				' When making a new product, so no PID exists yet.
		Val = "0"		' This is the generic value to place in all fields of form2 when PID = "", i.e. when making a NEW product / record.
	End If
	%>


	<% 
	If (Action = "CreateNewProduct") OR (Action = "DisplayProductToEdit") OR (Action = "DisplayProductToDelete") Then 
	%>


	<!-- 
	12/17/11: I would rather use METHOD=POST here (like for FORM1), instead of METHOD=GET, to avoid a huge quesrystring,
	but passing the SID value that is chosen below does not seem to work when using METHOD=POST. 
	Actually, I HAVE to use METHOD=POST or the URL gets gigantic and causes a "URL too long" error. I will need to solve the SID passing problem some way other that using METHOD=GET.
	-->
	<% If (Action = "DisplayProductToEdit") OR (Action = "DisplayProductToDelete") OR (Action = "SubmitProductChanges") Then  ' In this case, PID and SID are known at this point.  %>
		<!-- <FORM Name="FORM2" ACTION="SanctionProduct.asp?Action=<%=NextAction%>&PID=<%=PID%>&SID=<%=SID%>" METHOD=POST> -->
		<FORM Name="FORM2" ACTION="SanctionProduct.asp?Action=<%=NextAction%>" METHOD=POST>
	<% ElseIf Action = "CreateNewProduct" Then  ' In this case, PID and SID are not known at this point. %>
		<FORM Name="FORM2" ACTION="SanctionProduct.asp?Action=<%=NextAction%>" METHOD=POST> 
	<% End If %>

	<input type="hidden" name="PID" value="<%=PID%>" />

	<% If PID <> "" Then    ' Happens when Action = "Display Product to Edit" OR Action = "Display Product to Delete". %>
	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">PID (Record Number):</font></td>
		<!-- <td width="100"><input type="text" size="30" name="PID" value="<%=RVS("PID")%>"></td> -->
		<td width="100"><font size="2" face="Arial"><%=RVS("PID")%></font></td>
		<td><INPUT type="submit"  NAME="Action" VALUE="<%=NextAction%>"></td>
	</tr>
	<% End If	' PID <> "" %>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Product Name:</font></td>
		<% If PID <> "" Then Val = RVS("PName") End If %>
		<td width="100"><input type="text" size="30" name="PName" value="<%=Val%>"></td>
		<% If Action = "CreateNewProduct" Then %>
			<td><INPUT type="submit" NAME="Action" VALUE="<%=NextAction%>"></td>
		<% Else		' (Action = "DisplayProductToEdit") OR (Action = "DisplayProductToDelete") %>
			<td></td>
		<% End If %>
	</tr>
	
	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Sub-Area:</font></td>
		<td width="100">
			<%	SQLSubArea = "Select SubName, SID from SubArea ORDER BY Subname ASC"
				Set Conn = Server.CreateObject("ADODB.Connection")
    			Conn.Open Session("ConnectionString")
    			Set rsSubArea = Conn.Execute(SQLSubArea)
			%>

			<select name="SID" size="1" >
				<!-- <option value="<%=rsSubArea("SID")%>" ></option> -->
				<% 
				rsSubArea.MoveFirst
				Do While Not rsSubArea.EOF 
					'SID = rsSubArea("SID")
					If CStr(rsSubArea("SID")) = CStr(SID) Then %>
						<option value="<%=rsSubArea("SID")%>" selected="selected" > 
							<%=rsSubArea("SID")%>&nbsp;&nbsp;<%=rsSubArea("Subname")%> 
						</option>
				<%	Else %>
						<option value="<%=rsSubArea("SID")%>" > 
							<%=rsSubArea("SID")%>&nbsp;&nbsp;<%=rsSubArea("Subname")%> 
						</option>
				<%	End If 
				rsSubArea.MoveNext
				Loop
				   rsSubArea.Close %>
			</select>
		</td>
		<td></td>
	</tr>
		
	<tr>
	    <td align=right><font size="2" face="Arial" color="#FFBD00">Item ID <font color="white">(no spaces)</font>:</font></td>
		<% If PID <> "" Then Val = RVS("ITEMID") End If %>
	    <td width="100"><input type="text" size="30" name="ITEMID" value="<%=Val%>"></td>
        <td ></td>
	</tr>

	<tr>
	    <td align=right><font size="2" face="Arial" color="#FFBD00">UPC Code:</font></td>
		<% If PID <> "" Then Val = RVS("UPC") End If %>
	    <td width="100"><input type="text" size="30" name="UPC" value="<%=Val%>"></td>
        <td ></td>
	</tr>

	<tr>
	    <td align=right><font size="2" face="Arial" color="#FFBD00">Serial #:</font></td>
		<% If PID <> "" Then Val = RVS("SerialNumber") End If %>
	    <td width="100"><input type="text" size="30" name="SerialNumber" value="<%=Val%>"></td>
        <td ></td>
	</tr>

	<tr>
	    <td align=right><font size="2" face="Arial" color="#FFBD00">Description (250):</font></td>
		<% If PID <> "" Then Val = RVS("Descr") End If %>
		<td width="100"><textarea cols="110" name="Descr" rows="3"><%=Val%></textarea></td>
        <td ></td>
	</tr>

	<tr>
		<td valign='top' align=right>
		<font size="2" face="Arial" color="#FFBD00">Text1:</font> <font size="1" face="Arial" color="#FFBD00"><br>65535 chars</font>
		</td>
		<% If PID <> "" Then Val = RVS("Text1") End If %>
		<td width="100"><textarea cols="110" rows="10" name="Text1"><%=Val%></textarea></td>                                                  
		<td></td>
	</tr>

	<tr>
		<td valign='top' align=right>
			<font size="2" face="Arial" color="#FFBD00">Text2:</font>
			<font size="1" face="Arial" color="#FFBD00"><br>65535 chars</font>
			<font size="1" face="Arial" color="#FFBD00"><br>Use an optional tilde <b>~</b> to add a bullet</font>
		</td>
		<% If PID <> "" Then Val = RVS("Text2") End If %>
		<td width="100"><textarea cols="110" rows="10" name="Text2"><%=Val%></textarea></td>
		<td></td>
	</tr>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Thumbnail (250):</font></td>
		<% If PID <> "" Then Val = RVS("Pic1") End If %>
		<td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Pic1" value="<%=Val%>"></td>
		<td ></td>
	</tr>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Image (250):</font></td>
		<% If PID <> "" Then Val = RVS("Pic2") End If %>
		<td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Pic2" value="<%=Val%>"></td>
		<td ></td>
	</tr>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Manufacturers Name (250):</font></td>
		<% If PID <> "" Then Val = RVS("Manufa") End If %>
		<td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="Manufa" value="<%=Val%>"></td>
		<td ></td>
	</tr>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Manufacturers Link (250):</font></td>
		<% If PID <> "" Then Val = RVS("ManURL") End If %>
		<td width="100"><font size="2" face="Arial"></font><input type="text" size="30" name="ManURL" value="<%=Val%>"></td>
		<td ></td>
	</tr>

	<tr>
	    <td align=right><font size="2" face="Arial" color="#FFBD00">Cost:</font></td>

	    <td width="100">
			<table border=0 align='left'>
			<tr>
				<% If PID <> "" Then Val = RVS("Cost") End If %>
				<td>
				<input type="text" size="30" name="Cost" value="<%=Val%>">
				</td>
				<td align='right'>
				<font size="2" face="Arial" color="#FFBD00">&nbsp;&nbsp;&nbsp;Vendors:</font>
				</td>
				<% If PID <> "" Then Val = RVS("Vendors") End If %>
				<td>
				<input type="text" size="99" maxlength=250 name="Vendors" value="<%=Val%>">
				</td>
			</tr>
			</table>
		</td>

		<td></td>
	</tr>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">MSL:</font></td>
		<% If PID <> "" Then Val = RVS("MSL") End If %>
		<td width="100"><input type="text" size="30" name="MSL" value="<%=Val%>"></td>
		<td ></td>
	</tr>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">GPM:</font></td>
		<% If PID <> "" Then Val = RVS("GPM") End If %>
		<td width="100"><input type="text" size="30" name="GPM" value="<%=Val%>"></td>
		<td ></td>
	</tr>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Duty:</font></td>
		<% If PID <> "" Then Val = RVS("Duty") End If %>
		<td width="100"><input type="text" size="30" name="Duty" value="<%=Val%>"></td>
		<td ></td>
	</tr>
	
	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Gross Weight:</font></td>
		<% If PID <> "" Then Val = RVS("Weight") End If %>
		<td width="100"><input type="text" size="30" name="Weight" value="<%=Val%>"></td>
		<td ></td>
	</tr>

	<tr>  <!-- [BN, 3/2/04]. Added this row. -->
		<td align=right><font size="2" face="Arial" color="#FFBD00">Show Price:</font></td>
		<% If PID <> "" Then %>
			<td width="100"><input type="checkbox" size="30" name="ShowPrice" value="<%=True%>" <% If RVS("ShowPrice") Then Response.Write(" checked") %> ></td>
		<% Else %>
			<td width="100"><input type="checkbox" size="30" name="ShowPrice" value="<%=False%>" ></td>
		<% End If %>
		<td ></td>
	</tr>

	<tr>
		<td align=right><font size="2" face="Arial" color="#FFBD00">Special:</font></td>
		<% If PID <> "" Then %>
			<td width="100"><input type="checkbox" size="30" name="Special" value="<%=True%>" <% If RVS("Special") Then Response.Write(" checked") %> ></td>
		<% Else %>
			<td width="100"><input type="checkbox" size="30" name="Special" value="<%=False%>" ></td>
		<% End If %>
		<td ></td>
	</tr>

	<tr>
	    <td align=right><font size="2" face="Arial" color="#FFBD00">New Product:</font></td>
		<% If PID <> "" Then %>
			<td width="100"><input type="checkbox" size="30" name="NewProduct" value="<%=True%>" <% If RVS("NewProduct") Then Response.Write(" checked") %> ></td>
		<% Else %>
			<td width="100"><input type="checkbox" size="30" name="NewProduct" value="<%=False%>" ></td>
		<% End If %>
        <td ></td>
	</tr>

	<tr>  <!-- [BN, 12/15/11] Added this row.  -->
	    <td align=right><font size="2" face="Arial" color="#FFBD00">Has Rebate:</font></td>
		<% If PID <> "" Then %>
			<td width="100"><input type="checkbox" size="30" name="Rebated" value="<%=True%>" <% If RVS("Rebated") Then Response.Write(" checked") %> ></td>
		<% Else %>
			<td width="100"><input type="checkbox" size="30" name="Rebated" value="<%=False%>" ></td>
		<% End If %>
        <td ></td>
	</tr>

	<tr>  <!-- [BN, 12/15/11] Added this row.  -->
		<td align=right valign='top'>
			<font size="2" face="Arial" color="#FFBD00">Rebate Description:</font>
			<br /><font size="1" face="Arial" color="#FFBD00">(255 Chars):</font>
		</td>
		<% If PID <> "" Then Val = RVS("RebateDescr") End If %>
		<td width="100"><textarea cols="110" name="RebateDescr" rows="3"><%=Val%></textarea></td>
		<td></td>
	</tr>
	
	<tr>
        <td align=right><font size="2" face="Arial" color="#FFBD00">Over Sized:</font></td>
		<% If PID <> "" Then %>
			<td width="100"><input type="checkbox" size="30" name="OverSize" value="<%=True%>" <% If RVS("OverSize") Then Response.Write(" checked") %> ></td>
		<% Else %>
			<td width="100"><input type="checkbox" size="30" name="OverSize" value="<%=False%>" ></td>
		<% End If %>
        <td ></td>
	</tr>

	<tr>
        <td align=right><font size="2" face="Arial" color="white"><b>Logically Deleted:</b></font></td>
		<% If PID <> "" Then %>
			<td width="100"><input type="checkbox" size="30" name="Deleted" value="<%=True%>" <% If RVS("Deleted") Then Response.Write(" checked") %> >
		<% Else %>
			<td width="100"><input type="checkbox" size="30" name="Deleted" value="<%=False%>" >
		<% End If %>
           <font size="2">Careful: check this only for products that are obsolete but that you want to keep in the database.</font></td>
		<td></td>
	</tr>
                
	<tr>  <!-- [BN, 2/22/06] Added this row.  -->
		<td align=right valign="top">
			<font size="2" face="Arial" color="#FFBD00">Has Accessories:</font>
			<font size="1" face="Arial" color="#FFBD00"><br>Use double tilde <b>~~</b> to add a heading.</font>
			<font size="1" face="Arial" color="#FFBD00"><br>No commas allowed in headings.</font>
		</td>
		<% If PID <> "" Then Val = RVS("HasAccessories") End If %>
		<td width="100"><textarea cols="110" name="HasAccessories" rows="3"><%=Val%></textarea></td>
		<td></td>
	</tr>
	
	<tr>  <!-- [BN, 2/23/06] Added this row.  -->
		<td align=right valign="top">
			<font size="2" face="Arial" color="#FFBD00">Is an Accessory of:</font>
			<font size="1" face="Arial" color="#FFBD00"><br>Use double tilde <b>~~</b> to add a heading.</font>
			<font size="1" face="Arial" color="#FFBD00"><br>No commas allowed in headings.</font>
		</td>
		<% If PID <> "" Then Val = RVS("IsAccessoryOf") End If %>
		<td width="100"><textarea cols="110" name="IsAccessoryOf" rows="3"><%=Val%></textarea></td>
		<td></td>
	</tr>

	<%  
	If PID <> "" Then 
		RVS.Close 
	End If 
	%>

	<tr><td>&nbsp;</td></tr>
		
	<tr>
		<td align=center bgcolor=bbbbbb colspan=3><INPUT type="submit" NAME="Action" VALUE="<%=NextAction%>" ></td>
	</tr>
                
	<tr>
		<td colspan="3"><br><br></td>
	</tr>

	</FORM>  <!-- End of form2. -->
<% End If   ' (Action = "CreateNewProduct") OR (Action = "DisplayProductTo Edit") OR (Action = "DisplayProductToDelete") %>


	<% If FALSE Then	' If Action = "DeleteOldProduct" Then %>
		<!-- <FORM Name="form3" ACTION="SanctionProductInsert.asp" METHOD=POST>  -->
		<FORM Name="FORM3" ACTION="SanctionProduct.asp" METHOD="GET">       
		<tr>
			<td bgcolor="#790000"><font size="2" face="Arial" color="#FFFFFF"><b>PRODUCT TO PHYSICALLY DELETE:<b></font></td>

			<td width="100" bgcolor="#790000"><font size="2" face="Arial"></font>
			<%	SQL = "Select PName, PID, ITEMID from PRODUCT ORDER BY ITEMID ASC"
				Set conn = Server.CreateObject("ADODB.Connection")
	    		Conn.Open Session("ConnectionString")
	    		Set RS = Conn.Execute(SQL)
			%>
			
			<select name="PID" size="1" >
				<% Do While Not RS.EOF %>
				<option value="<%=RS("PID")%><%=Request("PID")%>">
				<font face ="arial" size="1"><%=RS("ITEMID")%>&nbsp;&nbsp;<%=RS("PName")%></font>
				</option>
				<% RS.MoveNext
				Loop
				RS.Close %>
			</select>
			</td>

			<td  bgcolor="#790000"><INPUT type="submit"  NAME="Action" VALUE="Delete"></td>
		</tr>
		</FORM>    
	<% End If   ' FALSE %>   

	</table>
	
</td>
</tr>

</table>


</body>

</html>
