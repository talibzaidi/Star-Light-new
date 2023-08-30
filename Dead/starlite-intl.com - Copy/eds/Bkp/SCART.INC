<% session("initq") = 1 %>
<%
'on error resume next

sAction = CStr(Request.QueryString("action"))
iItemID = (Request.QueryString("item"))
iItemCount = CInt(Request.QueryString("count"))

if (sar = "New%20Products" OR sar = "New Products" OR sar = "New!") then 


	 set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  NewProduct = yes ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4	



elseif sar ="Specials" then 

	response.redirect "scartstart.asp"

elseif sar ="Manufa" then 

	set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE  Manufa Like '" & ReQuest("Manufa") & "' ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4	
	ar="Search"
	

else


	 set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  SID = " + sid +   " ORDER BY MSL ASC ", "DSN=STAREC1" , 1, 4
	
end if
	countChoc = RS.RecordCount
	set RXS = CreateObject("ADODB.Recordset")
	
        sar = Replace( sar, " ", "%20") 
	


%>

<% 
' but this sample always just adds one.  If you wish to add different
' quantities simply replace the value of the Querystring parameter count.

Sub AddItemToCart(iItemID, iItemCount)
 	If dictCart.Exists(iItemID) Then
		dictCart(iItemID) = dictCart(iItemID) + iItemCount
	Else
		dictCart.Add iItemID, iItemCount
	End If
	
End Sub


Sub NewItemsCart(iItemID, iItemCount)
 	If dictCart.Exists(iItemID) Then
		dictCart(iItemID) = iItemCount
	Else
		dictCart.Add iItemID, iItemCount
	End If
	
End Sub


Sub RemoveItemFromCart(iItemID, iItemCount)
 	If dictCart.Exists(iItemID) Then
		If dictCart(iItemID) <= iItemCount Then
			dictCart.Remove iItemID
		Else
			dictCart(iItemID) = dictCart(iItemID) - iItemCount
		End If
		
	Else
		Response.Write("<font face=tahoma size=2><b>")
		Response.Write "Couldn't find any of that item your cart.<BR><BR>" & vbCrLf
		Response.Write("</font></b>")
	End If
End Sub
'57*******************************************************************************************************************************************************
Sub ShowItemsInCart()
Dim Key
Dim bParameters ' as Variant (Array)
Dim sTotal, sShipping
	
	%>
	<TABLE Border=0 CellPadding=3 CellSpacing=2>
	    <TR>
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>ID #</b></font></TD>
	 
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Description</b></font></TD>
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Select Quantity</b></font></TD>
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Price</b></font></TD>
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Totals</b></font></TD>
	    </TR>
	<%
	sTotal = 0
	For Each Key in dictCart
		
		bParameters = asGetItemParameters(Key)
		
		%>

<% ar = Replace( ar, " ", "%20") %>
<% sar = Replace( sar, " ", "%20") %>
<% Area = Replace( Area, " ", "%20") %>
		<TR>
			<TD ALIGN="Center" bgcolor="#EEEEEE"><font face="tahoma" size="2"><b><%= Key %></b></font></TD>
			<TD ALIGN="Left" bgcolor="#EEEEEE"><b><font face="tahoma" size="2"><a href="../detail.asp?pid=<%=(bParameters(5))%>"><%= bParameters(1) %></font></a></b></TD>
			
			<TD width=165 ALIGN="Left" bgcolor="#EEEEEE"><font face="tahoma" size="1">
		
		<form><input type="text" size="2" name="Qty"  value="<%= dictCart(Key) %>">
	
		<input type=hidden name="action" value="qty">
		<input type=hidden name="item" value="<%=Key%>">
		<input type=hidden name="sid" value="<%=sid%>">
		<input type=hidden name="Area" value="<%=Area%>">
		<input type=hidden name="Sar" value="<%=Sar%>">
		<input type=image src="../Images/button1.gif" name="Qtybut" value="Submit" border=0>

		</form></A></b></TD>
			


			<TD ALIGN="Right" bgcolor="#EEEEEE"><font face="tahoma" size="2"><b><%= bParameters(4) %></font></b></TD>
			<TD ALIGN="Right" bgcolor="#EEEEEE"><font face="tahoma" size="2"><b>$<%= FormatNumber(dictCart(Key) * CSng(bParameters(4)),2) %></font></b></TD>
		</TR>

		<%
		sTotal = sTotal + (dictCart(Key) * CSng(bParameters(4)))
		RXS.CLose
	Next
	
		%>
	<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 ><B> Sub Total:</B></font></TD><TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$<%= FormatNumber(sTotal,2) %></b></font></TD></TR>
<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 ><B> OverSize Charge:</B></font></TD><TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$<%= FormatNumber(osize,2) %></b></font></TD></TR>

<% iTotal = sTotal%>
<% sTotal = sTotal + osize%>
<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><B>Freight and Insurance:</B></TD><TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$<%= FormatNumber((iTotal * 0.04),2) %></b></font></TD></TR>
<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><B>Total:</B></TD><TD ALIGN="Right" 
bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$
<% gtotal = sTotal + (iTotal * 0.04)%>
<%= FormatNumber(gTotal,2) %></b></font></TD></TR>
<tr><td colspan=6 ALIGN="Center" bgcolor="#DDDDDD">
<font face="tahoma" size="2"><a href="./scart.asp?area=Accessories&amp;sid=0"><b>
Need Accessories?
</b></a>
</td></tr>	

<tr><td colspan=6><center><font face=tahoma size=2>
<br><b><img src="hand2.gif"> Please choose a category from the left to continue to shop.</b><br><Br>

<br><b>All information provided is secured through SSL (Secure Socket Layers) and is kept confidential. Tax will be added according to user input and added automatically. <br><br> Order online or by phone at: 1-800-387-8535</b></font></center></td></tr>


</TABLE><br>

<%
End Sub
'92******************************************************************************************************************************************************
Sub ShowFullCatalog()
Dim aParameters 
Dim I
Dim iItemCount ' 111 111 111 111 Number of items we sell

Response.Write(Session("Countdown"))
'on error resume next	
	iItemCount = countChoc 
	
	%>
<% ar = Replace( ar, "%20", " ") %>
<% sar = Replace( sar, "%20", " ") %>
<font face="tahoma" size="4"><b>
<% if ar <> "New!" then %>
<%=ar%>:<br>
<%end if%>
<%if sar = "Manufa" then %>
<% = ReQuest("Manufa") %>
<%else%>
<%=sar%>
<%end if%>

<b></font>
<%'*********************************%>
<% if not(ar = "New!" OR ar = "Search") then%>
<% if sid = 0 then %>
<%
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	dim ssdsqstring
	ssdsqstring = "select AreaDesc from Area51 WHERE  AreaName Like '" & ar & "'"
        Set RSS = Conn.Execute(ssdsqstring)
%>
<% 'if RSS("AreaDesc") <> "" then %>
<center><table width="380" cellpadding="3" border="1" cellspacing="0" bordercolor="#000099"><tr><td >
<font face="tahoma" size="2"><b>&nbsp;&nbsp;<%=RSS("AreaDesc")%></b></font>
</td></tr></table></center>
<% 'end if%>
<% RSS.Close %>
<% else%>
<%
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	dim sdsqstring
	sdsqstring = "select SubDesc from SubArea WHERE  Subname Like '" & sar & "'"
        	Set RSS = Conn.Execute(sdsqstring)
%>
<% 'if RSS("SubDesc") <> "" then %>
<center><table width="380" cellpadding="3" border="1" cellspacing="0" bordercolor="#000099"><tr><td >
<font face="tahoma" size="2"><b>&nbsp;&nbsp;<%=RSS("SubDesc")%></b></font>
</td></tr></table></center>
<% 'end if%>
<% RSS.Close %>
<% end if %>
<% end if%>
<%'********************************************%>
<font face="tahoma" size="1"><br><br></font>

<%
	if sar <> "Manufa" then
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	dim sfqstring
	sfqstring = "select AID from Area51 WHERE AreaName LIKE '" + ar  +"'   and AreaVisible = yes "
        	Set RDS = Conn.Execute(sfqstring)
	tempvar = Int(RDS("AID"))
	
	sqll = "select * FROM SubArea WHERE AID = " & tempvar & " and Subvisible = yes ORDER BY Subname ASC "
	Set RDS = Conn.Execute(sqll)
%>
<% if sid = 0 then '************************ELIMINATE SUB FOLLOW LINKS***************** %>
<% do while not rDs.eof %>
<% subname = Replace( RDS("Subname"), " ", "%20") %>
<% ar = Replace( ar, " ", "%20") %>
<a href="scart.asp?sar=<%=subname%>&amp;area=<%=ar%>&amp;sid=<%=RDS("SID")%>"><font face="tahoma" size="2"><b><%=RDS("SubName")%>
</b></a></font><font face=tahoma size=2 color="#b70000">&nbsp;&nbsp;</font>
<% rDs.movenext
      loop
      rDs.close  
      conn.close
%> 
<% end if%>
<% end if%>
<br>
	<font face=arial size=2 color=#0000CC><b>Click on an image to see a detailed description and larger image
for that product<b></font>
<br></center>
 <% if (SID = 0)  AND ((ar <> "Search") or (sar = "New!")) Then%>
 <div align=center><center>

   <!--#include file="SPECIALK.INC"-->
 <br><br>

</div></center>
<%end if%>
	<TABLE Border=0 CellPadding=3 CellSpacing=1 width="100%">
	  
	<%
	For I = 1 to iItemCount
		aParameters = GetItemParameters(I)
		%>
	    <TR>

<%
      xxxx = (aParameters(5))
%>
	        <TD valign="top" width="100"><A href="../Detail.asp?pid=<%=xxxx%>&amp;Key="><IMG SRC="<%= aParameters(0) %>" border="0" width="100"></a></TD>


<% area1 = Replace( ar, " ", "%20") %>
<% sarea1 = Replace( sar, " ", "%20") %>


<% if sar = "Manufa" or sar = "New!" then %>
    <TD valign="top"><A HREF="./scart.asp?action=add&item=<%=aParameters(6)%>&count=1&amp;sid=<%=aParameters(7)%>&amp;Area=<%=area1%>&amp;sar=<%=sarea1%>"><img src="../images/order.gif" border="0"></A><br>  
                       
<%else%>
 <TD valign="top"><A HREF="./scart.asp?action=add&item=<%=aParameters(6)%>&count=1&amp;sid=<%=sid%>&amp;Area=<%=area1%>&amp;sar=<%=sarea1%>"><img src="../images/order.gif" border="0"></A><br>  
<% end if %>
	      <Font face="Tahoma" size="1"><b><u>ID # <%=aParameters(6)%></b></u></font><br>
	      <Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2"><%= aParameters(3) %></font>
	      <br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%= aParameters(4) %></i></b></font>
				      

                       
	       </TD>

	        <TD valign="top" >&nbsp;&nbsp;<Font face="Tahoma" size="2"><b><%=aParameters(2)%></b> </font></TD>
   </tr><tr>
	        <TD valign="top" colspan="3">&nbsp;&nbsp;<Font face="Tahoma" size="3"><b><i><%= aParameters(1) %></i></b> </font><br><br>
</td></TD>
	    </TR>
		<%
	 RS.MoveNext
	Next 'I
'172
	%>



	</TABLE>
	<%
End Sub

Sub PlaceOrder()
Dim Key
Dim aParameters ' as Variant (Array)
Dim sTotal, sShipping
	
	%>
	<TABLE Border=0 CellPadding=3 CellSpacing=2><tr>
	   <TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>ID #</b></font></TD>
	 
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Description</b></font></TD>
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Qty.</b></font></TD>
			
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Price</b></font></TD>
			<TD bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Totals</b></font></TD></tr>
	<%
	sTotal = 0
	For Each Key in dictCart
		
		aParameters = asGetItemParameters(Key)
		%>
		<TR>
			<TD ALIGN="Center" bgcolor="#EEEEEE"><%= aParameters(6) %></TD>
			<TD ALIGN="Left" bgcolor="#EEEEEE"><%= aParameters(1) %></TD>
			<TD ALIGN="Center" bgcolor="#EEEEEE"><%= dictCart(Key) %></TD>
			<TD ALIGN="Right" bgcolor="#EEEEEE"><%= aParameters(4) %></TD>
			<TD ALIGN="Right" bgcolor="#EEEEEE">$<%= FormatNumber(dictCart(Key) * CSng(aParameters(4)),2) %></TD>
		</TR>
		<%
		sTotal = sTotal + (dictCart(Key) * CSng(aParameters(4)))
	RXS.Close
	Next
	
	
	%>
	
<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 ><B> Sub Total:</B></font></TD><TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$<%= FormatNumber(sTotal,2) %></b></font></TD></TR>
<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 ><B> OverSize Charge:</B></font></TD><TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$<%= FormatNumber(osize,2) %></b></font></TD></TR>

<% iTotal = sTotal%>
<% sTotal = sTotal + osize%>
<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><B>Freight and Insurance:</B></TD><TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$<%= FormatNumber((iTotal * 0.04),2) %></b></font></TD></TR>
<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><B>Total:</B></TD><TD ALIGN="Right" 
bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$
<% gtotal = sTotal + (iTotal * 0.04)%>
<%= FormatNumber(gTotal,2) %></b></font></TD></TR>
<tr><td colspan=5 ALIGN="center" bgcolor="#DDDDDD">
<font face="tahoma" size="2"><a href="./scart.asp?area=Accessories&amp;sid=0"><b>
Need Accessories?
</b></a>
</td></tr>	
<tr><td colspan=5><br><b><img src="hand2.gif"> Please choose a category from the left to continue to shop.</b><br><Br>
</td></tr>
	</TABLE>
	<%
End Sub

'217 We implemented this this way so if you attach it to a database you'd only need one call per item
'*****************************************************************************************************************************
' 
Function asGetItemParameters(iItemID)

Dim bParameters 

                                 if Session("Country") = "USA" then       
                              RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  ITEMID = " + iItemID +   " ", "DSN=STAREC1" , 1, 4

		 bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"),RXS("ITEMID") )
else
                 RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  ITEMID = " + iItemID +   " ", "DSN=STAREC1" , 1, 4
		 bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Freight")*RXS("Exch")), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")*(1/(1-(RXS("GPM"))))),RXS("PID"),RXS("ITEMID") )
		'bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"),RXS("ITEMID") )

end if

	
' Return array containing product info.
asGetItemParameters = bParameters

End Function


'***************************************************************************************************************************************************
Function GetItemParameters(iItemID)

Dim aParameters 
                                 if Session("Country") = "USA" then       
                              
		 aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("PName") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Freight")), formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM"))))), RS("PID"),RS("ITEMID"),RS("SID") )
	else
	 	 aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("Pname") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Duty")*RS("Freight")*RS("Exch")), formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM"))))),RS("PID"),RS("ITEMID"),RS("SID") )
		

end if

		 
' Return array containing product info.
GetItemParameters = aParameters
End Function
%>




<% ' ***** Begin the infamous runtime script *****
' Declare our Vars
Dim dictCart ' as dictionary
Dim sAction ' as string
Dim iItemID ' as integer
Dim iItemCount ' as integer

' Get a reference to the cart if it exists otherwise create it
If IsObject(Session("cart")) Then
	Set dictCart = Session("cart")
Else
	' We use a dictionary so we can name our keys to correspond to our
	' item numbers and then use their value to hold the quantity.  An
	' array would also work, but would be a little more complex and
	' probably not as easy for readers to follow.
	Set dictCart = Server.CreateObject("Scripting.Dictionary")
End If

' Get all the parameters passed to the script

%>
<center>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>
<TR><TD>
<% ar = Replace( ar, " ", "%20") %>
<% sar = Replace( sar, " ", "%20") %>
<% Area = Replace( Area, " ", "%20") %>
<% sid = Replace( sid, " ", "%20") %>
<% '************************QUANTITY AREA **** QTY AREA *******************************
' Select action based on user input
Select Case sAction
	Case "add"
		AddItemToCart iItemID, iItemCount
		ShowItemsInCart
		%>
		</TD></TR>
		<TR><TD ALIGN="right">
		<A HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><IMG SRC="../images/shop_checkout.gif" BORDER=0  ALT="Checkout"></A><BR><font face=tahoma size=1><b>CHECKOUT</b></font><br><br><br>
		<%
	Case "del"
		RemoveItemFromCart iItemID, iItemCount
		ShowItemsInCart
		%>
		<%
	Case "qty"

		NewItemsCart iItemID, CDbl(Request("Qty"))
		'AddItemToCart iItemID, CDbl(Request("Qty"))
		ShowItemsInCart
		%>
		</TD></TR>
		<TR><TD ALIGN="right">


		<A HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><IMG SRC="../images/shop_checkout.gif" BORDER=0  ALT="Checkout"></A><BR><font face=tahoma size=1><b>CHECKOUT</b></font><br><br><br>
		<%
	Case "viewcart"
		ShowItemsInCart
		%>
		</TD></TR>
		<TR><TD ALIGN="right">


		<A HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><IMG SRC="../images/shop_checkout.gif" BORDER=0  ALT="Checkout"></A><BR><font face=tahoma size=1><b>CHECKOUT</b></font><br><br><br>
		<%
	Case "checkout"
		PlaceOrder
		%>
		</TD></TR>
		<TR><TD ALIGN="left">
		<BR>
		<b>Please enter the following information to complete your order. <br> Your order will be shipped when method of payement is verified or personal check clears. In the event that there is a problem in processing your order you will be notified via phone or e-mail.<br> For further inquiries call: 1-800-387-8535 or contact us <a href="mailto:starlite@starlite-intl.com">here</a>.</b>
		<form action="sendmail.asp">
<p>&nbsp;</p>
    <div align="center"><center><table border="0" width="350">
        <tr>
            <td><font size="2" face="Tahoma"><strong>Name:</strong></font></td>
            <td><strong><input
            type="text" size="25" name="Name"></strong></td>
        </tr>
         <tr>
            <td><font size="2" face="Tahoma"><strong>E-Mail Address:</strong></font></td>
            <td><input
            type="text" size="25" name="Email" value="email@domain.com"></td>
        </tr>
	 <tr>
            <td><font size="2" face="Tahoma"><strong>Phone Number:</strong></font></td>
            <td><input
            type="text" size="25" name="Phone" value=""></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Street
            Address:</strong></font></td>
            <td><input
            type="text" size="25" name="Address"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>City:</strong></font></td>
            <td><input
            type="text" size="25" name="City"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>State/Province:</strong></font></td>
            <td><input
            type="text" size="25" name="StateProv"></td>
        </tr>
         <tr>
            <td valign=top><font size="2" face="Tahoma"><strong>If you live in Alaska, Nunavut, or Northwest Territories Check this Box:</strong></font></td>
            <td valign=top><input
            type="checkbox" name="BigShip"><font size="2" face="Tahoma"><strong>If you live in Alaska, Nunavut, or Northwest Territories: freight will be calculated at 3 times the above freight amount.</strong></font></td>
        </tr>
   <tr>
            <td><font size="2" face="Tahoma"><strong>Postal Code:</strong></font></td>
            <td><input
            type="text" size="25" name="Postal"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Country:</strong></font></td>
            <td><input
            type="text" size="25" name="Country"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Payment By:</strong></font></td>
            <td><select
            name="Payment" size="1">
                <option selected value>Choose an Option</option>
                <option value="Visa">Visa</option>
                <option value="Master Card">Master Card</option>
                <option value="American Express">American Express</option>
                <option value="Cheque">Personal Cheque</option>
                <option value="Money Order">Money Order</option>
            </select></td>
        </tr>
       
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>	
	<tr>
            <td colspan=2><font size="2" face="Tahoma"><strong> Taxes will be applied to total order plus freight.</strong></font>&nbsp;</td>
            
        </tr>
        
        <tr>
            <td><strong>State / Provincial Tax:</strong></td>
            <td><input
            type="text" size="3" name="Taxx1" value="0">% </td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Federal Tax:</strong></font></td>
            <td><input
            type="text" size="3" name="Taxx2" value="0">%</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        <tr>
            <td align="center" colspan="2"><font size="2"
            face="Tahoma"><strong><input type="submit" name="B3"
            value="Place My Order"></strong></font></td>
        </tr>
    </table>
    </center></div>



		</form>

		<%
	Case Else ' Shop
		ShowFullCatalog
		%>
		</TD></TR>
		<TR><TD ALIGN="right">
		
		<%
End Select

' Return cart to Session for storage
Set Session("cart") = dictCart
%>
</TD></TR>
</TABLE>
</center>