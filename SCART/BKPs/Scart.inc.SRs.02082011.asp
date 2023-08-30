

<% ' 4/22/10: For clarity I factored this file, of all and only the Subs and Functions, out from the file Scart.inc. %>

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

' *************************************************************************************


Sub NewItemsCart(iItemID, iItemCount)
     If dictCart.Exists(iItemID) Then
        dictCart(iItemID) = iItemCount
    Else
        dictCart.Add iItemID, iItemCount
    End If  
End Sub

' *************************************************************************************


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

' *************************************************************************************


Sub ShowItemsInCart()
Dim Key
Dim bParameters ' as Variant (Array)
Dim sTotal, sShipping
%>

<br>
<table Border="0" CellPadding="7" CellSpacing="0" align="left">
    <tr bgcolor="#DDDDDD">
        <td align=center><font face="tahoma" size="2"><b>ID #</b></font></td>
        <td align=center><font face="tahoma" size="2"><b>Description</b></font></td>
        <td><font face="tahoma" size="2"><b>Select Quantity</b></font></td>
        <td align=center><font face="tahoma" size="2"><b>Price</b></font></td>
        <td align=center><font face="tahoma" size="2"><b>Totals</b></font></td>
        <td align=center><font face="tahoma" size="2"><b>Add?</b></font></td>
    </tr>

    <%
    sTotal = 0  
    weightTotal = 0  
    For Each Key in dictCart    ' Loop over all products in shopping cart.
        if Key = "" then
        dictCart.Remove Key
        exit for
        end if
        bParameters = asGetItemParameters(Key)
    %>

	<% ar = Replace( ar, " ", "%20") %>
	<% sar = Replace( sar, " ", "%20") %>
	<% Area = Replace( Area, " ", "%20") %>

	<tr>
		<td ALIGN="left" bgcolor="#EEEEEE">		<!-- [BN] ID # field. -->
		<font face="tahoma" size="2"><b><%= Key %></b></font>
		</td>

        <td ALIGN="Left" bgcolor="#EEEEEE"> 		<!-- [BN] Description field. -->
		<b><font face="tahoma" size="2"><a href="../detail.asp?pid=<%=(bParameters(5))%>"><%= bParameters(1) %></font>
		</a></b>
		</td>         
            
		<form>
		<td  ALIGN="Left" bgcolor="#EEEEEE" valign=middle>   <!-- [BN] Quantity field. -->
        	<input type="text" size="2" name="Qty" value="<%= dictCart(Key) %>">&nbsp;<input type="image" src="../Images/button1.gif" name="Qtybut" value="Submit" border="0" WIDTH="50" HEIGHT="14">
        	<input type="hidden" name="action" value="qty">
        	<input type="hidden" name="item" value="<%=Key%>">
        	<input type="hidden" name="sid" value="<%=sid%>">
        	<input type="hidden" name="Area" value="<%=Area%>">
        	<input type="hidden" name="Sar" value="<%=Sar%>">
		</td>
        </form>
            
		<td ALIGN="Right" bgcolor="#EEEEEE"> 		<!-- [BN] Price field -->
			<font face="tahoma" size="2"><b><%= bParameters(4)%></b></font>
		</td>

		<td ALIGN="Right" bgcolor="#EEEEEE">		<!-- [BN] Total Price field. -->
			<%	rowPrice = dictCart(Key) * CSng(bParameters(4)) 
				sTotal = sTotal + rowPrice 
			%>
			<font face="tahoma" size="2"><b>$<%=FormatNumber(rowPrice, 2)%></b></font>
		</td>
		
		<td bgcolor="#EEEEEE">
			<% 
			' 5/4/07, BN: Determine if there are ANY accessories for this product.
			' Recordset rsThisProduct is essentially that used in Details.asp (the product details page) to display a product's accessories.
			set rsThisProduct = CreateObject("ADODB.Recordset")
			AccessoriesSQL = "SELECT HasAccessories, SID FROM Product WHERE  PID = " + CStr(bParameters(5)) +   " "
			rsThisProduct.Open AccessoriesSQL, "DSN=STAREC1" , 1, 4
			' If this product has one or more accessories then ...
			If NOT IsNULL(rsThisProduct("HasAccessories")) AND Trim(rsThisProduct("HasAccessories")) <> "" Then
			%>
			<a href="../detail.asp?pid=<%=bParameters(5)%>#HasAccessories">
			<font face="tahoma" size="2">Accessories?</font>
			</a>
			<br>
			<%
			End If
			
			' 5/4/07, BN: Determine if there are ANY warranties for this product.
			' Recordset rsSubArea is essentially that used in Warranties/Warranties.INC included in Details.inc (the product details page) to display a product's warranties.
			SubAreaID = rsThisProduct("SID")   ' RS("SID")
			'Response.Write "<br>SubAreaID = " & SubAreaID
			Set rsSubArea = CreateObject("ADODB.Recordset")
			rsSubArea.Open "SELECT * FROM SubArea WHERE SID = " & SubAreaID  &   " ", "DSN=STAREC1" , 1, 4
			Warranties = rsSubArea("Warranties")
			'Response.Write Warranties
			If Warranties <> "" AND NOT IsNull(Warranties) Then		' If this product (actually, if the subarea it belongs to) has one or more warranties
			%>
			<a href="../detail.asp?pid=<%=bParameters(5)%>#Warranties">
			<font face="tahoma" size="2">Warranty?</font>
			</a>
			<%
			End If
			%>
		</td>

		<!--  <TD ALIGN="Right" bgcolor="#EEEEEE">   [BN] Total Weight field.	<font face="tahoma" size="2"><b><%=FormatNumber(dictCart(Key) * CSng(bParameters(8)),2) %> lbs.		</font></b>		</TD> -->
	</tr>
        
    <% 
    if bParameters(7) = true then
		osize = 8.75
	end if

    ' sTotal = sTotal + (dictCart(Key) * CSng(bParameters(4)))
    ' [BN, 4/29/04]: Added ...
    weightTotal = weightTotal + (dictCart(Key) * CSng(bParameters(8)))
	ExchangeRate = bParameters(9)    ' This is actually the same for each product in shopping cart, but what the heck.

    RXS.Close 
    Next	' End of For loop for outputing rows for the products in the shopping cart.
	%>

    <!-- [BN, 3/3/04] Begin output of fixed number of remaining rows of shopping cart display table. -->
    <tr>
    	<td COLSPAN="4" ALIGN="Right" bgcolor="white">
		<font face="tahoma" size="2"><b> Sub Total:</b></font>
		</td>
		<td ALIGN="Right" bgcolor="#EEEEEE">
		<font face="tahoma" size="2" color="#b9000"><b>$<%= FormatNumber(sTotal,2) %></b></font>
		</td>
	</tr>

	<tr>
		<td COLSPAN="4" ALIGN="Right" bgcolor="white"><font face="tahoma" size="2">
		<b> OverSize Charge:</b></font>
		</td>
		<td ALIGN="Right" bgcolor="#EEEEEE">
		  <% osizeTotal = osize * ExchangeRate %>
		<font face="tahoma" size="2" color="#b9000"><b>$<%= FormatNumber(osizeTotal , 2) %></b></font>
		</td>
	</tr>

	<% 
	iTotal = sTotal 
	sTotalOld = sTotal
	   sTotal = sTotal + osize

	'if sTotal = 0 then
	'    fTotal= FormatNumber(0,2)
	'else
	'    fTotal = FormatNumber((((iTotal+osize) * 0.0375)+7.95),2) 
	'end if

	fTotal = SandH(weightTotal, ExchangeRate)
	If Session("Country") = "Canada" Then
		ExtraShippingAmountForCanada = 4.0		' 4/18/07 BN: Sani wanted to charge more for shipping and handling to Canada.
		fTotal = fTotal + ExtraShippingAmountForCanada
	End If 
	If weightTotal = 0 Then fTotal = 0.00   ' [BN] Not worth the trouble to distinguish between U.S. and Canada excgange rate.
	%>

	<tr>
		<td COLSPAN="4" ALIGN="Right" bgcolor="white"><font face="tahoma" size="2">
		<b>Shipping and Handling (within North America):</b></font>
		</td>
		<td ALIGN="Right" bgcolor="#EEEEEE"><font face="tahoma" size="2" color="#b9000"><b>
		$<%=FormatNumber(fTotal,2)%>
		</b></font>
		</td>
	</tr>

	<tr>	
		<td COLSPAN="4" ALIGN="Right" bgcolor="white"><font face="tahoma" size="2">
		<b>Total:</b></font>
		</td>
		<td ALIGN="Right" bgcolor="#EEEEEE"> <font face="tahoma" size="2" color="#b9000">
		<% 
		If sTotal = 0 Then
		    gTotal = 0
		Else
		    ' gTotal = sTotal + (((iTotal+osize) * 0.0375)+7.95)
		    gTotal = sTotalOld + fTotal + osizeTotal 
		End If
		%>
		<b>$<%=FormatNumber(gTotal,2)%></b>
		</font>
		</td>
	</tr>

	<% If fTotal = MaxSandH Then %>
		<tr>
		<td COLSPAN="5" ALIGN="center" bgcolor="#DDDDDD"><font face="tahoma" size="2" color="#b9000"> <b>
		Shipping and Handling may need to be adjusted.<br>We will notify you by email.
		</b></font>
		</td></tr>
	<% End If %>


	<!--	<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><B>Weight Total:</B></TD>		<TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>		<%=FormatNumber(weightTotal,2)%>		</b>		</font>		</TD>	</TR>-->

	<tr>
		<td colspan="6" ALIGN="Center" bgcolor="#DDDDDD">
		<font face="tahoma" size="2"><a href="https://www.starlite-intl.com"><b>
		Continue Shopping?
		</b></a>
		</td>
	</tr>    

<% If gTotal <> 0 Then %>
	<tr>
		<td colspan="5" align='right'>
		&nbsp;
		</td>
		<td align=center>
		<a HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>">
		<img SRC="https://www.starlite-intl.com/images/shop_checkout.gif" BORDER="0" ALT="Checkout" WIDTH="46" HEIGHT="46">
		</a>
		<br><font face="tahoma" size="1"><b>CHECKOUT</b></font>
		</td>
	</tr>
<% Else %>
	<tr height=20><td></td></tr>
<% End If %>

	<tr>
		<td colspan="6">
		<center><font face="tahoma" size="2">
		<div style="text-align: center;"><span style="font-family: helvetica,arial,sans-serif;">
		All information provided is secured through SSL (Secure Socket Layers) and is kept confidential.</span><br>
		<span style="font-family: helvetica,arial,sans-serif;">Tax will be added according to user input and added to the bottom line automatically.</span><br>
		<span style="font-family: helvetica,arial,sans-serif;">Terms and Conditions of Sale apply.</span>
		</div>
		<br><big><span style="font-family: helvetica,arial,sans-serif;">
		Order online or by phone at: 1-800-387-8535</span></big>
		<br><br> 
		<!--  <script language=JavaScript src='https://seal.XRamp.com/seal.asp?type=G'></script> -->
		<script type="text/javascript" src="https://seal.XRamp.com/seal.asp?type=H"></script>
		</center>
		</td>
	</tr>
</table>

<br>


<%
End Sub

' *************************************************************************************

' 3/3/06, BN: Displays the "Category" page with the subCategory drop-down menu,
' or the "subCategory" page of products in a given subcategory SID.

Sub ShowFullCatalog()
Dim aParameters 
Dim I
Dim iItemCount ' 111 111 111 111 Number of items we sell

Response.Write(Session("Countdown"))
'on error resume next    
    iItemCount = countChoc 

ar	= Replace( ar, "%20", " ")		' Area (Category)
sar = Replace( sar, "%20", " ")		' SubArea (subCategory) 

if not(ar = "New!" OR ar = "Search") then 
	if sid = 0 then ' BN: i.e. We are on an Area (Category) page, not a subArea (SubCategory) page. 
			%>
			
			<table align='center' cellpadding=5 border=0 width='100%'>
			<tr>
				<td valign='top'  align=right>
				<font face="tahoma" size="4">
				<b>Category:</b>
				</font>
				</td>
				<td valign='top'>
				<font face="tahoma" size="4">
				<% if ar <> "New!" then %>
				<%=ar%>
				<%end if%>
				</font>
				</td>
			</tr>
			</table>
			
	<% Else ' BN: i.e. We are on a subArea (SubCategory) page, not an Area (Category) page. %>
		<%
		    Set conn = Server.CreateObject("ADODB.Connection")
		    Conn.Open Session("ConnectionString")
		    dim sdsqstring
		    ' sdsqstring = "select SubDesc from SubArea WHERE  Subname Like '" & sar & "'"
		    ' 3/3/06, BN: Select by SID, not SAR. Cleaner and more reliable, especially since SAR is wrong, but SID is correct,
		    ' when using my new method below, of 3/2/06, of choosing subArea from a drop-down menu.
		    sdsqstring = "select SubDesc, Subname from SubArea WHERE SID LIKE '" & sid & "'"
		    'Response.Write "<br>sdsqstring = " & sdsqstring
		    Set RSS = Conn.Execute(sdsqstring)
		%>

		<table align='left' cellpadding=5 border=0 width='100%'>
		<tr>
			<td valign='top' width='140'>
			<font face="tahoma" size="4">
			<b>Category:</b>
			</font>
			</td>
			
			<td valign='top'>
			<font face="tahoma" size="4">
			<% if ar <> "New!" then %>
			<%=ar%>
			<%end if%>
			</font>
			</td>
		</tr>

		<tr>
			<td valign='top'>
			<font face="tahoma" size="4">
			<b>&nbsp;&nbsp;&nbsp;&nbsp;Subcategory:</b>
			</font>
			</td>
			
			<td valign='top'>
			<font face="tahoma" size="4">
			<%if sar = "Manufa" then %>
				&nbsp;&nbsp;&nbsp;&nbsp;<% = ReQuest("Manufa") %>
			<%else%>
				&nbsp;&nbsp;&nbsp;&nbsp;<%=RSS("Subname")%>
			<%end if%>
			</font>
			</td>
		</tr>
		
		</table>
		<br><br><br><br><br>
		

		<%  SubAreaDescription = RSS("SubDesc")   ' !! Must transfer RSS("SubDesc") to a variable, as here, or else get weird behavior.
		    if (SubAreaDescription <> "") AND NOT IsNull(SubAreaDescription) then 
			' BN: Display subArea description ... 
			%>
			<table align='center' width="380" cellpadding="3" border="1" cellspacing="0" >
			<tr>
				<td align='center'>
				<font face="tahoma" size="2"><b><%=SubAreaDescription%></b></font>
				</td>
			</tr>
			</table>
		<% end if %>
		<% RSS.Close %>
				
	<% end if %>
		
<% end if %>


<% '******************************************** %>

<%
if sar <> "Manufa" then
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	dim sfqstring
	if ar = "New!" then
		ar = "New Products"
	end if
	sfqstring = "select AID from Area51 WHERE AreaName LIKE '" + ar  +"'   and AreaVisible = yes "
	Set RDS = Conn.Execute(sfqstring)
	tempvar = Int(RDS("AID"))
    
	sqll = "select * FROM SubArea WHERE AID = " & tempvar & " and Subvisible = yes ORDER BY Subname ASC "
	Set RDS = Conn.Execute(sqll)
%>

<% 
	if sid = 0 then '************************ELIMINATE SUB FOLLOW LINKS***************** 
	' i.e. We are on an Area (Category) page, not a SubArea (SubCategory) page. %>
	
	<% If False Then    ' BN: Old ugly method, that just lists subcategories, each with a hyperlink (no drop-down menu). %>
		<% do while not rDs.eof %>
		<% subname = Replace( RDS("Subname"), " ", "%20") %>
		<%  ar = Replace( ar, " ", "%20") %>
			<a href="scart.asp?sar=<%=subname%>&amp;area=<%=ar%>&amp;sid=<%=RDS("SID")%>">
			<font face="tahoma" size="2"><b><%=RDS("SubName")%></b>
			</a>
			</font>
			<font face="tahoma" size="2" color="#b70000">&nbsp;&nbsp;</font>
		<% rDs.movenext
		      loop
		      rDs.close  
		      conn.close
		%> 
	<% End If ' False %>


	<% ' 3/2/06: BN added this drop-down menu, to replace earlier hyperlinked text. %>
	<form action="https://www.starlite-intl.com/scart/scart.asp" method="GET" name="PID">
	<%
	    Set conn2 = Server.CreateObject("ADODB.Connection")
	    Conn2.Open Session("ConnectionString")
	    dim sfqstring2
	    if ar = "New!" then
	        ar = "New Products"
	    end if
	    ar = Replace( ar, "%20", " ")
	    sfqstring2 = "select AID from Area51 WHERE AreaName LIKE '" + ar  +"'   and AreaVisible = yes "
		Set RDS = Conn2.Execute(sfqstring2)
	    tempvar = Int(RDS("AID"))
	    
	    sqll = "select * FROM SubArea WHERE AID = " & tempvar & " and Subvisible = yes ORDER BY Subname ASC "
	    Set RDS = Conn2.Execute(sqll)
	
	%>

	<center>
	<b><font size=4 face=Tahoma>Subcategories:</font></b>&nbsp;&nbsp;
	<select name="sid">
	<option>Please Select <%  	    
		rDs.movefirst()
	    do while not rDs.eof 
	%>

	<option value='<%=RDS("SID")%>'><%=RDS("SubName")%> <% rDs.movenext
	      loop
	      rDs.close  
	      conn2.close
	%> 
	</select>
	
	<input type='hidden' name='sar' value='<%=subname%>'> <% ' BN: This is ignored. SAR should be retrieved according to the SID. %>
	<input type='hidden' name='area' value='<%=ar%>'>
	<input type="submit" value="Submit">
	</center>
	</form>
	

	<% 
	Set conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	dim ssdsqstring2
	ssdsqstring2 = "select AreaDesc from Area51 WHERE  AreaName Like '" & ar & "'"
	Set RSS = Conn.Execute(ssdsqstring2)
		
	AreaDescription = RSS("AreaDesc")	  ' !! Must transfer RSS("AreaDesc") to a variable, as here, or else get weird behavior.
	If (AreaDescription <> "") AND NOT IsNull(AreaDescription) Then 
		' BN: Display Area description ... %>
		<table align='center' width="380" cellpadding="3" border="0" cellspacing="0">
		<tr>
			<td align='center'>
			<font face="tahoma" size="2"><%=AreaDescription%></font>
			</td>
		</tr>
		</table>
	<% Else 
		'Response.Write "<br><br>"
	End If %>
	<% RSS.Close %>
			

	<% end if%>
<% end if%>


 <% if (SID = 0)  AND ((ar <> "Search") or (sar = "New!")) Then%>
 <div align="center"><center>

   <!--#include file="SPECIALK.INC"-->
 <br><br>

</div>
</center>
<%end if%>

    <table Border="0" CellPadding="3" CellSpacing="10" width="100%">
      
    <%
    If sid <> "0" AND iItemCount > 0 Then   ' i.e. We are on a SubCategory page and there are (more than 0) products.
		'Response.Write "<br>iItemCount = " & iItemCount 
	%>
		<font face="Tahoma" size="2" color="navy">
		<br><%=iItemCount%> Products. <b>Click on an image</b> to see a detailed description and larger image for that product</font>
	<% End If 
    For I = 1 to iItemCount
        aParameters = GetItemParameters(I)
        %>
        <tr>

<%
      xxxx = (aParameters(5))
      If FALSE Then
      	Response.Write "<br>i = " & i & " xxxx = " & xxxx
      	Response.Write " Par(0) = " & aParameters(0)
      	Response.Write " Par(1) = " & aParameters(1)
      	Response.Write " Par(2) = " & aParameters(2)
      	Response.Write " Par(3) = " & aParameters(3)
      	Response.Write " Par(4) = " & aParameters(4)
      	Response.Write " Par(5) = " & aParameters(5)
      	Response.Write " Par(6) = " & aParameters(6)
      	Response.Write " Par(7) = " & aParameters(7)
      	Response.Write " Par(8) = " & aParameters(8)
      End If
%>


		<td valign='top'><font size=2 face='Arial'><%=I%></font></td>
		<td valign="top" width="100">
		<font face="Tahoma" size="1"><b><i><%= aParameters(1) %></i></b></font><br>
		<a href="../Detail.asp?pid=<%=xxxx%>&amp;Key=">
		<img SRC="<%= aParameters(0) %>" border="0" width="50"></a>
		<% 
		NewProduct = RS("NewProduct")
		If NewProduct Then
			NewIcon = "https://www.starlite-intl.com/imi/new1.gif"
			Response.Write "<img src='" & NewIcon & "' style='border: 0px solid ;' >"
		End If
		%>
		</td>


<% area1 = Replace( ar, " ", "%20") %>
<% sarea1 = Replace( sar, " ", "%20") %>

<!-- 5/5/05, BN: Original of case 1: 		<TD valign="top"><A HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=aParameters(6)%>&count=1&amp;sid=<%=aParameters(7)%>&amp;Area=<%=area1%>&amp;sar=<%=sarea1%>"><img src="https://www.starlite-intl.com/images/order.gif" border="0" height="15px" ></A><br>         '' 5/5/05, BN: Original of case 2:  		<TD valign="top"><A HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=aParameters(6)%>&count=1&amp;sid=<%=sid%>&amp;Area=<%=area1%>&amp;sar=<%=sarea1%>"><img src="https://www.starlite-intl.com/images/order.gif" border="0"></A><br>    -->
<% if sar = "Manufa" or sar = "New!" then %>
		<td valign="top"><a HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&amp;item=<%=aParameters(6)%>&amp;count=1&amp;sid=<%=aParameters(7)%>&amp;Area=&amp;sar="><img src="https://www.starlite-intl.com/images/order.gif" border="0" ></a><br>
<%else%>
 		<td valign="top"><a HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&amp;item=<%=aParameters(6)%>&amp;count=1&amp;sid=<%=sid%>&amp;Area=&amp;sar="><img src="https://www.starlite-intl.com/images/order.gif" border="0" ></a><br>
<% end if %>
          <font face="Tahoma" size="1"><b><u>ID # <%=aParameters(6)%></b></u></font><br>
          <font face="Tahoma" size="1">Reg. Price </font> <font face="Tahoma" size="2"><%= aParameters(3) %></font><br>
			<font face="Tahoma" size="2"><b>Our Price </b></font> 
			<font face="Tahoma" size="2" color="#B90000"><b><i>
<% ' [BN, 3/1/04]
'= aParameters(4)
If aParameters(8) <> true Then %>
   <%=aParameters(4) %>
   <%Else%>
   <font face="Tahoma" size="1">
   <b><i>Click Order Button to order or see our price.</i></b></font>
<% End If %>
</i></b></font>
                          
           	</td>

            <td valign="top"><font face="Tahoma" size="2"><%=aParameters(2)%> </font>
		</td>
   	</tr>

        <%
     RS.MoveNext
    Next 'I
'172
    %>



    </table>
    <%
End Sub   ' ShowFullCatalog()

' *************************************************************************************


Sub PlaceOrder()
Dim Key
Dim aParameters ' as Variant (Array)
Dim sTotal, sShipping

    %>
    <table Border="0" CellPadding="3" CellSpacing="2"><tr>
       <td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>ID #</b></font></td>
     
            <td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Description</b></font></td>
            <td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Qty.</b></font></td>
            
            <td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Price</b></font></td>
            <td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Totals</b></font></td></tr>
    <%
    sTotal = 0    
    For Each Key in dictCart
        
        aParameters = asGetItemParameters(Key)
        %>
        <tr>
            <td ALIGN="Left" bgcolor="#EEEEEE"><%= aParameters(6) %></td>
            <td ALIGN="Left" bgcolor="#EEEEEE"><%= aParameters(1) %></td>
            <td ALIGN="Center" bgcolor="#EEEEEE"><%= dictCart(Key) %></td>
            <td ALIGN="Right" bgcolor="#EEEEEE"><%= aParameters(4) %></td>
            <td ALIGN="Right" bgcolor="#EEEEEE">$<%= FormatNumber(dictCart(Key) * CSng(aParameters(4)),2) %></td>
        </tr>
        
        <%
        if aParameters(7) = true then
              osize = 8.75
               end if
        %> 
        
        <%
        sTotal = sTotal + (dictCart(Key) * CSng(aParameters(4)))
        
        ' [BN, 4/29/04]: Added ...
        weightTotal = weightTotal + (dictCart(Key) * CSng(aParameters(8)))
	  ExchangeRate = aParameters(9)    ' This is actually the same for each product in shopping cart, but what the heck.

    RXS.Close
    Next
    
    
    %>
    
<tr><td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2"><b> Sub Total:</b></font></td><td ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2" color="#b9000"><b>$<%= FormatNumber(sTotal,2) %></b></font></td></tr>

	<tr>
	<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2">
	<b> OverSize Charge:</b></font></td><td ALIGN="Right" bgcolor="#DDDDDD">
      <% osizeTotal = osize * ExchangeRate %>
	<font face="tahoma" size="2" color="#b9000"><b>$<%= FormatNumber(osizeTotal , 2) %></b></font></td>
	</tr>


<% iTotal = sTotal%>
<% sTotalOld = sTotal
   sTotal = sTotal + osize%>


<% 
'if sTotal = 0 then
'    fTotal= FormatNumber(0,2)
'else
'    fTotal = FormatNumber((((iTotal+osize) * 0.0375)+7.95),2) 
'end if
%>

<% 
fTotal = SandH(weightTotal, ExchangeRate) 
If Session("Country") = "Canada" Then
	ExtraShippingAmountForCanada = 4.0		' 4/18/07 BN: Sani wanted to charge more for shipping and handling to Canada.
	fTotal = fTotal + ExtraShippingAmountForCanada
End If 
If weightTotal = 0 Then fTotal = 0.00   ' [BN] Not worth the trouble to distinguish between U.S. and Canada exchange rate.
%>

	<tr>
	<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><b>Shipping and Handling (within North America):</b></td>
    	<td ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2" color="#b9000"><b>$<%=FormatNumber(fTotal,2)%>
      </b></font></td>
	</tr>


	<tr>	<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><b>Total:</b></td>
		<td ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2" color="#b9000"><b>$<% 
if sTotal = 0 then
    gTotal = 0
else
    ' gTotal = sTotal + (((iTotal+osize) * 0.0375)+7.95)
    gTotal = sTotalOld + fTotal + osizeTotal
end if%>
<%=FormatNumber(gTotal,2)%>
		</b>
		</font>
		</td>
	</tr>


<% If fTotal = MaxSandH  Then %>
	<tr>
	<td COLSPAN="5" ALIGN="center" bgcolor="#DDDDDD"><font face="tahoma" size="2" color="#b9000"> <b>
	Shipping and Handling may need to be adjusted.<br>We will notify you by email.
	</b></font>
	</td></tr>
<% End If %>


<!--	<TR>	<TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><B>Weight Total:</B></TD>		<TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>		<%=FormatNumber(weightTotal,2)%>		</b>		</font>		</TD>	</TR>-->


<tr><td colspan="5" ALIGN="center" bgcolor="#DDDDDD">
<font face="tahoma" size="2"><a href="scart.asp?area=Accessories&amp;sid=0"><b>
Need Accessories?
</b></a>
</td></tr>    
<tr><td colspan="5"><br><b>Please choose from the menus above to continue to shop.</b><br><br>
</td></tr>
    </table>

    <%
End Sub

' *************************************************************************************


' We implemented this this way so if you attach it to a database you'd only need one call per item
Function asGetItemParameters(iItemID)

Dim bParameters

if Session("Country") = "USA" then
       	csql = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID +   "' "         
       		'Response.Write csql & "<br>"     
     		RXS.Open csql, "DSN=STAREC1" , 1, 4
     	   ' Response.Write "RXS(ITEMID) = " & RXS("ITEMID") & "<br>"
		' 6/18/06, commented out, BN: bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"), RXS("ITEMID"),RXS("OverSize"), RXS("Weight"), 1 )
		               bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"), RXS("ITEMID"),RXS("OverSize"), RXS("Weight"), 1 )
else
            csql = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID +   "' "
           	RXS.Open csql, "DSN=STAREC1" , 1, 4
         	bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Exch") ), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")*(1/(1-(RXS("GPM"))))),RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), RXS("Exch") )
           ' 6/18/06, commented out, BN: bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Freight")* RXS("Exch") ), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")*(1/(1-(RXS("GPM"))))),RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), RXS("Exch") )
           'bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"), RXS("ITEMID"), RXS("OverSize") )
end if

    
' Return array containing product info.
asGetItemParameters = bParameters

End Function

' *************************************************************************************


Function GetItemParameters(iItemID)
Dim aParameters 
' [BN, 3/3/04] Appended ShowPrice field.

if Session("Country") = "USA" then       
 	    ' 6/18/06, commented out, BN: aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("PName") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Freight")), formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM"))))), RS("PID"),RS("ITEMID"),RS("SID"), RS("ShowPrice"))
									  aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("PName") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")), formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM"))))), RS("PID"),RS("ITEMID"),RS("SID"), RS("ShowPrice"))
else
    	' 6/18/06, commented out, BN: aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("Pname") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Duty")*RS("Freight")*RS("Exch")), formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM"))))),RS("PID"),RS("ITEMID"),RS("SID"),  RS("ShowPrice"))
    								  aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("Pname") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Duty")*RS("Exch")), formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM"))))),RS("PID"),RS("ITEMID"),RS("SID"),  RS("ShowPrice"))
end if     
' Return array containing product info.
GetItemParameters = aParameters
End Function

' *************************************************************************************
%>