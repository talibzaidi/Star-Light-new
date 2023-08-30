

<% ' 4/22/10: For clarity I factored this file, of all and only the Subs and Functions, out from the file Scart.inc. %>

<% 
' but this sample always just adds one.  If you wish to add different
' quantities simply replace the value of the Querystring parameter count.

Sub AddItemToCart(iItemID, iItemCount)
    iItemID = Trim(iItemID)     ' Added 9/23/11.
	If dictCart.Exists(iItemID) Then
        dictCart(iItemID) = dictCart(iItemID) + iItemCount
    Else
        dictCart.Add iItemID, iItemCount
    End If
%>


<!-- Google Code for Shopping Cart Conversion Page -->
<!-- BN, 10/23/13: Updates count at Google whenever someone actually adds something to the shopping cart. 
     This code comes from Sani's Google AdWords campaign, under "Conversions".
-->
<script type="text/javascript">
/* <![CDATA[ */
var google_conversion_id = 950708232;
var google_conversion_language = "en";
var google_conversion_format = "3";
var google_conversion_color = "ffffff";
var google_conversion_label = "sFKtCKjBiwwQiNCqxQM";
var google_conversion_value = 0;
var google_remarketing_only = false;
/* ]]> */
</script>
<script type="text/javascript"  
src="//www.googleadservices.com/pagead/conversion.js">
</script>
<noscript>
<div style="display:inline;">
<img height="1" width="1" style="border-style:none;" alt=""  
src="//www.googleadservices.com/pagead/conversion/950708232/?value=0&amp;label=sFKtCKjBiwwQiNCqxQM&amp;guid=ON&amp;script=0"/>
</div>
</noscript>

<%

%>
<!-- Google code for shopping cart conversion Page -->
<!-- SZ, 06/16/2015: Updates count at Google whenever someone actually adds something to the shopping cart. 
     This code comes from Sani's Google Adwords 'Shopping Campaign' under conversions measuring "ROI".

-->
<!-- Google Code for Cart - WSM Conversion Page -->
<script type="text/javascript">
/* <![CDATA[ */
var google_conversion_id = 947736516;
var google_conversion_language = "en";
var google_conversion_format = "3";
var google_conversion_color = "ffffff";
var google_conversion_label = "ish2CN2Uo10QxJ_1wwM";
var google_remarketing_only = false;
/* ]]> */
</script>
<script type="text/javascript" src="//www.googleadservices.com/pagead/conversion.js">
</script>
<noscript>
<div style="display:inline;">
<img height="1" width="1" style="border-style:none;" alt="" src="//www.googleadservices.com/pagead/conversion/947736516/?label=ish2CN2Uo10QxJ_1wwM&amp;guid=ON&amp;script=0"/>
</div>
</noscript>

<%
End Sub		' AddItemToCart


' *************************************************************************************

Sub NewItemsCart(iItemID, iItemCount)
    iItemID = Trim(iItemID)     ' Added 9/23/11.
    If dictCart.Exists(iItemID) Then
        dictCart(iItemID) = iItemCount
    Else
        dictCart.Add iItemID, iItemCount
    End If  
End Sub		' NewItemsCart


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
End Sub		' RemoveItemFromCart


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
		</tr>
        
		<% 
		if bParameters(7) = true then
			' osize = 15.00   ' [BN, 6/22/15] Commented out because was apparently an error. Should have been 10.00; see osize below.
            osize = 10.00
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
			<td COLSPAN="4" ALIGN="Right" bgcolor="white">
				<font face="tahoma" size="2"><b> OverSize Charge:</b></font>
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

		fTotal = SandH(weightTotal, ExchangeRate)
		If Session("Country") = "Canada" Then
			ExtraShippingAmountForCanada = 4.0		' 4/18/07 BN: Sani wanted to charge more for shipping and handling to Canada.
			fTotal = fTotal + ExtraShippingAmountForCanada
		End If 
		If weightTotal = 0 Then fTotal = 0.00   ' [BN] Not worth the trouble to distinguish between U.S. and Canada excgange rate.
		%>

		<tr>
			<td COLSPAN="4" ALIGN="Right" bgcolor="white">
				<font face="tahoma" size="2"><b>Shipping and Handling (within North America):</b></font>
			</td>
			<td ALIGN="Right" bgcolor="#EEEEEE">
				<font face="tahoma" size="2" color="#b9000"><b>$<%=FormatNumber(fTotal,2)%></b></font>
			</td>
		</tr>

		<tr>	
			<td COLSPAN="4" ALIGN="Right" bgcolor="white"><font face="tahoma" size="2"><b>Total:</b></font></td>
			<td ALIGN="Right" bgcolor="#EEEEEE"> 
				<font face="tahoma" size="2" color="#b9000">
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
				<td COLSPAN="5" ALIGN="center" bgcolor="#DDDDDD">
					<font face="tahoma" size="2" color="#b9000"> <b>
					Shipping and Handling may need to be adjusted.<br>We will notify you by email.
					</b></font>
				</td>
			</tr>
		<% End If %>

		<tr>
			<td colspan="6" ALIGN="Center" bgcolor="#DDDDDD">
				<font face="tahoma" size="2"><a href="https://www.starlite-intl.com"><b>Continue Shopping?</b></a></font>
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
				<center>
					<font face="tahoma" size="2">
					<div style="text-align: center;">
					<span style="font-family: helvetica,arial,sans-serif;">All information provided is secured through SSL (Secure Socket Layers) and is kept confidential.</span><br>
					<span style="font-family: helvetica,arial,sans-serif;">Tax will be added according to user input and added to the bottom line automatically.</span><br>
					<span style="font-family: helvetica,arial,sans-serif;">Terms and Conditions of Sale apply.</span>
					</div>
					<br><big><span style="font-family: helvetica,arial,sans-serif;">Order online or by phone at: 1-800-387-8535<br><br><a href="mailto:sales@starlite-intl.com">Contact us</a> for volume discounts.</span></big>
					</font>
					<br><br>
                    <!-- (c) 2005, 2012. Authorize.Net is a registered trademark of CyberSource Corporation --> <div class="AuthorizeNetSeal"> <script type="text/javascript" language="javascript">                                                                                                                                                   var ANS_customer_id = "30dd88ea-13d1-4a1f-bafd-84fd68507946";</script> <script type="text/javascript" language="javascript" src="//verify.authorize.net/anetseal/seal.js" ></script> <a href="http://www.authorize.net/" id="AuthorizeNetText" target="_blank">Transaction Processing</a> </div>   
					<script type="text/javascript" src="https://seal.XRamp.com/seal.asp?type=H"></script>
					<a title="Star Lite International, LLC BBB Business Review" href="https://www.bbb.org/eastern-michigan/business-reviews/electronic-equipment-and-supplies-wholesale-and-manufacturers/star-lite-international-llc-in-southfield-mi-45003227/#bbbonlineclick">
						<img alt="Star Lite International, LLC BBB Business Review" style="border: 0;" src="https://seal-easternmichigan.bbb.org/seals/blue-seal-96-50-star-lite-international-llc-45003227.png" />
					</a>
				</center>
			</td>
		</tr>
	</table>

	<br>

<%
End Sub    ' ShowItemsInCart
%>



<%
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
						<% end if%>
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
				<% end if%>
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
					<% if sar = "Manufa" then %>
						&nbsp;&nbsp;&nbsp;&nbsp;<% = Request("Manufa") %>
					<% else%>
						&nbsp;&nbsp;&nbsp;&nbsp;<%=RSS("Subname")%>
					<% end if%>
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
	If sar <> "Manufa" Then
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
		If sid = 0 Then '************************ELIMINATE SUB FOLLOW LINKS***************** 
		' i.e. We are on an Area (Category) page, not a SubArea (SubCategory) page. %>
	
		<% If False Then    ' BN: Old ugly method, that just lists subcategories, each with a hyperlink (no drop-down menu). %>
			<% do while not rDs.eof %>
			<% subname = Replace( RDS("Subname"), " ", "%20") %>
			<%  ar = Replace( ar, " ", "%20") %>
				<a href="scart.asp?sar=<%=subname%>&amp;area=<%=ar%>&amp;sid=<%=RDS("SID")%>">
				<font face="tahoma" size="2"><b><%=RDS("SubName")%></b></font>
				</a>
				<font face="tahoma" size="2" color="#b70000">&nbsp;&nbsp;</font>
			<% rDs.movenext
				  loop
				  rDs.close  
				  conn.close
			%> 
		<% End If ' False %>


		<% ' 3/2/06: BN added this drop-down menu, to replace earlier hyperlinked text. %>
		<FORM action="https://www.starlite-intl.com/scart/scart.asp" method="GET" name="PID">
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
				<option>Please Select</option>
				<%  	    
				rDs.movefirst()
				do while not rDs.eof 
				%>

				<option value='<%=RDS("SID")%>'><%=RDS("SubName")%></option>
				<% 
				rDs.movenext
				Loop
				rDs.close  
				conn2.close
				%> 
			</select>
	
			<input type='hidden' name='sar' value='<%=subname%>'> <% ' BN: This is ignored. SAR should be retrieved according to the SID. %>
			<input type='hidden' name='area' value='<%=ar%>'>
			<input type="submit" value="Submit">
			</center>
		</FORM>
	

		<% 
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("ConnectionString")
		dim ssdsqstring2
		ssdsqstring2 = "Select AreaDesc from Area51 WHERE  AreaName Like '" & ar & "'"
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
		End If 
		%>
		<% RSS.Close %>
			

		<% End If %>
	<% End If %>


	<% if (SID = 0)  AND ((ar <> "Search") or (sar = "New!")) Then %>
		<center>
			<div align="center">
			<!--#include file="SPECIALK.INC"-->
			<br><br>
			</div>
		</center>
	<% End If %>


	<% 
	If FALSE Then	' For debugging.
		Response.Write "<br>sid = "					& sid 
		Response.Write "<br>iItemCount = "			& iItemCount 
		Response.Write "<br>sar = "					& sar 
		Response.Write "<br>NewProductsSubgroup = "	& NewProductsSubgroup
		Response.Write "<br>SpecialsSubgroup = "	& SpecialsSubgroup 
		Response.Write "<br>RebatesSubgroup = "		& RebatesSubgroup 
	End If
	%>

      
    <% ' Do Heading
	'If iItemCount = 0 Then
	'	Response.Write "<br><br><br><br><center>"
	'	Response.Write "<font face='Tahoma' size='4' color='navy'>No items found.</font>"  ' This gets underlined. Don't know why!
	'	Response.Write "</center>"
	'ElseIf sid <> "0" AND iItemCount > 0 Then   ' i.e. We are on a SubCategory page and there are (more than 0) products.
	If sid <> "0" Then   ' i.e. We are on a SubCategory page and there are (more than 0) products.
		'Response.Write "<br>sar = " & sar
		'Response.Write "<br>iItemCount = " & iItemCount 

		' NewProductsSubgroup, SpecialsSubgroup and RebatesSubgroup actually holds the index value for the subgroup.
		If sar = "New Products" Then
			If NewProductsSubgroup = 1000 Then
				Heading = "All New Arrivals"
			Else
				TagsSQL = "Select * from Tags WHERE GroupIndex = 1 AND SubgroupIndex = " & NewProductsSubgroup   
				'Response.Write "<br>TagsSQL = " & TagsSQL
				Set rsTags = Conn.Execute(TagsSQL)
				Heading = rsTags("SubgroupDescr")
			End If

		ElseIf sar = "Specials" Then
			If SpecialsSubgroup = 1000 Then
				Heading = "All Specials"
			Else
				TagsSQL = "Select * from Tags WHERE GroupIndex = 2 AND SubgroupIndex = " & SpecialsSubgroup   
				Set rsTags = Conn.Execute(TagsSQL)
				Heading = rsTags("SubgroupDescr")
			End If
		ElseIf sar = "Rebated" Then
			If RebatesSubgroup = 1000 Then
				Heading = "All Products with Rebates"
			Else
				TagsSQL = "Select * from Tags WHERE GroupIndex = 3 AND SubgroupIndex = " & RebatesSubgroup   
				Set rsTags = Conn.Execute(TagsSQL)
				Heading = rsTags("SubgroupDescr")
			End If
		End If

			Response.Write "<font face='Tahoma' size='4' color='navy'><center>"
			Response.Write "<br>" & Heading
			Response.Write "<div style='margin:10pt'>"
			If iItemCount > 0 Then
				If sar = "Rebated" Then 
					Response.Write	"<div style='font-size:11pt'>(Rebates are redeemable only by residents of the U.S. and/or Canada and only for products purchased from this company.)</div><br>"
				End If
				Response.Write		"<font size='2'>" & iItemCount & " Products. <b>Click on a graphic</b> to see a detailed description of the product and its rebate, and a larger image for the product</font>"
				Response.Write "</div>"
			Else
				Response.Write "<br>No Items Found"
			End IF
		Response.Write "</center></font>"

	End If 
%>


<%	' Do Table.
	Response.Write "<table Border='0' CellPadding='3' CellSpacing='10' width='100%'>"

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


			<td valign='top'>
				<font size=2 face='Arial'><%=I%></font>
			</td>

			<td valign="top" width="100">
				<font face="Tahoma" size="1"><b><i><%= aParameters(1) %></i></b></font><br>
				<a href="../Detail.asp?pid=<%=xxxx%>&amp;Key=">
				<img SRC="<%= aParameters(0) %>" alt="<%= aParameters(1) %>" border="0" width="50" align="left"></a>
				<% 
				NewProductsSubgroup = RS("NewProductsSubgroup")
				If NewProductsSubgroup > 0 Then
					NewIcon = "https://www.starlite-intl.com/imi/new1.gif"
					Response.Write "<img src='" & NewIcon & "' style='border: 0px solid ;' >"
				End If
				RebatesSubgroup = RS("RebatesSubgroup")
				If RebatesSubgroup > 0 Then
					NewIcon = "https://www.starlite-intl.com/imi/Rebate.png"
					Response.Write "<img src='" & NewIcon & "' style='border: 0px solid ;' >"
				End If
				%>
			</td>


				<% area1 = Replace( ar, " ", "%20") %>
				<% sarea1 = Replace( sar, " ", "%20") %>

				<!-- 5/5/05, BN: Original of case 1: 		<TD valign="top"><A HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=aParameters(6)%>&count=1&amp;sid=<%=aParameters(7)%>&amp;Area=<%=area1%>&amp;sar=<%=sarea1%>"><img src="https://www.starlite-intl.com/images/order.gif" border="0" height="15px" ></A><br>         '' 5/5/05, BN: Original of case 2:  		<TD valign="top"><A HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=aParameters(6)%>&count=1&amp;sid=<%=sid%>&amp;Area=<%=area1%>&amp;sar=<%=sarea1%>"><img src="https://www.starlite-intl.com/images/order.gif" border="0"></A><br>    -->
				<% If sar = "Manufa" OR sar = "New!" Then %>
						<td valign="top"><a HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&amp;item=<%=aParameters(6)%>&amp;count=1&amp;sid=<%=aParameters(7)%>&amp;Area=&amp;sar="><img src="https://www.starlite-intl.com/images/order.gif" border="0" ></a><br>
				<% Else %>
 						<td valign="top"><a HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&amp;item=<%=aParameters(6)%>&amp;count=1&amp;sid=<%=sid%>&amp;Area=&amp;sar="><img src="https://www.starlite-intl.com/images/order.gif" border="0" ></a><br>
				<% End If %>
						  <font face="Tahoma" size="1"><b><u>ID # <%=aParameters(6)%></b></u></font><br>
						  <font face="Tahoma" size="1">Reg. Price </font> <font face="Tahoma" size="2"><%= aParameters(3) %></font><br>
							<font face="Tahoma" size="2"><b>Our Price </b></font> 
							<font face="Tahoma" size="2" color="#B90000"><b><i>
				<% ' [BN, 3/1/04]
					'= aParameters(4)
					If aParameters(8) <> TRUE Then %>
					   <%=aParameters(4) %>
					   <% Else %>
					   <font face="Tahoma" size="1">
					   <b><i>Click Order Button to order or see our price.</i></b></font>
					<% End If %>
					</i></b></font>
                          
           	</td>

            <td valign="top">
				<font face="Tahoma" size="2"><%=aParameters(2)%> </font>
			</td>

			<% If sar = "Rebated" Then   ' Add an extra column to describe the rebate for this product. %>
			<td valign="top" width='30%'>
				<font face="Tahoma" size="2"><% =RS("RebateDescr") %></font>
			</td>
			<% End If    ' sar = "Rebated" %>
   		</tr>

	<%
	RS.MoveNext
    Next	' For I = 1 to iItemCount

    Response.Write "</table>"

End Sub   ' ShowFullCatalog()
%>



<%
' *************************************************************************************

Sub PlaceOrder()
	Dim Key
	Dim aParameters ' as Variant (Array)
	Dim sTotal, sShipping

    %>
    <table Border="0" CellPadding="3" CellSpacing="2">
	<tr>
		<td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>ID #</b></font></td>
		<td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Description</b></font></td>
		<td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Qty.</b></font></td>
		<td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Price</b></font></td>
		<td bgcolor="#BBBBBB"><font face="tahoma" size="2"><b>Totals</b></font></td>
	</tr>
    
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
        
        <% If aParameters(7) = TRUE Then osize = 10.00 End If %> 
        
        <%
        sTotal = sTotal + (dictCart(Key) * CSng(aParameters(4)))
        
        ' [BN, 4/29/04]: Added ...
        weightTotal = weightTotal + (dictCart(Key) * CSng(aParameters(8)))
		ExchangeRate = aParameters(9)    ' This is actually the same for each product in shopping cart, but what the heck.

		RXS.Close
    Next
    %>
    
	<tr>
		<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD">
			<font face="tahoma" size="2"><b> Sub Total:</b></font>
		</td>
		<td ALIGN="Right" bgcolor="#DDDDDD">
			<font face="tahoma" size="2" color="#b9000"><b>$<%= FormatNumber(sTotal,2) %></b></font>
		</td>
	</tr>


	<tr>
		<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2"><b> OverSize Charge:</b></font></td>
		<td ALIGN="Right" bgcolor="#DDDDDD">
			<% osizeTotal = osize * ExchangeRate %>
			<font face="tahoma" size="2" color="#b9000"><b>$<%= FormatNumber(osizeTotal , 2) %></b></font>
		</td>
	</tr>


	<% 
	iTotal = sTotal
	sTotalOld = sTotal
	sTotal = sTotal + osize 
 
	fTotal = SandH(weightTotal, ExchangeRate) 
	If Session("Country") = "Canada" Then
		ExtraShippingAmountForCanada = 4.0		' 4/18/07 BN: Sani wanted to charge more for shipping and handling to Canada.
		fTotal = fTotal + ExtraShippingAmountForCanada
	End If 
	If weightTotal = 0 Then fTotal = 0.00   ' [BN] Not worth the trouble to distinguish between U.S. and Canada exchange rate.
	%>

	<tr>
		<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD">
			<b>Shipping and Handling (within North America):</b>
		</td>
		<td ALIGN="Right" bgcolor="#DDDDDD">
			<font face="tahoma" size="2" color="#b9000"><b>$<%=FormatNumber(fTotal,2)%></b></font>
		</td>
	</tr>


	<tr>	<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><b>Total:</b></td>
		<td ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2" color="#b9000"><b>$
			<% 
			If sTotal = 0 Then
				gTotal = 0
			Else
				' gTotal = sTotal + (((iTotal+osize) * 0.0375)+7.95)
				gTotal = sTotalOld + fTotal + osizeTotal
			End If 
			%>
			<%=FormatNumber(gTotal,2)%>
			</b>
			</font>
		</td>
	</tr>


	<% If fTotal = MaxSandH  Then %>
		<tr>
			<td COLSPAN="5" ALIGN="center" bgcolor="#DDDDDD">
				<font face="tahoma" size="2" color="#b9000"><b>
				Shipping and Handling may need to be adjusted.<br>We will notify you by email.
				</b></font>
			</td>
		</tr>
	<% End If %>


		<tr>
			<td colspan="5" ALIGN="center" bgcolor="#DDDDDD">
				<font face="tahoma" size="2"><a href="scart.asp?area=Accessories&amp;sid=0"><b>Need Accessories?</b></a></font>
			</td>
		</tr> 
		   
		<tr>
			<td colspan="5"><br>
				<b>Please choose from the menus above to continue to shop.</b><br><br>
			</td>
		</tr>
    </table>

<%
End Sub		' PlaceOrder
%>




<%
' *************************************************************************************

' We implemented this this way so if you attach it to a database you'd only need one call per item
Function asGetItemParameters(iItemID)
    Dim bParameters
    iItemID = Trim(iItemID)     ' Added 9/23/11.

    If Session("Country") = "USA" Then
       	    csql = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID +   "' "         
       		    'Response.Write csql & "<br>"     
     		    RXS.Open csql, "DSN=STAREC1" , 1, 4
     	        ' Response.Write "RXS(ITEMID) = " & RXS("ITEMID") & "<br>"
		        ' 6/18/06, commented out, BN: bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"), RXS("ITEMID"),RXS("OverSize"), RXS("Weight"), 1 )
		        bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"), Trim(RXS("ITEMID")),RXS("OverSize"), RXS("Weight"), 1 )
    Else
                csql = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID +   "' "
           	    RXS.Open csql, "DSN=STAREC1" , 1, 4
         	    bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Exch") ), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")*(1/(1-(RXS("GPM"))))),RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), RXS("Exch") )
               ' 6/18/06, commented out, BN: bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Freight")* RXS("Exch") ), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")*(1/(1-(RXS("GPM"))))),RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), RXS("Exch") )
               'bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"), RXS("ITEMID"), RXS("OverSize") )
    End If
    ' Return array containing product info.
    asGetItemParameters = bParameters
End Function	' asGetItemParameters

' *************************************************************************************

Function GetItemParameters(iItemID)
    Dim aParameters 
    ' [BN, 3/3/04] Appended ShowPrice field.
    iItemID = Trim(iItemID)     ' Added 9/23/11.
    If Session("Country") = "USA" Then       
		' 6/18/06, commented out, BN: aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("PName") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Freight")), formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM"))))), RS("PID"),RS("ITEMID"),RS("SID"), RS("ShowPrice"))
		aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("PName") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")), formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM"))))), RS("PID"),RS("ITEMID"),RS("SID"), RS("ShowPrice"))
    Else
		' 6/18/06, commented out, BN: aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("Pname") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Duty")*RS("Freight")*RS("Exch")), formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM"))))),RS("PID"),RS("ITEMID"),RS("SID"),  RS("ShowPrice"))
		aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("Pname") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Duty")*RS("Exch")), formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM"))))),RS("PID"),RS("ITEMID"),RS("SID"),  RS("ShowPrice"))
    End If     
    ' Return array containing product info.
    GetItemParameters = aParameters
End Function	' GetItemParameters

' *************************************************************************************
%>