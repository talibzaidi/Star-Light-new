
<!--#include file="RATES.INC"-->


<% session("initq") = 1 %>
<%

'on error resume next

sAction = CStr(Request.QueryString("action"))
iItemID = (Request.QueryString("item"))
iItemCount = CInt(Request.QueryString("count"))

if (sar = "New%20Products" OR sar = "New Products" OR sar = "New!") then 


     set RS = CreateObject("ADODB.Recordset")
         RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  NewProduct = yes ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4    

     'ar = "New Products"    

elseif sar ="Specials" then 

    response.redirect "scartstart.asp"

elseif sar ="Manufa" then 

    set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE  Manufa Like '" & ReQuest("Manufa") & "' ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4    
    ar="Search"
    

else
	 'Response.Write "Got to here."
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
    <table Border="0" CellPadding="3" CellSpacing="2" align="left">
        <tr>
            <td bgcolor="#BBBBBB"><font face="tahoma" width="100px" size="2"><b>ID #</b></font></td>
            <td bgcolor="#BBBBBB"><font face="tahoma" width="200px" size="2"><b>Description</b></font></td>
            <td bgcolor="#BBBBBB"><font face="tahoma" width="100px" size="2"><b>Select Quantity</b></font></td>
            <td bgcolor="#BBBBBB"><font face="tahoma" width="100px" size="2"><b>Price</b></font></td>
            <td bgcolor="#BBBBBB"><font face="tahoma" width="300px" size="2"><b>Totals </b></font></td>
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
            

		<td width="165" ALIGN="Left" bgcolor="#EEEEEE">   <!-- [BN] Quantity field. -->
		<font face="tahoma" size="1">
        	<form><input type="text" size="2" name="Qty" value="<%= dictCart(Key) %>">
        	<input type="hidden" name="action" value="qty">
        	<input type="hidden" name="item" value="<%=Key%>">
        	<input type="hidden" name="sid" value="<%=sid%>">
        	<input type="hidden" name="Area" value="<%=Area%>">
        	<input type="hidden" name="Sar" value="<%=Sar%>">
        	<input type="image" src="../Images/button1.gif" name="Qtybut" value="Submit" border="0" WIDTH="50" HEIGHT="14">
        	</form></a></b>
		</td>
            


            <td ALIGN="Right" bgcolor="#EEEEEE"> 		<!-- [BN] Price field -->
		<font face="tahoma" size="2"><b><%= bParameters(4) %></font></b>
		</td>


            <td ALIGN="Right" bgcolor="#EEEEEE">		<!-- [BN] Total Price field. -->
            <% rowPrice = dictCart(Key) * CSng(bParameters(4)) 
               sTotal = sTotal + rowPrice %>
		<font face="tahoma" size="2"><b>$<%=FormatNumber(rowPrice, 2) %>
		</font></b>
		</td>

 <!--           <TD ALIGN="Right" bgcolor="#EEEEEE">		 [BN] Total Weight field.		<font face="tahoma" size="2"><b><%=FormatNumber(dictCart(Key) * CSng(bParameters(8)),2) %> lbs.		</font></b>		</TD> -->

        </tr>
   
        
        <% 
        if bParameters(7) = true then
              osize = 8.75
            end if
        %> 
        
        <%
        ' sTotal = sTotal + (dictCart(Key) * CSng(bParameters(4)))

        ' [BN, 4/29/04]: Added ...
        weightTotal = weightTotal + (dictCart(Key) * CSng(bParameters(8)))
	  ExchangeRate = bParameters(9)    ' This is actually the same for each product in shopping cart, but what the heck.

    RXS.Close 
    Next	' End of For loop for outputing rows for the products in the shopping cart.
        %>


    	<!-- [BN, 3/3/04] Begin output of fixed number of remaining rows of shopping cart display table. -->
    	<tr><td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD">
	<font face="tahoma" size="2"><b> Sub Total:</b></font></td>
	<td ALIGN="Right" bgcolor="#DDDDDD">
	<font face="tahoma" size="2" color="#b9000"><b>$<%= FormatNumber(sTotal,2) %></b></font>
	</td>
	</tr>


	<tr>
	<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2">
	<b> OverSize Charge:</b></font></td><td ALIGN="Right" bgcolor="#DDDDDD">
      <% osizeTotal = osize * ExchangeRate %>
	<font face="tahoma" size="2" color="#b9000"><b>$<%= FormatNumber(osizeTotal , 2) %></b></font></td>
	</tr>


<% iTotal = sTotal%>
<% sTotalOld = sTotal
   sTotal = sTotal + osize
if sTotal = 0 then
    fTotal= FormatNumber(0,2)
else
    fTotal = FormatNumber((((iTotal+osize) * 0.0375)+7.95),2) 
end if
%>

<% 
fTotal = SandH(weightTotal, ExchangeRate) 
If weightTotal = 0 Then fTotal = 0.89   ' [BN] Not worth the trouble to distinguish between U.S. and Canada excgange rate.
%>

	<tr>
	<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2">
		<b>Shipping and Handling:</b></font></td>
    	<td ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2" color="#b9000"><b>
	$<%=FormatNumber(fTotal,2)%>
      </b></font></td>
	</tr>


	<tr>	<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><font face="tahoma" size="2">
			<b>Total:</b></font></td>
		<td ALIGN="Right" bgcolor="#DDDDDD"> <font face="tahoma" size="2" color="#b9000"> <b>

<% 
if sTotal = 0 then
    gTotal = 0
else
    ' gTotal = sTotal + (((iTotal+osize) * 0.0375)+7.95)
    gTotal = sTotalOld + fTotal + osizeTotal 
end if
%>

	$<%=FormatNumber(gTotal,2)%>  
		</b>
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


<!--	<TR>	<TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><B>Weight Total:</B></TD>		<TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>		<%=FormatNumber(weightTotal,2)%>		</b>		</font>		</TD>	</TR>-->

	<tr><td colspan="6" ALIGN="Center" bgcolor="#DDDDDD">
	<font face="tahoma" size="2"><a href="scart.asp?area=Accessories&amp;sid=0"><b>
	Need Accessories?
	</b></a>
	</td>
	</tr>    

	<tr>
	<td colspan="6" align='right'>
    <a HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><img SRC="../images/shop_checkout.gif" BORDER="0" ALT="Checkout" WIDTH="46" HEIGHT="46"></a><br><font face="tahoma" size="1"><b>CHECKOUT</b></font>
	</td>
	</tr>

	<tr>
	<td colspan="6"><center><font face="tahoma" size="2">
	<br><b>Please choose from the menus above to continue shopping.</b><br><br>


	<br>
	<div style="text-align: center;"><span style="font-family: helvetica,arial,sans-serif;">
	All information provided is secured through SSL (Secure Socket Layers) and is kept confidential.</span><br>
	<span style="font-family: helvetica,arial,sans-serif;">Tax will be added according to user input and added to the bottom line automatically.</span><br>
	<span style="font-family: helvetica,arial,sans-serif;">Terms and Conditions of Sale apply.</span>
	</div>
	
	<br><big><span style="font-family: helvetica,arial,sans-serif;">
	Order online or by phone at: 1-800-387-8535</span></big>
	
	<br><br> 
	<script language=JavaScript src='https://seal.XRamp.com/seal.asp?type=G'></script>

	</center>
	</td>
	</tr>


</table>

<br>

<%
End Sub



'92******************************************************************************************************************************************************

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

'*********************************
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
			<td valign='top' width='120'>
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
			<b>Subcategory:</b>
			</font>
			</td>
			
			<td valign='top'>
			<font face="tahoma" size="4">
			<%if sar = "Manufa" then %>
				<% = ReQuest("Manufa") %>
			<%else%>
				<%=RSS("Subname")%>
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
		
<% end if%>


<%'********************************************%>

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
	<form action="http://www.starlite-intl.com/scart/scart.asp" method="GET" name="PID">
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
	<option>Please Select
	<%  	    
		rDs.movefirst()
	    do while not rDs.eof 
	%>

	<option value='<%=RDS("SID")%>'><%=RDS("SubName")%>
	<% rDs.movenext
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
</center>


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
		<br><%=iItemCount%> Products. Click on an image to see a detailed description and larger image for that product</font>
	<% End If 
    For I = 1 to iItemCount
        aParameters = GetItemParameters(I)
        %>
        <tr>

<%
      xxxx = (aParameters(5))
%>


		<td valign='top'><font size=2 face='Arial'><%=I%></font></td>
		<td valign="top" width="100">
		<font face="Tahoma" size="1"><b><i><%= aParameters(1) %></i></b></font><br>
		<a href="../Detail.asp?pid=<%=xxxx%>&amp;Key=">
		<img SRC="<%= aParameters(0) %>" border="0" width="50"></a>
		</td>


<% area1 = Replace( ar, " ", "%20") %>
<% sarea1 = Replace( sar, " ", "%20") %>

<!-- 5/5/05, BN: Original of case 1: 		<TD valign="top"><A HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=aParameters(6)%>&count=1&amp;sid=<%=aParameters(7)%>&amp;Area=<%=area1%>&amp;sar=<%=sarea1%>"><img src="../images/order.gif" border="0"></A><br>         '' 5/5/05, BN: Original of case 2:  		<TD valign="top"><A HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=aParameters(6)%>&count=1&amp;sid=<%=sid%>&amp;Area=<%=area1%>&amp;sar=<%=sarea1%>"><img src="../images/order.gif" border="0"></A><br>    -->
<% if sar = "Manufa" or sar = "New!" then %>
		<td valign="top"><a HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&amp;item=<%=aParameters(6)%>&amp;count=1&amp;sid=<%=aParameters(7)%>&amp;Area=&amp;sar="><img src="../images/order.gif" border="0" WIDTH="63" HEIGHT="21"></a><br>
<%else%>
 		<td valign="top"><a HREF="https://www.starlite-intl.com/scart/scart.asp?action=add&amp;item=<%=aParameters(6)%>&amp;count=1&amp;sid=<%=sid%>&amp;Area=&amp;sar="><img src="../images/order.gif" border="0" WIDTH="63" HEIGHT="21"></a><br>
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

' *************************************************************************************************************88

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
if sTotal = 0 then
    fTotal= FormatNumber(0,2)
else
    fTotal = FormatNumber((((iTotal+osize) * 0.0375)+7.95),2) 
end if
%>

<% 
fTotal = SandH(weightTotal, ExchangeRate) 
If weightTotal = 0 Then fTotal = 0.89   ' [BN] Not worth the trouble to distinguish between U.S. and Canada excgange rate.
%>

	<tr>
	<td COLSPAN="4" ALIGN="Right" bgcolor="#DDDDDD"><b>Shipping and Handling:</b></td>
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



'217 We implemented this this way so if you attach it to a database you'd only need one call per item
'*****************************************************************************************************************************
' 
Function asGetItemParameters(iItemID)

Dim bParameters

if Session("Country") = "USA" then
       	csql = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID +   "' "         
       		'Response.Write csql & "<br>"     
     		RXS.Open csql, "DSN=STAREC1" , 1, 4
     	   ' Response.Write "RXS(ITEMID) = " & RXS("ITEMID") & "<br>"
		bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"), RXS("ITEMID"),RXS("OverSize"), RXS("Weight"), 1 )
else
            csql = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE ITEMID LIKE '" + iItemID +   "' "
           	RXS.Open csql, "DSN=STAREC1" , 1, 4
         	bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Duty")*RXS("Freight")* RXS("Exch") ), formatcurrency(RXS("Cost")*RXS("Duty")*RXS("Freight")*RXS("Exch")*(1/(1-(RXS("GPM"))))),RXS("PID"),RXS("ITEMID"), RXS("OverSize"), RXS("Weight"), RXS("Exch") )
        	'bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")*(1/(1-(RXS("GPM"))))), RXS("PID"), RXS("ITEMID"), RXS("OverSize") )
end if

    
' Return array containing product info.
asGetItemParameters = bParameters

End Function


'***************************************************************************************************************************************************
Function GetItemParameters(iItemID)
Dim aParameters 
' [BN, 3/3/04] Appended ShowPrice field.
if Session("Country") = "USA" then       
 	aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("PName") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Freight")), formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM"))))), RS("PID"),RS("ITEMID"),RS("SID"), RS("ShowPrice"))
else
    	aParameters = Array("../imi/" +RS("Pic1") +"","" +RS("Pname") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Duty")*RS("Freight")*RS("Exch")), formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM"))))),RS("PID"),RS("ITEMID"),RS("SID"),  RS("ShowPrice"))
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
<table BORDER="0" CELLSPACING="0" CELLPADDING="0">
<tr><td>
<% ' Response.Write "sAction = " & sAction & "<br>" %>
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
        </td></tr>
        
        <tr><td ALIGN="right">
<!--        <a HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><img SRC="../images/shop_checkout.gif" BORDER="0" ALT="Checkout" WIDTH="46" HEIGHT="46"></a><br><font face="tahoma" size="1"><b>CHECKOUT</b></font><br><br><br>
-->
        <%
    Case "del"
        RemoveItemFromCart iItemID, iItemCount
        ShowItemsInCart
        %>
        <%
    Case "qty"

        NewItemsCart iItemID, CDbl(Request("Qty"))
        'AddItemToCart iItemID, CDbl(Request("Qty"))
        if CDbl(Request("Qty")) = 0 then
            RemoveItemFromCart iItemID, CDbl(Request("Qty"))
        end if
        ShowItemsInCart
        %>
        </td></tr>
        <tr><td ALIGN="right">
<!--        <a HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><img SRC="../images/shop_checkout.gif" BORDER="0" ALT="Checkout" WIDTH="46" HEIGHT="46"></a><br><font face="tahoma" size="1"><b>CHECKOUT</b></font><br><br><br>
-->
        <%
    Case "viewcart"
        ShowItemsInCart
        %>
        </td></tr>
        <tr><td ALIGN="right">


<!--        <a HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><img SRC="../images/shop_checkout.gif" BORDER="0" ALT="Checkout" WIDTH="46" HEIGHT="46"></a><br><font face="tahoma" size="1"><b>CHECKOUT</b></font><br><br><br>
-->
        <%
    Case "checkout"
        ' PlaceOrder  [BN, 9/9/05]: I commeneted this out. It doesn't seem to belong here. Don't know why it was here.
        %>
        </td></tr>
        <tr><td ALIGN="left">
        <br>
        <span style="font-family: helvetica,arial,sans-serif;">
Please enter the following information to complete your order. </span><br style="font-family: helvetica,arial,sans-serif;">
<span style="font-family: helvetica,arial,sans-serif;"> 
        Your order will be shipped when method of payement is verified or personal check clears. 
        </span><br style="font-family: helvetica,arial,sans-serif;">
<span style="font-family: helvetica,arial,sans-serif;">In the event that there is a problem in processing your order you will be notified via phone 
        or e-mail.</span><br style="font-family: helvetica,arial,sans-serif;">
<span style="font-family: helvetica,arial,sans-serif;"> For further inquiries call: 1-800-387-8535 or contact us 
        </span><a href="mailto:starlite@starlite-intl.com" style="font-family: helvetica,arial,sans-serif;">here</a><span style="font-family: helvetica,arial,sans-serif;">.</span>
                
        <form action="sendmail.asp" method="post">
		<p>&nbsp;</p>
		<div align="center"><center>
		
		<table border="0" width="350">
    
        <tr>
            <td><font size="2" face="Tahoma"><strong>Name:</strong></font></td>
            <td><strong>
            <input type="text" size="25" name="Name"></strong></td>
        </tr>
        
		<tr>
            <td><font size="2" face="Tahoma"><strong>E-Mail Address:</strong></font></td>
            <td>
            <input type="text" size="25" name="Email" value="email@domain.com"></td>
        </tr>
        
         <tr>
            <td valign="top"></td>
            <td valign="top">
            <input type="checkbox" name="OptIntoEmailing">
            <font size="2" face="Tahoma">
            <strong>If you do not want to receive information on specials and new products, check this box.</strong>
            </font>
            </td>
        </tr>
        
		<tr>
            <td><font size="2" face="Tahoma"><strong>Phone Number:</strong></font></td>
            <td><input type="text" size="25" name="Phone" value></td>
        </tr>
        
        <tr>
            <td><font size="2" face="Tahoma"><strong>Street
            Address:</strong></font></td>
            <td>
            <input type="text" size="25" name="Address"></td>
        </tr>
        
        <tr>
            <td><font size="2" face="Tahoma"><strong>City:</strong></font></td>
            <td>
            <input type="text" size="25" name="City"></td>
        </tr>
        
        <tr>
            <td><font size="2" face="Tahoma"><strong>State/Province:</strong></font></td>
            <td><input type="text" size="25" name="StateProv"></td>
        </tr>
        
         <tr>
            <td valign="top"></td>
            <td valign="top">
            <input type="checkbox" name="BigShip">
            <font size="2" face="Tahoma">
            <strong>Check this box if 
            you live in Alaska, Hawaii, Nunavut, or Northwest Territories.</strong><br>
            Your checkout freight will be 
            recalculated at 3 times the above Shipping and Handling amount.
            </font><br>
            </td>
        </tr>
        
   <tr>
            <td><font size="2" face="Tahoma"><strong>Postal Code:</strong></font></td>
            <td><input type="text" size="25" name="Postal"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Country:</strong></font></td>
            <td>
<select name="Country" id="Country">
<option value="SELECT COUNTRY" SELECTED>SELECT COUNTRY</option>
<option value="Afghanistan">Afghanistan</option>
<option value="Albania">Albania</option>
<option value="Algeria">Algeria</option>
<option value="American Somoa">American Samoa</OPTION>
<option value="Andorra">Andorra</option>
<option value="Angola">Angola</option>
<option value="Anguilla">Anguilla</option>
<option value="Antarctica">Antarctica</option>
<option value="Aantigua and Barbuda">Antigua and Barbuda</option>
<option value="Argentina">Argentina</option>
<option value="Armenia">Armenia</option>
<option value="Aruba">Aruba</option>
<option value="Australia">Australia</option>
<option value="Austria">Austria</option>
<option value="Azerbbaijan">Azerbaijan</option>
<option value="Bahamas">Bahamas</option>
<option value="Bahrain">Bahrain</option>
<option value="Bangedesh">Bangladesh</option>
<option value="Barbados">Barbados</option>
<option value="Belarus">Belarus</option>
<option value="Belguim">Belgium</option>
<option value="Belize">Belize</option>
<option value="Benin">Benin</option>
<option value="Bermuda">Bermuda</option>
<option value="Bhutan">Bhutan</option>
<option value="Bolivia">Bolivia</option>
<option value="Bosnia and Herzegovina">Bosnia and Herzegovina</OPTION>
<option value="Botswana">Botswana</option>
<option value="Bouvet Island">Bouvet Island</option>
<option value="Brazil">Brazil</option>
<option value="British Indian Ocean Territory">British Indian Ocean Territory</option>
<option value="Brunei">Brunei</option>
<option value="Bulgaria">Bulgaria</option>
<option value="Burkina Faso">Burkina Faso</option>
<option value="Burundi">Burundic</option>
<option value="Cambodia">Cambodia</option>
<option value="Cameroon">Cameroon</option>
<option value="Canada">Canada</option>
<option value="Cape Verde">Cape Verde</option>
<option value="Cayman Islands">Cayman Islands</OPTION>
<option value="Central African Republic">Central African Republic</option>
<option value="Chad">Chad</option>
<option value="Chile">Chile</option>
<option value="China">China</option>
<option value="Christmas Island">Christmas Island</option>
<option value="Cocos Islands">Cocos Islands</option>
<option value="Columbia">Colombia</option>
<option value="Comoros">Comoros</option>
<option value="Congo">Congo</option>
<option value="Cook Islands">Cook Islands</option>
<option value="Costa Rica">Costa Rica</option>
<option value="Cote d&#146;Iviore">Côte d'Ivoire</option>
<option value="Croatia">Croatia</option>
<option value="Cuba">Cuba</option>
<option value="Cyrus">Cyprus</option>
<option value="Czech Republic">Czech Republic</option>
<option value="Congo (DRC)">Congo (DRC) </option>
<option value="Denmark">Denmark</option>
<option value="Djibouti">Djibouti</option>
<option value="Dominica">Dominica</option>
<option value="Dominion Republic">Dominican Republic</option>
<option value="East Timor">East Timor</option>
<option value="Ecuador">Ecuador</option>
<option value="Eqypt">Egypt</option>
<option value="El Salvador">El Salvador</option>
<option value="Equatorial Guinea">Equatorial Guinea</option>
<option value="Erutrea">Eritrea</option>
<option value="Estonia">Estonia</option>
<option value="Ethiopia">Ethiopia</option>
<option value="Falkland Islands">Falkland Islands</option>
<option value="Faroe Islands">Faroe Islands</option>
<option value="Fiji Islands">Fiji Islands</option>
<option value="Finland">Finland</option>
<option value="France">France</option>
<option value="French Guiana">French Guiana</option>
<option value="French Polynesia">French Polynesia</option>
<option value=" French Southern and Antarctic Lands ">French Southern and Antarctic Lands</option>
<option value="Gabon">Gabon</option>
<option value="Gambia">Gambia</option>
<option value="Georgia">Georgia</option>
<option value="Germany">Germany</option>
<option value="Ghana">Ghana</option>
<option value="Gibraltar">Gibraltar</option>
<option value="Greece">Greece</option>
<option value="Greenland">Greenland</option>
<option value="Grenada">Grenada</option>
<option value="Guadeloupe">Guadeloupe</option>
<option value="Gua,">Guam</option>
<option value="Guatemala">Guatemala</option>
<option value="Guinea">Guinea</option>
<option value="GuineaBissau">GuineaBissau</option>
<option value="Guyana">Guyana</option>
<option value="Haiti">Haiti</option>
<option value="Heard Island">Heard Island</option>
<option value="Honduras">Honduras</option>
<option value="Hong Kong">Hong Kong</option>
<option value="Hungary">Hungary</option>
<option value="Iceland">Iceland</option>
<option value="India">India</option>
<option value="Indonesia">Indonesia</option>
<option value="Iran">Iran</option>
<option value="Iraq">Iraq</option>
<option value="Ireland">Ireland</option>
<option value="Israel">Israel</option>
<option value="Italy">Italy</option>
<option value="Jamaica">Jamaica</option>
<option value="Japan">Japan</option>
<option value="Jordan">Jordan</option>
<option value="Kazakhstan">Kazakhstan</option>
<option value="Kenya">Kenya</option>
<option value="Kiribati">Kiribati</option>
<option value="Korea">Korea</option>
<option value="Kuwait">Kuwait</option>
<option value="Kyrgyzstan">Kyrgyzstan</option>
<option value="Laos">Laos</option>
<option value="Latvia">Latvia</option>
<option value="Lebanon">Lebanon</option>
<option value="Lesotho">Lesotho</option>
<option value="Liberia">Liberia</option>
<option value="Libya">Libya</option>
<option value="Liechtenstein">Liechtenstein</option>
<option value="Lthuania">Lithuania</option>
<option value="Luxembourg">Luxembourg</option>
<option value="Macau">Macau</option>
<option value="Macedonia">Macedonia</option>
<option value="Madagascar">Madagascar</option>
<option value="Malawi">Malawi</option>
<option value="Malaysia">Malaysia</option>
<option value="Maldives">Maldives</option>
<option value="Mali">Mali</option>
<option value="Malta">Malta</option>
<option value="Marshall Islands">Marshall Islands</option>
<option value="Martinique">Martinique</option>
<option value="Mauritania">Mauritania</option>
<option value="Mauritius">Mauritius</option>
<option value="Mayotte">Mayotte</option>
<option value="Mexico">Mexico</option>
<option value="Micronesia">Micronesia</option>
<option value="Moldova">Moldova</option>
<option value="Monaco">Monaco</option>
<option value="Mongolia">Mongolia</option>
<option value="Montserrat">Montserrat</option>
<option value="Morocco">Morocco</option>
<option value="Mozambique">Mozambique</option>
<option value="Myanmar">Myanmar</option>
<option value="Namibia">Namibia</option>
<option value="Nauru">Nauru</option>
<option value="Nepal">Nepal</option>
<option value="Netherlands">Netherlands</option>
<option value="Netherlands Antilles">Netherlands Antilles</option>
<option value="New Caledonia">New Caledonia</option>
<option value="New Zealand">New Zealand</option>
<option value="Nicaragua">Nicaragua</option>
<option value="Niger">Niger</option>
<option value="Nigeria">Nigeria</option>
<option value="Niue">Niue</option>
<option value="Norfolk Island">Norfolk Island</option>
<option value="North Korea">North Korea</option>
<option value="Northern Mariana Islands">Northern Mariana Islands</option>
<option value="Norway">Norway</option>
<option value="Oman">Oman</option>
<option value="Pakistan">Pakistan</option>
<option value="Palau">Palau</option>
<option value="Panama">Panama</option>
<option value="Papua New Guinea">Papua New Guinea</option>
<option value="Paraguay">Paraguay</option>
<option value="Peru">Peru</option>
<option value="Philippines">Philippines</option>
<option value="Pitcairn Islands">Pitcairn Islands</option>
<option value="Poland">Poland</option>
<option value="Portugual">Portugal</option>
<option value="Puerto Rico">Puerto Rico</option>
<option value="Qatar">Qatar</option>
<option value="Reunion">Reunion</option>
<option value="Romania">Romania</option>
<option value="Russia">Russia</option>
<option value="Rwanda">Rwanda</option>
<option value="St. Kits and Nevis">St. Kitts and Nevis</option>
<option value="St. Lucia">St. Lucia</option>
<option value="Samoa">Samoa</option>
<option value="San Marino">San Marino</option>
<option value="Saudi Arabia">Saudi Arabia</option>
<option value="Senegal">Senegal</option>
<option value="Seychelles">Seychelles</option>
<option value="Sierra Leone">Sierra Leone</option>
<option value="Sinapore">Singapore</option>
<option value="Slovakia">Slovakia</option>
<option value="Slovenia">Slovenia</option>
<option value="Solomon Islands">Solomon Islands</option>
<option value="Somalia">Somalia</option>
<option value="South Africa">South Africa</option>
<option value="South Georgia">South Georgia </option>
<option value="Spain">Spain</option>
<option value="Sri Lanka">Sri Lanka</option>
<option value="St. Helena">St. Helena</option>
<option value="Sudan">Sudan</option>
<option value="Suriname">Suriname</option>
<option value="Swaziland">Swaziland</option>
<option value="Sweden">Sweden</option>
<option value="Switzerland">Switzerland</option>
<option value="Syria">Syria</option>
<option value="Taiwan">Taiwan</option>
<option value="Tajikistan">Tajikistan</option>
<option value="Tanzania">Tanzania</option>
<option value="Thailand">Thailand</option>
<option value="Togo">Togo</option>
<option value="Tokelau">Tokelau</option>
<option value="Tonga">Tonga</option>
<option value="Trinidad and Tobago">Trinidad and Tobago</option>
<option value="Tunisia">Tunisia</option>
<option value="Turkey">Turkey</option>
<option value="Turkmenistan">Turkmenistan</option>
<option value="Tuvalu">Tuvalu</option>
<option value="Uganda">Uganda</option>
<option value="Ukraine">Ukraine</option>
<option value="United Kingdom">United Kingdom</option>
<option value="United States">United States</option>
<option value="Uruguay">Uruguay</option>
<option value="Uzbekistan">Uzbekistan</option>
<option value="Vanuartu">Vanuatu</option>
<option value="Vatican City">Vatican City</option>
<option value="Venezuela">Venezuela</option>
<option value="Viet Name">Viet Nam</option>
<option value="Virgin Islands">Virgin Islands</option>
<option value="Yemen">Yemen</option>
<option value="Yugoslavia">Yugoslavia</option>
<option value="Zambia">Zambia</option>
<option value="Zimbabwe">Zimbabwe</option>
</select></td>
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
            <td>
            <select name="Payment" size="1">
                <option selected value>Choose an Option</option>
                <option value="Visa">Visa</option>
                <option value="Master Card">Master Card</option>
                <option value="Discover">Discover</option>
				<option value="American Express">American Express</option>
                <option value="Cheque">Personal Cheque</option>
                <option value="Money Order">Money Order</option>
                <option value="Pay Pal">Pay Pal</option>
            </select></td>
        </tr>
       
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>    
    <tr>
            <td colspan="2"><font size="2" face="Tahoma"><strong> Taxes will be applied to total order plus freight.</strong></font>&nbsp;</td>
            
        </tr>
        
        <tr>
            <td><font size="2" face="Tahoma"><strong>State / Provincial Tax:</strong></font></td>
            <td><input type="text" size="3" name="Taxx1" value="0">&nbsp;%</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Federal Tax:</strong></font></td>
            <td><input type="text" size="3" name="Taxx2" value="0">&nbsp;%</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        <tr>
            <td align="center" colspan="2"><font size="2" face="Tahoma"><strong><input type="submit" name="B3" value="Place My Order"></strong></font></td>
        </tr>
		
    </table>
    </center></div>



        </form>

        <%
    Case Else ' Shop
        ShowFullCatalog
        %>
        </td></tr>
        <tr><td ALIGN="right">
        
        <%
End Select

' Return cart to Session for storage
Set Session("cart") = dictCart
%>
</td></tr>
</table>
</center>
