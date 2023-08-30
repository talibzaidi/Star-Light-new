<%
'set RS = CreateObject("ADODB.Recordset")
'RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE  PID = " + PID  +   " ", "DSN=STAREC1" , 1, 4
'on error resume next
'do while not rs.eof
%>

<% 
NewProduct	= RS("NewProduct") 
Deleted		= RS("Deleted")
'Response.Write "<br>Deleted = " & Deleted
%>

<table border="0" cellspacing="0" width="100%">

    <tr>
        <td valign="top" align='center'>
		<img src="./Imi/<%=RS("Pic2")%>" border="0">
		</td>
	</tr>
	
	<tr>
		<td align='center' valign='middle'>
			<% If (RS("Manufa") <> "General Information") AND (NOT Deleted) Then   ' Turn off the ORDER button if Manufacturer = General Information e.g. "What is WAAS?" %>
				<a href="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=RS("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> 
				<img src="Images/order.gif" border="0" hspace='10'></a><font size="4">
			<% End If %>

			<%
			If CBool(Deleted) Then
				Response.Write "<font color='#B90000' size='3'>No Longer Available</font>&nbsp;&nbsp;&nbsp;" 
			End If
			%>

			<b><%=RS("PName")%></b>&nbsp;(<%=RS("Manufa")%>)</font>
			<%
			If NewProduct Then
				NewIcon = "https://www.starlite-intl.com/imi/new1.gif"
				Response.Write "<img src='" & NewIcon & "' style='border: 0px solid ;' hspace='5'>"
			End If
			%>
		<td>
	</tr>
	
	<tr>	
		<td align='center'>
		<b>ID # <%=RS("ITEMID")%></b>
				
		<% 
		If RS("Manufa") <> "General Information" Then   ' Turn off the ORDER button if Manufacturer = "General Information" e.g. the "What is WAAS?" "product" %>
		
			<%	
			If Not Deleted Then	
				If Session("Country") = "USA" then    
					If RS("Manufa") <> "RepairMaster" Then
						Response.Write "&nbsp;&nbsp;&nbsp;Reg. Price: "
						' =formatcurrency(RS("MSL")*RS("Freight")) ' Old version 
						Response.Write formatcurrency(RS("MSL"))
						End If
					%>
					&nbsp;&nbsp;&nbsp;
					<b>Our Price: </b>
					<Font color="#B90000">
					<b><i>
					<% If RS("ShowPrice") = true Then '  "ShowPrice" really means "Don't Show Price" %>
						<b>
						Click ORDER to order or see our price.</b>
						<% Else %> 
							<%=formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM")))))%></i>
						<% End If %>
					</b>
					</font>
		 
				<% Else       
					If RS("Manufa") <> "RepairMaster" Then
						Response.Write "&nbsp;&nbsp;&nbsp;Reg. Price: "
						' =formatcurrency(((RS("MSL")*RS("Duty"))*RS("Freight"))*RS("Exch"))' Old version 
						Response.Write formatcurrency(((RS("MSL")*RS("Duty")))*RS("Exch"))
					End If
					%>
					&nbsp;&nbsp;&nbsp;
					<b>Our Price </b> 
					<Font color="#B90000">
					<b><i>
					<% 'If RS("ShowPrice")= true Then ' "ShowPrice" really means "Don't Show Price". 5/21/07, BN: It's ok to always show price for Canada case. %>
						<% ' Else %> 
							<%=formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM")))))%></i>
						<% ' End If %>
					</b>
					</font>
		
			<% 
				End If		' Session("Country") = "USA"		
			End If			' Not Deleted
		End If				' RS("Manufa") <> "General Information" 
		%>

        </td>
    </tr>
    
   
    
    <tr colspan='2'>
        <td valign="top">
		
		<% if rs("Text1") <> "" Then %>
			&nbsp;&nbsp;<%=RS("Text1")%><br>
		<% end if %>
		
		<%
		count = 0
		if rs("Text2") <> "" then
		bullet = RS("Text2")
		Response.Write "<br>bullet = " & bullet
		loobol = len(bullet)%>
		<br>&nbsp;&nbsp;&nbsp;<b>•</b>&nbsp;&nbsp;
		<%
		loobol = loobol

		while loobol > 0 
		    if left(bullet,1) = "~" then%>
		        <br>&nbsp;&nbsp;&nbsp;<b>•</b>&nbsp;&nbsp;
		    <%bullet = right(bullet,(len(bullet)-1))
		    end if
		    bullpos = inStr(bullet,"~")
		    if bullpos = 0 then%>
		        <%=bullet%>
		    <%bullet = ""
		    else
		        bulletdisplay = left(bullet, bullpos-1)%>
		        <%=bulletdisplay%>
		        <% bullet = right(bullet,(len(bullet) - len(bulletdisplay)))
		    end if
		    loobol = len(bulllet)
		wend%>
		<br>
		<% end if %>
		
		<HR>
    	<center>
		Specifications, price and availability are subject to change without notice.<br>
		<% if RS("ManURL") <> "http://" then %>
			<a href="outsider.asp?WarpURL=<%=RS("ManURL")%>&amp;Warp=<%=RS("PID")%>">
			Link to manufacturer's site.</a>
		<% end if %>
		
		Gross Weight: <%=RS("Weight")%>&nbsp; lbs.
		</center>

		
		<!--#include file="Warranties/Warranties.INC"-->
		
		<% 
		MakeTable("Accessories")			' From Detail.asp
		'MakeTable("AccessoryOf")			' From Detail.asp
		ItemID = RS("ITEMID")
		'Response.Write "<br>ItemID = " & ItemID
		MakeTable3(ItemID)					' From Detail.asp
		%>

        </td>
        
        
    </tr>
    
</table>
               

<%	
'rs.movenext
'loop
rs.Close 
%>