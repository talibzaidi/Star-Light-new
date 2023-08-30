<%
'set RS = CreateObject("ADODB.Recordset")
'RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE  PID = " + PID  +   " ", "DSN=STAREC1" , 1, 4
'on error resume next
'do while not rs.eof
%>

<% 
NewProductsSubgroup	= RS("NewProductsSubgroup") 
RebatesSubgroup		= RS("RebatesSubgroup") 
Deleted				= RS("Deleted")
'Response.Write "<br>Deleted = " & Deleted

USARegPrice = RS("MSL")
USAOurPrice = RS("Cost")*RS("Freight")*(1/(1-(RS("GPM"))))
USAPercentagePriceDiff = (abs(USARegPrice - USAOurPrice) / USARegPrice) * 100
'Response.Write "<br>USAPercentagePriceDiff = " & USAPercentagePriceDiff

CanadaRegPrice = RS("MSL")*RS("Duty")*RS("Exch")									' = USARegPrice*RS("Duty")*RS("Exch")
CanadaOurPrice = RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM"))))	' = USAOurPrice*RS("Duty")*RS("Exch")
CanadaPercentagePriceDiff = (abs(CanadaRegPrice - CanadaOurPrice) / CanadaRegPrice) * 100
'Response.Write "<br>CanadaPercentagePriceDiff = " & CanadaPercentagePriceDiff
%>


<table border="0" cellspacing="0" width="100%">

    <tr>
		<td align="center">
			<table border="0">
			<tr>
				<td valign="top">
					<img src="/Imi/<%=RS("Pic2")%>" border="0" XXXheight="70" align="middle" style="width:100%;">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	
	<tr>
		<td align='center' valign='middle'>
			<% If TRUE AND (RS("Manufa") <> "General Information") AND (NOT Deleted) Then   ' Turn off the ORDER button if Manufacturer = General Information e.g. "What is WAAS?" %>
				<a href="https://www.starlite-intl.com/mobile1/scart/scart.asp?action=add&item=<%=RS("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> 
				<img src='../Images/order.gif' border="0" hspace='10'></a><font size="4">
			<% End If %>

			<%
			If CBool(Deleted) Then
				Response.Write "<font color='#B90000' size='3'>No Longer Available</font>&nbsp;&nbsp;&nbsp;" 
			End If
			%>

			<b><%=RS("PName")%></b>&nbsp;(<%=RS("Manufa")%>)</font>
			<%
			If TRUE AND NewProductsSubgroup Then
				NewIcon = "https://www.starlite-intl.com/imi/new1.gif"
				Response.Write "&nbsp;&nbsp;<img src='" & NewIcon & "' valign='left' style='border: 0px solid ;' hspace='5'>"
			End If
			If RebatesSubgroup Then
				NewIcon = "https://www.starlite-intl.com/imi/Rebate.png"
				Response.Write "&nbsp;&nbsp;<img src='" & NewIcon & "' valign='left' style='border: 0px solid ;' >"
			End If	
			%>
			<br><br><b>&nbsp;&nbsp;&nbsp;&nbsp;ID# <%=RS("ITEMID")%></b>
		</td>
	</tr>
	
	<tr>	
		<td align='center'>
				
		<% 
		If RS("Manufa") <> "General Information" Then   ' Turn off the ORDER button if Manufacturer = "General Information" e.g. the "What is WAAS?" "product" %>

			<%	
			'Session("Country") = "Canada"
			'Response.Write "<br>Session('Country') = " & Session("Country") & "<br>"

			If TRUE AND Not Deleted Then	
				If Session("Country") = "USA" then    
					If RS("Manufa") <> "RepairMaster" AND USAPercentagePriceDiff > 3 Then
						Response.Write "&nbsp;&nbsp;&nbsp;Reg. Price: "
						Response.Write formatcurrency(USARegPrice)
					End If
					%>
					&nbsp;&nbsp;&nbsp;
					<b>Our Price: </b>
					<font color="#B90000">
					<b><i>
					<% If RS("ShowPrice") = true Then '  "ShowPrice" really means "Don't Show Price" %>
						<b>
						Click ORDER to order or see our price.</b>
					<% Else %> 
						<%=formatcurrency(USAOurPrice)%></i>
					<% End If %>
					</b>
					</font>

			  <% Else       ' [BN, 12/11/17] i.e if Canada
					If RS("Manufa") <> "RepairMaster" AND CanadaPercentagePriceDiff > 3 Then
						Response.Write "&nbsp;&nbsp;&nbsp;Reg. Price: "
						Response.Write formatcurrency(CanadaRegPrice)
					End If
					%>
					&nbsp;&nbsp;&nbsp;
					<b>Our Price: </b>    <% ' [BN, 5/21/07] It's ok to always show Our Price for Canada case. %> 
					<font color="#B90000">
					<b><i><%=formatcurrency(CanadaOurPrice)%></i></b>
					</font>
		
			<% 
				End If		' Session("Country") = "USA"		
			End If			' Not Deleted
		End If				' RS("Manufa") <> "General Information" 
		%>

        </td>
    </tr>

	<% ' [BN, 12/6/17] Added this table row, to show some at least the Descr field for the product.  %>
	<tr>
		<td align="center">
		<% Response.Write "<br>" & RS("Descr") %>
		</td>
	</tr>
    
   
    <tr XXXcolspan='2' >
        <td valign="top" align='center' style='border: 0px solid red;'>
		
		<% 
		If rs("Text1") <> "" Then 
			'Response.Write "&nbsp;&nbsp;" & RS("Text1") & "<br>"
			'Response.Write "<HR><br>"
		End If 
		%>

		<%
		If rs("Text2") <> "" Then
			Response.Write "<table align='center' width='100%' style='border: 0px solid blue;'><tr><td align='left'>"
			Txt2 = RS("Text2")
			lngth = len(Txt2)
			While lngth > 0 
				If left(Txt2,1) = "~" Then 
					Response.Write "<br><br>&nbsp;&nbsp;&nbsp;<b>&#149;</b>&nbsp;&nbsp;"	' Write out a bullet in place of the tilde "~".
					Txt2 = right(Txt2,(len(Txt2)-1))										' Remove the tilde.
				End If

				bullpos = Instr(Txt2,"~")
				If bullpos = 0 Then
					'Response.Write Txt2
					Txt2 = ""
				Else
					Txt2display = left(Txt2, bullpos-1)
					Response.Write Txt2display
					Txt2 = right(Txt2,(len(Txt2) - len(Txt2display)))
				End If
				lngth = len(Txt2)
			Wend
			Response.Write "<br>"
			Response.Write "</td></tr></table><br>"
		End If 
		%>

		
		<!--# include file="Warranties/Warranties.INC"-->
		
		<table align='center' style='border:0px solid black;' ><tr><td>
		<% 
		MakeTable("Accessories")			' From Detail.asp
		'MakeTable("AccessoryOf")			' From Detail.asp
		ItemID = RS("ITEMID")
		'Response.Write "<br>ItemID = " & ItemID
		MakeTable3(ItemID)					' From Detail.asp
		%>
		</td></tr></table>

        </td>
        
        
    </tr>
    
</table>
               

<%	
'rs.movenext
'loop
rs.Close 
%>