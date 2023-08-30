
<%
' lines 387 to 416 ...

		Response.Write "<tr bgcolor='" & color & "'>"
		'Response.Write "<td valign='top'><font size='1'>" & row & "</font></td>" 
		Response.Write "<td valign='middle' align='left' width='200'>" 
			Response.Write "<a href='https://starlite-intl.com/mobile1/detail.asp?pid=" & PID & "'>"
			Response.Write "<img hspace='20' align='left' border='0' src='" & graphicFile & "'>"
			If NewProductsSubgroup Then
				NewIcon = "https://www.starlite-intl.com/imi/new1.gif"
				Response.Write "<img src='" & NewIcon & "' valign=left style='border: 0px solid ;' >"
			End If
			If RebatesSubgroup Then
				NewIcon = "https://www.starlite-intl.com/imi/Rebate.png"
				Response.Write "<img src='" & NewIcon & "' valign=left style='border: 0px solid ;' >"
			End If		
		Response.Write "</a></td>" 
		'Response.Write "<td valign='top'><b><font color='indigo'>" & ProductName & "</font></b></a><br><br><font size=1>" & Manufacturer & " " & ItemID & "</font></td>" 
		
		Response.Write "<td valign='top'>"
		Response.Write "<b><font color='indigo'>" & ProductName & "</font></b></a><br><font size=1>ID# " & ItemID & "</font>"

		If (Manufacturer <> "General Information") AND (NOT Deleted) Then   ' Turn off the ORDER button if Manufacturer = General Information e.g. "What is WAAS?"
		%>
		<br /><br />
		<center>
		<a href="https://www.starlite-intl.com/mobile1/scart/scart.asp?action=add&item=<%=ItemID%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> 
				<img src="http://www.starlite-intl.com/Images/order.gif" border="0" hspace='10'></a>
		</center>
		<%
		End If

		Response.Write "</td>" 
		
		Response.Write "</tr>"


' Lines 417 to 516 ...

		
		Response.Write "<tr bgcolor='" & color & "'>"
		Response.Write "<td valign='top' colspan='2'>" & Description 
		If TRUE Then
			'Response.Write "<br><br>PID = " & PID

			Set RS = CreateObject("ADODB.Recordset")
			'RS.Open "SELECT *, Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  PID = " & PID  &   " ", "DSN=STAREC1" , 1, 4
	
            ' 11/18/15: I replaced DSN method of line above with the folllowing Connection String method. 	
			RS_SQL   = "SELECT *, Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  PID = " & PID  &   " "
            Set Conn = Server.CreateObject("ADODB.Connection") 
	        Conn.Open Session("ConnectionString")
	        Set RS 	 = Conn.Execute(RS_SQL)	
	        Set Conn = Nothing

            'Response.Write "<br><br>ITEMID = " & RS("ITEMID")
			'Response.Write "<br>Exch = " & RS("Exch")
			'Response.Write "<br>Freight = " & RS("Freight")


' [BN, 2/18/18] Added this block of variables ...
USARegPrice = RS("MSL")
USAOurPrice = RS("Cost")*RS("Freight")*(1/(1-(RS("GPM"))))
USAPercentagePriceDiff = (abs(USARegPrice - USAOurPrice) / USARegPrice) * 100
'Response.Write "<br>USAPercentagePriceDiff = " & USAPercentagePriceDiff

CanadaRegPrice = RS("MSL")*RS("Duty")*RS("Exch")									' = USARegPrice*RS("Duty")*RS("Exch")
CanadaOurPrice = RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM"))))	' = USAOurPrice*RS("Duty")*RS("Exch")
CanadaPercentagePriceDiff = (abs(CanadaRegPrice - CanadaOurPrice) / CanadaRegPrice) * 100
'Response.Write "<br>CanadaPercentagePriceDiff = " & CanadaPercentagePriceDiff

			'**********************************************************

		If RS("Manufa") <> "General Information" Then   ' Turn off the ORDER button if Manufacturer = "General Information" e.g. the "What is WAAS?" "product" 
		
			If Not Deleted Then	
                Response.Write "<br>"

				If Session("Country") = "USA" Then    
            	    If RS("Manufa") <> "RepairMaster" AND USAPercentagePriceDiff > 3 Then
						Response.Write "<br>Reg. Price: "
						Response.Write formatcurrency(USARegPrice)
                        'Response.Write "<br>" & formatcurrency(RS("MSL"))
				    End If
					%>
					&nbsp;&nbsp;&nbsp;
					<br /><b>Our Price: </b>
					<Font color="#B90000">
					<b><i>
					<% If RS("ShowPrice") = true Then '  "ShowPrice" really means "Don't Show Price" %>
						<b>
						Click ORDER to order or see our price.</b>
						<% Else %> 
							<%
                            Response.Write formatcurrency(USAOurPrice)
                            'Response.Write "<br>" & formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM")))))
                            %>
					   </i>
						<% End If %>
					</b>
					</font>

				<% Else     ' [BN, 12/11/17] i.e if Canada
					If RS("Manufa") <> "RepairMaster" AND CanadaPercentagePriceDiff > 3 Then
						Response.Write "<br>Reg. Price: "
                        Response.Write formatcurrency(CanadaRegPrice)
						'Response.Write "<br>" & formatcurrency(((RS("MSL")*RS("Duty")))*RS("Exch"))
					End If
					%>
					&nbsp;&nbsp;&nbsp;
					<br /><b>Our Price </b> 
					<Font color="#B90000">
					<b><i>
					<% 'If RS("ShowPrice")= true Then ' "ShowPrice" really means "Don't Show Price". 5/21/07, BN: It's ok to always show price for Canada case. %>
						<% ' Else %> 
							<%
                            Response.Write formatcurrency(CanadaOurPrice)
                            'Response.Write "<br>" & formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM")))))
                            %>
					</i></b>
						<% ' End If %>
					</font>
		
			<% 
				End If		' Session("Country") = "USA"		
			End If			' Not Deleted
			%>
				<% 
				' 5/4/07, BN: Determine if there are ANY accessories for this product.
				' Recordset rsThisProduct is essentially that used in Details.asp (the product details page) to display a product's accessories.
				Set rsThisProduct = CreateObject("ADODB.Recordset")
				'AccessoriesSQL = "SELECT HasAccessories, SID FROM Product WHERE  PID = " + CStr(bParameters(5)) +   " "
				AccessoriesSQL = "SELECT HasAccessories, SID FROM Product WHERE  PID = " + CStr(PID) +   " "
				'rsThisProduct.Open AccessoriesSQL, "DSN=STAREC1" , 1, 4

                ' 11/18/15: I replaced DSN method of line above with the folllowing Connection String method. 	
                Set Conn = Server.CreateObject("ADODB.Connection") 
	            Conn.Open Session("ConnectionString")
	            Set rsThisProduct = Conn.Execute(AccessoriesSQL)	
	            Set Conn = Nothing

				' If this product has one or more accessories then ...
				'Response.Write "<br>Trim(rsThisProduct('HasAccessories')) = " & Trim(rsThisProduct("HasAccessories"))
				'If NOT IsNULL(rsThisProduct("HasAccessories")) AND (Trim(rsThisProduct("HasAccessories")) <> "" AND Trim(CStr(rsThisProduct("HasAccessories"))) <> "0") Then
				'If (Trim(CStr(rsThisProduct("HasAccessories"))) <> "0") Then
				strVal = Trim(rsThisProduct("HasAccessories") & "") ' This converts rsThisProduct("HasAccessories") to a string, even if rsThisProduct("HasAccessories") is NULL.
				If (strVal <> "" AND strVal <> "0") Then
				%>
					<br />
					<a href="https://starlite-intl.com/mobile1/detail.asp?pid=<%=PID%>#HasAccessories">
					<font face="tahoma" size="2">See Accessories</font>
					</a>
					<br>
				<%
				End If
				%>

		<%
		End If				' RS("Manufa") <> "General Information" 

		'**********************************************************
		End If
		Response.Write "</td>" 
        'If CBool(Deleted) Then
        '    Response.Write "<td valign='top' align='center' style='vertical-align: middle;' width='120'>" 
        '    Response.Write "<font color='#B90000' size='2'>No Longer Available.</font><br>"
        '    'Response.Write "<font color='indigo' size='1'><a href='https://www.starlite-intl.com/Detail.asp?pid=" & PID & "'>Click for possible alternatives</a></font></td>" 
	    '    Response.Write "<a href=https://www.starlite-intl.com/Detail.asp?pid=" & PID & "><font color='navy' size='1'>Click for possible alternatives and accessories</font></a></td>"
        'Else
        '    Response.Write "<td valign='top'></td>" 
        'End If
		'Response.Write "<td valign='top'>$" & Cost & "</td>" 

		Response.Write "</tr>"

		%>