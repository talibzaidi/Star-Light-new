
<HR>
	
<%

SubAreaID = RS("SID")
'Response.Write "<br>SubAreaID = " & SubAreaID
Set rsSubArea = CreateObject("ADODB.Recordset")
rsSubArea.Open "SELECT * FROM SubArea WHERE SID = " & SubAreaID  &   " ", "DSN=STAREC1" , 1, 4
Warranties = rsSubArea("Warranties")
'Response.Write Warranties

OurPrice = RS("Cost")*RS("Freight")*(1/(1-(RS("GPM")))) 
'Response.Write "<br>OurPrice = " & OurPrice

If Warranties <> "" AND NOT IsNull(Warranties) Then

	Response.Write "<table border=0 width='100%'>"
	Response.Write "<tr><td colspan=3><font color=navy size=4><b>Warranty Options ...</b></font></td></tr>"
	
	Warr = Split(Warranties, ",")
	For i = 0 To UBound(Warr)
		W = Trim(Warr(i))	 ' W is actually the code for the appropriate warranty table, eg. "M" or "V" or "EXP".
		'Response.Write "<br>" & W

		Select Case W
		Case "M"
			Set rsM_Warranties = CreateObject("ADODB.Recordset")
			rsM_Warranties.Open "SELECT * FROM WarrantyMobileElectronics", "DSN=STAREC1" , 1, 4	
			
			Do While NOT rsM_Warranties.EOF 
				PriceThreshold = rsM_Warranties("UnderOurProductPrice")
				'Response.Write PriceThreshold & " "
				If OurPrice < PriceThreshold Then 
					Exit Do
				End If
				rsM_Warranties.MoveNext
			Loop		
					
			If NOT rsM_Warranties.EOF Then	
				Response.Write "<tr><td>"
				Response.Write "<font size=1>" & rsM_Warranties("3YrDOP") & "</font> "
				Response.Write "</td><td>"
				Response.Write "3-Year Warranty"
				Response.Write "</td><td>"
				Response.Write "$" & rsM_Warranties("OurWarrantyPrice3")
				Response.Write "</td></tr>"
			
				Response.Write "<tr><td>"
				Response.Write "<font size=1>" & rsM_Warranties("5YrDOP") & "</font> "
				Response.Write "</td><td>"
				Response.Write "5-Year Warranty"
				Response.Write "</td><td>"
				Response.Write "$" & rsM_Warranties("OurWarrantyPrice5")
				Response.Write "</td></tr>"
			End If
			
		Case "V"
			Set rsV_Warranties = CreateObject("ADODB.Recordset")
			rsV_Warranties.Open "SELECT * FROM WarrantyVideoEquipment", "DSN=STAREC1" , 1, 4	
			
			Do While NOT rsV_Warranties.EOF 
				PriceThreshold = rsV_Warranties("UnderOurProductPrice")
				'Response.Write PriceThreshold & " "
				If OurPrice < PriceThreshold Then 
					Exit Do
				End If
				rsV_Warranties.MoveNext
			Loop		
					
			If NOT rsV_Warranties.EOF Then	
				Response.Write "<tr><td>"
				Response.Write "<font size=1>" & rsV_Warranties("3YrDOP") & "</font> "
				Response.Write "</td><td>"
				Response.Write "3-Year Warranty"
				Response.Write "</td><td>"
				Response.Write "$" & rsV_Warranties("OurWarrantyPrice3")
				Response.Write "</td></tr>"
			
				Response.Write "<tr><td>"
				Response.Write "<font size=1>" & rsV_Warranties("5YrDOP") & "</font> "
				Response.Write "</td><td>"
				Response.Write "5-Year Warranty"
				Response.Write "</td><td>"
				Response.Write "$" & rsV_Warranties("OurWarrantyPrice5")
				Response.Write "</td></tr>"
			End If
			
		Case "EXP"
			Set rsEXP_Warranties = CreateObject("ADODB.Recordset")
			rsEXP_Warranties.Open "SELECT * FROM WarrantyExportCoverage", "DSN=STAREC1" , 1, 4	
			
			Do While NOT rsEXP_Warranties.EOF 
				PriceThreshold = rsEXP_Warranties("UnderOurProductPrice")
				'Response.Write PriceThreshold & " "
				If OurPrice < PriceThreshold Then 
					Exit Do
				End If
				rsEXP_Warranties.MoveNext
			Loop		
					
			If NOT rsEXP_Warranties.EOF Then	
				Response.Write "<tr><td>"
				Response.Write "<font size=1>" & rsEXP_Warranties("1Yr") & "</font> "
				Response.Write "</td><td>"
				Response.Write "1 Year Export Warranty"
				Response.Write "</td><td>"
				Response.Write "$" & rsEXP_Warranties("OurWarrantyPrice1")
				Response.Write "</td></tr>"
			End If
			
		Case "SP"
			Set rsSP_Warranties = CreateObject("ADODB.Recordset")
			rsSP_Warranties.Open "SELECT * FROM WarrantyScreenDamageProtection", "DSN=STAREC1" , 1, 4
			
			Response.Write "<tr><td>"
			Response.Write "<font size=1>" & rsSP_Warranties("2YrDOP") & "</font> "
			Response.Write "</td><td>"
			Response.Write "2-Year Screen Protection Warranty"
			Response.Write "</td><td>"
			Response.Write "$" & rsSP_Warranties("OurWarrantyPrice2")
			Response.Write "</td></tr>"
			
			Response.Write "<tr><td>"
			Response.Write "<font size=1>" & rsSP_Warranties("3YrDOP") & "</font> "
			Response.Write "</td><td>"
			Response.Write "3-Year Screen Protection Warranty"
			Response.Write "</td><td>"
			Response.Write "$" & rsSP_Warranties("OurWarrantyPrice3")
			Response.Write "</td></tr>"
			
			Response.Write "<tr><td>"
			Response.Write "<font size=1>" & rsSP_Warranties("5YrDOP") & "</font> "
			Response.Write "</td><td>"
			Response.Write "5-Year Screen Protection Warranty"
			Response.Write "</td><td>"
			Response.Write "$" & rsSP_Warranties("OurWarrantyPrice5")
			Response.Write "</td></tr>"
			
		Case "X"
			Set rsX_Warranties = CreateObject("ADODB.Recordset")
			rsX_Warranties.Open "SELECT * FROM WarrantyReplacements", "DSN=STAREC1" , 1, 4
	
			Do While NOT rsX_Warranties.EOF 
				PriceThreshold = rsX_Warranties("UnderOurProductPrice")
				'Response.Write PriceThreshold & " "
				If OurPrice < PriceThreshold Then 
					Exit Do
				End If
				rsX_Warranties.MoveNext
			Loop		
							
			If NOT rsX_Warranties.EOF Then
				Response.Write "<tr><td>"
				Response.Write "<font size=1>" & rsX_Warranties("1YrExtension") & "</font> "
				Response.Write "</td><td>"
				Response.Write "1-Year Replacement Warranty"
				Response.Write "</td><td>"
				Response.Write "$" & rsX_Warranties("OurWarrantyPrice1")
				Response.Write "</td></tr>"
			End If

		End Select
	Next

	Response.Write "</table>"

End If
%>
		
<HR>

