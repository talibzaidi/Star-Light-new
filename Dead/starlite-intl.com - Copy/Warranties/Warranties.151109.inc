
<HR>
	
<%
backgroundColor = "palegoldenrod"

SubAreaID = RS("SID")
'Response.Write "<br>SubAreaID = " & SubAreaID
Set rsSubArea = CreateObject("ADODB.Recordset")
rsSubArea.Open "SELECT * FROM SubArea WHERE SID = " & SubAreaID  &   " ", "DSN=STAREC1" , 1, 4

Warranties = rsSubArea("Warranties")
'Response.Write " Warranties = " & Warranties

OurPrice = RS("Cost")*RS("Freight")*(1/(1-(RS("GPM")))) 
'Response.Write "<br>OurPrice = " & OurPrice

If Warranties <> "" AND NOT IsNull(Warranties) Then 

	Response.Write "<table border=0 width='100%'>"
	Response.Write "<tr><td colspan=3 align=left><font color=navy size=4><b><a name='Warranties'>Extended Warranty Options ...</a></b></font></td></tr>"
	
	Warr = Split(Warranties, ",")
	For i = 0 To UBound(Warr)
		W = Trim(Warr(i))	 ' W is actually the code for the appropriate warranty table, eg. "M" or "V" or "EXP".
		'Response.Write "<br>" & W & "<br>"

		Select Case W
		Case "M"
			Set rsM_Warranties = CreateObject("ADODB.Recordset")
			rsM_Warranties.Open "SELECT * FROM WarrantyMobileElectronics", "DSN=STAREC1" , 1, 4			
			
			Do While NOT rsM_Warranties.EOF 
				PriceThreshold = rsM_Warranties("UnderOurProductPrice")
				'Response.Write "<br>PriceThreshold = " & PriceThreshold 
				If OurPrice < PriceThreshold Then 
					Exit Do
				End If
				rsM_Warranties.MoveNext
			Loop		
					
			If NOT rsM_Warranties.EOF Then	
				If backgroundColor = "palegoldenrod" Then
					backgroundColor = "white"
				Else
					backgroundColor = "palegoldenrod"
				End If
				Response.Write "<tr><td>"
				Response.Write "<table align=center border=0 width='100%' bgcolor='" & backgroundColor & "'>"		
				OutputWarrantyRow rsM_Warranties("3YrDOP"), "3-Year Warranty"	
				OutputWarrantyRow rsM_Warranties("5YrDOP"), "5-Year Warranty"
				Response.Write "</table>"
				Response.Write "</td><tr>"
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
				If backgroundColor = "palegoldenrod" Then
					backgroundColor = "white"
				Else
					backgroundColor = "palegoldenrod"
				End If
				Response.Write "<tr><td>"
				Response.Write "<table align=center border=0 width='100%' bgcolor='" & backgroundColor & "'>"	
				OutputWarrantyRow rsV_Warranties("3YrDOP"), "3-Year Warranty"	
				OutputWarrantyRow rsV_Warranties("5YrDOP"), "5-Year Warranty"
				Response.Write "</table>"
				Response.Write "</td><tr>"
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
				If backgroundColor = "palegoldenrod" Then
					backgroundColor = "white"
				Else
					backgroundColor = "palegoldenrod"
				End If
				Response.Write "<tr><td>"
				Response.Write "<table align=center border=0 width='100%' bgcolor='" & backgroundColor & "'>"
				OutputWarrantyRow rsEXP_Warranties("1Yr"), "1 Year Export Warranty"	
				Response.Write "</table>"
				Response.Write "</td><tr>"
			End If
			
		Case "SP"   ' Screen Protection Warranties
			Set rsSP_Warranties = CreateObject("ADODB.Recordset")
			rsSP_Warranties.Open "SELECT * FROM WarrantyScreenDamageProtection", "DSN=STAREC1" , 1, 4
			
			If backgroundColor = "palegoldenrod" Then
				backgroundColor = "white"
			Else
				backgroundColor = "palegoldenrod"
			End If
			Response.Write "<tr><td>"
			Response.Write "<table align=center border=0 width='100%' bgcolor='" & backgroundColor & "'>"
			OutputWarrantyRow rsSP_Warranties("2YrDOP"), "2-Year Screen Protection Warranty"	
			OutputWarrantyRow rsSP_Warranties("3YrDOP"), "3-Year Screen Protection Warranty"				
			OutputWarrantyRow rsSP_Warranties("5YrDOP"), "5-Year Screen Protection Warranty"	
			Response.Write "</table>"
			Response.Write "</td><tr>"
			
		Case "X"   ' Replacement Warranties
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
				If backgroundColor = "palegoldenrod" Then
					backgroundColor = "white"
				Else
					backgroundColor = "palegoldenrod"
				End If
				Response.Write "<tr><td>"
				Response.Write "<table align=center border=0 width='100%' bgcolor='" & backgroundColor & "'>"
				OutputWarrantyRow rsX_Warranties("1YrExtension"), "1-Year Replacement Warranty"
				Response.Write "</table>"
				Response.Write "</td><tr>"
			End If

		Case "LMP"
			Set rsLMP_Warranties = CreateObject("ADODB.Recordset")
			rsLMP_Warranties.Open "SELECT * FROM WarrantyLampReplacement", "DSN=STAREC1" , 1, 4
	
			If backgroundColor = "palegoldenrod" Then
				backgroundColor = "white"
			Else
				backgroundColor = "palegoldenrod"
			End If
			Response.Write "<tr><td>"
			Response.Write "<table align=center border=0 width='100%' bgcolor='" & backgroundColor & "'>"
			OutputWarrantyRow rsLMP_Warranties("3YrDOP"), "3-Year Lamp Replacement Warranty"
			Response.Write "</table>"
			Response.Write "</td><tr>"

		End Select
	Next

	Response.Write "</table>"


End If

%>
		
<HR>

