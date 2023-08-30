<% 
If (Request.form("Canada") <> "" OR Request.form("  USA  ") <> "") Then
	If Request.form("Canada") <> "" Then
		Session("Country") = "Canada"
	Else
		Session("Country") = "USA"
	End if
End If
%>

<form method="post" name="Country" action="<%=request.servervariables("URL")%>">
    <p align='center'>
    <font face="Tahoma" size="2">
    You are currently a 
    <% If Session("Country") = "Canada" Then%>
		<img src="Images/can1.gif" title="Canada"> 
	<% Else                            ' Previously: Elseif Session("Country") = "USA" Then 
		Session("Country") = "USA"
	%>
		<img src="Images/USA1.gif" title="USA"> 
	<% End If %> customer. Click on a button below to change countries.
	</font></p>
                                
    <p align='center'>
    <font face="Tahoma">
    <input type="submit" name="Canada" value="Canada">
    </font><font face="Tahoma">
    <input type="submit" name="  USA  " value="USA">
    </font>
    </p>
</form>
                 
<center>
<font style="font-size:8pt">
<a href="http://www.starlite-intl.com/Misc/AuthorizedDealerFor.asp" style="text-decoration:none">Authorized Dealer for Garmin, USGlobalSat</a></font></center>
				                   

<br>
<center>
<font color="#BB0000" size="4" face=Tahoma>
<strong>TODAY'S SPECIALS!</strong><br><br>
</font>
</center>
                                
<%
' **************************************************************************************************

If Err.number <> 0 then
     response.redirect "error.asp"
End If


'set rsSpecials = CreateObject("ADODB.Recordset")
'rsSpecials.Open "SELECT *, Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE SpecialsSubgroup > 0 ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4

' 11/9/15: Using the connection string method instead of the DSN method above.
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open Session("ConnectionString")
SQLstring = "SELECT *, Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE SpecialsSubgroup > 0 ORDER BY MSL ASC"
Set rsSpecials = Server.CreateObject("ADODB.Recordset")
'rsSpecials.Open SQLstring, Conn1, adOpenStatic, adLockOptimistic, adCmdText 
rsSpecials.Open SQLstring, Conn1, 3, 3, &H0001 


rsSpecials.MoveLast

numrec = rsSpecials.Recordcount		' BN: numrec is the number of available specials.
'Response.Write "numrec = " & numrec 
aryRS = array(0,0,0,0,0,0,0,0,0,0,0,0,0)
count = 0
               
               
Do Until count = 6		' 4/26/07, BN:   The RHS value is is the number of specials to display.		Old: spec
' Apparently in each cycle it needs to make sure that the random special we choose does not duplicate an earlier special.

	Randomize
	MinIndex = 1
	MaxIndex = numrec
	'Special = Int(Rnd * numrec)  
	Special = Int((MaxIndex - MinIndex + 1) * Rnd + MinIndex) 
	If True Then		' This is the old code.   
	For z = 0 to count
		For i = 0 to count       
			While aryRS(i) = Special		' BN: While new special is a duplicate of any earlier special.
				'Special = Int(Rnd*numrec)	' BN: Choose a new spacial.
				Special = Int((MaxIndex - MinIndex + 1) * Rnd + MinIndex)      
			Wend
		Next
	Next
	End If   ' True / False
	

aryRS(count) = Special		' BN: Save new special into the array.
'Response.Write "<br>" & Special & "<br>"
'For j = 0 to 6			' BN: To just confirm 
'	Response.Write aryRS(j) & " "
'Next

rsSpecials.MoveFirst
rsSpecials.Move(Special - 1)	' BN: Get ready to display the current cycle's special.
count = count + 1				' BN: Get ready to look for next special.
                 
If Err.number <> 0 Then
   Response.Redirect "error.asp"
End If
%>


<center>


<table border="0" cellspacing="5" >
<tr>
	<td valign="top" width="50%">
	<a href="Detail.asp?pid=<%=rsSpecials("PID")%>">
	<img src="Imi/<%=rsSpecials("Pic1")%>" title="<%=rsSpecials("PName")%>" width="100" border="0"></a>

	<font face="Tahoma" size="2"><b><%=rsSpecials("PName")%></b></font>  
    </td>
    
	<td valign="top" width="50%">
	<a href="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=rsSpecials("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> 
	<img src="Images/order.gif"  border="0"></a>
	
<% ' [BN, 12/18/17] This section was based on a copy from a recently-added section in Detail.inc file.

USARegPrice = formatcurrency(rsSpecials("MSL"))
USAOurPrice = formatcurrency(rsSpecials("Cost")*rsSpecials("Freight")*(1/(1-(rsSpecials("GPM")))))
USAPercentagePriceDiff = (abs(USARegPrice - USAOurPrice) / USARegPrice) * 100

If False Then
	Response.Write "<br><font size='2'>USARegPrice = " & USARegPrice & "</font>"
	Response.Write "<br><font size='2'>USAOurPrice = " & USAOurPrice & "</font>"
	Response.Write "<br><font size='2'>USAPercentagePriceDiff = " & USAPercentagePriceDiff & "</font>"
End If 

CanadaRegPrice = formatcurrency(rsSpecials("MSL")*rsSpecials("Duty")*rsSpecials("Exch"))		' = USARegPrice*rsSpecials("Duty")*rsSpecials("Exch")
CanadaOurPrice = formatcurrency(rsSpecials("Cost")*rsSpecials("Duty")*rsSpecials("Freight")*rsSpecials("Exch")*(1/(1-(rsSpecials("GPM")))))	' = USAOurPrice*rsSpecials("Duty")*rsSpecials("Exch")
CanadaPercentagePriceDiff = (abs(CanadaRegPrice - CanadaOurPrice) / CanadaRegPrice) * 100

If False Then
	Response.Write "<br><font size='2'>CanadaRegPrice = " & CanadaRegPrice & "</font>"
	Response.Write "<br><font size='2'>CanadaOurPrice = " & CanadaOurPrice & "</font>"
	Response.Write "<br><font size='2'>CanadaPercentagePriceDiff = " & CanadaPercentagePriceDiff & "</font>"
End If
%>
	<br>
	<font face="Tahoma" size="1"><b><u>ID # <%=rsSpecials("ITEMID")%></u></b></font>

	<% If Session("Country") = "USA" Then %>
		<% 
		If USAPercentagePriceDiff > 3 Then
			Response.Write "<br><font face='Tahoma' size='1'>Reg. Price: </font>"
			Response.Write "<font face='Tahoma' size='2'>" & USARegPrice & "</font>"
		End If
		%>
		<br>
		<font face="Tahoma" size="2"><b>Our Price</b></font> 		
		<font face="Tahoma" size="2" color="#B90000">
		<b><i>
		<% If rsSpecials("ShowPrice") = true Then '  "ShowPrice" really means "Don't Show Price" %>
			Click ORDER to order or see our price.
			<% Else %> 
			<% Response.Write USAOurPrice%>
			<% End If %>
		</i></b>
		</font>
		<% 
		If False Then
			Response.Write "<br><font size='2'>USARegPrice = " & USARegPrice & "</font>" 
			Response.Write "<br><font size='2'>USAOurPrice = " & USAOurPrice & "</font>"
			Response.Write "<br><font size='2'>USAPercentagePriceDiff = " & USAPercentagePriceDiff & "</font>"
		End If
		%>


	<% Else   ' [BN, 12/18/17] i.e if Canada %> 
		<% 
		If CanadaPercentagePriceDiff > 3 Then
			Response.Write "<br><font face='Tahoma' size='1'>Reg. Price: </font>"
			Response.Write "<font face='Tahoma' size='2'>" & CanadaRegPrice & "</font>"
		End If
		%>
		<br>
		<font face="Tahoma" size="2"><b>Our Price </b></font>   <% ' [BN, 5/21/07] It's ok to always show Our Price for Canada case. %> 
		<font face="Tahoma" size="2" color="#B90000">
		<b><i><% Response.Write CanadaOurPrice%></i></b>
		</font>
		<%
		If False Then
			Response.Write "<br><font size='1'>CanadaRegPrice = " & CanadaRegPrice & "</font>"
			Response.Write "<br><font size='1'>CanadaOurPrice = " & CanadaOurPrice & "</font>"
			Response.Write "<br><font size='1'>CanadaPercentagePriceDiff = " & CanadaPercentagePriceDiff & "</font>"
		End If
		%>
	<% End If %>

	</td>
</tr>

</table>
</center>
<br>

<% 
Loop

rsSpecials.Close 
%>