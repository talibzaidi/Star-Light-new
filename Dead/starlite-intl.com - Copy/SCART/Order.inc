<%
'on error resume next
sAction = CStr(Request.QueryString("action"))
iItemID = (Request.QueryString("item"))
iItemCount = CInt(Request.QueryString("count"))

if sar = "New Products" then 

	 set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates WHERE  NewProduct = yes ", "DSN=STAREC1" , 1, 4	
elseif sar ="Specials" then 

	 set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates WHERE  Special = Yes ", "DSN=STAREC1" , 1, 4

else


	 set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates WHERE  SID = " + sid +   " ", "DSN=STAREC1" , 1, 4
	
end if
	countChoc = RS.RecordCount
	set RXS = CreateObject("ADODB.Recordset")
	
%>
<%
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
	
	<TR><TD COLSPAN=4 ALIGN="Right" bgcolor="#DDDDDD"><B>Total:</B></TD><TD ALIGN="Right" bgcolor="#DDDDDD"><font face=tahoma size=2 color=#b9000><b>$<%= FormatNumber(sTotal,2) %></b></font></TD></TR>
	</TABLE>
	<%


' We implemented this this way so if you attach it to a database you'd only need one call per item
'*****************************************************************************************************************************
' 220
Function asGetItemParameters(iItemID)

Dim bParameters 

                                 if Session("Country") = "USA" then       
                              RXS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates WHERE  ITEMID = " + iItemID +   " ", "DSN=STAREC1" , 1, 4

		 bParameters = Array("../imi/" +RXS("Pic1") +"","" +RXS("PName") +"", "" +RXS("Descr") +"",formatcurrency(RXS("MSL")*RXS("Freight")), formatcurrency(RXS("Cost")*RXS("Freight")/RXS("GPM")), RXS("PID"),RXS("ITEMID") )
		
else

 bParameters = Array("../imi/" +RS("Pic1") +"","" +RS("Pname") +"", "" +RS("Descr") +"",formatcurrency(RS("MSL")*RS("Duty")*RS("Freight")*RS("Exch")), formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")/RS("GPM")),RS("PID"),RS("ITEMID") )
		

end if

	
' Return array containing product info.
asGetItemParameters = bParameters

End Function

%>