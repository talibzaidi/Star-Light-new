<%

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
	set RXS = CreateObject("ADODB.Recordset")
	
%>
<%
Sub PlaceOrder()
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

End Sub
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
<%
	PlaceOrder
%>