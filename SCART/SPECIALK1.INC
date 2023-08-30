<%
Set KRS = CreateObject("ADODB.Recordset")
KRS.Open "SELECT *, Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE Special = yes ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4
%>


<table border="0" cellpadding="10" cellspacing="7" align="center">   <% ' Begin Table 1 %>
<tr>
	<td valign="top">
	
<% 
	KRS.MoveFirst
	do while not KRS.EOF
        on error resume next
%>

	<table border="0" valign=top cellspacing="0" width="300">		<% ' Begin Table 1.1 %>
		<tr>
			<td valign="top">
			<A href="../Detail.asp?pid=<%=KRS("PID")%>"> <img src="../Imi/<%=KRS("Pic1")%>" width="100" border="0"></a>
            </td>
			<td valign="top" width=180>
			<A href="./scart.asp?action=add&item=<%=KRS("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> <img src="../Images/order.gif"  border="0"></a>
			<br>
			<Font face="Tahoma" size="1"><u>ID # <%=KRS("ITEMID")%></u></font>
			<% If Session("Country") = "USA" Then %>
				<br>
				<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2">
				<% ' 6/18/06, commented out, BN: =formatcurrency(KRS("MSL")*KRS("Freight"))%></font>
				<% =formatcurrency(KRS("MSL"))%></font>
				<br>
				<font face="Tahoma" size="2"><b>Our Price </b></font>
				<font face="Tahoma" size="2" color="#B90000">
				<% If KRS("ShowPrice") = true Then  '  "ShowPrice" really means "Don't Show Price" %>
				<b>Click ORDER to order or see our price.</b>
				<% Else %>
				<b><i><%=formatcurrency(KRS("Cost")*KRS("Freight")*(1/(1-(KRS("GPM")))))%></i></b>
				<% End If %>
				</font>
			<% Else %> 
				<br>
				<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2">
				<% ' 6/18/06, commented out, BN: =formatcurrency(KRS("MSL")*KRS("Duty")*KRS("Freight")*KRS("Exch"))%></font>
				<% =formatcurrency(KRS("MSL")*KRS("Duty")*KRS("Exch"))%></font>
				<br>
				<Font face="Tahoma" size="2"><b>Our Price </b></font>
				<Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(KRS("Cost")*KRS("Duty")*KRS("Freight")*KRS("Exch")*(1/(1-(KRS("GPM")))))%></i></b></font>
			<% End If %>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
			<Font face="Tahoma" size="2"><b> <%=KRS("PName")%></b> </font>
			</td> 
		</tr>
	</table>		<% ' End Table 1.1 %>
	
	</td>

	<td valign="top">
<%   
'******************************************
   KRS.MoveNext 
   if KRS.EOF <> True then 
'******************************************%>

	<table border="0" valign=top cellspacing="0" width="300">		<% ' Begin Table 1.2 %>
		<tr>
			<td valign="top">
			<A href="../Detail.asp?pid=<%=KRS("PID")%>"> <img src="../Imi/<%=KRS("Pic1")%>"  width="100" border="0"></a>
			</td>
			<td valign="top"  width=180>
			<A href="./scart.asp?action=add&item=<%=KRS("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> <img src="../Images/order.gif"  border="0"></a>
			<br>
			<Font face="Tahoma" size="1"><u>ID # <%=KRS("ITEMID")%></u></font>
			<% If Session("Country") = "USA" then %>
				<br>
				<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2">
				<% ' 6/18/06, commented out, BN: =formatcurrency(KRS("MSL")*KRS("Freight"))%></font>
				<% =formatcurrency(KRS("MSL"))%></font>
				<br>
				<Font face="Tahoma" size="2"><b>Our Price </b></font> 
				<font face="Tahoma" size="2" color="#B90000">
				<% If KRS("ShowPrice") = true Then    '  "ShowPrice" really means "Don't Show Price" %>
				<b>Click ORDER to order or see our price.</b>
				<% Else %>
				<b><i><%=formatcurrency(KRS("Cost")*KRS("Freight")*(1/(1-(KRS("GPM")))))%></i></b>
				<% End If %>
				</font>
				
			<% Else %> 
				<br>
				<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2">
				<% ' 6/18/06, commented out, BN:  =formatcurrency(KRS("MSL")*KRS("Duty")*KRS("Freight")*KRS("Exch"))%></font>
				<% =formatcurrency(KRS("MSL")*KRS("Duty")*KRS("Exch"))%></font>
				<br>
				<Font face="Tahoma" size="2"><b>Our Price </b></font> 
				<Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(KRS("Cost")*KRS("Duty")*KRS("Freight")*KRS("Exch")*(1/(1-(KRS("GPM")))))%></i></b></font>
			<% End If %>
			</td>
		</tr>
		<tr>
		<td colspan="2" align="left">
		<Font face="Tahoma" size="2"><b> <%=KRS("PName")%></b> </font>
		</td>
		</tr>
	</table>		<% ' End Table 1.2 %>

	</td>
</tr>

<tr>
	<td valign=top>
<%
   else
   end if
   KRS.MoveNext
   loop
   KRS.Close 
%>


	</td>
</tr>
</table>	<% ' End Table 1 %>