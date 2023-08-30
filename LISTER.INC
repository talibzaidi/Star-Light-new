

<%
if sar = "New Products" then 

	 set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates WHERE  NewProduct = yes ", "DSN=STAREC1" , 1, 4
	
elseif sar ="Specials" then 

	 set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates WHERE  Special = Yes ", "DSN=STAREC1" , 1, 4

else

	 set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Duty AS Duty, Rates.GPM AS GPM, Rates.Freight AS Freight FROM Product, Rates WHERE  SID = " + SID  +   " ", "DSN=STAREC1" , 1, 4

end if
	do while not rs.eof
%>

<div    align="center"><center>


<table border="0"
                cellspacing="0" width="100%">


                    <tr>
                        <td colspan="3"><font color="#FF0000" size="4"><strong></strong></font><p><font
                        color="#BB0000" size="4"><strong><br></strong></font>
</td>





</tr>
<tr><td valign="top">

<A href="Detail.asp?pid=<%=RS("PID")%>"> <img src="Imi/<%=RS("Pic1")%>"  width="100" border="0"></a>


                        </td>
<td valign="top">

<A href="SCART/SCART.asp?pid=<%=RS("PID")%>"> <img src="Images/order.gif"  border="0"></a>
<br>
<Font face="Tahoma" size="1"><b><u>ID # <%=RS("ITEMID")%></b></u></font>



<%if Session("Country") = "USA" then %>
<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2"><%=formatcurrency(RS("MSL")*RS("Freight"))%></font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(RS("Cost")*RS("Freight")/RS("GPM"))%></i></b></font>

<% else %> 

<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2"><%=formatcurrency(RS("MSL")*RS("Duty")*RS("Freight")*RS("Exch"))%></font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")/RS("GPM"))%></i></b></font>


<% end if %>

</td>
<td valign="top" width="200" >
&nbsp;&nbsp;<Font face="Tahoma" size="2"><b><%=RS("Descr")%></b> </font>
</td>

                    </tr>
<tr>
<td colspan="3" align="left">
<Font face="Tahoma" size="4"><b><u><i> <%=RS("PName")%></b></i></u></font>
</td>

</tr>
                </table>
                </center></div>

<% 
       rs.movenext
       loop
       rs.Close %>