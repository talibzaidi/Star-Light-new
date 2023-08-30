<%
	 set KRS = CreateObject("ADODB.Recordset")
                 KRS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  Special = yes ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4
	

%>


<div    align="center"><center>


<table border="0" cellpadding="0" cellspacing="7" width="380"><tr><td valign="top">
<% 
	KRS.MoveFirst
	do while not KRS.EOF
        on error resume next
      
%>
<table border="0"  cellspacing="0" width="168">


                    <tr>
                        <td colspan="2"><font color="#FF0000" size="4"><strong></strong></font><p><font
                        color="#BB0000" size="4"><strong><br>TODAYS
                        SPECIALS</strong></font>
</td></tr>
<tr><td valign="top">

<A href="../Detail.asp?pid=<%=KRS("PID")%>"> <img src="../Imi/<%=KRS("Pic1")%>"  width="100" border="0"></a>


                        </td>
<td valign="top">

<A href="./scart.asp?action=add&item=<%=KRS("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> <img src="../Images/order.gif"  border="0"></a>
<br>
<Font face="Tahoma" size="1"><u>ID # <%=KRS("ITEMID")%></u></font>
<%if Session("Country") = "USA" then %>
<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2"><%=formatcurrency(KRS("MSL")*KRS("Freight"))%></font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(KRS("Cost")*KRS("Freight")*(1/(1-(KRS("GPM")))))%></i></b></font>

<% else %> 

<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2"><%=formatcurrency(KRS("MSL")*KRS("Duty")*KRS("Freight")*KRS("Exch"))%></font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(KRS("Cost")*KRS("Duty")*KRS("Freight")*KRS("Exch")*(1/(1-(KRS("GPM")))))%></i></b></font>


<% end if %>

</td>
                    </tr>
<tr>
<td colspan="2" align="center">
<Font face="Tahoma" size="2"><b> <%=KRS("PName")%></b> </font>
</td>
</tr>
        

                </table>
</td><td>
<%   
'******************************************
   KRS.MoveNext 
   if KRS.EOF <> True then 
'******************************************%>
<table border="0"  cellspacing="0" width="168">


                    <tr>
                        <td colspan="2"><font color="#FF0000" size="4"><strong></strong></font><p><font
                        color="#BB0000" size="4"><strong><br>TODAYS
                        SPECIALS</strong></font>
</td></tr>
<tr><td valign="top">

<A href="../Detail.asp?pid=<%=KRS("PID")%>"> <img src="../Imi/<%=KRS("Pic1")%>"  width="100" border="0"></a>


                        </td>
<td valign="top">

<A href="./scart.asp?action=add&item=<%=KRS("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> <img src="../Images/order.gif"  border="0"></a>
<br>
<Font face="Tahoma" size="1"><u>ID # <%=KRS("ITEMID")%></u></font>
<%if Session("Country") = "USA" then %>
<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2"><%=formatcurrency(KRS("MSL")*KRS("Freight"))%></font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(KRS("Cost")*KRS("Freight")*(1/(1-(KRS("GPM")))))%></i></b></font>

<% else %> 

<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2"><%=formatcurrency(KRS("MSL")*KRS("Duty")*KRS("Freight")*KRS("Exch"))%></font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(KRS("Cost")*KRS("Duty")*KRS("Freight")*KRS("Exch")*(1/(1-(KRS("GPM")))))%></i></b></font>


<% end if %>

</td>
                    </tr>
<tr>
<td colspan="2" align="center">
<Font face="Tahoma" size="2"><b> <%=KRS("PName")%></b> </font>
</td>
</tr>
 
                </table>




</td></tr><tr><td>

<%
   else
   end if
   KRS.MoveNext
    loop
      KRS.Close %>
                </center></div>

</td></tr></table>