<%
    
     set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE  Special = yes ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4
     
    if not ( rs.eof or rs.bof ) then
    RS.MoveLast
                 numrec = RS.Recordcount
                 aryRS = array(0,0,0,0,0,0,0,0,0,0,0,0,0)
         count = 0
           
     do until count = spec
                 Randomize
                 Special = Int(Rnd*numrec)
              
                 for z =  0 to count
                 for i = 0 to count    
                
     while aryRS(i) = Special
                     Special = Int(Rnd*numrec)
     wend
     
     next
     next

                 aryRS(count) = Special
                 'Response.Write(Special)
                 RS.MoveFirst
     ' RS.Move(Special)
                 count = count +1
                 loop
                 
     gigaloop = 0
end if    

%>
<% do while gigaloop <> spec 
      RS.MoveFirst
      RS.Move(aryRS(gigaloop))
%>

<div    align="center"><center>


<table border="0" cellpadding="0" cellspacing="9"><tr><td valign="top">
<table border="0"
                cellspacing="0" width="168">


                    <tr>
                        <td colspan="2"><font color="#FF0000" size="4"><strong></strong></font><p><font
                        color="#BB0000" size="4"><strong><br>TODAY'S
                        SPECIALS</strong></font>
</td></tr>
<tr><td valign="top">

<A href="Detail.asp?pid=<%=RS("PID")%>"> <img src="Imi/<%=RS("Pic1")%>"  width="100" border="0"></a>


                        </td>
<td valign="top">

<A href="https://www.starlite-intl.com/scart/scart.asp?action=add&item=<%=RS("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> <img src="Images/order.gif"  border="0"></a>
<br>
<Font face="Tahoma" size="1"><b><u>ID # <%=RS("ITEMID")%></b></u></font>
<%if Session("Country") = "USA" then %>
<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2">
<% ' 6/18/06, commented out, BN:  =formatcurrency(RS("MSL")*RS("Freight"))%>
<% =formatcurrency(RS("MSL"))%>
</font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM")))))%></i></b></font>

<% else %> 

<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2">
<% ' 6/18/06, commented out, BN: =formatcurrency(RS("MSL")*RS("Duty")*RS("Freight")*RS("Exch"))%>
<% =formatcurrency(RS("MSL")*RS("Duty")*RS("Exch"))%>
</font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM")))))%></i></b></font>


<% end if %>

</td>
                    </tr>
<tr>
<td colspan="2" align="center">
<Font face="Tahoma" size="2"><b> <%=RS("PName")%></b></font>
</td>
</tr>
                </table>
</td><td valign="top">

<% gigaloop = gigaloop+1
     RS.MoveFirst
     RS.Move(aryRS(gigaloop))
     gigaloop = gigaloop+1
%>
<table border="0"
                cellspacing="0" width="168">


                    <tr>
                        <td colspan="2"><font color="#FF0000" size="4"><strong></strong></font><p><font
                        color="#BB0000" size="4"><strong><br>TODAY'S
                        SPECIALS</strong></font>
</td></tr>
<tr><td valign="top">

<A href="Detail.asp?pid=<%=RS("PID")%>"> <img src="Imi/<%=RS("Pic1")%>"  width="100" border="0"></a>


                        </td>
<td valign="top">

<A href="https://10.1.1.1/starlite/scart/scart.asp?action=add&item=<%=RS("ITEMID")%>&count=1&amp;sid=<%=0%>&amp;Area=<%=Area%>&amp;sar=<%="Special"%>"> <img src="Images/order.gif"  border="0"></a>
<br>
<Font face="Tahoma" size="1"><b><u>ID # <%=RS("ITEMID")%></b></u></font>


<%if Session("Country") = "USA" then %>
<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2">
<% ' 6/18/06, commented out, BN: =formatcurrency(RS("MSL")*RS("Freight"))%>
<%=formatcurrency(RS("MSL"))%>
</font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(RS("Cost")*RS("Freight")*(1/(1-(RS("GPM")))))%></i></b></font>

<% else %> 

<br>
<Font face="Tahoma" size="1">Reg. Price </font> <Font face="Tahoma" size="2">
<% ' 6/18/06, commented out, BN: =formatcurrency(RS("MSL")*RS("Duty")*RS("Freight")*RS("Exch"))%>
<% =formatcurrency(RS("MSL")*RS("Duty")*RS("Exch"))%>
</font>
<br>
<Font face="Tahoma" size="2"><b>Our Price </b></font> <Font face="Tahoma" size="2" color="#B90000"><b><i><%=formatcurrency(RS("Cost")*RS("Duty")*RS("Freight")*RS("Exch")*(1/(1-(RS("GPM")))))%></i></b></font>


<% end if %>


</td>
                    </tr>
<tr>
<td colspan="2" align="center">
<Font face="Tahoma" size="2"><b> <%=RS("PName")%></b></font>
</td>
</tr>
                </table>
</td><tr></table>


<% loop
      RS.Close %>
                </center></div>

