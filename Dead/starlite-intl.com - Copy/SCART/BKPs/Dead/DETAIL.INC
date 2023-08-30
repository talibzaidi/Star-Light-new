<%
     set RS = CreateObject("ADODB.Recordset")
                 RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE  PName Like 'Terms And Conditions' ", "DSN=STAREC1" , 1, 4
    on error resume next
    do while not rs.eof
%>

<div    align="center"><center>


<table border="0"
                cellspacing="0" width="380">


               

<tr>
</td>
<td colspan="3" valign="top"  >
&nbsp;&nbsp;<Font face="Tahoma" size="2"><b><%=RS("Descr")%></b> </font><br><br>
<% if rs("Text1") <> "" Then %>
&nbsp;&nbsp;<Font face="Tahoma" size="2"><b><%=RS("Text1")%></b> </font><br><br>
<% end if %>
&nbsp;&nbsp;<Font face="Tahoma" size="2"><b>
<%
count = 0
if rs("Text2") <> "" then
bullet = RS("Text2")
loobol = len(bullet)
Response.Write("<br>")
Response.Write("<font face =arial size=3>")
response.write("&nbsp;&nbsp;&nbsp;")
Response.Write("<b>•</b>")
response.write("&nbsp;&nbsp;")
loobol = loobol + 1

do Until count = loobol 

 charz = left(bullet,count)
 char = right(charz,1)
 if char = "~" then
 response.write("<br>")
 response.write("&nbsp;&nbsp;&nbsp;")
 response.write("<b>•</b>")
 response.write("&nbsp;&nbsp;")
 else
 response.write("<b>")
 response.write(char)
 response.write("</b>")
 end if
 count = (count+1)
loop
Response.Write("<br>")
Response.Write("<br>")
Response.Write("</font>")
end if
%><br>

</td>

</tr>
                </table>
                </center></div>

<% 
       rs.movenext
       loop
       rs.Close %>