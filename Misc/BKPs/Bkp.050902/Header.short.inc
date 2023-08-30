
<%	' [BN, 6/10/05] See similar code (that I blanked out using IF-FALSE) for searching by product category in file Details.asp.
	' It may provide a model for how AIC did this kind of thing slightly differently or slighlty better (in particular, without (?) 
	' the "if RSS("SerialNumber") <> "" then" bug that they had below).
%>

<% If True Then %>
		<table border="0" bordercolor="green"  align='center' width='100%'>		<% ' Start Table Header 1 %>
            <tr>
                <td background="http://www.starlite-intl.com/Images/topback.gif" bgcolor="blue">
                					
						<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"> <% ' Start Table Header 1.1.1 %>
                        <tr>
							<td></td>
							<td><img src="http://www.starlite-intl.com/Images/toptitle.jpg" WIDTH="411" HEIGHT="29"></td>
						</tr>
						
						<tr>
							<td></td>
							<td><img src="http://www.starlite-intl.com/Images/bottitle.JPG" WIDTH="411" HEIGHT="14"></td>
						</tr>
						</table>   <% ' End Table Header 1.1.1 %>
					
				</td>
			</tr>
			
			<tr>	
                <td bgcolor="blue">
							
						<table align='center' border="0" bordercolor="red" cellpadding="0" cellspacing="0">   <% ' Start Table Header 1.1.2 %>
						<tr>
							<td>
							&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/index.asp"><font color='white' face='Verdana'>Home</font></a>
							</td>
							<td>
							&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/search/search.asp">
							<font color='white' face='Verdana'>Advanced Search</font></a>
							</td>
							<td>
							&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/scart/scartstart.asp?pid=0&sid=11&area=Specials&sar=Specials">
							<font color='white' face='Verdana'>Specials</font></a>
							</td>
							<td>
							&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/Misc2/contact.asp"><font color='white' face='Verdana'>Contact Us</font></a>
							</td>
							<td>
							&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/Misc2/Terms_and_Conditions.asp"><font color='white' face='Verdana'>Terms &amp; Conditions</font></a>
							</td>
							<td>
							&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/Misc2/LinkList.asp"><font color='white' face='Verdana'>Links</font></a>
							</td>
							<td>
							&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/scart/scart.asp?action=viewcart&pid=0&sid=11&">
							<font color='white' face='Verdana'>Shopping Cart</font></a>
							</td>
						</tr>
						</table>   <% ' End Table Header 1.1.2 %>
						
				
				</td>
            </tr>
		</table>	<% ' End Table Header 1 %>
            
<% End If %>
            

