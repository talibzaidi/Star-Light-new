<%	' [BN, 6/10/05] See similar code (that I blanked out using IF-FALSE) for searching by product category in file Details.asp.
	' It may provide a model for how AIC did this kind of thing slightly differently or slighlty better (in particular, without (?) 
	' the "if RSS("SerialNumber") <> "" then" bug that they had below).
%>


<% If TRUE Then %>
<table align='center' cellpadding=1 bgcolor='blue'>
<tr>
<td>
		<table border="0" bordercolor="green"  align='center' width='940'>		<% ' Start Table Header 1 %>
            <tr>
                <td background="http://www.starlite-intl.com/Images/topback.gif" bgcolor="blue">
                					
						<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"> <% ' Start Table Header 1.1.1 %>
                        <tr>
							<td></td>
							<td><img alt="Star Lite International, LLC - GPS Sensors" src="https://www.starlite-intl.com/Images/TitleStarLiteInternational.jpg" WIDTH="411" HEIGHT="43"></td>
						</tr>
						
						
						</table>   <% ' End Table Header 1.1.1 %>
					
				</td>
			</tr>
			
			<tr>	
                <td bgcolor="blue">
							
						<table align="center" border="0" bordercolor="red" cellpadding="0" cellspacing="0">   <% ' Start Table Header 1.1.2 %>
						<tr>
							<td><br>
							</td>
							<td style="text-align: center; font-family: arial;">
							<a href="http://www.starlite-intl.com/index.asp">
							<font color="white" size=2>Home</font></a>
							</td>
							<td style="font-family: arial;">
							&nbsp;&nbsp;<font color="white">|</font>&nbsp;&nbsp;
							</td>
							<td style="font-family: arial;">
							<a href="http://www.starlite-intl.com/search/search.asp">
							<font color="white" size=2>Products</font></a>
							</td>
							<td style="font-family: arial;">
							&nbsp;&nbsp;<font color="white">|</font>&nbsp;&nbsp;
							</td>
							<td style="font-family: arial;">
							<a href="http://www.starlite-intl.com/scart/scartstart.asp?pid=0&amp;sid=11&amp;area=Specials&amp;sar=Specials">
							<font color="white" size=2>Specials</font></a>
							</td>
							<td style="font-family: arial;">
							&nbsp;&nbsp;<font color="white">|</font>&nbsp;&nbsp;
							</td>
							<td style="font-family: arial;">
							<a href="http://www.starlite-intl.com/scart/scart.asp?pid=0&amp;sid=11&amp;area=New%21&amp;sar=New%20Products">
							<font color="white" size=2>What's New?</font></a>
							</td>
							<td style="font-family: arial;">
							&nbsp;&nbsp;<font color="white">|</font>&nbsp;&nbsp;
							</td>
							<td style="font-family: arial;">
							<a href="http://www.starlite-intl.com/Misc2/contact.asp">
							<font color="white" size=2>Contact Us</font></a>
							</td>
							<td style="font-family: arial;">
							&nbsp;&nbsp;<font color="white">|</font>&nbsp;&nbsp;
							</td>
							<td style="font-family: arial;">
							<a href="http://www.starlite-intl.com/Misc2/Terms_and_Conditions.asp">
							<font color="white" size=2>Terms &amp; Conditions</font></a>
							</td>
							<td style="font-family: arial;">
							&nbsp;&nbsp;<font color="white">|</font>&nbsp;&nbsp;
							</td>
							<td style="font-family: arial;">
							<a href="http://www.starlite-intl.com/Misc2/LinkList.asp">
							<font color="white" size=2>Links</font></a>
							</td>
							<td style="font-family: arial;">
							&nbsp;&nbsp;<font color="white">|</font>&nbsp;&nbsp;
							</td>
							<td style="font-family: arial;">
							<a href="http://www.starlite-intl.com/scart/scart.asp?sar=Certificates&amp;area=Gift%20Certificates&amp;sid=257">
							<font color="white" size=2>Gift Certificates</font></a>
							</td>
							<td style="font-family: arial;">
							&nbsp;&nbsp;<font color="white">|</font>&nbsp;&nbsp;
							</td>
							<td style="font-family: arial;">
							<a href="http://www.starlite-intl.com/scart/scart.asp?action=viewcart&amp;pid=0&amp;sid=11&amp;">
							<font color="white" size=2>Shopping Cart</font></a>
							</td>
						</tr>
						</table>					<% ' End Table Header 1.1.2 %>

										
				</td>
            </tr>
		</table>	<% ' End Table Header 1 %>
            
</td>
</tr>
</table>
<% End If    ' TRUE / FALSE %>
            

<!-- #INCLUDE virtual = "/Misc/Navbars/NavBar1.inc.asp" --> 