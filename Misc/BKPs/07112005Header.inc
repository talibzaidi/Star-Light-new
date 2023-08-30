
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
							&nbsp;&nbsp;<font color='white'>|</font>&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/search/search.asp">
							<font color='white' face='Verdana'>Advanced Search</font></a>
							</td>
							<td>
							&nbsp;&nbsp;<font color='white'>|</font>&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/scart/scartstart.asp?pid=0&sid=11&area=Specials&sar=Specials">
							<font color='white' face='Verdana'>Specials</font></a>
							</td>
							<td>
							&nbsp;&nbsp;<font color='white'>|</font>&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/scart/scart.asp?pid=0&sid=11&area=New!&sar=New%20Products">
							<font color='white' face='Verdana'>New Products</font></a>
							</td>
							<td>
							&nbsp;&nbsp;<font color='white'>|</font>&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/Misc2/contact.asp"><font color='white' face='Verdana'>Contact Us</font></a>
							</td>
							<td>
							&nbsp;&nbsp;<font color='white'>|</font>&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/Misc2/Terms_and_Conditions.asp"><font color='white' face='Verdana'>Terms &amp; Conditions</font></a>
							</td>
							<td>
							&nbsp;&nbsp;<font color='white'>|</font>&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/Misc2/LinkList.asp"><font color='white' face='Verdana'>Links</font></a>
							</td>
							<td>
							&nbsp;&nbsp;<font color='white'>|</font>&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/scart/scart.asp?sar=Certificates&area=Gift%20Certificates&sid=257"><font color='white' face='Verdana'>Gift Certificates</font></a>
							</td>
							<td>
							&nbsp;&nbsp;<font color='white'>|</font>&nbsp;&nbsp;
							</td>
							<td>
							<a href="http://www.starlite-intl.com/scart/scart.asp?action=viewcart&pid=0&sid=11&">
							<font color='white' face='Verdana'>Shopping Cart</font></a>
							</td>
						</tr>
						</table>   <% ' End Table Header 1.1.2 %>
						


						<table border=0 bordercolor='red' align='center' width='100%' >	<% ' Start Table Header 1.1.3 %>
						<tr>
						<td align='center'>
																			<% ' Start Table Header 1.1.3.1 %>
							<table  border="0" cellpadding="5" cellspacing="0" bordercolor="#000000" bordercolordark="#000000" bordercolorlight="#000000">
                           
                           
							<form action="http://www.starlite-intl.com/scart/scart.asp" method="GET" id="form1" name="form1">             
                            
                            <tr>
                            <%
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							dim sfffqstring
							sfffqstring = "SELECT ITEMID FROM PRODUCT "
							Set RHS = Conn.Execute(sfffqstring)
							%>
							<td>
							
							<%
							SfSQL = "Select Distinct Manufa from PRODUCT ORDER BY Manufa ASC"
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							Set RSS = Conn.Execute(SfSQL)
							%>

							<select name="Manufa" size="1">
								<option selected value>Search by Manufacturer
								<%	Do While Not RSS.EOF%>
								
									<option value="<%=RSS("Manufa")%>">
									<font face="arial" size="1"><%=RSS("Manufa")%></font>
								<%	RSS.MoveNext
								
									Loop
									RSS.Close
								%>
							</select>
							</td>
							
							<td valign="bottom">
							<input type='hidden' name='sar' value='Manufa'>
							<input type='hidden' name='SID' value='0'>
							<input type="submit" value="Submit" id=submit1 name=submit1>
							</td>
							</tr>
							
                            </form>
							
                            
							</table>	<% ' End Table Header 1.1.3.1 %>
							
						</td>
						
						<td align='center'>
				
																						<% ' Start Table Header 1.1.3.2 %>
							<table  border="0" cellpadding="5" cellspacing="0" bordercolor="#000000" bordercolordark="#000000" bordercolorlight="#000000" bgcolor="blue">
                         										
							<form action="http://www.starlite-intl.com/scart/scart.asp" method="GET" name="PID">
                
                            <tr>
                            <%
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							dim sqstring2
							sqstring2 = "SELECT ITEMID FROM PRODUCT "
							Set RHS = Conn.Execute(sqstring2)
							%>
							
							<td>
							<%
							SSQL = "Select  AID, AreaName from Area51 ORDER BY AreaName ASC"
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							Set RSS = Conn.Execute(SSQL)
							%>
							
							<select name="Manufat" size="1">
								<option selected value>Search by Product Category
								<option>
								<%	Do While Not RSS.EOF
								
									manufatar = RSS("AreaName")
									if manufatar = "New!" then
									manufatar= "New Products"
									end if
								%>

								<%	' if RSS("SerialNumber") <> "" then   [BN, 6/10/05] ERROR IN ORIGINAL CODE BY IAC (?)
									if RSS("AID") <> "" then %>
									<option value="<%=manufatar%>">
									<font face="arial" size="1"><%=manufatar%></font>
								<%	end if %>
        
								<%	RSS.MoveNext
								
									Loop
									RSS.Close 
								%>        
							</select>
							</td>
							
							<td valign="bottom"> 
							<input type='hidden' name='Area' value='iii'>
							<input type='hidden' name='SID' value='0'>
							<input type="submit" value="Submit" id=submit2 name=submit2>
							</td>
                            </tr>
                            
                            </form>
                            
							</table>	<% ' End Table Header 1.1.3.2 %>
							
						</td>
						</tr>
						</table>	<% ' End Table Header 1.1.3 %>
				
				</td>
            </tr>
		</table>	<% ' End Table Header 1 %>
            
<% End If %>
            

