
<%	' [BN, 6/10/05] See similar code (that I blanked out using IF-FALSE) for searching by product category in file Details.asp.
	' It may provide a model for how AIC did this kind of thing slightly differently or slighlty better (in particular, without (?) 
	' the "if RSS("SerialNumber") <> "" then" bug that they had below).
%>

<% If True Then %>
<table bgcolor='blue' align='center'width='100%'>
<tr>
<td>
		<table border="0" bordercolor="green"  cellpadding=0 bgcolor='blue' align='center' width='100%'>		<% ' Start Table Header 1 %>
            <tr>
                <td background="https://www.starlite-intl.com/Images/topback.gif" bgcolor="blue">
                					
						<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"> <% ' Start Table Header 1.1.1 %>
                        <tr>
							<td></td>
							<td><img src="https://www.starlite-intl.com/Images/toptitle.jpg" WIDTH="411" HEIGHT="29"></td>
						</tr>
						
						<tr>
							<td></td>
							<td><img src="https://www.starlite-intl.com/Images/bottitle.JPG" WIDTH="411" HEIGHT="14"></td>
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
							<font color="white" size=2>Advanced Search</font></a>
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
							<a href="https://www.starlite-intl.com/scart/scart.asp?action=viewcart&amp;pid=0&amp;sid=11&amp;">
							<font color="white" size=2>Shopping Cart</font></a>
							</td>
						</tr>
						</table>					<% ' End Table Header 1.1.2 %>
						

<% If ShowHeaderMenus = "" Then ' Empty string acts as True, non-empty as False. Used, for example, in GWays/GPS.asp. %>
						<table border=0 bordercolor='red' align='center' width='100%' >	<% ' Start Table Header 1.1.3 %>
						<tr>
						<td align='center'>
																			<% ' Start Table Header 1.1.3.1 %>
							<table  border="0" cellpadding="5" cellspacing="0" bordercolor="#000000" bordercolordark="#000000" bordercolorlight="#000000">
                           
							<form action="http://www.starlite-intl.com/scart/scart.asp" method="GET" id="form1" name="form1">             
                            
                            <tr>
                            
                            <%
							'Set conn = Server.CreateObject("ADODB.Connection")
							'Conn.Open Session("ConnectionString")
							'dim sfffqstring
							'sfffqstring = "SELECT ITEMID FROM PRODUCT "
							'Set RHS = Conn.Execute(sfffqstring)
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
								<%	Do While Not RSS.EOF 
									If NOT RSS("Manufa") = "RepairMaster" Then   ' Don't list manufaturer of warranties. %>
										<option value="<%=RSS("Manufa")%>">
										<font face="arial" size="1"><%=RSS("Manufa")%></font>
								<%	End If
									RSS.MoveNext
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
							'Set conn = Server.CreateObject("ADODB.Connection")
							'Conn.Open Session("ConnectionString")
							'dim sqstring2
							'sqstring2 = "SELECT ITEMID FROM PRODUCT "
							'Set RHS = Conn.Execute(sqstring2)
							%>
							
							<td>
							
							<%
							SSQL = "Select AID, AreaName from Area51 ORDER BY AreaName ASC"
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							Set RSS = Conn.Execute(SSQL)
							%>
							
							<select name="Manufat" size="1">
								<option selected value>Search by Product Category
								<option>
								<%	Do While Not RSS.EOF
									manufatar = RSS("AreaName")
									If manufatar = "New!" Then
									manufatar= "New Products"
									End If
									' if RSS("SerialNumber") <> "" then   [BN, 6/10/05] ERROR IN ORIGINAL CODE BY IAC (?)
									If RSS("AID") <> "" Then
										If NOT manufatar = "Warranties" Then	' Don't list Warranties. %>
											<option value="<%=manufatar%>">
											<font face="arial" size="1"><%=manufatar%></font>
								<%		End If
									End If 
									RSS.MoveNext
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
<% End If   ' False %>
				
				</td>
            </tr>
		</table>	<% ' End Table Header 1 %>
            
</td>
</tr>
</table>
<% End If %>
            

