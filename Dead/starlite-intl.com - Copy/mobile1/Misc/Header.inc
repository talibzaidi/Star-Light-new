<%	' [BN, 6/10/05] See similar code (that I blanked out using IF-FALSE) for searching by product category in file Details.asp.
	' It may provide a model for how AIC did this kind of thing slightly differently or slighlty better (in particular, without (?) 
	' the "if RSS("SerialNumber") <> "" then" bug that they had below).
%>


<% 
' Some Global Parameters ...
' PageWidth = 1120 
PageWidth = "100%"
'Response.Write "<br>***action = " & action 
'Response.Write "<br>***InArea = " & InArea 
%>


<% If FALSE Then %>
<!-- AddThis Button (from https://www.addthis.com/get/sharing) BEGIN -->
	<div class="addthis_toolbox addthis_floating_style addthis_16x16_style" style="left:35px; top:130px;">
	Share
	<!-- 
	<a class="addthis_button_preferred_3"></a>
	<a class="addthis_button_preferred_2"></a>
	<a class="addthis_button_preferred_1"></a>
	<a class="addthis_button_linkedin"></a>
	<a class="addthis_button_preferred_4"></a>
	<a class="addthis_button_compact"></a>
	-->

	<a class="addthis_button_preferred_1"></a>
	<a class="addthis_button_preferred_2"></a>
	<a class="addthis_button_preferred_3"></a>
	<a class="addthis_button_preferred_4"></a>
	<a class="addthis_button_preferred_5"></a>
	<a class="addthis_button_compact"></a>
	</div>
	<script type="text/javascript">	var addthis_config = { "data_track_addressbar": true };</script>
	<script type="text/javascript" src="//s7.addthis.com/js/300/addthis_widget.js#pubid=ra-50dd11947c476138"></script>
<!-- AddThis Button END -->
<% End If %>



<% If TRUE Then 	 %>
<table align='center' width=<%=PageWidth%> border="0">

<tr>
<td>
		<table border="0" bordercolor="red" cellpadding=0 bgcolor=#0080FF align='center' width='100%'>		<% ' Start Table Header 1 %>
            <tr>
                <td>
                					
				<table align="center" border="0" color="green" cellpadding="0" cellspacing="0" width="100%"> <% ' Start Table Header 1.1.1 %>
                        
				<tr>
					<td width="4" align="center">&nbsp;</td>
					<td align="center">
						<a href="https://www.starlite-intl.com/mobile1">
						<img border="0" alt="Star Lite International, LLC - GPS Sensors" src="https://www.starlite-intl.com/Images/StarLogo100.png" title="Star Lite International, LLC - GPS Sensors" align="top" width="80">
						</a>
					</td>
					<td width="12" align="center">&nbsp;</td>
					<td align="left" height="90">
						<!-- <font style="color:#DAE8FB; font-size:9.0pt;"> -->
						<font style="color:#FFFFFF; font-size:9.0pt;">
						<em><b>Serving Business, Government, OEM, Education and the General Public since 1994</b></em>
						</font>
					</td>
					<td width="4" align="center">&nbsp;</td>
				</tr>

				<tr>
					<td colspan="5" style="left-margin:auto; right-margin:auto;">

						<center>
						<font style="font-size:11.0pt; text-decoration:none;">                                               
						<a href="https://www.starlite-intl.com/mobile1/index.asp"><font style="color:white; text-decoration:none;"><b>Home</b></font></a> 
						<font style="color:white; text-decoration:none;">&nbsp; &bull; &nbsp;</font>
						<a href="https://www.starlite-intl.com/mobile1/Misc2/contact.asp"><font style="color:white;"><b>Contact</b></font></a> 
						<font style="color:white; text-decoration:none;">&nbsp; &bull; &nbsp;</font>
						<a href="https://www.starlite-intl.com/mobile1/Misc2/Terms_and_Conditions.asp"><font style="color:white;"><b>Terms</b></font></a>
						<font style="color:white; text-decoration:none;">&nbsp; &bull; &nbsp;</font>
						<a href="https://www.starlite-intl.com/mobile1/scart/scart.asp?action=viewcart&pid=0&sid=11"><font style="color:white;"><b>Cart</b></font>
						    <img align="bottom" src="https://www.starlite-intl.com/images/SCRT.png" border="0" height="20" ></a> 
						</font>
						</center>

						<% If FALSE Then %>
						<center style="height:30px;"> 
								<form action="https://www.starlite-intl.com/mobile1/Misc/redirect.asp" method='GET' >
									<!-- <font style="color:white; font-size:9.5pt;">Browse:&nbsp;&nbsp;&nbsp;</font> -->
									<select name="target">
									    <option>Select a page</option>
										<option>Home</option>
										<option>Products</option>
										<option>OEM GPS Sensors</option>
										<option>Night Vision Optics</option>
										<option>Communications</option>
										<option value="Terms">Terms & Conditions</option>
										<option>Contact Us</option>
										<option>Shopping Cart</option>
									</select>&nbsp;<input type=submit value="Go">
								</form>
						</center>
						<% End If %>

						<center style="height:30px;"> 
								<form action="https://www.starlite-intl.com/mobile1/Search/SearchSummary.asp" method='GET' >
									<font style="color:white; font-size:11.5pt;"><b>Search:</b>&nbsp;</font>
									<input type=hidden name=CF value='KWS'> <!-- KWS = Keyword Search -->

									<input type="text" value=" Type a keyword" name="KW" id="KW" size="13" style="height:25px; font-size:18px; color:#AAAAAA;"
									  onblur="if (this.value == '') {this.value = ' Type a keyword'; this.style.color='#AAAAAA'}"
									  onfocus="if (this.value == ' Type a keyword') {this.value = '';} {this.style.color='#000000'}" />&nbsp;<input type=submit value="Go">   <!-- KW = Keyword -->
								</form>
						</center>
					</td>
				</tr>

				</table>   <% ' End Table Header 1.1.1 %>
					
				</td>
			</tr>
			
		</table>	<% ' End Table Header 1 %>

</td>
</tr>


			<tr>
				<td>
					<table align="center" cellpadding="0" cellspacing="0" width="100%">
						<tr>
							<td>
							<!-- 
							11/11/13: John of MenuCool emailed me ...
							Your amenu-source.html is retrieved by Ajax that requires to be the same domain.
							So the link to it should be: href="/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/amenu-source.html" to avoid dis-matching domains.
							-->
							<% If TRUE Then %>
							<a id="amenuLink" href="/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/amenu-source.test2.html"></a>
							<% End If ' TRUE/FALSE %>
							</td>
						</tr>
						<tr>
							<td>
							<!-- # include virtual = "mobile1/Misc/Navbars/NavBar1.inc.asp" -->
							</td>
						</tr>
					</table>
				</td>
			</tr>


</table>
<% End If %>
