<%	' [BN, 6/10/05] See similar code (that I blanked out using IF-FALSE) for searching by product category in file Details.asp.
	' It may provide a model for how AIC did this kind of thing slightly differently or slighlty better (in particular, without (?) 
	' the "if RSS("SerialNumber") <> "" then" bug that they had below).
%>


<% 
' Some Global Parameters ...
' PageWidth = 1120 
PageWidth = "100%"
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
<table align='center' width=<%=PageWidth%> XXXborder="1">

<tr>
<td>
		<table border="0" bordercolor="red" cellpadding=0 bgcolor='blue' align='center' width='100%'>		<% ' Start Table Header 1 %>
            <tr>
                <td>
                					
				<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"> <% ' Start Table Header 1.1.1 %>
                        
				<tr>
					<td width="4" align="center">&nbsp;</td>
					<td align="center">
						<img alt="Star Lite International, LLC - GPS Sensors" src="https://www.starlite-intl.com/Images/StarLogo100.png" title="Star Lite International, LLC - GPS Sensors" align="top" width="80">
					</td>
					<td width="12" align="center">&nbsp;</td>
					<td align="left" height="90">
						<font style="color:#DAE8FB; font-size:9.0pt;">
						<em><b>Serving Business, Government, OEM, Education and the General Public since 1994</b></em>
						</font>
					</td>
					<td width="4" align="center">&nbsp;</td>
				</tr>

				<tr>
					<td colspan="5" style="left-margin:auto; right-margin:auto;">
						<% If TRUE Then %>
						<center style="height:30px;"> 
								<form action="https://www.starlite-intl.com/mobile/Misc/redirect.asp" method='GET' >
									<!-- <font style="color:white; font-size:9.5pt;">Browse:&nbsp;&nbsp;&nbsp;</font> -->
									<select name="target">
									    <option>Select a page</option>
										<option>Home</option>
										<option>Products</option>
										<option>OEM GPS Sensors</option>
										<option>Night Vision Optics</option>
										<option value="Terms">Terms & Conditions</option>
										<option>Contact Us</option>
										<option>Shopping Cart</option>
									</select>&nbsp;<input type=submit value="Go">
								</form>
						</center>

						<center style="height:30px;"> 
								<form action="https://www.starlite-intl.com/Search/SearchSummary.asp" method='GET' >
									<!-- <font style="color:white; font-size:9.5pt;">Search:&nbsp;</font> -->
									<input type=hidden name=CF value='KWS'> <!-- KWS = Keyword Search -->
									<input type=text name="KW" value="Type a keyword">&nbsp;<input type=submit value="Go">   <!-- KW = Keyword -->
								</form>
						</center>
						<% End If %>
					</td>
				</tr>

				</table>   <% ' End Table Header 1.1.1 %>
					
				</td>
			</tr>
			
		</table>	<% ' End Table Header 1 %>

</td>
</tr>
</table>
<% End If %>
