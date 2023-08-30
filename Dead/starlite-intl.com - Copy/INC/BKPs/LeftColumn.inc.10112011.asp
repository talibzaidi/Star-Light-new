			<% ' Start Table 1.1.1 %>
			<table style="width:210px; border:0px solid red;" cellpadding="5px" cellspacing="0" >   
				<tr>
					<td align="center" valign="bottom">
						<form method="get" name="Country">
						<p><br><font face="Tahoma" size="2">You are currently a 
						<% If Session("Country") = "Canada" Then %>
							<img src="https://www.starlite-intl.com/Images/can1.gif" WIDTH="36" HEIGHT="18"> 
						<% Else				' Previously: ElseIf Session("Country") = "USA" Then 
							Session("Country") = "USA"
						%>
							<img src="https://www.starlite-intl.com/Images/USA1.gif" WIDTH="34" HEIGHT="18"> 
						<% End if %> 
						customer. Click on a button below to change countries.
						</font></p>
								
						<p>
                        <font face="Tahoma">
                        <strong>
                        <input type="submit" name="Canada" value="Canada">
                        </strong>
                        </font>

                        <font face="Tahoma">
                        <strong>
                        <input type="submit" name="  USA  " value="USA">
                        </strong>
                        </font>
                        </p>


						<input type="hidden" name="pid" value="<%=request("pid")%>">
						<input type="hidden" name="sid" value="<%=request("sid")%>">
						<input type="hidden" name="area" value="<%=request("area")%>">
						<input type="hidden" name="sar" value="<%=request("sar")%>">
						<input type="hidden" name="action" value="<%=request("action")%>">
						</form>

						<center>
						<font style="font-size:8pt">
						<a href="http://www.starlite-intl.com/Misc/AuthorizedDealerFor.asp" style="text-decoration:none">Authorized Dealer for Garmin, USGlobalSat, Pharos</a>
						</font>
						</center>

						<br />
						<form action="https://www.starlite-intl.com/scart/scart.asp?sid=197&amp;sar=&amp;area=GPS+Navigation%2C+GPS+Sensors%2C+OEM%2C+FishFinders%2C+Maps." method="Post" >
								<input type="Submit" value="Check for GPS Rebates">
						</form>
					</td>
				</tr>	
                
                <% ' *************************************************************************** 
                   ' Ads ...  %>
                
                <tr>
                    <td align="center" style="border:solid 0px red;">   

                    <table class="LeftColumnAds" style="border:0px solid red;"  cellspacing="15">

                    <tbody>

                    <tr >
                    <td >
                        <a href="https://www.starlite-intl.com/Detail.asp?pid=2846">
                        <img alt="GTU 10 tracking GPS" src="https://www.starlite-intl.com/Imi/GTU10t.jpg" height="58" width="74">
                        </a>
                        <b>GTU 10</b> by Garmin. High sensitivity tracking GPS, waterproof, U.S. national coverage, GSM wireless connection.
			 Combines a web-based tracking service with GPS technology to keep safe watch on children, pets and property.</font>.
                    </td>
                    </tr>

                    <tr>
                    <td >
                        <a href="https://www.starlite-intl.com/Detail.asp?pid=2437">
                        <img alt="PTL117 3.5G Smart phone and PDA" src="https://www.starlite-intl.com/Imi/PTL117t.gif" style="width: 65px; height: 80px;" align="left" height="128" width="100">
                        </a>
                        Pharos <b>Traveler 117 GPS</b>
                        - 3.5G smartphone - PDA combo. Featuring a flush 2.8 inch touch screen,
                        free live traffic, gas price, movie and weather information and more.
                    </td>
                    </tr>

                    <tr>
                    <td >
                        <a href="https://www.starlite-intl.com/Detail.asp?pid=2498&Key=">
                        <img alt="PTL 137 smartphone by Pharos" src="https://www.starlite-intl.com/Imi/PTL137t.jpg" height="82" width="77">
                        </a>
                        Pharos <b>Traveler 137</b> GPS Smartphone,
                        with 3.5'' flush touch-screen WVGA display. 3.5G communications capability based on a tri-band UMTS/HSDPA/HSUPA and a 
                        quad-band GSM/GPRS/EDGE cellular modems. On board two (2) cameras and mic.
                    </td>
                    </tr>

                    <tr >
                    <td >
                        <a href="https://www.starlite-intl.com/Detail.asp?pid=2763">
                        <img alt="Astro GPS Dog tracking by Garmin" src="https://www.starlite-intl.com/imi/010-00596-20t.jpg" height="68" width="64">
                        </a>
                        <a href="https://www.starlite-intl.com/Detail.asp?pid=2116&amp;Key=">
                        </a>
                        <b>Astro Bundle</b> by Garmin,
                        premier high sensitivity GPS-enabled <b>dog tracking system</b> 
                        for sporting dogs. This unique system pinpoints your dog's position and shows you exactly where he is, 
                        even when you can't see or hear him.
                    </td>
                    </tr>


		    <tr >
                    <td >
                        <a href="https://www.starlite-intl.com/Detail.asp?pid=2103&Key=">
                        <img alt="PTL 600e by Pharos" src="https://www.starlite-intl.com/Imi/PTL600et.jpg" height="68" width="64">
                        </a>
                        <a href="https://www.starlite-intl.com/Detail.asp?pid=2103&Key=">
                        </a>
                        <b>PTL 600e</b> by Pharos,
                        integrates GPS, Phone, PDA, WiFi, Bluetooth, Windows Mobile Office, Live Search, digital camera, FM radio, Photo Viewer, MP3 
			and video player and much more. <B> SPECIAL Price while still in our stock</b>, Maps not inclluded...
                    </td>
                    </tr>


                    <tr >
                    <td >
                        <a href="https://www.starlite-intl.com/Detail.asp?pid=2462&amp;Key=">
                        <img alt="night vision monocular by Night Owel Optics" src="https://www.starlite-intl.com/imi/NODS3t.jpg" height="79" width="65">
                        </a>
                        <b>NODS3</b> - Night Vision by Night Owel Optics.
                        Quality night vision monocular with a built-in infrared illuminator.
                        Features an accentuated palm grip with ergonomically placed soft-touch
                        operational buttons for easy one-hand operation and much more.
                    </td>
                    </tr>

                    </tbody>
                    </table>


                    </td>
                </tr>

                <% ' *************************************************************************** %>
                
                	
			</table>   
			<% ' End Table 1.1.1 %>
		
		
