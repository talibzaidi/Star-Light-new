<%@ LANGUAGE = VBScript %>

<%response.buffer=true%>

<% ar = Request("Area") %>
<% Area = Request("Area") %>

<% sar = ReQuest("sar") %>
<% SID = ReQuest("SID") %>
<% spec = 6 

	If (Request("Canada") <> "" OR Request("  USA  ") <> "") then
		If Request("Canada") <> "" then
			Session("Country") = "Canada"
		else
			Session("Country") = "USA"
		end if
	end if

%>


<html>


<head>
<link rel="stylesheet" type="text/css" href="../../Misc/StyleSheet1.css">
<title>Star Lite Intl. - GPS, GPS Sensors, GPS accessories, CB radios, 2-way radios, Marine electronics, Flash memory, Scanners, MP3, audio/video, hand tools</title>
<meta name="keywords" content="GPS, GPS navigation, GPS sensors, OEM GPS, GPS accessories, CB, cb radio, 2-way radios, Garmin gps, global positioning, Walkytalky, mobile tracking, fleet tracking, USglobasat gps, bluetooth, flash memory, gmrs, marine radios, navigation electronics, 2-way radios, radio scanners, marine radios, car audio, car stereos, power amplifiers, antennas, power supplies, regulated power supplies, dj, accessories, mechanics tools, hand tools, uniden, cobra, midland, mit, pyramid, pyle, solarcon">
<meta name="description" content="Large selection of - GPS, GPS sensors, GPS accessories, GPS OEM, PDA, tracking GPS, bluetooth GPS, Fish finders, sounders, cb radios and Walky-talky, 2-way radios, flash memory, MP3, radio scanners, digital cameras, car audio and car video, DJ, hand tools, mechanics tools.">
<% ' <meta name="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds"> %>

<script language="Javascript">
<!--
    once = new MakeArray(6)
    over = new MakeArray(6)
    under = new MakeArray(6)
    standard = new MakeArray(1)
    once[0].src = "../Images/question1.gif"
    once[1].src = "../Images/scart1.gif"
    once[2].src = "../Images/home1.gif"
    once[3].src = "../Images/new1.gif"
	once[4].src = "../Images/cat1.gif"
    once[5].src = "../Images/ex1.gif"    
    over[0].src = "../Images/question2.gif"
    over[1].src = "../Images/scart2.gif"
    over[2].src = "../Images/home2.gif"
    over[3].src = "../Images/new2.gif"
    over[4].src = "../Images/cat2.gif"
    over[5].src = "../Images/ex2.gif"
    under[0].src = "../Images/helpnav.gif"
    under[1].src = "../Images/shoppingcartnav.gif"
    under[2].src = "../Images/homenav.gif"
    under[3].src = "../Images/newproductsnav.gif"
    under[4].src = "../Images/onlinecataloguenav.gif"
    under[5].src = "../Images/specialsnav.gif"
    standard[0].src = "../Images/emptynav.jpg"
    
    
function MakeArray(n) 
    {
    this.length = n
    for (var i = 1; i<=n; i++) 
        {
        this[i-1] = new Image()
        }
    return this
    }

function msover(inum,d_inum) 
    {
        if ((over[inum].src != "")) 
            {
            document.images[d_inum].src = over[inum].src
            document.images[7].src = under[inum].src
            }
    }


function msout(inum,d_inum) 
    {
        if ((once[inum].src != "")) 
            {
            document.images[d_inum].src = once[inum].src
            document.images[7].src = standard[0].src
            }
    }

// -->
</script>
</head>



<body bgcolor="#FFFFFF" link="#000000" vlink="#000000" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<table border="0" bordercolor="green" bgcolor="slateblue" align="center">		<% ' Start Table 1 %>
<tr><td>

<!--#include file="../Misc/Header.INC"-->


<table border="0" bordercolor="red" cellpadding="0" cellspacing="0" align="center" width="940"> <% ' Start Table 1.1 %>

    <tr>
		<td background="../../Images/goldbackground222.jpg" width="223" valign="top" align="center">

			<table border="0" bordercolor="red" cellpadding="20" cellspacing="0">   <% ' Start Table 1.1.1 %>
				<tr>
					<td align="left" valign="top" align="center">      
						<form method="get" name="Country">
							<p align="center"><br>
							<font face="Tahoma">
							You are currently a 
							<% If Session("Country") = "Canada" Then%>
							<img src="../../Images/can1.gif" WIDTH="36" HEIGHT="18"> 
							<%elseif Session("Country") = "USA" Then %>
							<img src="../../Images/USA1.gif" WIDTH="34" HEIGHT="18"> 
							<%end if%>
							customer. Click on a country below to change countries.
							</font></p>
                                
						<p align="center"><font face="Tahoma">
						<input type="submit" name="Canada" value="Canada">
						</font>
						<font face="Tahoma">
						<input type="submit" name="  USA  " value="USA">
						</font></p>
						
						<input type="hidden" name="Area" value="<%=request("Area")%>">
						<input type="hidden" name="Manufat" value="<%=request("Manufat")%>">
						<input type="hidden" name="sar" value="<%=request("sar")%>">
						<input type="hidden" name="SID" value="<%=request("SID")%>">
						<input type="hidden" name="action" value="<%=request("action")%>">
						</form>
					</td>
				</tr>
			</table>	<% ' End Table 1.1.1 %>

		</td>


		<td background="../../Images/bluebackground.jpg">
            
					<table border="0" cellpadding="10" cellspacing="0" align="center">  <% ' Start Table 1.1.2 %>
					<tr>
						<td valign="top" align="center">
						<% if Request("sar") = "Specialx" then %>
							<!-- #INCLUDE FILE="SPECIALK.INC" -->
						<% else%>
							<!-- #INCLUDE FILE="SPECIALK1.INC" -->
						<%end if%>
						<br><br>
												<% ' Start Table 1.1.2.1 %>
                		<table border="0" cellpadding="3" cellspacing="0" width="95%" bordercolor="#000000">
                    		<tr>
                        		<td align="center"><a href="#top"><font size="3"><strong>Back to top</strong></font></a></td>
                    		</tr>
                		</table>				<% ' End Table 1.1.2.1 %>
                
						</td>
					</tr>
					</table>		<% ' End Table 1.1.2 %>
            		
            		
		</td>
            
	</tr>

			 
</table>		<% ' End Table 1.1 %>
        
	<!--#include file="../Misc/Footer.INC"-->
	
	</td>
    </tr>



</table>    <% ' End Table 1 %>


</body>


</html>
