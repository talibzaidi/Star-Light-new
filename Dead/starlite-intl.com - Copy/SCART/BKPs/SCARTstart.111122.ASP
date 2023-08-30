<%@ LANGUAGE = VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



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
<link rel="stylesheet" type="text/css" href="/Misc/StyleSheet1.css">

<title>Star Lite Intl.: GPS, GPS Sensors, GPS engines, Garmin GPS, USglobal GPS, CB radios, 2-way radios, Scanners, Marine electronics, Audio, Video, Hand tools</title>
<meta name="keywords" content="GPS, GPS sensors, GPS sensor, GPS engine, global positioning system, GPS navigation, oem GPS, GPS oem, GPS18 sensor, GPS 18, GPS18, GPS18lvc, GPS system, car GPS, auto GPS, GPS equipment, differential GPS, WAAS, GPS receivers, buy GPS, GPS now, streetpilot, portable GPS, handheld GPS, ique m3, quest, streetpilot i5, forerunner 201, forerunner201, bu303, bu353, mr350, gv101, marine GPS, GPS receiver, GPS accessories, fish finder, navigator, GPS navigator, GPS software, GPS equipment, best GPS, pda GPS, garmin GPS, bluetooth GPS, global positioning, tracking GPS, fleet tracking GPS, GPS antenna, antenna, antennas, usglobal GPS, discount GPS, GPS on sale, navigation electronics, CB radio, CB radios, CB, CB on sale, 2-way radios, two way radios, walky talky, marine radio, marine electronics, radar, network, uniden CB, cobra CB, Midland CB, CB antennas, cb antenna, amateur radios, galaxy radio, magnum radio, radio scanner, radio scanners, scanner, flash memory, secure digital, compact flash, digital cameras, car audio, car stereos, stereo, car video, power amplifier, scanner antennas,scanner antenna, power supplies, power supply, regulated power supplies, DJ equipment, DJ, accessories, hand tools, mechanics tools, MIT tools, pyramid, pyle, silicon power, solarcon, fuji, fujifilm, nikon, olympus, panasonic, motorola, ranger, cherokee, wilson, firestick, valor, workman, RF, Limited, Solarcon, solarcon, Antron, antron, A99, gpk1, B100, saturn, 82fl, gzero48, gzero, little wil, w1000, w5000, a99ck, bm921, bm922, International, coax, RG8, K40, k40 antenna, K-40, Maco, Shakespeare, Imax, star, starlite, inverters, microphone, microphones, SSB, AM, FM, roger, beep, 40channels, 40 channels, car, system, jvc, speaker, speakers, stereo, amplifier, amplifiers, CD, player, wholesale, gps wholesale, new, new products, stereo, ham, ham radio, amateur, citizens, band, radar, detector, detectors, sales, international, international sales, export, power, supply, waterproof, rechargable, nicad, nimh, battery, gtl, ltd, dx, fm, ssb, cw, clarifier, bearcat, bearcat products, power, supply, rx, tx, transmit, transmitter, receive, receiver, receivers, SWR, frequency, counter, 10 meter, 11 meter, meters, citizens band, citizen-band, band, amateur-radio, RF limited, hustler, pioneer, discount, inexpensive, cheap, sony, caramplifier, caramplifiers, linear amplifier, amps, mobile, radio, radios, beam antenna, m103c, beams">
<meta name="description" content="GPS: Full line of garmin GPS, usglobal GPS, GPS sensors, oem gps, GPS receivers, gps board, gps navigation, tracking GPS, bluetooth GPS, GPS network, fish finders. CB: cb radios, walky-talky, cb antennas, cb accessories, 2-way radios, scanners. Digital: digital cameras, flash memory. car audio and video, DJ, hand tools, mechanics tools.">
<meta name="author" content="Star Lite International, LLC">
<meta name="copyright" content="2005 Star Lite International, LLC">
<meta name="revisit-after" content="10 days">
<meta name="distribution" content="global">
<meta name="robots" content="index,follow">
<meta name="rating" content="general">
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
<meta http-equiv="content-language" content="en">
<meta name="DC.Title" content="Star Lite Intl.: GPS, GPS Sensors, GPS engines, Garmin GPS, USglobal GPS, CB radios, 2-way radios, Amateur radios, Scanners, Marine electronics, Audio/video, Hand tools">
<meta name="DC.Description" content="GPS: Full line of garmin GPS, usglobal GPS, GPS sensors, oem gps, GPS receivers, gps board, gps navigation, tracking GPS, bluetooth GPS, GPS network, fish finders. CB: cb radios, walky-talky, cb antennas, cb accessories, 2-way radios, scanners. Digital: digital cameras, flash memory. car audio and video, DJ, hand tools, mechanics tools.">
<meta name="abstract" content="Complete selection of  GPS, GPS sensors, oem GPS and GPS accessories. Also wide selection of cb radios, scanners, antennas and automotive electronics, digital cameras, memory cards and hand tools.">


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




<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">


<% InArea = "Specials" %>

<!--#include file="../Misc/Header.INC"-->


<table style="border-right:1px solid #84bff1;" width='1120' cellpadding="0" cellspacing="0" align="center" >   <% ' Start Table 1 %>
    <tr>
     	<td class="Gradient2" width="223" valign="top" align="left">
     	
     	<!--#include virtual="INC/LeftColumn.inc.asp"-->

		<form method="get" name="Country">		<% ' THis form was inherited from old version of this page, before I started INCLUDE-ing file INC/LeftColumn.inc.asp. I don't know if this form stuff is still needed or not. %>					
			<input type="hidden" name="Area" value="<%=request("Area")%>">
			<input type="hidden" name="Manufat" value="<%=request("Manufat")%>">
			<input type="hidden" name="sar" value="<%=request("sar")%>">
			<input type="hidden" name="SID" value="<%=request("SID")%>">
			<input type="hidden" name="action" value="<%=request("action")%>">
			<input type="hidden" name="Manufa" value="<%=request("Manufa")%>">
		</form>

      	</td>
								
	    <td valign='top'>
	    
	    			<% ' Start Table 1.1.2 %>
					<table border="0" cellpadding="10" cellspacing="0" align="center">  
					<tr>
						<td align=center>
						<font face=Tahoma size=4 color="#B90000"><b>SPECIALS - LIMITED TIME ONLY</b></font>
						</td>
					</tr>
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
					
					
            <% ' Start Table 1.1.2 %>
		    <table border="0" cellpadding="10" cellspacing="0" align="center">  
		    <tr>
			    <td valign="top" align="center">
			    <!--#include file="BANNER.INC"-->
			    <!--#include file="SCART.INC"-->
			    </td>
		    </tr>
		    </table>		
            <% ' End Table 1.1.2 %>	
            

        </td>
	</tr>

    <tr>
        <td colspan="2">
        	<!--#include file="../Misc/Footer.INC"-->
        </td>
    </tr>

</table>   
<% ' End Table 1 %>



</body>


</html>
