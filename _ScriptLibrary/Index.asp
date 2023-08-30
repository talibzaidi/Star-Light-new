<%@ Language=VBScript%>


<SCRIPT id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</SCRIPT>


<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>




<% response.buffer=true %> 

<% spec = 2 %> 

<% sar = "Home"
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
<title>Star Lite Intl.: GPS, GPS Sensors, GPS engines, Garmin GPS, USglobal GPS, CB radios, 2-way radios, Scanners, Marine electronics, Audio, Video, Hand tools</title>
<meta name="keywords" content="GPS, GPS sensors, GPS sensor, GPS engine, global positioning system, GPS navigation, oem GPS, GPS oem, GPS18 sensor, GPS 18, GPS18, GPS18lvc, GPS system, car GPS, auto GPS, GPS equipment, differential GPS, WAAS, GPS receivers, buy GPS, GPS now, streetpilot, portable GPS, handheld GPS, ique m3, quest, streetpilot i5, forerunner 201, forerunner201, bu303, bu353, mr350, gv101, marine GPS, GPS receiver, GPS accessories, fish finder, navigator, GPS navigator, GPS software, GPS equipment, best GPS, pda GPS, garmin GPS, bluetooth GPS, global positioning, tracking GPS, fleet tracking GPS, GPS antenna, antenna, antennas, usglobal GPS, discount GPS, GPS on sale, navigation electronics,  gps tracking, gps locator, avl, fleet tracking, fleet management, monitor fleet, passive gps, real-time gps, teen tracking, child tracking, equipment tracking, find stolen car, find stolen auto, automotive security, car security, lojack, monitor employees, monitor drivers, find my car, find my truck, locate my car with gps, gps child locator, portable gps, covert gps, automatic vehicle locator, gps navagation, gps dealers, gps resellers, gps business, resell gps, resell tracking gps, low cost gps, CB radio, CB radios, CB, CB on sale, 2-way radios, two way radios, walky talky, marine radio, marine electronics, radar, network, uniden CB, cobra CB, Midland CB, CB antennas, cb antenna, amateur radios, galaxy radio, magnum radio, radio scanner, radio scanners, scanner, flash memory, secure digital, compact flash, digital cameras, car audio, car stereos, stereo, car video, power amplifier, scanner antennas,scanner antenna, power supplies, power supply, regulated power supplies, DJ equipment, DJ, accessories, hand tools, mechanics tools, MIT tools, pyramid, pyle, silicon power, solarcon, fuji, fujifilm, nikon, olympus, panasonic, motorola, ranger, cherokee, wilson, firestick, valor, workman, RF, Limited, Solarcon, solarcon, Antron, antron, A99, gpk1, B100, saturn, 82fl, gzero48, gzero, little wil, w1000, w5000, a99ck, bm921, bm922, International, coax, RG8, K40, k40 antenna, K-40, Maco, Shakespeare, Imax, star, starlite, inverters, microphone, microphones, SSB, AM, FM, roger, beep, 40channels, 40 channels, car, system, jvc, speaker, speakers, stereo, amplifier, amplifiers, CD, player, wholesale, gps wholesale, new, new products, stereo, ham, ham radio, amateur, citizens, band, radar, detector, detectors, sales, international, international sales, export, power, supply, waterproof, rechargable, nicad, nimh, battery, gtl, ltd, dx, fm, ssb, cw, clarifier, bearcat, bearcat products, power, supply, rx, tx, transmit, transmitter, receive, receiver, receivers, SWR, frequency, counter, 10 meter, 11 meter, meters, citizens band, citizen-band, band, amateur-radio, RF limited, hustler, pioneer, discount, inexpensive, cheap, sony, caramplifier, caramplifiers, linear amplifier, amps, mobile, radio, radios, beam antenna, m103c, beams">
<meta name="description" content="GPS: Full line of garmin GPS, usglobal GPS, GPS sensors, oem gps, GPS receivers, gps board, gps navigation, tracking GPS, bluetooth GPS, GPS network, fish finders. CB: cb radios, walky-talky, cb antennas, cb accessories, 2-way radios, scanners. Digital: digital cameras, flash memory. car audio and video, DJ, hand tools, mechanics tools.">
<meta name="author" content="Star Lite International, LLC">
<meta name="copyright" content="1994-2005 Star Lite International, LLC">
<meta name="revisit-after" content="10 days">
<meta name="distribution" content="global">
<meta name="robots" content="all,index,follow">
<meta name="rating" content="general">
<meta http-equiv="content-language" content="en">
<META name="Classification" content="GPS, GPS sensors and GPS engines, GPS Tracking, communication and electronic equipment">
<meta name="DC.Title" content="Star Lite Intl.: GPS, GPS Sensors, GPS engines, Garmin GPS, USglobal GPS, CB radios, 2-way radios, Amateur radios, Scanners, Marine electronics, Audio/video, Hand tools">
<meta name="DC.Description" content="GPS: Full line of garmin GPS, usglobal GPS, GPS sensors, oem gps, GPS receivers, gps board, gps navigation, tracking GPS, bluetooth GPS, GPS network, GPS antenna, fish finders. CB: cb radios, walky-talky, cb antennas, cb accessories, 2-way radios, scanners. Digital: digital cameras, flash memory. car audio and video, DJ, hand tools, mechanics tools.">
<meta name="abstract" content="Complete selection of  GPS, GPS sensors, oem GPS and GPS accessories. Also wide selection of cb radios, scanners, antennas and automotive electronics, digital cameras, memory cards and hand tools.">
<link rel="stylesheet" href="style_new.css" type="text/css">
<link rel="shortcut icon" href="favicon.ico" TYPE="image/ico">
<LINK REL="SHORTCUT ICON" HREF="favicon.ico">



<script language="Javascript"><!--
once = new MakeArray(6)
    over = new MakeArray(6)
    under = new MakeArray(6)
    standard = new MakeArray(1)
 
    once[0].src = "Images/question1.gif"
    once[1].src = "Images/scart1.gif"
    once[2].src = "Images/home1.gif"
    once[3].src = "Images/new1.gif"
    once[4].src = "Images/cat1.gif"
    once[5].src = "Images/ex1.gif"    
        
    over[0].src = "Images/question2.gif"
    over[1].src = "Images/scart2.gif"
    over[2].src = "Images/home2.gif"
    over[3].src = "Images/new2.gif"
    over[4].src = "Images/cat2.gif"
    over[5].src = "Images/ex2.gif"

    under[0].src = "Images/helpnav.gif"
    under[1].src = "Images/shoppingcartnav.gif"
    under[2].src = "Images/homenav.gif"
    under[3].src = "Images/newproductsnav.gif"
    under[4].src = "Images/onlinecataloguenav.gif"
    under[5].src = "Images/specialsnav.gif"

    standard[0].src = "Images/emptynav.jpg"


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



<!--Start Tracker Code//-->            
<script language="JavaScript"><!--
var myPage;
var myReferrer;
var subPage;
var subPage2;
var slashcount;

myReferrer=document.referrer;
myPage = location.href;
subPage = String(myPage).substring(7,myPage.length);
for(x=0;x<subPage.length;x++)
{
    if(subPage.charAt(x) == "/")
    {
    slashcount = x;
    break;
    }
}        
subPage2 = String(myPage).substring(0,slashcount+7);
subPage2 = subPage2+"/stats/record.asp?page="+myPage+"&ref="+myReferrer;
//mywindow = window.open(subPage2,'recorder','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=1,height=1');

self.focus();    
//--></script>
<!--End Tracker Code//-->

</head>




<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
       
<table border="0" width='900' bordercolor="green" bgcolor="slateblue" align='center'>		<% ' Start Table 1 %>
<tr><td>

<!--#include virtual="Misc/Header.INC"-->

<table border="0" bordercolor="red" cellpadding="0" cellspacing="0" align="center" width="900"> <% ' Start Table 1.1 %>
            
            <tr>
				<td background="Images/goldbackground222.jpg" width="223" valign="top" align="center">
				
				<table cellpadding="20">	<% ' Start of Table 1.1.4 %>
					<tr>
					<td>
						<center><img src="Images/logo.gif" WIDTH="135" HEIGHT="145"></center>
						<!--#include virtual="INC/SPECIAL.INC"-->
					</td>
					</tr>
				</table>				<% ' End Table 1.1.4 %>
				
				</td>
				
				<td background="Images/bluebackground.jpg">
					<table border="0" cellpadding="10" cellspacing="0" align="center">  <% ' Start Table 1.1.5 %>
					<tr>
						
						<td valign="top" align="center">

						<!--#include virtual="INC/BANNER.INC"-->
                                

                        <%
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							shqstring = "SELECT Text1, Text2 FROM Company "
							Set RHS = Conn.Execute(shqstring)
						%>
						</p>
                        
                        <p>	<font size="2" face="Tahoma">
							<strong>&nbsp;&nbsp;&nbsp;<%=RHS("Text1")%></strong>
							</font>
                        </p>
                        
						<p>	<font size="2" face="Tahoma">
							<strong>&nbsp;&nbsp;&nbsp;<%=RHS("Text2")%></strong>
							</font>
							
							<br>
                               
						<% rhs.close %>
                        
															<% ' Start Table 1.1.5.2 %>
							<table border="0" cellpadding="3" cellspacing="0" width="95%" bordercolor="#000000">
							<tr>
								<td align="center">
								<a href="#top">
								Back to top
								</a>
								</td>
							</tr>
							</table>						<% ' End Table 1.1.5.2 %>
						
						
					</td>
					</tr>

					</table>		<% ' End Table 1.1.5 %>
			
				</td>
           
			</tr>
                        
            
</table>	<% ' End Table 1.1 %>


<!--#include file="Misc/Footer.INC"-->

      
</td>
</tr>
</table>	<% ' End Table 1 %>

<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>


<a href="http://www.digitalpoint.com/tools/geovisitors/">
<img src="http://geo.digitalpoint.com/a.png" alt="" style="border:0">
</a>

        
</body>


<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>




</html>