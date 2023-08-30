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
<link rel="stylesheet" type="text/css" href="../../Misc/StyleSheet1.css">
<link rel="shortcut icon" href="../favicon.ico" TYPE="image/ico">
<LINK REL="SHORTCUT ICON" HREF="../favicon.ico">



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
       

<table border="0" width='900' bordercolor="green"  align='center'>		<% ' Start Table 1 %>
<tr><td>

<!--#include virtual="Misc/Header.INC"-->

<table border="0" bordercolor="red" cellpadding="0" cellspacing="0" align="center" width="900"> <% ' Start Table 1.1 %>
            
            <tr>
				<td>
					<table border="0" cellpadding="10" cellspacing="0" align="center">  <% ' Start Table 1.1.5 %>
					<tr>
						
						<td valign="top" align="center">

						<!--#include virtual="INC/BANNER.INC"-->
                                
                      
<% 
MaxNumRows = 2
MaxNumCols = 5
Products = "1705,1776,1389,1602,1438,1213,1382,1319,1479,1354" 
Product = split(Products, ",")
row = 1
Set RS = CreateObject("ADODB.Recordset")
Response.Write "<table border=1 align='center'>"
Do While row <= MaxNumRows
	Response.Write "<tr>"
	col = 1
	Do While col <= MaxNumCols
	Response.Write "<td>"
	'Response.Write row & "-" & col & "-"
	pid = Product((row-1) * MaxNumCols + col - 1)
	'Response.Write pid & " "
	SQL = "SELECT * FROM Product WHERE pid = " & pid
	'Response.Write SQL
	RS.Open SQL, "DSN=STAREC1" , 1, 4
	Pic1URL = "http://www.starlite-intl.com/imi/" & RS("Pic1")
	'Response.Write Pic1URL
	'Response.Write "<img src='" & Pic1URL & "' style='border: 0px solid ; width: 67px; height: 57px;' align='left' hspace='5'>"
	Response.Write "<img src='" & Pic1URL & "' style='border: 0px solid ;' align='left' hspace='5'>"
	Response.Write "<font size=1>" & RS("Descr") & "</font>"
	RS.Close 
	Response.Write "</td>"
	col = col + 1
	Loop
	Response.Write "</tr>"
row = row + 1
Loop
Response.Write "</table>"
%>  
							

<table style="text-align: left; width: 740px; height: 781px;" border="1" cellspacing="0">

<tr>

<td style="height: 102px; width: 175px;"><p align="center">
<font face="Arial"><img style="border: 0px solid ; width: 100px; height: 89px;" src="http://www.starlite-intl.com/imi/authDealer.gif" title="" alt="Garmin Authorized Dealer" align="middle" height="89" hspace="10" width="100"></font></p>
</td>

<td colspan="3" bgcolor="#000080" height="102" width="510">
<p align="center"><span style="color: rgb(255, 255, 255);"><font face="Arial" size="3">We
offer a <span style="font-weight: bold;">full line</span> of GPS products: </font></span></p>
<p align="center"><font face="Arial"><span style="color: rgb(255, 255, 255);"><span style="color: rgb(255, 255, 255);"><font size="2">GPS
sensors, tracking GPS, OEM GPS, GPS PDA, GPS accessories, </font></span><font size="2"><big style="color: rgb(255, 255, 255);">
</big></font></span><font size="3"><span style="color: rgb(255, 255, 51);"><big style="font-weight: bold;"><big style="color: rgb(255, 255, 255);"><img title="" style="width: 235px; height: 20px;" alt="Scroll Down" src="http://www.starlite-intl.com/imi/scroll-down.gif" align="top" height="20" width="235"><br>
</big></big></span><big style="color: rgb(255, 255, 255);"><big style="color: rgb(255, 255, 255);">Featured Products</big>
</big></font><big style="color: rgb(255, 255, 255);"><font size="2">(click on
a picture for details)</font><font size="3"><b><br>
</b></font></big><big style="background-color: rgb(255, 255, 0);"><font color="#000080" size="1">
<a href="http://www.starlite-intl.com/scart/scart.asp?sar=Rebates%20for%20GPS%20units&amp;area=Navigation:%20GPS,%20Sensors,%20OEM,%20FishFinders,%20CDs&amp;sid=197">Click
here to check for available REBATES.</a></font></big></font></p>
</td>

<td style="width: 160px; height: 102px;">
<p align="center"><font face="Arial"><img style="border: 0px solid ; width: 120px; height: 27px;" src="http://www.starlite-intl.com/imi/Auth_Dealer_USGLogo.jpg" title="" alt="USglobal Authorized Dealer" height="34" width="150"></font></p>
</td>

</tr>


<tr>
<td style="height: 153px; vertical-align: top; width: 175px; font-family: helvetica,arial,sans-serif;">
<a href="http://www.starlite-intl.com/Detail.asp?pid=1705&amp;Key=">
<img src="http://www.starlite-intl.com/imi/nuvitt.jpg" alt="StreetPilot 7200" style="border: 0px solid ; width: 67px; height: 57px;" align="left" hspace="1">
</a>
<a href="http://www.starlite-intl.com/Detail.asp?pid=1705&amp;Key="><font size="2"><font size="2">
<img src="http://www.starlite-intl.com/imi/new1.gif" alt="" style="border: 0px solid ; width: 31px; height: 31px;" align="top">
</font></font></a><font size="2">
<font size="2">Garmin,</font></font>
<br>

<a style="font-weight: bold;" href="http://www.starlite-intl.com/Detail.asp?pid=1705&amp;Key=">
<font size="2"><font size="2">Nuvi 360</font></font></a><font size="2"> Translator. Entertainer. Tour Guide.
Pocket-sized Personal Travel Assistant comes with hands-free
<span style="color: rgb(0, 0, 153); font-weight: bold;"><span style="color: rgb(51, 51, 255);">Bluetooth</span> </span><span style="color: rgb(0, 0, 153);"><span style="color: rgb(0, 0, 0);">GPS, touch screen, voice prompt navigation</span></span></font><font size="2"><strong><font size="2">
</font></strong></font>
</td>

<td style="width: 173px; height: 153px; text-align: left; vertical-align: top; font-family: helvetica,arial,sans-serif;">
<font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1776">
<img alt="StreetPilot c530" src="http://www.starlite-intl.com/imi/spc530tt.gif" style="border: 0px solid ; width: 70px; height: 63px;" align="left" hspace="1"></a></font><font size="2"><font size="2"><img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" height="31" width="31"></font></font><font size="2">Garmin, GPS <span style="font-weight: bold;">StreetPilot c530</span>. A portable automotive navigator. Preloaded with maps<small> </small></font><small>and plenty of options.<font size="2"><small>.<big> Voice,</big> </small></font>touch-screen</small>
</td>

<td style="height: 153px; vertical-align: top; font-family: helvetica,arial,sans-serif;" width="156"><font size="2">
<a href="http://www.starlite-intl.com/Detail.asp?pid=1389&amp;Key=">
<img title="" style="border: 0px solid ; width: 50px; height: 80px;" alt="GPS10 Deluxe" src="http://www.starlite-intl.com/imi/gps10tt.jpg" align="left" height="80" width="50">
</a></font>
<p><font size="2">Garmin, <span style="font-weight: bold;">GPS10
Deluxe</span> <span style="font-weight: bold; color: rgb(51, 51, 255);">Bluetooth
GPS</span> enabled, GPS receiver. Waterproof. Voice prompt navigation</font></p>
</td>

<td style="vertical-align: top; width: 186px; font-family: helvetica,arial,sans-serif;" height="153">
<font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1602&amp;Key="><span style="text-decoration: underline;"><img src="http://www.starlite-intl.com/imi/FR350tt.jpg" title="" alt="GPSMAP 60CS" style="border: 0px solid ; width: 55px; height: 66px;" align="left"></span></a><b> </b></font><font size="2"><img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" height="31" width="31"></font><small>Garmin,</small><font size="2"><b> Forerunner 305 </b></font><small>personal trainer and&nbsp; high sensitivity GPS. For
athletes of all levels. See also Forerunner <a href="http://www.starlite-intl.com/Detail.asp?pid=1382">301</a><small>&nbsp;<big> </big></small></small><font size="2"> </font>
</td>

<td style="width: 161px; height: 153px; vertical-align: top; font-family: helvetica,arial,sans-serif;">
<font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1438&amp;Key="><img title="" style="border: 0px solid ; width: 50px; height: 75px;" alt="bt338" src="http://www.starlite-intl.com/imi//BT338tt.jpg" align="left" hspace="1" vspace="1"></a>
USGlobalsat,&nbsp; <span style="font-weight: bold;">BT338&nbsp;</span> <span style="font-weight: bold; color: rgb(0, 0, 153);"><span style="color: rgb(51, 51, 255);">Bluetooth&nbsp;</span>
</span><span style="font-weight: bold; color: rgb(51, 51, 255);">GPS</span>
receiver. High sensitivity, fast satellite&nbsp; acquisition.</font>
</td>

</tr>

<tr>

<td style="vertical-align: top; height: 155px; width: 175px; font-family: helvetica,arial,sans-serif;">
<font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1213&amp;Key=">
<img title="" style="border: 0px solid ; width: 50px; height: 75px;" alt="GPS18USB" src="http://www.starlite-intl.com/imi/gps18usbtt.jpg" align="left" height="75" width="50"></a>Garmin,
<span style="font-weight: bold;">GPS18 USB </span>For a variety of OEM
applications. This GPS sensor receives power and data via a single USB
connection!</font>
</td>

<td style="vertical-align: top; height: 155px; font-family: helvetica,arial,sans-serif;" width="173">
<font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1382&amp;Key="><img title="" style="border: 0px solid ; width: 50px; height: 57px;" alt="Forerunner301" src="http://www.starlite-intl.com/imi/forerunner301tt.jpg" align="left"></a><img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" align="left">Garmin
Training partner. <span style="font-weight: bold;">Forerunner301</span>,
This </font><font size="2">GPS sensor</font><font size="2"> continuously monitors heart rate, speed, distance, pace and calories
burned<br>
</font>
</td>

<td style="vertical-align: top; height: 155px; font-family: helvetica,arial,sans-serif;" width="156">
<font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1319&amp;Key="><img title="" style="border: 0px solid ; width: 50px; height: 75px;" alt="GPS18 Deluxe" src="http://www.starlite-intl.com/imi/gps18usbtt.jpg" align="left" height="75" width="50"></a>Garmin,<span style="font-weight: bold;">
GPS18 Deluxe USB</span> <big><small>GPS </small> sensor</big> with nRoute™ software for
turn-by-turn directions, voice prompts</font>
</td>

<td style="vertical-align: top; width: 186px; height: 155px; font-family: helvetica,arial,sans-serif;">
<a href="http://www.starlite-intl.com/Detail.asp?pid=1479&amp;Key=">
<font size="2"><img title="" style="border-style: solid; border-width: 0px; width: 52px; height: 52px;" alt="EM406" src="http://www.starlite-intl.com/imi/EM-406tt.jpg" align="left"></font></a><font size="2"><img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" align="left">USGlobalsat,
<span style="font-weight: bold;">EM406</span> SiRFIII GPS Engine Board for
OEM application. SiRF Star III GPS Chipset. Extremely
fast TTFF.</font>
</td>

<td style="width: 161px; vertical-align: top; height: 155px; font-family: helvetica,arial,sans-serif;">
<font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1354&amp;Key=">
<img title="" style="border: 0px solid ; width: 47px; height: 80px;" alt="MR350" src="http://www.starlite-intl.com/imi/MR-350tt.jpg" align="left" height="75" width="50"></a><span style="font-weight: bold;">MR350
</span>is a&nbsp; “mini” external&nbsp; weatherproof&nbsp; GPS receiver&nbsp; by&nbsp; <span style="font-weight: bold;"></span>USGlobalsat. For a more permanent mounting solution in&nbsp; OEM
applications. </font><span><font size="2">SiRF Star III</font></span><font size="2">
chip set</font>
</td>

</tr>


<tr>

<td style="vertical-align: top; height: 147px; width: 175px; font-family: helvetica,arial,sans-serif;"><font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1233&amp;Key="><img title="" style="border: 0px solid ; width: 50px; height: 75px;" alt="GPS18PC" src="http://www.starlite-intl.com/imi/gps18pctt.jpg" align="left" height="75" width="50"></a>Garmin,<span style="font-weight: bold;">
GPS18 PC</span>, GPS receiver sensor with DB-9 pin serial connector with
12-volt cigarette lighter adapter</font>
</td>

<td style="vertical-align: top; height: 147px; font-family: helvetica,arial,sans-serif;" width="173">
<a href="http://www.starlite-intl.com/Detail.asp?pid=1599">
<img alt="Forerunner201" src="http://www.starlite-intl.com/imi/edge305t.jpg" style="border: 0px solid ; width: 52px; height: 57px;" align="left"></a><font size="2"><font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1705&amp;Key="><font size="2"><font size="2"><img src="http://www.starlite-intl.com/imi/new1.gif" alt="" style="border: 0px solid ; width: 31px; height: 31px;" align="top"></font></font></a></font></font><small>Garmin</small>, <small><small><font size="4"><small><small><b>Edge&nbsp; 305HR </b>with<b> </b><span style="font-weight: bold;">heart rate monitor, Speed Cadence Sensor.</span></small></small></font></small></small><span style="font-weight: bold;"> </span><small>Personal trainer and cycle computer. High sensitivity GPS ...<br>
</small>
</td>

<td style="vertical-align: top; height: 147px; font-family: helvetica,arial,sans-serif;" width="156"><a href="http://www.starlite-intl.com/Detail.asp?pid=543&amp;Key="><img title="" style="border: 0px solid ; width: 70px; height: 50px;" alt="GPS35HVS" src="http://www.starlite-intl.com/imi/gps35hvstt.jpg" align="left" height="50" hspace="1" vspace="1" width="70"></a><font size="2">Garmin,
<span style="font-weight: bold;">GPS</span> <span style="font-weight: bold;">35HVS</span>
OEM GPS sensor, High voltage supply 6 to 40 VDC, RS-232, 5 m cable with bare wires</font>
</td>

<td style="vertical-align: top; width: 186px; height: 147px; font-family: helvetica,arial,sans-serif;"><font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1488&amp;Key="><img title="" style="border-style: solid; border-width: 0px; width: 50px; height: 80px;" alt="iQue M3" src="http://www.starlite-intl.com/imi/iQueM3tt.gif" align="left" height="80" width="50"></a>G<font size="2">armin,<span style="font-weight: bold;">
iQue M3 </span></font><small><big>easy in-car GPS navigation with
PDA Pocket PC applications. </big><span style="color: rgb(0, 0, 0);"><big>Microsoft’s Windows Mobile 2003 Pro.<br>
</big></span></small>
</font>
</td>

<td style="vertical-align: top; height: 147px; font-family: helvetica,arial,sans-serif;" width="161"><font size="2">
<a href="http://www.starlite-intl.com/Detail.asp?pid=1278&amp;Key=">
<img title="" style="border: 0px solid ; width: 50px; height: 75px;" alt="BU303" src="http://www.starlite-intl.com/imi/bu303tt.jpg" align="left" height="75" width="50">
</a></font>
<img src="http://www.starlite-intl.com/imi/Sale.jpg">
<br>
<br>
<font size="2">USGlobalsat,<b>
BU303 </b><span id="lblProductOverview">GPS Receiver. Connects to your
laptop USB port</span></font>
</td>

</tr>


<tr>
<td style="vertical-align: top; height: 175px; width: 175px; font-family: helvetica,arial,sans-serif;">
<a href="http://www.starlite-intl.com/Detail.asp?pid=1526"><font size="2">
<img alt="streetpilot i5" src="http://www.starlite-intl.com/imi/SPi5t.jpg" style="border: 0px solid ; width: 47px; height: 47px;" align="left" hspace="1"></font></a><font size="2">
<img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" height="31" width="31">Garmin, GPS navigator </font>
<font><font size="2"><font style="font-family: helvetica,arial,sans-serif;" size="2">
<font size="2"><span style="font-weight: bold;">StreetPilot i5</span></font>Preloaded
with detailed maps of North America<font size="2">, </font>automatic route calculation. include 3-D map
graphics, provides voice-prompted turn-by-turn directions through a
built-in speaker</font></font></font><font size="2">
</font>
<font style="font-family: helvetica,arial,sans-serif;" size="2"><p><font size="2"><span style="font-weight: bold;"></span></font>
</p></font>
</td>

<td style="vertical-align: top; height: 175px; font-family: helvetica,arial,sans-serif;" width="173"><span style="text-decoration: underline;"><a href="http://www.starlite-intl.com/Detail.asp?pid=1481&amp;Key="><img src="http://www.starlite-intl.com/imi/quest2tt.jpg" title="" alt="Quest 2" style="border: 0px solid ; width: 46px; height: 42px;" align="left" hspace="1" vspace="1"></a></span><font size="2"><img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" align="left" height="31" width="31"><br>
Garmin, <span style="font-weight: bold;">Quest 2. </span>This<span style="font-weight: bold;"> GPS</span> is
pre-loaded with City Select North America NT, full coverage of the
entire U.S.A., Canada, and Puerto Rico, nearly six million points
of interest</font>
</td>

<td style="vertical-align: top; height: 175px; font-family: helvetica,arial,sans-serif;" width="156"><font size="2"><a href="http://www.starlite-intl.com/Detail.asp?pid=1486&amp;Key="><img title="" style="border: 0px solid ; width: 65px; height: 57px;" alt="StreetPilot2720" src="http://www.starlite-intl.com/imi/sp2720tt.jpg" align="left" height="57" width="65"></a><img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" height="31" width="31">Garmin,
</font><small><font size="3"><small><b>StreetPilot 2720 </b></small></font></small><font size="2">a
premium GPS automotive navigator. Offers text-to-speech and traffic
interface capabilities. touch screen display</font>
</td>

<td style="vertical-align: top; width: 186px; height: 175px; font-family: helvetica,arial,sans-serif;"><a href="http://www.starlite-intl.com/Detail.asp?pid=1277&amp;Key="><img title="" style="border: 0px solid ; width: 50px; height: 75px;" alt="GPSMAP3010C" src="http://www.starlite-intl.com/imi/gpsmap3010ctt.jpg" align="left" height="75" width="50"></a><font size="2"><img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" height="31" width="31">
Garmin, <span style="font-weight: bold;">GPSMAP 3010C</span> get into the
world of Marine networking, with plug-and-play systems puts GPS,
weather, sonar, radar, and other important data at boaters’...</font>
</td>

<td style="width: 161px; vertical-align: top; height: 175px; font-family: helvetica,arial,sans-serif;">
<a href="http://www.starlite-intl.com/Detail.asp?pid=1530&amp;Key="><font size="2"><img title="" style="border: 0px solid ; width: 67px; height: 58px;" alt="Nuvi350" src="http://www.starlite-intl.com/imi/nuvitt.jpg" align="left" height="58" width="67"></font></a><font size="2"><img title="" style="width: 31px; height: 31px;" alt="" src="http://www.starlite-intl.com/imi/new1.gif" align="left">Garmin,&nbsp;</font>
<p><font size="2"><span style="font-weight: bold;">n&uuml;vi
350</span> is a portable&nbsp; GPS navigator, traveler’s reference, digital
entertainment system, automatic routing,
turn- by-turn&nbsp; voice directions, finger-touch- screen ...</font></p>
</td>

</tr>

</table>
                        
						
						
					</td>
					</tr>

					</table>		<% ' End Table 1.1.5 %>
			
				</td>
           
			</tr>
                        
            
</table>	<% ' End Table 1.1 %>


<!--#include file="../Misc/Footer.INC"-->

      
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