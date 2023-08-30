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
<title>GPS @ Star Lite Intl.: GPS, GPS Sensors, GPS engines, OEM GPS, Garmin GPS, USglobal GPS, Marine radar, Cartography</title>
<meta name="keywords" content="GPS sensors,GPS sensor,GPS engine,OEM GPS,GPS OEM,GPS board,GPS navigation,GPS15 OEM sensor,GPS16 OEM sensor,GPS16 HVS,GPS16 LVS,GPS17 HVS,GPS18 sensor,GPS 18,GPS18,GPS 18LVC,GPS system,car GPS,global positioning system,auto GPS,GPS equipment,Street Pilot GPS,WAAS GPS,GMR406,GMR 41,GMR404,GDL 30,GXM 30,GTM 12,GTM10,GTM20,GVN52,GPSMAP,bt 359 GPS sensor,GPS receivers,portable GPS,handheld GPS,ique m3,quest,Nuvi350 GPS,Nuvi 360 GPS,forerunner 201,forerunner201,br355 GPS,bu353 GPS,mr350,em 406 GPS engine,em 408 GPS engine,marine GPS,GPS receiver,GPS accessories,fish finder,sounders,transducers,navigator,GPS navigator,cartography,GPS software,GPS equipment,buy GPS,GPS now,best GPS,pda GPS,Garmin GPS,bluetooth GPS,global positioning,tracking GPS,fleet tracking GPS,GPS antenna,antenna,antennas,usglobal GPS,usglobalsat,discount GPS,GPS on sale,navigation electronics,gps tracking,gps locator,AVL,fleet tracking,fleet management,monitor fleet,passive gps,real-time gps,active antenna,teen tracking,child tracking,equipment tracking,find stolen car,find stolen auto,automotive security,car security,lojack,monitor employees,monitor drivers,find my car,find my truck,locate my car with gps,gps child locator,portable gps,covert gps,automatic vehicle locator,GPS navagation,marine networking, marine navigation,SIRF,marine radar,GPS dealers,GPS resellers,GPS business,low cost GPS">
<meta name="description" content="GPS: Full line of Garmin GPS and USglobal GPS. GPS sensors, OEM GPS, GPS antennas, GPS boards, GPS accessories, tracking GPS, bluetooth GPS, GPS network, Marine radar, fish finders and sounders.">
<meta name="author" content="Star Lite International, LLC">
<meta name="copyright" content="1994-2006 Star Lite International, LLC">
<meta name="revisit-after" content="5 days">
<meta name="distribution" content="global">
<meta name="robots" content="all,index,follow">
<meta name="rating" content="general">
<meta http-equiv="content-language" content="en">
<META name="Classification" content="GPS, GPS sensors and GPS engines, GPS Tracking, OEM GPS and navigation electronic equipment">
<meta name="DC.Title" content="Star Lite Intl.: GPS, GPS Sensors, GPS engines, OEM GPS, Marine radar electronics, Garmin GPS, USglobal GPS">
<meta name="DC.Description" content="GPS: Full line of garmin GPS, usglobal GPS, GPS sensors, oem gps, GPS receivers, gps board, gps navigation, tracking GPS, bluetooth GPS, GPS network, GPS antenna, fish finders.">
<meta name="abstract" content="Complete selection of  GPS, GPS sensors, oem GPS, cartography and GPS accessories.">
<link rel="stylesheet" type="text/css" href="../Misc/StyleSheet1.css">
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
       
<table border="0" width='900' bordercolor="green"  align='center'>		<% ' Start Table 1 %>

<tr>
	<td colspan=2>
	<% ShowHeaderMenus = "No" %>
	<!--#include virtual="Misc/Header.INC"-->
	</td>
</tr>


<tr>
	<td valign='top'>
	<% ' MyStyle="'background-color: rgb(204, 255, 255);'" 
	   MyStyle="'background-color: rgb(208, 226, 246);'"
	%>
	<table style=<%=MyStyle%> border="0" width="180">
	<tr>
		<td>
		GPS Accessories    
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=171&sar=&area=Accessories'>
		GPS A/C Adapters
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=211&sar=&area=Accessories'>
		GPS Accessory Kits
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=168&sar=&area=Accessories'>
		GPS Antennas
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=208&sar=&area=Accessories'>
		GPS Battery Packs
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=165&sar=&area=Accessories'>
		GPS Data Cards (Blank)
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=165&sar=&area=Accessories'>
		GPS Data Card Programmer
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp; 
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=166&sar=&area=Accessories'>
		GPS MapSource Software
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp; 
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=166&sar=&area=Accessories'>
		Marine Blue Charts
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp; 
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=166&sar=&area=Accessories'>
		Marine Cartography
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp; 
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=191&sar=&area=Accessories'>
		Cartography Unlock Certificates
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp; 
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=261&sar=&area=Accessories'>
		GPS Preprogrammed Data Cards
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp; 
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=240&sar=&area=Accessories'>
		GPS Microphones / HeadSets
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=243&sar=&area=Accessories'>
		Miscellaneous GPS items
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=234&sar=&area=Accessories'>
		GPS Network Accessories
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=254&sar=&area=Accessories'>
		GPS Traffic Receivers
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=199&sar=&area=Accessories'>
		Marine Transducers
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=162&sar=&area=Accessories'>
		GPS Mounts
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=163&sar=&area=Accessories'>
		GPS Cables & Adapters
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=164&sar=&area=Accessories'>
		Carrying Cases
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		&nbsp;&nbsp;&nbsp;
		<font size=1>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=207&sar=&area=Accessories'>
		GPS Instructional Videos
		</font>
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=220&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS Blue Tooth
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=77&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS Fish Finders
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=77&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS Sounders
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=73&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS Fixed Mount
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=73&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		 Chart Plotters
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=72&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS Handheld
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=72&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS Portable
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=72&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS / PDA Combo
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=173&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS Sensors
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=173&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS OEM
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=173&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS OEM Boards
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=173&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS OEM TracPaks
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=193&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		GPS 2-Way Radios
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=237&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		Marine Radar
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=221&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		Marine Weather Receivers
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=219&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		Tracking GPS
		</a>
		</td>
	</tr>
	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=203&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		Software Upgrades
		</a>
		</td>
	</tr>

	<tr>
		<td>
		<a href='http://www.starlite-intl.com/scart/scart.asp?sid=197&sar=&area=Navigation%3A+GPS%2C+Sensors%2C+OEM%2C+FishFinders%2C+CDs.'>
		Rebates for GPS
		</a>
		</td>
	</tr>
	
	
	</table>
	
	</td>

	<td valign='top'> 
		<table style=<%=MyStyle%> border=0 align=center width='100%'>
		<tr>
			<td>
			<center>
			<font size=2 color=navy>
			<b>
			Serving: Businesses, Government, Educational Institutions and the General Public.
			</b>
			</font>
			</center>
			<font size=1>
			The source for GPS, GPS sensors, OEM GPS, GPS OEM boards, tracking GPS, GPS / PDA combo, GPS antennas,
			GPS accessories, Fixed and portable GPS, Fish Finders, Sounders, Marine Radar, Cartography, Chart Plotters, 
			Marine networking, GPS - 2-way radios and much more.
			Select a GPS OEM sensor, GPS OEM Engine board for your aviation or any OEM application.
			</font>
			</td>			
		</tr>
		</table>                
	<% 
	MaxNumRows = 6
	MaxNumCols = 5
	Products =	"1705,1776,1389,1602,1827," & _
				"1213,1382,1319,1479,1354," & _
				"1233,1599,1025,1488,1472," & _
				"1526,1481,1766,1277,1530," & _
				"540,1386,1507,1830,1765," & _
				"1167,1616,1509,1539,457"
	Product = split(Products, ",")
	row = 1
	Set RS = CreateObject("ADODB.Recordset")
	Response.Write "<table border=1 align='center'>"
	Do While row <= MaxNumRows
		Response.Write "<tr>"
		col = 1
		Do While col <= MaxNumCols
		Response.Write "<td valign='top'>"
		'Response.Write row & "-" & col & "-"
		pid = Product((row-1) * MaxNumCols + col - 1)
		'Response.Write pid & " "
		SQL = "SELECT * FROM Product WHERE pid = " & pid
		'Response.Write SQL
		RS.Open SQL, "DSN=STAREC1" , 1, 4
		Pic1URL = "http://www.starlite-intl.com/imi/" & RS("Pic1")
		TargetURL = "http://www.starlite-intl.com/Detail.asp?pid=" & pid
		Descr = RS("Descr")
		ProductName = RS("PName")
		NewProduct = RS("NewProduct")
		length = len(Descr)
		'Response.Write Pic1URL
		'Response.Write "<img src='" & Pic1URL & "' style='border: 0px solid ; width: 67px; height: 57px;' align='left' hspace='5'>"
		Response.Write "<a href='" & TargetURL & "'>"
		Response.Write "<font color='navy' style='text-decoration:none' size=1><b>" & ProductName & "</b></font>"	
		Response.Write "<img src='" & Pic1URL & "' style='border: 0px solid ; width: 67px;' align='left' hspace='5'>"
		Response.Write "</a>"
		If NewProduct Then
			NewIcon = "http://www.starlite-intl.com/imi/new1.gif"
			Response.Write "<img src='" & NewIcon & "' style='border: 0px solid ;' align='left' hspace='5'>"
		End If
		If length <= 120 Then
			Response.Write "<font size=1>&nbsp;" & Descr & "</font>"
		Else
			length = 120
			Response.Write "<font size=1>&nbsp;" & left(Descr,length) & "&nbsp; ..." & "</font>"
		End If
		RS.Close 
		Response.Write "</td>"
		col = col + 1
		Loop
		Response.Write "</tr>"
	row = row + 1
	Loop
	Response.Write "</table>"
	%>  
</td>
</tr>                      					


<tr>
	<td colspan=2>
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