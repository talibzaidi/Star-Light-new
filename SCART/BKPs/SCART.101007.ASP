<%@ LANGUAGE = VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<html>

<head>
<link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->

<title>GPS Best Source: Garmin GPS, USglobalSat GPS, Pharos GPS, GPS Sensors, GPS engines, OEM GPS, GPS boards, Marine radar, Deluo GPS, Cartography</title>
<meta name="keywords" content="GPS,GPS sensors,GPS sensor,GPS engine,OEM GPS,GPS OEM,GPS board,GPS engine boards,garmin GPS,USGlobalsat GPS,pharos GPS,Deluo GPS,GPS navigation,GPS 15,GPS 16,GPS 18,GPS17,PDA,mr350,sensor,GPS18PC,em-408,em411,bu 353,GPS 16LVS,GPS 17HVS,GPS18 sensor,GPS 18LVC,GPS systems,auto GPS,GPS equipment,StreetPilot,bluetooth GPS,BT 359,handheld GPS,EM318,ET318,ET333,fixed mount GPS,Nuvi350,Nuvi 360,Nuvi 660,forerunner 201,br355,bu353,marine GPS,GPS receivers,GPS accessories,fish finder,sounders,transducers,cartography,GPS mapping software,GMR 41,GMR404,GDL 30,GXM 30,GTM 12,GTM10,GTM20,GVN52,GPSMAP,tracking GPS,navigation electronics,gps tracking,low cost GPS,GPS specials,AVL,fleet tracking GPS,fleet management,passive GPS,real-time GPS,active antenna,equipment tracking,covert gps,automatic vehicle locator,GPS navagation,marine networking,marine navigation,SIRF,marine radars,Authorized GPS dealers,GPS resellers,GPS business,gps wholesale">
<meta name="description" content="The GPS source for Garmin GPS and USglobal GPS. GPS sensors, OEM GPS, GPS boards, GPS accessories, tracking GPS, bluetooth GPS, GPS network systems, GPS antennas, Marine GPS, Marine radar, Navigation Mapping software, fish finders and sounders.">
<meta name="author" content="Star Lite International, LLC">
<meta name="copyright" content="2008 Star Lite International, LLC">
<meta name="revisit-after" content="20 days">
<meta name="distribution" content="global">
<meta name="robots" content="index,follow">
<meta name="rating" content="general">
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
<meta http-equiv="content-language" content="en">
<meta name="DC.Title" content="Star Lite International, LLC">
<meta name="DC.Description" content="The best GPS source for Garmin GPS, USglobalsat GPS, Pharos GPS and Deluo GPS. GPS sensors, oem gps, GPS receivers, gps boards, tracking GPS, bluetooth GPS, GPS networks, GPS antennas, fish finders and sounders, GPS accessories, tracking GPS, Marine GPS, Marine radars">
<meta name="abstract" content="The source for  GPS, GPS sensors, oem GPS, OEM engine boards, hand held and fixed mount GPS, cartography and GPS accessories">


<% response.buffer=true %>

<% ar = Request("Area") %>
<% Area = Request("Area") %>

<% sar = ReQuest("sar") %>
<% 'sar = Replace( sar, " ", "%20") %>

<% SID = ReQuest("SID") %>

<% if Area="iii" then
   Area = Request("Manufat")
   ar = Request("Manufat")
   end if
   
   if Area = "New Products" then
     response.redirect "./scart.asp?pid=0&sid=11&area=New!&sar=New%20Products"
   end if    

   if Area="Choose a catalog area." then
   response.redirect "https://www.starlite-intl.com/index.asp"
   end if
   
   If (Request("Canada") <> "" OR Request("  USA  ") <> "") then
		If Request("Canada") <> "" then
			Session("Country") = "Canada"
		else
			Session("Country") = "USA"
		end if
	end if
%>



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



<body >


<% 
InArea = ""   ' Refers to area of website, not area of products.
	
action = request("action")
'Response.Write "<br>action = " & action 
'Response.Write "<br>Area = " & Area 

If (action = "add") OR (action = "viewcart") OR (action = "checkout") Then
	InArea = "ShoppingCart" 
ElseIf Area = "New!" Then
	InArea = "WhatsNew"
ElseIf Area = "Gift Certificates" Then
	InArea = "GiftCertificates"
End If

'Response.Write "<br>InArea = " & InArea
%>

<!-- #include file="../Misc/Header.INC" -->


<% ' Start Table 1 %>
<table style="border-right:1px solid #84bff1;" width='1120' cellpadding="0" cellspacing="0" align="center" >   
    <tr>
     	<td class="Gradient2" width="223" valign="top" align="left">
      
      	<!--#include virtual="INC/LeftColumn.inc.asp"-->

		<form method="get" name="Country">		<% ' This form was inherited from old version of this page, before I started INCLUDE-ing file INC/LeftColumn.inc.asp. I don't know if this form stuff is still needed or not. %>					
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