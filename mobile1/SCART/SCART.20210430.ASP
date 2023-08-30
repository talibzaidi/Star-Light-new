<%@ LANGUAGE = VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<html>
 

<head>
	<title>GPS sensors | GPS engine boards | OemGPS | GPS tracking | two-way communication | Garmin | USglobalSat | Pharos</title>
	<meta name="keywords" content="GPS,GPS sensors,GPS engine,OEM GPS,GPS OEM,GPS board,GPS engine boards,garmin GPS,USGlobalsat GPS,pharos GPS,GPS navigation,GPS 15x,GPS16x,GPS18x,GPS19xHVS,PDA,mr350,sensor,em408,em411,bu353,GPS16xLVS,GPS 18xLVC,GPS systems,auto GPS,GPS accessories,bluetooth GPS,BT 359,handheld GPS,EM318,ET318,ET333,fixed mount GPS,Nuvi,forerunner,br355,bu353,marine GPS,GPS receivers,fish finder,sounders,transducers,cartography,GPS mapping software,GMR,GPSMAP,tracking GPS,navigation,low cost GPS,GPS specials,AVL,fleet tracking GPS,fleet management,active GPS antenna,equipment,vehicle locator,marine networking,marine navigation,SIRF,marine radars,Authorized GPS dealer,GPS resellers,GPS business,gps wholesale">
	<meta name="description" content="The GPS source for Garmin GPS and USglobal GPS. GPS sensors, OEM GPS, GPS boards, GPS accessories, tracking GPS, bluetooth GPS, GPS network systems, GPS antennas, Marine GPS, Marine radar, Navigation Mapping software, fish finders and sounders.">
	<meta name="author" content="Star Lite International, LLC">
	<meta name="copyright" content="1994 - 2012 Star Lite International, LLC">
	<meta name="revisit-after" content="15 days">
	<meta name="distribution" content="global">
	<meta name="robots" content="index,follow">
	<meta name="rating" content="general">
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
	<meta http-equiv="content-language" content="en">
	<meta name="DC.Title" content="Star Lite International, LLC">
	<meta name="DC.Description" content="The best GPS source for GPS sensors, oem gps, GPS receivers, gps engine boards, tracking GPS, GPS antennas, GPS accessories, tracking GPS, Garmin GPS, USglobalsat GPS, Pharos GPS. bluetooth GPS, GPS networks, fish finders and sounders, Marine GPS">
	<meta name="abstract" content="The source for  GPS, GPS sensors, oem GPS, OEM engine boards, hand held and fixed mount GPS, cartography and GPS accessories">
	<meta http-equiv="content-language" content="en">
	
    <link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/mobile1/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
	<meta charset=utf-8>

	<meta name="viewport" content="width=device-width; initial-scale=1.0">
	<!-- foneFrame.css is the stylesheet with comments, so it is readable.
	     foneFrame-min.css is the minimized version; it is smaller and loads faster. -->
	<link href="https://www.starlite-intl.com/mobile1/foneFrame.css" rel="stylesheet" type="text/css">
	<!-- The following 2 lines are not strict HTML5. -->
	<meta name="HandheldFriendly" content="true"/>
	<meta name="MobileOptimized" content="320"/>


<% response.buffer=true %>

<% ar = Request("Area") %>
<% Area = Request("Area") %>

<% 
sar		= Request("sar") 
NewProductsSubgroup	= Request.Querystring("NewProductsSubgroup")
SpecialsSubgroup	= Request.Querystring("SpecialsSubgroup")
RebatesSubgroup		= Request.Querystring("RebatesSubgroup")

If FALSE Then
	Response.Write "<br>sar = "					& sar 
	Response.Write "<br>NewProductsSubgroup = "	& NewProductsSubgroup
	Response.Write "<br>SpecialsSubgroup = "	& SpecialsSubgroup 
	Response.Write "<br>RebatesSubgroup = "		& RebatesSubgroup 
End If
%>


<% SID = ReQuest("SID") %>

<% if Area="iii" then
   Area = Request("Manufat")
   ar = Request("Manufat")
   end if
   
   'If Area = "New Products" then
   '  Response.Redirect "./scart.asp?pid=0&sid=11&area=New!&sar=New%20Products"
   'End If    

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


<!-- 11/10/13: For the accordion menu from menucool.com, where its HTML is in a separate file, and does not have to be repeated in each webpage that has the menu. -->
<link href="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.css" rel="stylesheet" type="text/css" />
<script src="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.js" type="text/javascript"></script>

</head>


<body>

<%
' [BN, 2/1/18] The following script is an experiment to try to close the amenu whenever this file is entered.
' I am not really sure that this is correct.
' I also do this in file Misc > MenuCoolAmenuOneForAll > amenu > amenu-source.html
%>
<script type="text/javascript"> 
    amenu.close(true);                          // This seems to actually WORK!!
    // amenu.open("Communications", true);     // This actually WORKS!! (if comment out "amenu.close(true);")
</script>

<% 
InArea = "SCART"   ' Refers to area of website, not area of products.
'Response.Write "<br>InArea = " & InArea 
	
action = request("action")
'Response.Write "<br>action = " & action 
'Response.Write "<br>Area = " & Area 

If (action = "add") OR (action = "viewcart") OR (action = "checkout") Then
	InArea = "ShoppingCart" 
'ElseIf Area = "New!" Then
'	InArea = "WhatsNew"
ElseIf sar = "New Products" Then
	InArea = "WhatsNew"
ElseIf sar = "Specials" Then
	InArea = "Specials"
ElseIf sar = "Rebated" Then
	InArea = "Specials"
ElseIf Area = "Gift Certificates" Then
	InArea = "GiftCertificates"
End If

'Response.Write "<br>InArea = " & InArea
%>


<!-- #include virtual="mobile1/Misc/Header.INC" -->

<% 
' I added this 'If action <> "checkout"' so as to avoid the unnecessary clutter of the USA/Canada flag stuff appearing at 
' top of checkout form on a mobile phone, where space is at a premium.
If action <> "checkout" Then 
%>
    <!-- This buttons form was copied from INC/LeftColumn.inc.asp file of original, non-mobile version of this site. -->
    <form method="get" name="Country">
    <center>
    <br /><font face="Tahoma" size="2">You are currently a 
    <% If Session("Country") = "Canada" Then %>
	    <img src="https://www.starlite-intl.com/Images/can1.gif" WIDTH="36" HEIGHT="18"> 
    <% Else				' Previously: ElseIf Session("Country") = "USA" Then 
	    Session("Country") = "USA"
    %>
	    <img src="https://www.starlite-intl.com/Images/USA1.gif" WIDTH="34" HEIGHT="18"> 
    <% End if %> 
    customer. Click on a button to change countries:
    </font>					
    <input type="submit" name="Canada" value="Canada">
    <input type="submit" name="  USA  " value="USA">
    <input type="hidden" name="pid" value="<%=request("pid")%>">
    <input type="hidden" name="sid" value="<%=request("sid")%>">
    <input type="hidden" name="area" value="<%=request("area")%>">
    <input type="hidden" name="sar" value="<%=request("sar")%>">
    <input type="hidden" name="action" value="<%=request("action")%>">
    </center>
    </form>
<% End If     ' action <> "checkout" %>


<% ' Start Table 1 %>
<table style="border-right:0px solid #84bff1; border:0px solid black;" XXXwidth='1120' cellpadding="0" cellspacing="0" align="center" >   
    <tr>
     	<td class="Gradient2" XXXwidth="223" valign="top" align="left">
      
      	<!--# include virtual="INC/LeftColumn.inc.asp"-->

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
		    <table border="0" cellpadding="0" cellspacing="0" align="center">  
		    <tr>
			    <td valign="top" align="center">
			    <!-- #include file="BANNER.INC" -->
			    <!-- #include file="SCART.INC" -->
			    </td>
		    </tr>
		    </table>		
            <% ' End Table 1.1.2 %>	

        </td>
	</tr>

    <tr>
        <td colspan="2">
        	&nbsp;<!-- #include file="../Misc/Footer.INC" -->
        </td>
    </tr>

</table>   
<% ' End Table 1 %>


</body>


</html>