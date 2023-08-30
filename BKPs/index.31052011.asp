<%@ Language=VBScript %>


<SCRIPT id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</SCRIPT>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->




<html>


<head>
<link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->

<title>GPS Global Positioning Systems: Garmin, USglobalSat, Pharos. GMRS/VHF/UHF/CB Radios, Scanners. Tools</title>
<meta name="keywords" content="GPS,garmin,usglobalsat,pharos,GPS sensor,GPS engine,oem,GPS navigation,nuvi,GPS tracking,Smartphone,AVL,PDA,GMRS,CB Radio,2-way radio,scanner,Camera,tools">
<meta name="description" content="GPS at Star Lite Intl. Best source for Garmin GPS, USGlobalsat GPS, Pharos GPS, GPS sensors, GPS boards, OEM GPS, GPS engines, GPS tracking and Smartphone. 2-way radios, CB radios, scanners, antennas and accessories. Digital cameras, audio/video equipment, tools.">
<meta name="author" content="Star Lite International, LLC">
<meta name="copyright" content="1994-2010 Star Lite International, LLC">
<meta name="revisit-after" content="15 days">
<meta http-equiv="refresh" content="120; url=https://www.starlite-intl.com">
<meta name="distribution" content="global">
<meta name="robots" content="index,follow">
<meta name="rating" content="general">
<meta http-equiv="content-type" content="text/html;charset=western (Iso-8859-1)">
<meta http-equiv="content-language" content="en">
<META name="Classification" content="GPS, GPS sensors and GPS engines, GPS Tracking, oem GPS, communication and electronic equipment">
<meta name="DC.Title" content="Star Lite International, LLC">
<meta name="DC.Description" content="GPS: Full line of Garmin GPS, USglobalsat GPS, Pharos GPS, sensors, oem gps, GPS engines, gps board, gps navigation, tracking GPS, bluetooth GPS, GPS network, GPS antenna, fish finders. CB: cb radios, walky-talky, cb antennas, cb accessories, 2-way radios, scanners. Digital: digital cameras, flash memory. car audio and video, DJ, hand tools, mechanics tools.">
<meta name="abstract" content="Complete selection of GPS, GPS sensors, oem GPS, GPS engine boards, GPS antennas and accessories. Also wide selection of cb radios, scanners, antennas and automotive electronics, digital cameras, memory cards and hand tools.">
<link rel="stylesheet" href="style_new.css" type="text/css">
<link rel="shortcut icon" href="favicon.ico" TYPE="image/ico">
<LINK REL="SHORTCUT ICON" HREF="favicon.ico">




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
//var myPage;
//var myReferrer;
//var subPage;
//var subPage2;
//var slashcount;

//myReferrer=document.referrer;
//myPage = location.href;
//subPage = String(myPage).substring(7,myPage.length);
//for(x=0;x<subPage.length;x++)
//{
//    if(subPage.charAt(x) == "/")
//    {
//    slashcount = x;
//    break;
//    }
//}        
//subPage2 = String(myPage).substring(0,slashcount+7);
//subPage2 = subPage2+"/stats/record.asp?page="+myPage+"&ref="+myReferrer;
//mywindow = window.open(subPage2,'recorder','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=1,height=1');

//self.focus();    
//--></script>
<!--End Tracker Code//-->


<style>
#HomePageTable tbody tr td
{
    border:1px solid #aaa;
    padding: 10px;
    font-family: Tahoma, Arial, Sans-Serif;
}

#HomePageTable tbody tr td img
{
    border:0px solid red;
    margin: 5px;
}
</style>


</head>




<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<table style="border:0px solid green;" width='1100'  bgcolor="" align='center'>		<% ' Start Table 1 %>
<tr><td>

<% InArea="Home" %>

<!--#include virtual="Misc/Header.INC"-->


<% '*********************************************************************************************************************** %>


<table style="border-right:1px solid #84bff1;" width='1120' bordercolor="blue" cellpadding="0" cellspacing="0" align="center" > <% ' Start Table 1.1 %>
            
            <tr>
				<!-- <td background="Images/goldbackground222.jpg" width="223" valign="top" align="center"> -->
                <td class="Gradient2" width="223" valign="top" align="center">
				
				<table border=0 cellpadding="20">	<% ' Start of Table 1.1.4 %>
					<tr>
					<td>
						<center><img src="Images/logo.gif" WIDTH="135" HEIGHT="145"></center>
						<!--#include virtual="INC/SPECIAL.INC"-->
					</td>
					</tr>
				</table>				<% ' End Table 1.1.4 %>
				
				</td>
				
				<!-- <td background="Images/bluebackground2.jpg">  -->
                <td background="">

					<table border="0" cellpadding="10" cellspacing="0" align="center">  <% ' Start Table 1.1.5 %>
					<tr>
						
						<td valign="top" align="center">

						<!--#include virtual="INC/BANNER.INC"-->
                                
						<!-- [9/28/06, BN] This was the old way of including the "Text1" field of the Company table. 
						     Now we are using a copy of the Text1 data stored in html file Index3.htm.
                        <%
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							shqstring = "SELECT Text1, Text2 FROM Company "
							Set RHS = Conn.Execute(shqstring)
						%>

						
                        <p>	<font size="2" face="Tahoma">
							<strong>&nbsp;&nbsp;&nbsp;<%=RHS("Text1")%></strong>
							</font>
                        </p>
                       
						<% rhs.close %>
						-->
						
                        
                        <%
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							shqstring = "SELECT Text1, Text2 FROM Company "
							Set RHS = Conn.Execute(shqstring)
						%>
						
                        <p>	<font size="2" face="Tahoma">
							<strong>&nbsp;&nbsp;&nbsp;</strong>
							<!--#include virtual="Index3.asp"-->
							</font>
                        </p>
                        
						<p>	<font size="2" face="Tahoma">
							<strong>&nbsp;&nbsp;&nbsp;<%=RHS("Text2")%></strong>
							</font>
                        </p>
							
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



        
</body>


<% ' VI 6.0 Scripting Object Model Enabled %><% EndPageProcessing() %></html>

