<% @ Language=VBScript %> 


<SCRIPT id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</SCRIPT>

<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->


<html>


<head>

	<!-- <meta http-equiv="content-type" content="text/html; charset=UTF-8"> -->
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->

	<title>GPS Engine Boards | OEM GPS | CB Radios | Star Lite International, 
	LLC</title>
	<meta name="keywords" content="CB radios, GPS boards, GPS engine, GPS, GPS engine boards, OEM GPS sensors, GPS engine boards, GPS tracking, PDA, GPS smart phones, GPS antennas, CB, amateur and marine radios, radio scanners, fish finders">
	<meta name="description" content="Star Lite International,LLC is a supplier of GPS Engine Boards for the OEM industry. We stock a wide range of brand name CB radio and accessories in our online shop">
	<meta name="author" content="Star Lite International, LLC">
	<meta name="copyright" content="1994-2016 Star Lite International, LLC">
	<meta name="revisit-after" content="7 days">
	<meta name="distribution" content="global">
	<meta name="robots" content="index,follow">
	<meta name="robots" content="all">
	<meta name="rating" content="general">
	<meta http-equiv="content-language" content="en">
	<meta name="mssmarttagspreventparsing" content="True">
	<meta name="generator" content="Microsoft FrontPage 12.0">
	<meta name="abstract" content="Complete selection of GPS, GPS sensors, oem GPS, GPS engine boards, GPS modules, GPS antennas and accessories, marine electronics, sailing electronics, night vision and a wide selection of cb radios, radio scanners, antennas and automotive electronics, hand tools.">
	<meta name="google-site-verification" content="l1Q8MeNa6MWukWUWvkzDPjeHPnuuRAiDQiIVQwovhHE" />
	<meta name="google-site-verification" content="YQEJxbNZcOScCN1JPylYkrig4-ONJsFbLcG1Z7e8_I0" />
    
	<!-- <script type="text/javascript" src="http://www.ginwiz.com/mobileDetectionScript/redirection_mobile.min.js"></script> <script type="text/javascript">
                  SA.redirection_mobile({ noredirection_param: "noredirection", mobile_url: "starlite-intl.ginwiz.com", cookie_hours: "2", keep_path: "true" });
    </script> -->

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


	<!-- 
	1/5/14, BN: This js code is for detecting mobile devices and tablets and redirecting them to the mobile version of our site: startlite-intl.com/mobile1.
	It was free and comes from https://www.handsetdetection.com/sites/setup/48009/js . They apparently store the source URL startlite-intl.com
	and the target URL startlite-intl.com/mobile1, in their db.
	I log into my account at https://www.handsetdetection.com using email = my bn2 address, pwd = the usual.
	This js code is to be pasted into the <head> section, before any other js code.
	-->
	<script type='text/javascript'>
		(function () {
			'use strict';
			var hd, internal;
			internal = document.referrer.search(document.domain);
			if (internal === -1) {
				hd = document.createElement('script'); hd.type = 'text/javascript'; hd.async = true;
				hd.src = ('https:' == document.location.protocol ? 'https://' : 'http://') + 'api.handsetdetection.com/sites/js/48009.js';
				document.getElementsByTagName("head")[0].appendChild(hd);
			}
		} ());
	</script>


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
		border:1px solid #DDD;
		padding: 10px;
		font-family: Tahoma, Arial, Sans-Serif;
	}

	#HomePageTable tbody tr td img
	{
		border:0px solid red;
		margin: 5px;
	}
	</style>


<!-- The folowing script added on July 22, 2015. -->

	<script>
	(function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
	(i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
	m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
	})(window,document,'script','//www.google-analytics.com/analytics.js','ga');
	
	ga('create', 'UA-27883837-2', 'auto');
	ga('send', 'pageview');
	
	</script>

	<!-- <script type="text/javascript">
  		var _gaq = _gaq || [];
		_gaq.push(['_setAccount', 'UA-27883837-2']);
  		_gaq.push(['_setDomainName', 'starlite-intl.com']);
  		_gaq.push(['_trackPageview']);

  		(function() {
    		var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    		ga.src = ('https:' == document.location.protocol ? 'https://' : 'http://') + 'stats.g.doubleclick.net/dc.js';
    		var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  		})();

	</script> -->


</head>




<body>

<table style="border:0px solid green;" width='1100'  bgcolor="" align='center'>		<% ' Start Table 1 %>
<tr><td>

<% InArea="Home" %>

<!-- #include virtual="Misc/Header.INC" -->


<% '*********************************************************************************************************************** %>


<table style="border-right:1px solid #84bff1;" width='1120' bordercolor="blue" cellpadding="0" cellspacing="0" align="center" > <% ' Start Table 1.1 %>
            
            <tr>
				<!-- <td background="Images/goldbackground222.jpg" width="223" valign="top" align="center"> -->
                <td class="Gradient2" width="223" valign="top" align="center">
				
				<table border=0 cellpadding="20">	<% ' Start of Table 1.1.4 %>
					<tr>
					<td>
						<center><XXXimg src="Images/StarLogo.png" WIDTH="140" HEIGHT="140"></center>
						<!--#Include virtual="INC/SPECIAL.INC"-->
					</td>
					</tr>
				</table>				<% ' End Table 1.1.4 %>
				
				</td>
				
				<!-- <td background="Images/bluebackground2.jpg">  -->
                <td background="" valign="top">

					<table border="0" cellpadding="0" cellspacing="0" align="center">  <% ' Start Table 1.1.5 %>
					<tr>
						
						<td valign="top" align="center">

						<!--#Include virtual="INC/BANNER.INC"-->
                                
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
						
                        <font size="2" face="Tahoma">
							<strong>&nbsp;&nbsp;&nbsp;</strong>
							<!-- #include virtual="Index3.inc.asp"-->
						</font>
                        
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
								<font size=2>Back to top</font>
								</a>
								<br /><br />
								</td>
							</tr>
							</table>						<% ' End Table 1.1.5.2 %>
						
						
					</td>
					</tr>

					</table>		<% ' End Table 1.1.5 %>
			
				</td>
           
			</tr>
                        
            
</table>	<% ' End Table 1.1 %>


	<!-- #include file="Misc/Footer.INC" -->
      
</td>
</tr>
</table>	<% ' End Table 1 %>


<br>

        
</body>


<% ' VI 6.0 Scripting Object Model Enabled %><% EndPageProcessing() %></html>

