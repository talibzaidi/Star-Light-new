
<%@ LANGUAGE = VBScript %>

<!doctype html> 

<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->


<html lang=en> 
<!--  Some of this was based on foneFrame v1.0.1 Copyright 2011 Azalea Software, Inc. www.QRdvark.com/foneFrame/ 31aug11  
foneFrame is a mobile framework that creates web pages for smartphones like Android & iPhone. 
I have stored my foneFrame stuff on my laptop in folder at: 
C:\Users\owner\Google Drive\Work\Websites - Clients\Starlite\Mobile\foneframe
Where can I learn more about building mobile websites?
http://qrdvark.com/templates/best-practices/
[12/5/20:] I discovered that qrdvark.com no longer exists.

See article: "Foneframe Is Now A Kickstarter Project; Mobile Web Sites Will Never Be The Same Again"
at https://www.newswire.com/foneframe-is-now-a-kickstarter/152651
www.kickstarter.com/projects/JetCityOrange/foneframe-a-mobile-web-framework
I found that foneFrame no longer exists!

Google "mobile device detect and redirect"
-->

<!--
<article class=i>
	<p class=iDes><b>Is my browser HTML compliant?</b><br />
	Check to see if your browser is <a href="http://www.html5test.com/"/>HTML5 compliant</a>.</p>

	<p class=iDes><b>Where do I validate my HTML5 code?</b><br />
	Use the <a href="http://validator.w3.org/unicorn/">W3 HTML5 validator</a> to insure that your files are correct. 
	Certain sections of this template aren&rsquo;t strict HTML5, no surprise given that HTML5 is new and subject to change.</p>

	<p class=iDes><b>How can I support non-HTML5 browsers?</b><br />
	You can display a mobile page that doesn&rsquo;t require HTML5 by replacing the first two lines of this template with:</p>
	<p class=iDes><blockquote><code>
&lt;?xml version="1.0" ?><br />
<!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.2//EN"<br />
"http://www.openmobilealliance.org/tech/DTD/xhtml-mobile12.dtd"><br />
&lt;html xmlns="http://www.w3.org/1999/xhtml"><br />
</blockquote></p>
</article>
-->


<head>

	<link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/mobile1/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
	<meta name="keywords" content="GPS, GPS navigation, GPS sensors, OEM GPS, GPS accessories, CB, CB radio, CB radios, Garmin gps, global positioning, WalkyTalky, mobile tracking, fleet tracking, USglobasat gps, bluetooth, flash memory, gmrs, marine radios, navigation electronics, 2-way radios, radio scanners, marine radios, car audio, car stereos, power amplifiers, antennas, power supplies, regulated power supplies, DJ, accessories, hand tools, mechanics tools, Uniden, Cobra, Midland, MIT, Pyramid, Pyle, Solarcon">
	<meta name="description" content="Large selection of - GPS, GPS sensors, GPS accessories, GPS OEM, PDA, tracking gps, bluetooth gps, fish finders, sounders, CB radios and Walky-talky, flash memory, MP3, radio scanners, digital cameras, car audio and car video, DJ, hand tools and mechanics tools">
	<meta charset=utf-8>
	<XXXmeta name="keywords" content="foneFrame, mobile web design, smartphone web template, open source mobile web design, cell phone web pages, mobile phone template, HTML5 mobile phone template, how to make a mobile web page using HTML5, how to make a mobile web page using CSS3, mobile phone framework, HTML5, CSS3, HTML5 for mobile phones, CSS3 for mobile phones, HTML5 phone framework, CSS3 phone framework, QRdvark, Azalea Software">
	<XXXmeta name="description" content="foneFrame is a mobile phone framework built with HTML5 & CSS3. The foneFrame mobile web template is open source released under Creative Commons (CC BY 3.0) that designs web pages for smartphones like Android and iPhone.">
	
	<meta name="viewport" content="width=device-width; initial-scale=1.0">
	<!-- foneFrame.css is the stylesheet with comments, so it is readable.
	     foneFrame-min.css is the minimized version; it is smaller and loads faster. -->
	<link href="foneFrame.css" rel="stylesheet" type="text/css">
	<!-- The following 2 lines are not strict HTML5. -->
	<meta name="HandheldFriendly" content="true"/>
	<meta name="MobileOptimized" content="320"/>

	<!-- You can use different style sheets for mobile vs. computer browsers: -->
	<!--  <link href="style-mobile.css" rel="stylesheet" type="text/css" media="handheld"> -->
	<!--  <link href="style-computer.css" rel="stylesheet" type="text/css" media="screen"> -->
	<!-- The favicon & iOS home screen icon are both 57x57 PNG's. Use a full URL file path for Android devices.  -->
	<!--  <link rel="apple-touch-icon-precomposed" href="http://yoursite.com/apple-touch-icon.png">  -->
	<!--  <link rel="icon" type="image/vnd.microsoft.icon" href="http://yoursite.com/favicon.png">  -->
	<!-- Site maps help search spiders where to find your pages.  www.xml-sitemaps.com  -->
	<!--  <link rel="alternate" type="application/rss+xml" title="ROR" href="ror.xml"> -->
	<!-- Your Google Analytics code goes here, just before the </head> tag. -->

    <!-- 11/10/13: For the accordion menu from menucool.com, where its HTML is in a separate file, and does not have to be repeated in each webpage that has the menu. -->
    <link href="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.css" rel="stylesheet" type="text/css" />
    <script src="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.js" type="text/javascript"></script>
	<script type="text/javascript">amenu.close(true);</script>
</head>


<body>


<!-- Remove the comments below after reading them. Smaller files load faster.  -->
<!-- 	http://www.textfixer.com/html/compress-html-compression.php -->

	<!--
	<nav id=navBtn>
		<ul>
			<li><a href="index-alternate.html">alternate</a></li>
			<li><a href="http://QRdvark.com/foneFrame/">full version</a></li>
			<li><a href="http://QRdvark.com/qr-generator/">QR barcodes</a></li>
		</ul>
	</nav>
	-->
<!--  end of nav  -->

<%
' [12/6/20, BN] The following line is just a test. From https://stackoverflow.com/questions/7883552/asp-substitute-for-user-agent-php. Works fine.

' Response.Write Request.ServerVariables("HTTP_USER_AGENT")

' Experimenting with using this to check what the User Agent string contains for various mobile devices being used, to help me 
' decide what I might use to trigger redirection from main site's home page "index.asp" to this page /mobile1/index.asp.
' For example, when I use my Android phone, the word "Android" appears in the User Agent string. So on main site home page I could use:

'  Response.Write Request.ServerVariables("HTTP_USER_AGENT")
'  Android = InStr(Request.ServerVariables("HTTP_USER_AGENT"), "Android")
'  Response.Write "<br>Android = " & Android
'  If (Android) Then 
'	Response.Redirect "/mobile1/index.asp"
'	Response.End
'  End If 

' The above conditional redirect worked for my Android phone, but I probably won't want to use the User Agent to decide whether to redirect to mobile site, 
' bc would need too many separate tests of above type, one for each device that I want to redirect for. 
%>


<!-- #include virtual="mobile1/Misc/Header.INC" -->


	<% sar = "Home"
	If (Request("Canada") <> "" OR Request("  USA  ") <> "") then
		If Request("Canada") <> "" then
		Session("Country") = "Canada"
		else
		Session("Country") = "USA"
		end if
	end if

	'Response.Write "<br>Session('Country') = " & Session("Country")
	%>

	<form method="post" name="Country" action="<%=request.servervariables("URL")%>">
		<p align='center'>
		<font face="Tahoma" size="2">
		You are currently a 
		<% If Session("Country") = "Canada" Then%>
			<img src="https://www.starlite-intl.com/Images/can1.gif"> 
		<% Else                            ' Previously: Elseif Session("Country") = "USA" Then 
			Session("Country") = "USA"
		%>
			<img src="https://www.starlite-intl.com/Images/USA1.gif"> 
		<% End If %> customer.
		<br />Click on a button below to change countries.
		</font></p>
                                
		<center>
		<input type="submit" name="Canada" value="Canada">
		<input type="submit" name="  USA  " value="USA">
		</center>
	</form>


<table style="border:0px solid green;" XXXwidth='1100'  bgcolor="" align='center'>		<% ' Start Table 1 %>
<tr><td>

<% InArea="Home" %>


<% '*********************************************************************************************************************** %>


<table style="border:0px solid #84bff1;" XXXwidth='1120' bordercolor="blue" cellpadding="0" cellspacing="0" align="center" > <% ' Start Table 1.1 %>
            
            <tr>
                <td background="" valign="top">
					<table border="0" cellpadding="0" cellspacing="0" XXXalign="center">  <% ' Start Table 1.1.5 %>
					<tr>		
						<td valign="top" align="center">						
                        <%
						' See https://www.starlite-intl.com/mobile1/index.extraction.asp for what used to be here
						' in the FALSEd-out block, from mobile1/Index3.inc.asp and RHS("Text2") and "Back to Top".
						If TRUE Then
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							shqstring = "SELECT Text1, Text2 FROM Company "
							Set RHS = Conn.Execute(shqstring)
						%>


                        <font size="2" face="Tahoma">
							<strong>&nbsp;&nbsp;&nbsp;</strong>
							<!-- # include virtual="mobile1/Index3.inc.asp"-->
							
						</font>

        <center>
        <font face=Tahoma>
		
		<!-- (c) 2005, 2012. Authorize.Net is a registered trademark of CyberSource Corporation --> 
		<!-- OLD:
		<div class="AuthorizeNetSeal"> 
		<script type="text/javascript" language="javascript">var ANS_customer_id="30dd88ea-13d1-4a1f-bafd-84fd68507946";</script> 
		<script type="text/javascript" language="javascript" src="//verify.authorize.net/anetseal/seal.js" ></script> 
		<a href="https://www.authorize.net/" id="AuthorizeNetText" target="_blank">Transaction Processing</a> 
		</div> 
		-->
        
        <br>

		<table>
        <tr>
        <td valign="top">
		<!-- NEW 8/27/17: -->
		<img src="https://www.starlite-intl.com/images/usaepaysecure.gif">&nbsp;&nbsp;&nbsp;
        </td>
        <td valign="top">
		<script type="text/javascript" src="https://seal.XRamp.com/seal.asp?type=H"></script>&nbsp;&nbsp;&nbsp;
		</td>
        </tr>
		<tr>
		<td colspan="2" align="center">
		    <a title="Star Lite International, LLC BBB Business Review" href="https://www.bbb.org/eastern-michigan/business-reviews/electronic-equipment-and-supplies-wholesale-and-manufacturers/star-lite-international-llc-in-southfield-mi-45003227/#bbbonlineclick"><img alt="Star Lite International, LLC BBB Business Review" style="border: 0;" src="https://seal-easternmichigan.bbb.org/seals/blue-seal-250-52-star-lite-international-llc-45003227.png" /></a>
		</td>
        </tr>
		</table>


        </font>
        </center>
	
                        <!--
						<p>	<font size="2" face="Tahoma">
							<strong>&nbsp;&nbsp;&nbsp;<% '=RHS("Text2")%></strong>
							</font>
                        </p>
                        -->
						
						<% End If ' FALSE %>

                        <br />
                        <small>
                        USA and Canadian Customers are welcome to use any of the following payment methods ...
                        <br />
                        <img style="border: 0px solid;" src="https://www.starlite-intl.com/imi/AcceptMarks.jpg" title="American
Express, Discover, MasterCard, Visa" align="middle">
                        <img style=" width:
55px; height: 55px; border: 0px solid;" src="https://www.paypal.com/en_US/i/icon/verification_seal.gif" alt="Official PayPal Seal" align="middle">
						<br />Terms and volume discounts are available for pre-approved corporations.
                        <br /><br /><hr />
                         All trademarks and logos on our website are the property of their 
                         respective owners or companies.
						<br />
                        Copyright &copy; 1994-2020 Star Lite International, LLC
                        </small>
                        
                        </td>
					</tr>

					</table>		<% ' End Table 1.1.5 %>
			
				</td>
           
			</tr>
</table>	<% ' End Table 1.1 %>

    <br />
	<!-- #include file="Misc/Footer.INC" -->
     <br /><br />

</td>
</tr>
</table>	<% ' End Table 1 %>


<!--
Copy & paste this section to add a new item 
<article class=i>
	<p class=iTtl>Item Title</p>
	<p class=iDes>Item Description</p>
	<p class=iLnk>Item Link</p>
</article>
-->


</body>
</html>