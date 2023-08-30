
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

Where can I learn more about building mobile websites?
http://qrdvark.com/templates/best-practices/
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
	<link href="https://www.starlite-intl.com/mobile1/foneFrame.css" rel="stylesheet" type="text/css">
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
</head>


<body>

<span class=i>
	<p class=iDes>
	Some of this was based on foneFrame v1.0.1 Copyright 2011 Azalea Software, Inc. www.QRdvark.com/foneFrame/ 31aug11 
	</p>

	<p class=iDes>
	foneFrame is a mobile framework that creates web pages for smartphones like Android & iPhone. 
	</p>

	<p class=iDes>
	Where can I learn more about building mobile websites?
	http://qrdvark.com/templates/best-practices/
	</p>

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

	<!-- gradient box -->
	<p class=iDes>Begin by reading this HTML file &amp; <i>foneFrame.css</i>. 
	There are comments in both files to help you. <i>foneFrame.css</i> is easy to read but <i>foneFrame-min.css</i> 
	has been compressed so it will load faster. Develop and deploy respectively.</p>

	<p class=iDes><i>The content is in the HTML, the formatting is in the CSS.</i> 
	Edit the CSS file to change your site&rsquo;s look &amp; feel. 
	For example, the gradient boxes are defined in the <code>i</code> style. 
	Remove the style &amp; the page has <a href="index-alternate.html">a more open look</a>. 
	Below are examples of how the base styles can be used to create specific layouts.</p>
</span>


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


<!-- #include virtual="mobile1/Misc/Header.INC" -->


	<% sar = "Home"
	If (Request("Canada") <> "" OR Request("  USA  ") <> "") then
		If Request("Canada") <> "" then
		Session("Country") = "Canada"
		else
		Session("Country") = "USA"
		end if
	end if
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
		<br />Click to change countries.
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


<table style="border-right:1px solid #84bff1;" XXXwidth='1120' bordercolor="blue" cellpadding="0" cellspacing="0" align="center" > <% ' Start Table 1.1 %>
            
            <tr>
                <td background="" valign="top">

					<table border="0" cellpadding="0" cellspacing="0" XXXalign="center">  <% ' Start Table 1.1.5 %>
					<tr>
						
						<td valign="top" align="center">

						<!--# Include virtual="INC/BANNER.INC"-->
                                
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
							<!-- #include virtual="mobile1/Index3.inc.asp"-->
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


	<!-- # include file="Misc/Footer.INC" -->
      
</td>
</tr>
</table>	<% ' End Table 1 %>


</body>
</html>