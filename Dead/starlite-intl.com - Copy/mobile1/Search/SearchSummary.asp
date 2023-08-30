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



<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>




<script language="javascript">
// Based on summary page from www.futuresimchas.com.

	function pNext(sObj, NumMembersPerPage){
		if (sObj.selectedIndex < (sObj.options.length -1)){
			sObj.options[sObj.selectedIndex+1].text="Loading "+sObj.options[sObj.selectedIndex+1].text;
			sObj.options[sObj.selectedIndex+1].selected=true;
			pGo(sObj, NumMembersPerPage);
		} else {
			alert('End of Results Reached.');
		}
	}
	
	function pPrev(sObj, NumMembersPerPage){
		if (sObj.selectedIndex > 0){
			sObj.options[sObj.selectedIndex-1].text="Loading "+sObj.options[sObj.selectedIndex-1].text;
			sObj.options[sObj.selectedIndex-1].selected=true;
			pGo(sObj, NumMembersPerPage);
		} else {
			alert('Beginning of Results Reached.');
		}
	}
	
	function pGo(sObj, NumMembersPerPage){
	// The next line is because the values in each option of the menu are not pages per se, 
	// but the number (in consecutive order in those found by the search; not MemberID) of the first member on the page.
	SelectedPage = (sObj.options[sObj.selectedIndex].value - 1 )/ NumMembersPerPage + 1;   
	//location.href='searchsummary.asp?ShowPageNum='+SelectedPage;
	location.href='searchsummary.asp?ShowPageNum='+SelectedPage;
	}
	
</script>


	<script type="text/javascript">
  		var _gaq = _gaq || [];
		_gaq.push(['_setAccount', 'UA-27883837-2']);
  		_gaq.push(['_setDomainName', 'starlite-intl.com']);
  		_gaq.push(['_trackPageview']);

  		(function() {
    		var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    		ga.src = ('https:' == document.location.protocol ? 'https://' : 'http://') + 'stats.g.doubleclick.net/dc.js';
    		var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  		})();

	</script>


	<script type="text/javascript">
		// 11/8/2011: For Google Analytics for Sani's Google acct.
		var _gaq = _gaq || [];
		_gaq.push(['_setAccount', 'UA-4694351-1']);
		_gaq.push(['_trackPageview']);

		(function () {
			var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
			ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
			var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
		})();
	</script>


<HTML>


<head>
    <% ' The following stylesheet was causing trouble on the Details.asp page (of the mobile site) by indenting the left edge of the <body>.
       ' I have not investigated why, but at least for now I am just commenting out the stylesheet.
       ' 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus (on main, non-mobile, site) to work. 
       ' 12/6/17: But in any case, it is apparently not needed here (on the mobile site).
    %>
    <XXXlink rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/mobile1/Misc/StyleSheet1.css"> 
	<meta charset=utf-8>
	
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
	<script type="text/javascript">amenu.close(true);</script>
	<!-- <script type="text/javascript">amenu.open("Communications", true); -->
</head>



<body onload="amenu.close(true)">

<!-- 
<form>
<div>
    <p>Using the <span class="cn green">open</span> or <span class="cn green">close</span> function to 
	programmatically open or close a sub-menu. <a href="http://www.menucool.com/vertical/accordion-menu#open_close">details...</a></p>
    <input type="radio" name="ml" onclick="amenu.open('GPS', true)" /> Open GPS<br />
    <input type="radio" name="ml" onclick="amenu.open('Communications', true)" /> Open Communications<br />
    <input type="radio" name="ml" onclick="amenu.close(true)" /> Close
</div>
</form>
-->


<%
'ComingFrom = Request.QueryString("CF")
'Response.Write "<br>ComingFrom = "	& ComingFrom 

'ShowPageNum = Request.QueryString("ShowPageNum")
'If ShowPageNum = "" Then ShowPageNum = 1 End If
%>



<% InArea = "Products" %>

<table style="border:0px solid green;" XXXwidth='1100'  bgcolor="" align='center'>		
<tr><td>
<!-- #include virtual="mobile1/Misc/Header.INC" -->
</td></tr>
</table>


<% 
If (Request("Canada") <> "" OR Request("  USA  ") <> "") Then
	If Request("Canada") <> "" Then
		Session("Country") = "Canada"
	Else
		Session("Country") = "USA"
	End If
End If

'Response.Write "Request('Canada') = "		& Request("Canada")
'Response.Write "<br>Request('  USA  ') = "	& Request("  USA  ")
'Response.Write "<br>Session('Country') = "	& Session("Country")
%>


<!-- This buttons form was copied from INC/LeftColumn.inc.asp file of original, non-mobile version of this site. -->
<form method="get" name="Country">
<center>
	<p><font face="Tahoma" size="2">You are currently a 
	<% If Session("Country") = "Canada" Then %>
		<img src="https://www.starlite-intl.com/Images/can1.gif" WIDTH="36" HEIGHT="18"> 
	<% Else				' Previously: ElseIf Session("Country") = "USA" Then 
		Session("Country") = "USA"
	%>
		<img src="https://www.starlite-intl.com/Images/USA1.gif" WIDTH="34" HEIGHT="18"> 
	<% End if %> 
	customer.<br />
	Click to change countries.
	</font></p>
								

	<input type="submit" name="Canada" value="Canada">
	<input type="submit" name="  USA  " value="USA">

	<% 
	ComingFrom = Request.QueryString("CF")    ' "CSCS"
	AID = Request.QueryString("AID")
	SID = Request.QueryString("SID")
	KW = Request.QueryString("KW")  
    If False Then
        'ComingFrom = "CSCS"
        'AID = "45"
        'SID = "294"
	    Response.Write "<br>ComingFrom = "	& ComingFrom 
	    Response.Write "<br>AID = "		    & AID 
	    Response.Write "<br>SID = "		    & SID 
	    Response.Write "<br>KW = "			& KW 
    End If
	%>

	<input type="hidden" name="AID" value="<%=AID%>">
	<input type="hidden" name="SID" value="<%=SID%>">

	<!--
	<input type="hidden" name="pid" value="<%=request("pid")%>">
	<input type="hidden" name="sid" value="<%=request("sid")%>"> 
	<input type="hidden" name="area" value="<%=request("area")%>">
	<input type="hidden" name="sar" value="<%=request("sar")%>">
	<input type="hidden" name="action" value="<%=request("action")%>">
	-->

</center>
</form>


<table align='center' cellpadding='5' cellspacing='0' border='0' XXXwidth='<%=PageWidth%>' >
<!--
<tr>
	<td>
		<table align='center' width='70%'>
		<tr>
			<td>
			<br /><br />
			Select a GPS Sensor, OEM GPS, or GPS engine board from our wide selection of
			 OEM GPS sensors, Engine boards and OEM GPS accessories.
			 Surely you will find one here suitable for YOUR application.
			 We feature Garmin OEM GPS, and USGlobalsat OEM GPS products.  
			 <br /><br />
			 </td>
		 </tr>
		 </table>
	</td>
</tr>
-->

<tr>
	<td>
	<!-- #INCLUDE file="SearchSummary.inc.asp" -->
	</td>
</tr>
</table>


<table align='center' cellpadding='5' cellspacing='0' border='0' XXXwidth='<%=PageWidth%>' >
<tr>
	<td>
	<!-- # include virtual="Misc/Footer.INC" --> 
	</td>
</tr>
</table>


</BODY>



</HTML>

<% ' VI 6.0 Scripting Object Model Enabled %><% EndPageProcessing() %>
