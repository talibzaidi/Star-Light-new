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
<!--#include file="../_ScriptLibrary/pm.asp"-->
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


<HEAD>
	<!-- <meta http-equiv="Content-Type" content="text/html; charset=utf-8"> -->
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
    <meta http-equiv="content-type" content="text/html; charset=UTF-32">
    <meta http-equiv="content-language" content="en">
    <title>GPS sensors | GPS OEM engine boards | OEM GPS | OEM GPS module | starlite-intl.com</title>
    <meta NAME="Description" CONTENT="Complete line of GPS sensors, GPS oem boards, OEM GPS, GPS engine, GPS engine Boards, GPS tracking. Also GPS Smartphones, truck GPS, 2-way radios, CB radios, radio scanners, antennas and accessories, night vision optics./">
    <meta NAME="Keywords" CONTENT="GPS sensors,OEM GPS boards,GPS boards,GPS engine,GPS oem engine Boards,GPS tracking,GPSMap,GPS Smartphones,2-way radios,CB radios,radio scanners,antennas and accessories,night vision optics,Garmin,USGlobal,Pharos,Uniden,Cobra,Midland /">
</HEAD>



<body>

<%
'ComingFrom = Request.QueryString("CF")
'Response.Write "<br>ComingFrom = "	& ComingFrom 

'ShowPageNum = Request.QueryString("ShowPageNum")
'If ShowPageNum = "" Then ShowPageNum = 1 End If
%>



<% InArea = "Products" %>

<!-- #INCLUDE VIRTUAL = "Misc/Header.inc" -->

<table align='center' cellpadding='5' cellspacing='0' border='0' width='<%=PageWidth%>' >
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

	<!-- #INCLUDE file="searchsummaryOEM-GPS-sensors.inc.asp"    
	  Speacialized for:
	  AID = "45"	  ' GPS Navigation, GPS Sensors, OEM, FishFinders, Maps
      SID = "173"     ' GPS - OEM: Sensors / Boards / TracPacs
	-->

	</td>
</tr>
</table>


<table align='center' cellpadding='5' cellspacing='0' border='0' width='<%=PageWidth%>' >
<tr>
	<td>
	<!-- #include virtual="Misc/Footer.INC" --> 
	</td>
</tr>
</table>

<br>

</BODY>


<% ' VI 6.0 Scripting Object Model Enabled %><% EndPageProcessing() %>


</HTML>
