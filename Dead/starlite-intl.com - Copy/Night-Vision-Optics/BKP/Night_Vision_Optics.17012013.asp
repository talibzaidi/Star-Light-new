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



<HTML>


<HEAD>
	<!-- <meta http-equiv="Content-Type" content="text/html; charset=utf-8"> -->
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
    <meta http-equiv="content-type" content="text/html; charset=UTF-32">
    <meta http-equiv="content-language" content="en">
    <title>Night Vision Optics | GPS OEM engine boards | OEM GPS | OEM GPS module | starlite-intl.com</title>
    <meta NAME="Description" CONTENT="Night Vision Optics by Night Owl Optics and Bering Optics./">
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
<tr>
	<td>
	<XXXiframe name="inlineframe" src="https://www.starlite-intl.com/OEM_GPS_sensors/searchsummary.asp?CF=CSCS&AID=45&SID=173&ShowPageNum=1" frameborder="0" scrolling="auto" width="95%" height="600" marginwidth="5" marginheight="5" ></iframe> 

	<!-- #INCLUDE file="searchsummaryNight_Vision_Optics.inc.asp"    
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
