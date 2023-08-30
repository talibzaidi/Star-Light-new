<%@ LANGUAGE = VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<% 
ExternalSiteURL = ReQuest("URL") 
State = ReQuest("state")			' Needed only when coming from Misc/AuthorizedDealerFor.asp to display Starlite as a Garmin authoized dealer.
City = ReQuest("city")				' Needed only when coming from Misc/AuthorizedDealerFor.asp to display Starlite as a Garmin authoized dealer,
' in which case I need to build garmin URL:                        https://www8.garmin.com/cgi-bin/us_dealers.pl?state=MI&dealer_type=&city=OAK+PARK from
' URL to this file: https://www.starlite-intl.com/external.asp?URL=https://www8.garmin.com/cgi-bin/us_dealers.pl&state=MI&dealer_type=&city=OAK+PARK 
' Note required change from & to ? before state=
'Response.Write "<br>ExternalSiteURL = " & ExternalSiteURL
'Response.Write "<br>State = " & State
'Response.Write "<br>City = " & City

If City = "OAK PARK" Then
	City = "OAK+PARK"
	FullExternalSiteURL = ExternalSiteURL & "?state=" & State & "&city=" & City   ' Needed only when coming from Misc/AuthorizedDealerFor.asp to display Satrlite as a Garmin authoized dealer.
Else
	FullExternalSiteURL = ExternalSiteURL
End If

'Response.Write "<br>FullExternalSiteURL = " & FullExternalSiteURL
%>



<html>


<head>
	<link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
	<title></title>
	<meta name="keywords" content="">
	<meta http-equiv="content-type" content="text/html; charset=UTF-8">
	<meta http-equiv="content-language" content="en">
	<meta name="description" content="">
	<meta name="www.intelligineering.com" content="www.intelligineering.com">
</head>



<body >


<table style="border:0px solid blue;" width='1100'  bgcolor="" align='center'>		<% ' Start Table 1 %>
<tr><td>

<% InArea = "Products" %>

<!--#Include virtual="Misc/Header.INC"-->


<% '*********************************************************************************************************************** %>


<% ' Start Table 1.1 %>
<table style="border-right:1px solid #84bff1;" width='1120' cellpadding="0" cellspacing="0" align="center" > <% ' Start Table 1.1 %>
    <tr>
        <td class="Gradient2" width="223" valign="top" align="left">
            <!--#Include virtual="INC/LeftColumn.inc.asp"-->	
		</td>
					
					
        <!-- <td width="100%" background="Images/bluebackground2.jpg" valign=top> -->	
		<td width="100%" valign=top>
					<br>
					<% ' Start Table 1.1.2 %>
					<table border=0 cellpadding="0" cellspacing="0" align="center" width=100%>  
					<tr>
						<td valign="top" align="center">
						<iframe src=<%=FullExternalSiteURL%> width="98%" height="900">
							<p>Your browser does not support iframes.</p>
						</iframe> 
                		</td>
            		</tr>
            		</table>				
            		<% ' End Table 1.1.2 %>
            		   		
		</td>         
	</tr>
		 
</table>		
<% ' End Table 1.1 %>
       		
       		 
<!--#Include file="Misc/Footer.INC"-->

</td>
</tr>
	
	
</table>   
<% ' End Table 1 %>

</body>

</html>





