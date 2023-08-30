

<% 
'-----------------------------------------------------------------------------------------------------------------------------

' DROP-DOWN MENUS SECTION
' 5/31/09: Note: I had a great deal of trouble getting this menu bar to work properly in IE 7, although it DID work
' fine in Firefox and Chrome. In IE 7 the menu bar would render ok and top-level items on the bar could be clicked on, 
' but the drop-down menus would not open - even though the version in file 
' /mytests/navbars/suckerfish/sfdropdownmenus1.inc.asp
' did fully work in IE 7. Finally I realized that the crucial difference was that 
' /mytests/navbars/suckerfish/sfdropdownmenus1.inc.asp
' contained the following 2 lines at the top ...
' <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
' "https://www.w3.org/TR/html4/loose.dtd">
' whereas my Mit Mazel files that this navBar1.inc was being included into, did not contain them. I don't (yet) know 
' why those 2 lines make a difference to IE 7. Anyway, looks like I will need to add those 2 lines at the top of every 
' Mit Mazel file where I want this drop-down menu to appear.
%>


<%	' 6/3/09: Now that file navbar1a.inc.asp (old tabbed menu bars) is being phased out in favour of this file 
    ' navbar1b.inc.asp (new css-based drop-down menus), this section was copied from navbar1a.inc.asp just because some 
	' of these contants are still needed on some webpages. 

	
	TextColorActive = "white"    ' "silver"   '  "firebrick"
	FontFace="Verdana"

	' 2/11/09: This is a duplication of the Session.Timeout command in global.asa file, because it was not apparently taking effect from that file.
	Session.Timeout 		= 60	' minutes
	'Response.Write "<br>Session.Timeout = " & Session.Timeout
%>



<!-- 4/1/10, BN: See https://users.tpg.com.au/j_birch/plugins/superfish/# -->

<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.4.2/jquery.min.js"></script>

<script type="text/javascript" src="https://www.starlite-intl.com/Misc/Navbars/suckerfish/superfish.js"></script>
 
<!--[if lte IE 6]>
<script type="text/javascript" src="https://www.starlite-intl.com/Misc/Navbars/suckerfish/suckerfish_ie.js"></script>
<![endif]-->



<% 
' 3/20/10: To allow context-sensitive highlighting of tabs on the navbar.
ActiveHeaderColor = "slateblue"
If InArea="Home" 				Then HomeStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 				End If
If InArea="Products" 			Then ProductsStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 			End If
If InArea="Specials" 			Then SpecialsStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 			End If
If InArea="WhatsNew?" 			Then WhatsNewStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 			End If
If InArea="ContactUs" 			Then ContactUsStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 			End If
If InArea="Terms" 				Then TermsStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 				End If
If InArea="Links" 				Then LinksStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 				End If
'If InArea="GiftCertificates" 	Then GiftCertificatesStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 	End If
If InArea="ShoppingCart" 		Then ShoppingCartStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 		End If


' The following extra Session("UserType") = "Admin" text is needed for when viewing Login.asp because that file does NOT set InArea to Admin. If it did, user would get Admin nav bar BEFORE having logged in.
' A very nice fringe benefit is that the Admin tab will remain highlighted TOGETHER with the tab for any other (non-Admin) page when go there while logged in as Admin! 
If (InArea="Admin") OR (Session("UserType") = "Admin")		Then AdminStyle="style=BACKGROUND-COLOR:" & ActiveHeaderColor 		End If
%>



<table align='center'>
<tr>
	<td>
	<div <%=HomeStyle%>><a href="https://www.starlite-intl.com/mobile/"><font color="white">Home</font></a></div>
	<br />
	<div <%=XXXHomeStyle%>><a href="https://www.starlite-intl.com/mobile/OEM_GPS_sensors/OEM_GPS_sensors.asp"><font color="white">OEM GPS Sensors</font></a></div>
	<br />
	<div <%=ContactUsStyle%>><a href="https://www.starlite-intl.com/mobile/Misc2/contact.asp"><font color="white">Contact Us</font></a></div>
	<br />
	<div <%=TermsStyle%>><a href="https://www.starlite-intl.com/mobile/Misc2/Terms_and_Conditions.asp"><font color="white">Terms & Conditions</font></a></div>
	</td>
</tr>
</table>

