<%	' InArea = "Members"			' InArea is set separately on each web page that inludes the present file. 
	' InSubArea = "None"			' InSubArea is set separately on each web page that inludes the present file.
		
	TextColorActive = "white"    ' "silver"   '  "firebrick"
	FontFace="Verdana"
%>




<%
If TRUE Then 
%>
<!-- 
Suckerfish drop-down menus ...
#INCLUDE VIRTUAL = "mobile1/Misc/Navbars/navbar1b.inc.asp"  
	[BN, 11/16/20] I don't know why the above INCLUDE uses the mobile1 version. But it seems to be needed.
INCLUDE VIRTUAL = "/Misc/Navbars/navbar1b.inc.asp"
--> 
<% End If %>