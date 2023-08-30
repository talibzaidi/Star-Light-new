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
<FORM name=thisForm  method="POST">




<%
Sub btnFindKeyword_onclick()
	Keyword = Trim(Request.Form("txtKeyword"))  
	Response.Write "Keyword = " & Keyword & "<br>"
	'ProductSQL = "SELECT * FROM Product WHERE PName LIKE '%" & CStr(Keyword) & "%' AND Cost <> 0 ORDER BY Cost"
	'ProductSQL = "SELECT * FROM Product WHERE (	 PName  LIKE '%" & CStr(Keyword) & "%' OR " & _
	'											"Descr  LIKE '%" & CStr(Keyword) & "%' OR " & _
	'											"ITEMID LIKE '%" & CStr(Keyword) & "%' OR " & _
	'											"Text1  LIKE '%" & CStr(Keyword) & "%') AND " & _
	'											"Cost <> 0 ORDER BY Cost"
	ProductSQL = "SELECT * FROM Product WHERE (	 PName  LIKE '%" & CStr(Keyword) & "%' OR " & _
												"Descr  LIKE '%" & CStr(Keyword) & "%' OR " & _
												"ITEMID LIKE '%" & CStr(Keyword) & "%') AND " & _
												"Cost <> 0 ORDER BY Cost"
	' Get an error when try to include:			"Text2  LIKE '%" & CStr(Keyword) & "%' OR " & _
	' probably because Text2 field is (often) NULL?
	Response.Write "ProductSQL = " & ProductSQL & "<br>"
	'Response.End
	Session("ProductSQL")= ProductSQL
	
	'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp
	' (nor the rsProduct.getCount method, which is apparently only available when using the recordset DTC, which I
	' have not (yet?) figured out how to use on the MS Access database (not the SQL Server database that I am used to) 
	' that this website uses).
	'ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE (PName LIKE '%" & CStr(Keyword) & "%' OR Descr LIKE '%" & CStr(Keyword) & "%') AND Cost <> 0"
	'ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE " & _
	'											  "( PName  LIKE '%" & CStr(Keyword) & "%' OR " & _
	'												"Descr  LIKE '%" & CStr(Keyword) & "%' OR " & _
	'												"ITEMID LIKE '%" & CStr(Keyword) & "%' OR " & _
	'												"Text1  LIKE '%" & CStr(Keyword) & "%') AND " & _
	'												"Cost <> 0"
	ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE " & _
												  "( PName  LIKE '%" & CStr(Keyword) & "%' OR " & _
													"Descr  LIKE '%" & CStr(Keyword) & "%' OR " & _
													"ITEMID LIKE '%" & CStr(Keyword) & "%') AND " & _
													"Cost <> 0"
	' Get an error when try to include:			"Text2  LIKE '%" & CStr(Keyword) & "%' OR " & _
	' probably because Text2 field is (often) NULL?
	Response.Write "ProductCountSQL = " & ProductCountSQL & "<br>"
	Session("ProductCountSQL")= ProductCountSQL
	Session("SummaryHeading")= "Keyword: " & CStr(Keyword)
	Response.Write Session("SummaryHeading")
	'Response.End
	
	Response.Redirect "searchsummary.asp?ShowPageNum=1"     
End Sub		' btnFindKeyword_onclick()


'_________________________________________________________________________________________

Sub btnFindProductName_onclick()
	ProductName = Trim(Request.Form("txtProductName"))  
	'ProductName = replace(ProductName , "-", "") 
	'ProductName = replace(ProductName , " ", "") 
	Response.Write "ProductName = " & ProductName & "<br>"
	ProductSQL = "SELECT * FROM Product WHERE PName LIKE '%" & CStr(ProductName) & "%' AND Cost <> 0 ORDER BY Cost"
	'ProductSQL = "SELECT PID, ItemID, PName, Pic1, Manufa, Descr, REPLACE(PName, '-', '') As PName2 FROM Product WHERE PName2 LIKE '%" & CStr(ProductName) & "%' AND Cost <> 0 ORDER BY Cost"
	'ProductSQL = "SELECT * FROM vwProducts WHERE PName2 LIKE '%" & CStr(ProductName) & "%' AND Cost <> 0 ORDER BY Cost"
	Response.Write "ProductSQL = " & ProductSQL & "<br>"
	'Response.End
	Session("ProductSQL")= ProductSQL
	
	'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
	ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE PName LIKE '%" & CStr(ProductName) & "%' AND Cost <> 0"
	'ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM vwProducts WHERE PName2 LIKE '%" & CStr(ProductName) & "%' AND Cost <> 0"
	'Response.Write "ProductCountSQL = " & ProductCountSQL & "<br>"
	Session("ProductCountSQL")= ProductCountSQL
	Session("SummaryHeading")= "Product Name: " & CStr(ProductName)
	Response.Write Session("SummaryHeading")
	'Response.End
	
	Response.Redirect "searchsummary.asp?ShowPageNum=1"     
End Sub		' btnFindKeyword_onclick()


'_________________________________________________________________________________________

Sub btnFindManufacturer_onclick()
	Manufacturer = Trim(Request.Form("Manufa"))
	Response.Write "Manufacturer = " & Manufacturer & "<br>"
	ProductSQL = "SELECT * FROM Product WHERE Manufa LIKE '%" & CStr(Manufacturer) & "%' AND Cost <> 0 ORDER BY Cost"
	Response.Write "ProductSQL = " & ProductSQL & "<br>"
	'Response.End
	Session("ProductSQL")= ProductSQL
	
	'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
	ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE Manufa LIKE '%" & CStr(Manufacturer) & "%' AND Cost <> 0"
	Response.Write "ProductCountSQL = " & ProductCountSQL & "<br>"
	Session("ProductCountSQL")= ProductCountSQL
	Session("SummaryHeading")= "Maunfacturer: " & Manufacturer
	Response.Write Session("SummaryHeading")
	'Response.End
	
	Response.Redirect "searchsummary.asp?ShowPageNum=1"     
End Sub		' btnFindManufacturer_onclick()



Sub btnFindCatAndSubCat_onclick()
	CatAndSubCat = Trim(Request.Form("CatAndSubCat"))
	Response.Write "CatAndSubCat = " & CatAndSubCat & "<br>"

	' Parse the CatAndSubCat string ...
	p2 = Instr(CatAndSubCat, "-") + 1		' Beginning of SID
	p3 = Instr(CatAndSubCat, "~") + 1		' Beginning of Cat Name.
	p4 = Instr(CatAndSubCat, "+") + 1		' Beginning of SubCat Name.
	'Response.Write "p2 = " & p2 & "<br>"
	'Response.Write "p3 = " & p3 & "<br>"
	'Response.Write "p4 = " & p4 & "<br>"
	CategoryID = Mid(CatAndSubCat, 1, p2-2)
	SubCategoryID = Mid(CatAndSubCat, p2, p3-p2-1)
	CatName = Mid(CatAndSubCat, p3, p4-p3-1)
	SubCatName = Mid(CatAndSubCat, p4, Len(CatAndSubCat) - p4 + 1)
	'Response.Write "CategoryID = " & CategoryID & "<br>"
	'Response.Write "SubCategoryID = " & SubCategoryID & "<br>"
	'Response.Write "CatName = " & CatName & "<br>"
	'Response.Write "SubCatName = " & SubCatName & "<br>"
	
	If SubCategoryID <> "0" Then		' User selected a subcategory (not a category).
		ProductSQL = "SELECT * FROM Product WHERE SID = " & CStr(SubCategoryID) & " AND Cost <> 0 ORDER BY Cost"
		'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
		ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE SID = " & CStr(SubCategoryID) & " AND Cost <> 0"
		Session("SummaryHeading")= "Subcategory: " & SubCatName
	Else								' User selected a category (not a subcategory).
		ProductSQL = "SELECT * FROM Product INNER JOIN SubArea ON Product.SID = SubArea.SID WHERE SubArea.AID = " & CStr(CategoryID) & " AND Cost <> 0 ORDER BY Cost"
		'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
		ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product INNER JOIN SubArea ON Product.SID = SubArea.SID WHERE SubArea.AID = " & CStr(CategoryID) & " AND Cost <> 0"
		Session("SummaryHeading")= "Category: " & CatName
	End If
	Response.Write "ProductSQL = " & ProductSQL & "<br>"
	Response.Write "ProductCountSQL = " & ProductCountSQL & "<br>"
	Session("ProductSQL")= ProductSQL
	Session("ProductCountSQL")= ProductCountSQL
	Response.Write Session("SummaryHeading")
	'Response.End
	Response.Redirect "searchsummary.asp?ShowPageNum=1"     
End Sub		' btnFindCatAndSubCat_onclick()
%>



<HTML>


<head>
<link rel="stylesheet" type="text/css" href="../../Misc/StyleSheet1.css">
<title>Star Lite Intl.: GPS, GPS Sensors, GPS engines, Garmin GPS, USglobal GPS, CB radios, 2-way radios, Scanners, Marine electronics, Audio, Video, Hand tools</title>
<meta name="keywords" content="GPS, GPS sensors, GPS sensor, GPS engine, global positioning system, GPS navigation, oem GPS, GPS oem, GPS18 sensor, GPS 18, GPS18, GPS18lvc, GPS system, car GPS, auto GPS, GPS equipment, differential GPS, WAAS, GPS receivers, buy GPS, GPS now, streetpilot, portable GPS, handheld GPS, ique m3, quest, streetpilot i5, forerunner 201, forerunner201, bu303, bu353, mr350, gv101, marine GPS, GPS receiver, GPS accessories, fish finder, navigator, GPS navigator, GPS software, GPS equipment, best GPS, pda GPS, garmin GPS, bluetooth GPS, global positioning, tracking GPS, fleet tracking GPS, GPS antenna, antenna, antennas, usglobal GPS, discount GPS, GPS on sale, navigation electronics, CB radio, CB radios, CB, CB on sale, 2-way radios, two way radios, walky talky, marine radio, marine electronics, radar, network, uniden CB, cobra CB, Midland CB, CB antennas, cb antenna, amateur radios, galaxy radio, magnum radio, radio scanner, radio scanners, scanner, flash memory, secure digital, compact flash, digital cameras, car audio, car stereos, stereo, car video, power amplifier, scanner antennas,scanner antenna, power supplies, power supply, regulated power supplies, DJ equipment, DJ, accessories, hand tools, mechanics tools, MIT tools, pyramid, pyle, silicon power, solarcon, fuji, fujifilm, nikon, olympus, panasonic, motorola, ranger, cherokee, wilson, firestick, valor, workman, RF, Limited, Solarcon, solarcon, Antron, antron, A99, gpk1, B100, saturn, 82fl, gzero48, gzero, little wil, w1000, w5000, a99ck, bm921, bm922, International, coax, RG8, K40, k40 antenna, K-40, Maco, Shakespeare, Imax, star, starlite, inverters, microphone, microphones, SSB, AM, FM, roger, beep, 40channels, 40 channels, car, system, jvc, speaker, speakers, stereo, amplifier, amplifiers, CD, player, wholesale, gps wholesale, new, new products, stereo, ham, ham radio, amateur, citizens, band, radar, detector, detectors, sales, international, international sales, export, power, supply, waterproof, rechargable, nicad, nimh, battery, gtl, ltd, dx, fm, ssb, cw, clarifier, bearcat, bearcat products, power, supply, rx, tx, transmit, transmitter, receive, receiver, receivers, SWR, frequency, counter, 10 meter, 11 meter, meters, citizens band, citizen-band, band, amateur-radio, RF limited, hustler, pioneer, discount, inexpensive, cheap, sony, caramplifier, caramplifiers, linear amplifier, amps, mobile, radio, radios, beam antenna, m103c, beams">
<meta name="description" content="GPS: Full line of garmin GPS, usglobal GPS, GPS sensors, oem gps, GPS receivers, gps board, gps navigation, tracking GPS, bluetooth GPS, GPS network, fish finders. CB: cb radios, walky-talky, cb antennas, cb accessories, 2-way radios, scanners. Digital: digital cameras, flash memory. car audio and video, DJ, hand tools, mechanics tools.">
<meta name="author" content="Star Lite International, LLC">
<meta name="copyright" content="2005 Star Lite International, LLC">
<meta name="revisit-after" content="10 days">
<meta name="distribution" content="global">
<meta name="robots" content="index,follow">
<meta name="rating" content="general">
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
<meta http-equiv="content-language" content="en">
<meta name="DC.Title" content="Star Lite Intl.: GPS, GPS Sensors, GPS engines, Garmin GPS, USglobal GPS, CB radios, 2-way radios, Amateur radios, Scanners, Marine electronics, Audio/video, Hand tools">
<meta name="DC.Description" content="GPS: Full line of garmin GPS, usglobal GPS, GPS sensors, oem gps, GPS receivers, gps board, gps navigation, tracking GPS, bluetooth GPS, GPS network, fish finders. CB: cb radios, walky-talky, cb antennas, cb accessories, 2-way radios, scanners. Digital: digital cameras, flash memory. car audio and video, DJ, hand tools, mechanics tools.">
<meta name="abstract" content="Complete selection of  GPS, GPS sensors, oem GPS and GPS accessories. Also wide selection of cb radios, scanners, antennas and automotive electronics, digital cameras, memory cards and hand tools.">

<link rel="shortcut icon" href="../favicon.ico" TYPE="image/ico">
<LINK REL="SHORTCUT ICON" HREF="../favicon.ico"><% ' <meta name="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds"> %><% ' <meta name="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds"> %>
</head>




<body topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">

<% InArea = "Products" %>

<!-- #INCLUDE VIRTUAL = "Misc/Header.inc" -->

<br><br><br><br><br><br>


<table align='center' border='0' cellspacing='0' cellpadding=5 width='940'>		<% ' Start Outer Table %>

<tr bgcolor='blue'>
<td height='15'>     
<font color='white'><b>Search By ...</b></font>
</td>
<td>
</td>
<td>
</td>
</tr>

<tr>
<td height='20'>     
</td>
<td>
</td>
<td>
</td>
</tr>


<tr>
<td height='50'>     
<b>Keyword / phrase:</b>
</td>
<td>

<input id="txtKeyword" maxLength="30" name="txtKeyword" size="30">
</td>
<td>
<!--METADATA TYPE="DesignerControl" startspan
<object id="btnFindKeyword" style="LEFT: 0px; WIDTH: 46px; TOP: 0px; HEIGHT: 27px" height="27" width="46" classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" dtcid="1">
	<param NAME="_ExtentX" VALUE="1217">
	<param NAME="_ExtentY" VALUE="714">
	<param NAME="id" VALUE="btnFindKeyword">
	<param NAME="Caption" VALUE="Find">
	<param NAME="Image" VALUE>
	<param NAME="AltText" VALUE>
	<param NAME="Visible" VALUE="-1">
	<param NAME="Platform" VALUE="256">
	<param NAME="LocalPath" VALUE="../">
</object>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnFindKeyword()
{
	btnFindKeyword.value = 'Find';
	btnFindKeyword.setStyle(0);
}
function _btnFindKeyword_ctor()
{
	CreateButton('btnFindKeyword', _initbtnFindKeyword, null);
}
</script>
<% btnFindKeyword.display %>

<!--metadata TYPE="DesignerControl" endspan-->
</td>
</tr>


<tr>
<td height='50'>     
<b>Product Name:</b>
</td>
<td>

<input id="txtKeyword" maxLength="30" name="txtProductName" size="30">
</td>
<td>
<!--METADATA TYPE="DesignerControl" startspan
<object id="btnFindProductName" style="LEFT: 0px; WIDTH: 46px; TOP: 0px; HEIGHT: 27px" height="27" width="46" classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" dtcid="2">
	<param NAME="_ExtentX" VALUE="1217">
	<param NAME="_ExtentY" VALUE="714">
	<param NAME="id" VALUE="btnFindProductName">
	<param NAME="Caption" VALUE="Find">
	<param NAME="Image" VALUE>
	<param NAME="AltText" VALUE>
	<param NAME="Visible" VALUE="-1">
	<param NAME="Platform" VALUE="256">
	<param NAME="LocalPath" VALUE="../">
</object>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnFindProductName()
{
	btnFindProductName.value = 'Find';
	btnFindProductName.setStyle(0);
}
function _btnFindProductName_ctor()
{
	CreateButton('btnFindProductName', _initbtnFindProductName, null);
}
</script>
<% btnFindProductName.display %>

<!--metadata TYPE="DesignerControl" endspan-->
</td>
</tr>



<!--  <form action="http://www.starlite-intl.com/scart/scart.asp" method="GET" id="form1" name="form1">   -->          
<tr>
<td height='50'>     
<b>Manufacturer:</b>
</td>
<td>                       
							<%
							MenuSQL = "Select Distinct Manufa from PRODUCT ORDER BY Manufa ASC"
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							Set rsProduct = Conn.Execute(MenuSQL)
							Set conn = Nothing
							%>

							<select name="Manufa" size="1">
								<option>Choose ...
								<option> 
								<%	Do While Not rsProduct.EOF 
									Manufacturer = rsProduct("Manufa")
									If Manufacturer <> "" Then %>
									<option value="<%=Manufacturer%>"><%=Manufacturer%>
								<%	End If
									rsProduct.MoveNext
									Loop
									rsProduct.Close
								%>
							</select>
							
</td>
<td>
							<input type='hidden' name='sar' value='Manufa'>
							<input type='hidden' name='SID' value='0'>
							<!--METADATA TYPE="DesignerControl" startspan
							<object id="btnFindManufacturer" style="LEFT: 0px; WIDTH: 46px; TOP: 0px; HEIGHT: 27px" height="27" width="46" classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" dtcid="3">
								<param NAME="_ExtentX" VALUE="1217">
								<param NAME="_ExtentY" VALUE="714">
								<param NAME="id" VALUE="btnFindManufacturer">
								<param NAME="Caption" VALUE="Find">
								<param NAME="Image" VALUE>
								<param NAME="AltText" VALUE>
								<param NAME="Visible" VALUE="-1">
								<param NAME="Platform" VALUE="256">
								<param NAME="LocalPath" VALUE="../">
							</object>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnFindManufacturer()
{
	btnFindManufacturer.value = 'Find';
	btnFindManufacturer.setStyle(0);
}
function _btnFindManufacturer_ctor()
{
	CreateButton('btnFindManufacturer', _initbtnFindManufacturer, null);
}
</script>
<% btnFindManufacturer.display %>

							<!--metadata TYPE="DesignerControl" endspan-->
</td>
</tr>



<tr>
<td height='50'>     
<b>Product Category or Subcategory:</b>
</td>
<td>
							<%
							'MenuSQL = "SELECT * FROM Subarea INNER JOIN Area51 ON Subarea.AID = Area51.AID ORDER BY Subarea.AID ASC, Subarea.SID ASC"
							MenuSQL = "SELECT * FROM Subarea INNER JOIN Area51 ON Subarea.AID = Area51.AID ORDER BY AreaName ASC, Subname ASC"		
							Set conn = Server.CreateObject("ADODB.Connection")
							Conn.Open Session("ConnectionString")
							Set rsSubArea = Conn.Execute(MenuSQL)
							Set conn = Nothing
							AIDprevious =  -1
							%>
							
							<select name="CatAndSubCat" size="1">
								<option>Choose ...
								<%	Do While Not rsSubArea.EOF
									SID = rsSubArea("SID")				' i.e. ID of Product SubArea or SubCategory.
									AID = rsSubArea("AID")				' i.e. ID of Product Area or Category.
									AreaName = rsSubArea("AreaName")	' i.e. Name of Product Area or Category.
									SubCategorgyName = rsSubArea("Subname") 
									  
									If SID <> "" AND AID <> 0 AND SubCategorgyName <> "" AND SubCategorgyName <> "test" Then 
										If AID <> AIDprevious Then  
											Response.Write "<option value='-1'> " 
											Response.Write "<option value='" & AID & "-" & "0" & "~" & AreaName & "+" & "NULL" & "'>" & AreaName
											Response.Write "<option value='" & AID & "-" & SID & "~" & AreaName & "+" & SubCategorgyName & "'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & SubCategorgyName
										Else
											Response.Write "<option value='" & AID & "-" & SID & "~" & AreaName & "+" & SubCategorgyName & "'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & SubCategorgyName
										End If 
										AIDprevious = AID
									End If 
									
									rsSubArea.MoveNext
									Loop
									rsSubArea.Close 
								%>        
							</select>
</td>
<td>
							<input type='hidden' name='Area' value='iii'>
							<input type='hidden' name='SID' value='0'>
							<!--METADATA TYPE="DesignerControl" startspan
							<object id="btnFindCatAndSubCat" style="LEFT: 0px; WIDTH: 46px; TOP: 0px; HEIGHT: 27px" height="27" width="46" classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" dtcid="4">
								<param NAME="_ExtentX" VALUE="1217">
								<param NAME="_ExtentY" VALUE="714">
								<param NAME="id" VALUE="btnFindCatAndSubCat">
								<param NAME="Caption" VALUE="Find">
								<param NAME="Image" VALUE>
								<param NAME="AltText" VALUE>
								<param NAME="Visible" VALUE="-1">
								<param NAME="Platform" VALUE="256">
								<param NAME="LocalPath" VALUE="../">
							</object>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnFindCatAndSubCat()
{
	btnFindCatAndSubCat.value = 'Find';
	btnFindCatAndSubCat.setStyle(0);
}
function _btnFindCatAndSubCat_ctor()
{
	CreateButton('btnFindCatAndSubCat', _initbtnFindCatAndSubCat, null);
}
</script>
<% btnFindCatAndSubCat.display %>

							<!--metadata TYPE="DesignerControl" endspan-->

</td>
</tr>
				
							
</table>			<% ' End Outer Table %>


</BODY>


<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>


</HTML>
