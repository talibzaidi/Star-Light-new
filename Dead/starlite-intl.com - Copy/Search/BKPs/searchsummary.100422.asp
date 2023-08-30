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



<script LANGUAGE="vbscript" RUNAT="Server">

Sub thisPage_onenter()
End Sub   ' thisPage_onenter


' VB does not have a built-in ceiling function. I found this one at http://www.visualbasicforum.com/t54042.html.
' [6/21/04] Copied from my page Products.asp.
Function Celg(Number) 
   Celg = -Int(-Number)
End Function



Sub thisPage_onexit()
'If rsProduct.isOpen() Then rsProduct.close()
End Sub
</script>



<script language="javascript">
// Based on summary page from www.futuresimchas.com.

	function pNext(sObj, NumMembersPerPage, AID, SID){
		if (sObj.selectedIndex < (sObj.options.length -1)){
			sObj.options[sObj.selectedIndex+1].text="Loading "+sObj.options[sObj.selectedIndex+1].text;
			sObj.options[sObj.selectedIndex+1].selected=true;
			pGo(sObj, NumMembersPerPage, AID, SID);
		} else {
			alert('End of Results Reached.');
		}
	}
	
	function pPrev(sObj, NumMembersPerPage, AID, SID){
		if (sObj.selectedIndex > 0){
			sObj.options[sObj.selectedIndex-1].text="Loading "+sObj.options[sObj.selectedIndex-1].text;
			sObj.options[sObj.selectedIndex-1].selected=true;
			pGo(sObj, NumMembersPerPage, AID, SID);
		} else {
			alert('Beginning of Results Reached.');
		}
	}
	
	function pGo(sObj, NumMembersPerPage, AID, SID){
	// The next line is because the values in each option of the menu are not pages per se, 
	// but the number (in consecutive order in those found by the search; not MemberID) of the first member on the page.
	SelectedPage = (sObj.options[sObj.selectedIndex].value - 1 )/ NumMembersPerPage + 1;   
	//location.href='searchsummary.asp?ShowPageNum='+SelectedPage;
	location.href='searchsummary.asp?AID='+AID+'&SID='+SID+'&ShowPageNum='+SelectedPage;
	}
</script>


<script LANGUAGE="vbscript" RUNAT="Server">
' 4/21/10: Based on a copy of Sub btnFindCatAndSubCat_onclick() from search/search.asp. 
' This was added here so as to work with the new suckerfish drop-down menu for Products that I added on 4/20/10.
' This uses Session variables, which is klunky, but that is just because of how it worked in search/search.asp to pass stuff from there to here.
' Sub btnFindCatAndSubCat_onclick()
Sub PrepareNeededStuff(CategoryID, SubCategoryID, CatName, SubCatName)	

	If SubCategoryID <> 0 Then		' User selected a subcategory, not a category.
		ProductSQL = "SELECT * FROM Product WHERE SID = " & CStr(SubCategoryID) & " AND Cost <> 0 ORDER BY Cost"
		'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
		ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE SID = " & CStr(SubCategoryID) & " AND Cost <> 0"
		Session("SummaryHeading") = "Subcategory: " & SubCatName
	Else								' User selected a category, not a subcategory.
		ProductSQL = "SELECT * FROM Product INNER JOIN SubArea ON Product.SID = SubArea.SID WHERE SubArea.AID = " & CStr(CategoryID) & " AND Cost <> 0 ORDER BY Cost"
		'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
		ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product INNER JOIN SubArea ON Product.SID = SubArea.SID WHERE SubArea.AID = " & CStr(CategoryID) & " AND Cost <> 0"
		Session("SummaryHeading") = "Category: " & CatName
	End If
	
	Session("ProductSQL")= ProductSQL
	Session("ProductCountSQL")= ProductCountSQL 
	   
End Sub		' FindCatAndSubCat()
</script>



<HTML>


<HEAD>
<link rel="stylesheet" type="text/css" href="../../Misc/StyleSheet1.css">
<meta http-equiv="content-type" content="text/html;charset=iso-8859-1">
<title>Star Lite Intl. - GPS, GPS Sensors, GPS accessories, CB radios, 2-way radios, Marine electronics, Flash memory, MP3, audio/video, hand tools</title>
<meta NAME="Description" CONTENT="Large selection of - GPS, GPS sensors, GPS accessories, GPS OEM, PDA, tracking gps, bluetooth gps, fish finders, sounders, cb radios and Walky-talky, flash memory, MP3,  radio scanners, digital cameras, car audio and car video, DJ, hand tools, mechanics tools.">
<meta NAME="Keywords" CONTENT="gps, gps navigation, gps sensors, OEM gps, GPS accessories, cb radio, 2-way radios, cb radios, cb, garmin gps, global positioning, Walkytalky, mobile tracking, fleet  tracking, usglobasat gps, bluetooth, flash memory, gmrs, marine radios, navigation electronics, 2-way radios, radio scanners, marine radios, car audio, car stereos, power amplifiers, antennas, power  supplies, regulated power supplies, dj, accessories, mechanics tools, hand tools, uniden, cobra, midland, mit, pyramid, pyle, solarcon">
</HEAD>



<body topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">

<%
ShowPageNum = Request.QueryString("ShowPageNum")
AID 		= Request.QueryString("AID")		' i.e. ID of Product Area or Category.
SID 		= Request.QueryString("SID")		' i.e. ID of Product SubArea or SubCategory.
'Response.Write "<br>AID = "	& AID 	
'Response.Write "<br>SID = "	& SID	
If SID = "" Then SID = 0 End If					' Need an integer value for SID, for use in the javascript functions pPrev, pNext, pGo.
'Response.Write "<br>SID = "	& SID

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Session("ConnectionString")

If TRUE Then 	' This block is always valid, whether user selected a category or a subcategory.
	CatSQL 		= "SELECT AreaName FROM Area51 WHERE AID=" & AID
	'Response.Write "<br>CatSQL = " & CatSQL 
	Set rsCat 	= Conn.Execute(CatSQL)
	CatName 	= rsCat("AreaName")					' i.e. Name of Product Area or Category.
	'Response.Write "<br>CatName = "	& CatName 
End If

If SID <> 0 Then	' This occurs when user selected a subcategory, not a category.
	SubCatSQL 	= "SELECT Subname FROM Subarea WHERE AID=" & AID & " AND SID=" & SID
	'Response.Write "<br>SubCatSQL = " & SubCatSQL 
	Set rsSubCat = Conn.Execute(SubCatSQL)
	SubCatName 	= rsSubCat("Subname")			' i.e. Name of Product SubArea or SubCategory.
	'Response.Write "<br>SubCatName = "	& SubCatName 
End If

	 
PrepareNeededStuff AID, SID, CatName, SubCatName 
%>


<%
ProductSQL = Session("ProductSQL")
'Response.Write "ProductSQL = " & ProductSQL & "<br>"

Set rsProduct = Conn.Execute(ProductSQL)
'Response.Write "rsProduct.RecordCount = " & rsProduct.RecordCount & "<br>"
'Response.Write "rsProduct.PageCount = " & rsProduct.PageCount & "<br>"
'Response.Write "rsProduct.PageSize = " & rsProduct.PageSize & "<br>"
'Response.Write "Response.Write rsProduct(PName) = " & rsProduct("PName")  & "<br>"


'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount method to work
' (nor the rsProduct.getCount method, which is apparently only available when using the recordset DTC, which I
' have not (yet?) figured out how to use on the MS Access database (not the SQL Server database that I am used to) 
' that this website uses).
ProductCountSQL = Session("ProductCountSQL")
'Response.Write "ProductCountSQL = " & ProductCountSQL & "<br>"
Set rsProductCount = Conn.Execute(ProductCountSQL)
'Response.Write "<br>Response.Write rsProductCount(cnt) = " & rsProductCount("Cnt")  & "<br>"
TotalNumMembers = rsProductCount("Cnt")
'Response.Write "TotalNumMembers = " & TotalNumMembers & "<br>"
'Respond.End

' From http://www.asp101.com/samples/db_count.asp.
' This does say RecordCount will work, but it doesn't!
'If rsProduct.Supports(adBookmark) Or rsProduct.Supports(adApproxPosition) Then
'	Response.Write "RecordCount will work!"
'End If

Set Conn = Nothing

%>


<% InArea = "Products" %>

<!-- #INCLUDE VIRTUAL = "Misc/Header.inc" -->


<% ' Build drop-down menu of the page numbers for the user to hyperlink to different pages of the summary data.
    ' Based on summary page from www.futuresimchas.com.
    ' This section for the menu at the top of the page should be identical to the one below for the menu at 
    ' the bottom of the page, except for use of TopMenu instead of BottomMenu for the form name.

 ' TotalNumMembers = rsvwMembers.getCount()
 NumRows = 100        ' This is not necessarily the number of rows on the last page.
 'NumCols = 1
 MembersPerPage = NumRows    ' * NumCols
 'Response.Write "MembersPerPage = " & MembersPerPage 
 StartRecord = (ShowPageNum - 1) * NumRows + 1
 
 
 'SQL = Session("BooksSQL")
 'Response.Write "SQL = " & SQL
 
 'rsProduct2.setSQLText(ProductSQL)
 'If  rsProduct2.isOpen() Then rsProduct2.close() End If
 'rsProduct2.Open
 
'TotalNumMembers = rsProduct.getCount() ' rsvwMembersB.getParameter(4) ' The TOTAL number of members that match the SQL query. This is the first time I have used getParameters method of a recordset DTC.
'TotalNumMembers = rsProduct.RecordCount
'Response.Write "<br>TotalNumMembers = " &  TotalNumMembers 
If TotalNumMembers = 0 Then ' Response.Redirect "TrySearchAgain.asp" 
  Response.Write "<br><br><br><br><br><br><br><br>"
  Response.Write "<center>"
  Response.Write "<font size='4' color='blue'><b>Sorry. Nothing was found.</b></font><br><br>"
  Response.Write "<font size='4' color='blue'><b><a href='http://www.starlite-intl.com/search/search.asp'>Click here</a> to try a different search.</b></font>"
  Response.Write "</center>"
Else  ' i.e. TotalNumMembers > 0
%>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="95%">
	<tr><td height='10pt'></td></tr>
	<tr>
		<td align="middle">
		<font color='indigo' size=4><b><%=Session("SummaryHeading")%></b></font>
		<br><br>There are <strong> <font color="#000080"><% =TotalNumMembers %> </font> </strong> products. 
		Click on the picture to view a product's details and price.
		</td>
</table>
 
<% 
 NumPages = Celg(TotalNumMembers / NumRows)   ' Using Ceiling function I found on the web; see above.
'Response.Write "NumPages = " & NumPages & "<br>"
'Response.Write "ShowPageNum = " & ShowPageNum & "<br>"
 
 MembersOnLastPage = TotalNumMembers Mod NumRows
 If MembersOnLastPage = 0 Then
   MembersOnLastPage = MembersPerPage
 End If
 
 'If Int(ShowPageNum) = Int(NumPages) Then    ' i.e. If displaying the last page.
 '   ' Response.Write "On last page."
 '   NumRows = MembersOnLastPage
 'End If
 
       
       
 If NumPages > 1 Then
 	AID = Request.QueryString("AID")   	' I don't know why I have to do this again here. But If I don't, it gets the wrong AID from somewhere! Weird.
										' I think it's something to do with the order in which Javascript and ASP are evaluated on a page.
 	SID = Request.QueryString("SID")   	' Ditto. These re-evaluations of AID and SID are NOT needed in the bottom version (BottomMenu) of this below.	
	If SID = "" Then SID = 0 End If		' Need an integer value for SID, for use in the javascript functions pPrev, pNext, pGo.
 	'Response.Write "<br>AID=" & AID
 	'Response.Write "<br>SID=" & SID
 	
    Response.Write "<form name='TopMenu'>"
    Response.Write "<center>"
    
    Response.Write "<A href='javascript:pPrev(TopMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'><img border=0 hspace='5' align='middle' src='../images/NavImages/ArrowLeft.gif'></a>"
    Response.Write "<A href='javascript:pPrev(TopMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'>Previous Page</a>"
 	Response.Write "&nbsp;&nbsp;"  	
    Response.Write "<select name='Page' onChange='pGo(TopMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'>"
 	For i = 1 To NumPages - 1
 	If i <> CInt(ShowPageNum) Then
 		Response.Write "<option value='" & (i-1)*MembersPerPage + 1 & "'>Page " & i & " (" & (i-1)*MembersPerPage + 1 & "-" & i*MembersPerPage & ")</option>"
 		Else Response.Write "<option selected value='" & (i-1)*MembersPerPage + 1 & "'>Page " & i & " (" & (i-1)*MembersPerPage + 1 & "-" & i*MembersPerPage & ")</option>"
 	End If
 	Next
 	
 	i = NumPages    ' The last page is a special case because it may not be full.
 	If i <> CInt(ShowPageNum) Then
 		Response.Write "<option value='" & (i-1)*MembersPerPage + 1 & "'>Page " & i & " (" & (i-1)*MembersPerPage + 1 & "-" & (i-1)*MembersPerPage + MembersOnLastPage & ")</option>"
 		Else Response.Write "<option selected value='" & (i-1)*MembersPerPage + 1 & "'>Page " & i & " (" & (i-1)*MembersPerPage + 1 & "-" & (i-1)*MembersPerPage + MembersOnLastPage & ")</option>"
 	End If
 	
 	Response.Write "</select>"
 	Response.Write "&nbsp;&nbsp;" 
	Response.Write "<A href='javascript:pNext(TopMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'>Next Page</a>"
 	Response.Write "<A href='javascript:pNext(TopMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'><img border=0 hspace='5' align='middle' src='../images/NavImages/ArrowRight.gif'></a>"
 	Response.Write "</center>"
 	Response.Write "</form>"
 Else
	Response.Write "<br><br>"
 End If 
 
%>


<%
 StartRecordNumber = (ShowPageNum - 1) * MembersPerPage + 1				' ShowPageNum is set above using ShowPageNum = Request.QueryString("ShowPageNum").

'rsProduct.moveAbsolute(StartRecordNumber)
' 8/17/05: Iterate over ALL records, but only display the subset needed for this page.
' This is a kludge because I can't get rsProduct.moveAbsolute(StartRecordNumber) method to work.
' It is apparently only available when using the recordset DTC, which I
' have not (yet?) figured out how to use on the MS Access database (not the SQL Server database that I am used to) 
' that this website uses.
rsProduct.moveFirst
endRecordNumber = StartRecordNumber + NumRows - 1
parity = -1
Response.Write "<table align='center' cellpadding='5' cellspacing='0' border='0' width=" & PageWidth & ">"
For row = 1 to TotalNumMembers ' NumRows 
	If row >= StartRecordNumber AND row <= endRecordNumber Then
	graphicFile = "../Imi/" & rsProduct("Pic1")
	parity = - parity
	If parity = 1 Then color = "gainsboro" Else color = "white" End If
	PID = rsProduct("PID")
	ItemID = rsProduct("ItemID")
	ProductName = rsProduct("PName")
	Manufacturer = rsProduct("Manufa")
	Description = rsProduct("Descr")
	'Cost = rsProduct("Cost")
	Response.Write "<tr bgcolor='" & color & "'>"
	Response.Write "<td valign='top'><font size='1'>" & row & "</font></td>" 
	Response.Write "<td valign='top' align='left'><a href='http://www.starlite-intl.com/Detail.asp?pid=" & PID & "'><img hspace='20' align='left' border='0' src='" & graphicFile & "'></td>" 
	Response.Write "<td valign='top'><b><font color='indigo'>" & ProductName & "</font></b></a><br><br><font size=1>" & Manufacturer & " " & ItemID & "</font></td>" 
	Response.Write "<td valign='top'>" & Description & "</td>" 
	'Response.Write "<td valign='top'>$" & Cost & "</td>" 
	Response.Write "</tr>"
	End If
	rsProduct.moveNext
Next ' row
Response.Write "</table>"


' ******************************************
 NumPages = Celg(TotalNumMembers / NumRows)   ' Using Ceiling function I found on the web; see above.
'Response.Write "NumPages = " & NumPages & "<br>"
'Response.Write "ShowPageNum = " & ShowPageNum & "<br>"
 
 MembersOnLastPage = TotalNumMembers Mod NumRows
 If MembersOnLastPage = 0 Then
   MembersOnLastPage = MembersPerPage
 End If
 
 'If Int(ShowPageNum) = Int(NumPages) Then    ' i.e. If displaying the last page.
 '   ' Response.Write "On last page."
 '   NumRows = MembersOnLastPage
 'End If
 
       
 If NumPages > 1 Then
    Response.Write "<form name='BottomMenu'>"
    Response.Write "<center>"
    
    Response.Write "<A href='javascript:pPrev(BottomMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'><img border=0 hspace='5' align='middle' src='../images/NavImages/ArrowLeft.gif'></a>"
    Response.Write "<A href='javascript:pPrev(BottomMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'>Previous Page</a>"
 	Response.Write "&nbsp;&nbsp;"  	
    Response.Write "<select name='Page' onChange='pGo(BottomMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'>"
 	For i = 1 To NumPages - 1
 	If i <> CInt(ShowPageNum) Then
 		Response.Write "<option value='" & (i-1)*MembersPerPage + 1 & "'>Page " & i & " (" & (i-1)*MembersPerPage + 1 & "-" & i*MembersPerPage & ")</option>"
 		Else Response.Write "<option selected value='" & (i-1)*MembersPerPage + 1 & "'>Page " & i & " (" & (i-1)*MembersPerPage + 1 & "-" & i*MembersPerPage & ")</option>"
 	End If
 	Next
 	
 	i = NumPages    ' The last page is a special case because it may not be full.
 	If i <> CInt(ShowPageNum) Then
 		Response.Write "<option value='" & (i-1)*MembersPerPage + 1 & "'>Page " & i & " (" & (i-1)*MembersPerPage + 1 & "-" & (i-1)*MembersPerPage + MembersOnLastPage & ")</option>"
 		Else Response.Write "<option selected value='" & (i-1)*MembersPerPage + 1 & "'>Page " & i & " (" & (i-1)*MembersPerPage + 1 & "-" & (i-1)*MembersPerPage + MembersOnLastPage & ")</option>"
 	End If
 	
 	Response.Write "</select>"
 	Response.Write "&nbsp;&nbsp;" 
	Response.Write "<A href='javascript:pNext(BottomMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'>Next Page</a>"
	Response.Write "<A href='javascript:pNext(BottomMenu.Page," & MembersPerPage & ",AID=" & AID & ",SID=" & SID & ");'><img border=0 hspace='5' align='middle' src='../images/NavImages/ArrowRight.gif'></a>"
 	Response.Write "</center>"
 	Response.Write "</form>"
 Else
	Response.Write "<br><br>"
 End If 
 
 Response.Write "<br>"
' ******************************************




End If   ' This is the end of: If TotalNumMembers = 0 Then ...
%>


</BODY>


<% ' VI 6.0 Scripting Object Model Enabled %><% EndPageProcessing() %>


</HTML>
