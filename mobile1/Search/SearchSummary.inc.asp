


<% 
' 1/4/13: This file is essentially a copy of Search/searchsummary.asp, but specialized for  
' AID = "45"	  ' GPS Navigation, GPS Sensors, OEM, FishFinders, Maps
' SID = "173"     ' GPS - OEM: Sensors / Boards / TracPacs
%>



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


<%
' 4/21/10: Based on a copy of Sub btnFindCatAndSubCat_onclick() from search/search.asp. 
' This was added here so as to work with the new suckerfish drop-down menu for Products that I added on 4/20/10.

Sub InitializeForCatSubCatSearch(CategoryID, SubCategoryID)

    'Response.Write "<br><br>CategoryID = "	& CategoryID 
    'Response.Write "<br>SubCategoryID = "	& SubCategoryID 

	If TRUE Then 	' This block is always valid, whether user selected a category or a subcategory.
		Set Conn 	= Server.CreateObject("ADODB.Connection") 
		Conn.Open Session("ConnectionString")
		CatSQL 		= "SELECT AreaName, AreaDesc FROM Area51 WHERE AID=" & CategoryID 
	      'Response.Write "<br>CatSQL = " & CatSQL 
		Set rsCat 	    = Conn.Execute(CatSQL)

		' [11/26/20] Re following if-then-else see https://p2p.wrox.com/classic-asp-basics/93350-why-use-bof-eof.html
		' [12/3/20] The following if-then was added.
		if rsCat.EOF AND rsCat.BOF then						
			Response.Write "<div style='font-family:verdana; font-size:20px; text-align:center'>"
			Response.Write "<br><br><p>No products were found</p>"
			Response.Write "<br><a href='https://www.starlite-intl.com/mobile1/'>Go to home page</a>"
			Response.Write "</div>"
			Response.End
		else
			'Response.Write "<br>Records were found"
			CatName 	    = rsCat("AreaName")					' i.e. Name of Product Area or Category.
			CatDesc         = rsCat("AreaDesc")                 ' i.e. Description of Product Area or Category.
			'Response.Write "<br>CatName = "	& CatName 
			'Response.Write "<br>CatDesc = "	& CatDesc 
		end if


		CatName 	= rsCat("AreaName")					' i.e. Name of Product Area or Category.
        	CatDesc         = rsCat("AreaDesc")                 ' i.e. Description of Product Area or Category.
	       'Response.Write "<br><br>CatName = "	& CatName 
               'Response.Write "<br>CatDesc = "	& CatDesc 
	End If
	
	If SID <> 0 Then	' This occurs when user selected a subcategory, not a category.
		SubCatSQL 	= "SELECT Subname, SubDesc FROM Subarea WHERE AID=" & CategoryID & " AND SID=" & SubCategoryID
		'Response.Write "<br>SubCatSQL = " & SubCatSQL 
		Set rsSubCat    = Conn.Execute(SubCatSQL)
		SubCatName      = rsSubCat("Subname")			    ' i.e. Name of Product SubArea or SubCategory.
        	SubCatDesc 	= rsSubCat("SubDesc")			    ' i.e. Description Product SubArea or SubCategory.
	      'Response.Write "<br>SubCatName = "	& SubCatName 
              'Response.Write "<br>SubCatDesc = "	& SubCatDesc
	End If
	
	If SubCategoryID <> 0 Then		' User selected a subcategory, not a category.
		ProductSQL = "SELECT * FROM Product WHERE SID = " & CStr(SubCategoryID) & " AND Cost <> 0 ORDER BY Cost"
		'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
		ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE SID = " & CStr(SubCategoryID) & " AND Cost <> 0"
		'SummaryHeading = "<table><tr><td align=right><b>Category:</b></td><td>&nbsp;&nbsp;</td><td align=left>" & CatName & "</td></tr>" 
		'SummaryHeading = SummaryHeading & "<tr><td align=right><b>Subcategory:</b></td><td>&nbsp;&nbsp;</td><td align=left>" & SubCatName & "</td></tr></table>"
		'Session("SummaryHeading")   =  "OEM GPS Sensors" 'SummaryHeading 
		Session("SummaryHeading")   =  SubCatName
        	'Session("Description")      = SubCatDesc
	Else								' User selected a category, not a subcategory.
		ProductSQL = "SELECT * FROM Product INNER JOIN SubArea ON Product.SID = SubArea.SID WHERE SubArea.AID = " & CStr(CategoryID) & " AND Cost <> 0 ORDER BY Cost"
        	'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
		ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product INNER JOIN SubArea ON Product.SID = SubArea.SID WHERE SubArea.AID = " & CStr(CategoryID) & " AND Cost <> 0"
		'Session("SummaryHeading")   = "<b>Category:</b>&nbsp;&nbsp;" & CatName
		Session("SummaryHeading")   = CatName
        	Session("Description")      = CatDesc
	End If
	
    'Response.Write "<br>ProductSQL = "      & ProductSQL
    'Response.Write "<br>ProductCountSQL = " & ProductCountSQL
	Session("ProductSQL")		= ProductSQL
	Session("ProductCountSQL")	= ProductCountSQL
	Set Conn 	= Nothing
	
End Sub		' InitializeForCatSubCatSearch
%>



<%
' 4/25/10: Based on a copy of Sub btnFindKeyword_onclick() from search/search.asp.

' Keyword0 has already had its dashes and spaces removed. Keyword is the original version.
Sub InitializeForKeywordSearch(Keyword0, Keyword)  
	
	' Replace(PName, 'a', 'b') below causes an error.
	'From http://bytes.com/topic/access/answers/209056-using-replace-function-ado-access-db-visual-basic-6-a
	'> The ability to use functions like Replace within an update query is a
	'> 'special trick' which MS Access can do, but you cannot use this from
	'> vb/ado. Your options are:
	'> If you are replacing something easy like the first letter in the word,
	'> then you can use functions like left, right, mid, etc which will work from
	'> vb/ado.
	'> If the replace is more complicated, you will need to create an updateable
	'> recordset, looping through and updating each record. Depending on how
	'> many records you have, you may notice a drop in speed with this approach.
	'> However, if it's only a few thousand records, I guess you'll hardly notice
	'> the difference.
	
	' From http://www.keyongtech.com/398440-sql-replace-function-does-not ...
	'There is no 'REPLACE' function in Jet SQL. The query works in Access because
	'Access is using the VBA Replace function, but that function can not be used
	'in Jet queries when they are executed outside of the Microsoft Access
	'environment. You may be able to achieve the same result in a query using
	'string-chopping functions (Left, Right, Mid, etc) or you may need to do the
	'replace in code by opening a recordset and looping through the records.
	
	'ProductSQL = "SELECT PName, Replace(PName, 'a', 'b'), Descr, ITEMID, Manufa, Cost, Pic1, PID FROM Product " &_
	'ProductSQL = "SELECT PName, Descr, ITEMID, Manufa, Cost, Pic1, PID FROM Product " &_
	'	"WHERE ( PName  LIKE '%" & CStr(Keyword) & "%' OR " & _
	'	"Descr  LIKE '%" & CStr(Keyword) & "%' OR " & _
	'	"ITEMID LIKE '%" & CStr(Keyword) & "%' OR " & _
	'	"Manufa LIKE '%" & CStr(Keyword) & "%') AND " & _
	'	"Cost <> 0 ORDER BY Cost"
	
	
	' I cannot match product fields to Keyword adequately using just SQL, because of the lack of Replace capability mentioned above.
	' To compensate, I am forced to SELECT ALL products and do Keyword match in the 
	' VB loop below rather than using SQL SELECT statement ...
	'ProductSQL = "SELECT PName, Descr, ITEMID, Manufa, Cost, Pic1, PID FROM Product WHERE Cost <> 0 ORDER BY Cost"
	ProductSQL = "SELECT * FROM Product WHERE Cost <> 0 ORDER BY Cost"
		
	
	'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp
	' (nor the rsProduct.getCount method, which is apparently only available when using the recordset DTC, which I
	' have not (yet?) figured out how to use on the MS Access database (not the SQL Server database that I am used to) 
	' that this website uses).
	' 5/6/10: No longer needed, since am now filtering in VB Loop below and not via SQL SELECT statement.
	'ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product " &_
	'	"WHERE ( PName  LIKE '%" & CStr(Keyword) & "%' OR " & _
	'	"Descr  LIKE '%" & CStr(Keyword) & "%' OR " & _
	'	"ITEMID LIKE '%" & CStr(Keyword) & "%' OR " & _
	'	"Manufa LIKE '%" & CStr(Keyword) & "%') AND " & _
	'	"Cost <> 0"

	
	Session("ProductSQL")		= ProductSQL
	Session("ProductCountSQL")	= ProductCountSQL
	Session("SummaryHeading")	= "<b>Keyword:</b>&nbsp;&nbsp; " & CStr(Keyword)
	
End Sub		' InitializeForKeywordSearch
%>






<%
If False Then
	'ComingFrom = Request.QueryString("CF")   
	'AID = Request.QueryString("AID")
	'SID = Request.QueryString("SID")
	Response.Write "<br>ComingFrom = "	& ComingFrom 
	Response.Write "<br>AID = "	& AID 
	Response.Write "<br>SID = "	& SID 
End If


'Response.Write "<br>ComingFrom = " & ComingFrom 

If ComingFrom <> "" Then   ' Then initialize ProductSQL and ProductCountSQL, according to search type.

' 12/3/20: The following comment is from the main site version of this file. It is not true for this mobile site version.
' If  ComingFrom = "" it is because we have just entered this file by a pagination change, and so ProductSQL and ProductCountSQL have already
' been initialized and can be re-used via Session("ProductSQL") and Session("ProductCountSQL"). There is no need to call InitializeForKeywordSearch 
' or InitializeForCatSubCatSearch again.	
	
	Select Case ComingFrom
	
	Case "KWS"
		ComingFrom 	= "KeywordSearch"
		Keyword 	= Trim(Request.QueryString("KW"))
		Keyword0 	= LCase(Replace(Keyword, "-", ""))		' Remove dashes in Keyword; to allow for the kind of matching against Keyword that Sani wants.
		Keyword0 	= Replace(Keyword0, " ", "")			' Remove spacse in Keyword; to allow for the kind of matching against Keyword that Sani wants.
		'Response.End
		InitializeForKeywordSearch Keyword0, Keyword
		
	Case "CSCS"
		ComingFrom = "CatSubCatSearch"
		'AID 		= "45"   'Request.QueryString("AID")	' i.e. ID of Product Area or Category.
		'SID 		= "173"  'Request.QueryString("SID")	' i.e. ID of Product SubArea or SubCategory.

		'Response.Write "<br>AID = "	& AID 	
		'Response.Write "<br>SID = "	& SID	
		If SID = "" Then SID = 0 End If						' Need an integer value for SID, for use in the javascript functions pPrev, pNext, pGo.
		'Response.Write "<br>SID = "	& SID
		InitializeForCatSubCatSearch AID, SID
	
	End Select
	
End If 		'  ComingFrom <> ""


ShowPageNum = Request.QueryString("ShowPageNum")
If ShowPageNum = "" Then ShowPageNum = 1 End If
%>



<% InArea = "Products" %>


<%
If ComingFrom <> "KeywordSearch" Then 

	ProductSQL = Session("ProductSQL")
    	'Response.Write "<br>* ProductSQL = " & ProductSQL
		
	'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount method to work
	' (nor the rsProduct.getCount method, which is apparently only available when using the recordset DTC, which I
	' have not (yet?) figured out how to use on the MS Access database (not the SQL Server database that I am used to) 
	' that this website uses).
       ProductCountSQL = Session("ProductCountSQL")
       'Response.Write "<br>* ProductCountSQL = " & ProductCountSQL
	
	Set Conn 			= Server.CreateObject("ADODB.Connection")   
	Conn.Open Session("ConnectionString")
	Set rsProduct 		= Conn.Execute(ProductSQL)
	Set rsProductCount 	= Conn.Execute(ProductCountSQL)
	Set Conn 			= Nothing
	TotalNumMembers 	= rsProductCount("Cnt")

%>
	<table align="center" border="0" cellPadding="1" cellSpacing="1" width="95%">
		<tr>
			<td align="middle">
			<br /><b><font size='5'><%=Session("SummaryHeading")%></font></b>
			<!-- 
			<br>There are <strong> <font color="#000080"><% =TotalNumMembers %> </font> </strong> products. 
			Click on the picture to view a product's details and price.<br />
			-->
			</td>
	</table>
<%

	' Build drop-down menu of the page numbers for the user to hyperlink to different pages of the summary data.
	' Based on summary page from www.futuresimchas.com.
	' This section for the menu at the top of the page should be identical to the one below for the menu at 
	' the bottom of the page, except for use of TopMenu instead of BottomMenu for the form name.
	
	' TotalNumMembers = rsvwMembers.getCount()
	' NumRows = 10000 is a quick way to turn pagination off, by making allowed NumRows / page bigger than what will ever be needed.
	NumRows = 10000        '( This is not necessarily the number of rows on the last page.)
	'NumCols = 1
	MembersPerPage = NumRows    ' * NumCols
	'Response.Write "MembersPerPage = " & MembersPerPage 
	StartRecord = (ShowPageNum - 1) * NumRows + 1
	 
	 
	If TotalNumMembers = 0 Then 		' Response.Redirect "TrySearchAgain.asp" 
		Response.Write "<br>"
		Response.Write "<center>"
		Response.Write "<font size='3' color='blue'>Sorry. Nothing was found.</font><br><br>"
		Response.Write "<font size='3' color='blue'>Please try a different search.</font>"
		Response.Write "</center>"
		Response.End
	End If    ' TotalNumMembers = 0
	
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
	    Response.Write "<form name='TopMenu'>"
	    Response.Write "<center>"
	
	    Response.Write "<A href='javascript:pPrev(TopMenu.Page," & MembersPerPage & ");'><img border=0 hspace='5' align='middle' src='../images/NavImages/ArrowLeft.gif'></a>"
	    Response.Write "<A href='javascript:pPrev(TopMenu.Page," & MembersPerPage & ");'>Previous Page</a>"
	    Response.Write "&nbsp;&nbsp;"  	
	    Response.Write "<select name='Page' onChange='pGo(TopMenu.Page," & MembersPerPage & ");'>"
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
	    Response.Write "<A href='javascript:pNext(TopMenu.Page," & MembersPerPage & ");'>Next Page</a>"
	    Response.Write "<A href='javascript:pNext(TopMenu.Page," & MembersPerPage & ");'><img border=0 hspace='5' align='middle' src='../images/NavImages/ArrowRight.gif'></a>"
	    Response.Write "</center>"
	    Response.Write "</form>"
	    Response.Write "<br>"
	Else
	    'Response.Write "<br><br>"
	End If 
%>


<% If ShowPageNum = 1 Then %>
<!-- Output preliminary description text for the Category or Subcategory. -->
<div style="margin:20px auto 20px auto;">
<!-- <% =Session("Description")%><br /> -->
</div>
<% End If %>


<%
	StartRecordNumber = (ShowPageNum - 1) * MembersPerPage + 1				' ShowPageNum is set above using ShowPageNum = Request.QueryString("ShowPageNum").
	'Response.Write "<br>StartRecordNumber = " & StartRecordNumber 
	
	'rsProduct.moveAbsolute(StartRecordNumber)
	' 8/17/05: Iterate over ALL records, but only display the subset needed for this page.
	' This is a kludge because I can't get rsProduct.moveAbsolute(StartRecordNumber) method to work.
	' It is apparently only available when using the recordset DTC, which I
	' have not (yet?) figured out how to use on the MS Access database (not the SQL Server database that I am used to) that this website uses.
	rsProduct.moveFirst
	endRecordNumber = StartRecordNumber + NumRows - 1
	parity = -1
	color = "white"
	Response.Write "<table align='center' cellpadding='5' cellspacing='0' border='0' width=" & PageWidth & ">"
	
	For row = 1 to TotalNumMembers ' NumRows 
		If row >= StartRecordNumber AND row <= endRecordNumber Then
		graphicFile = "../../Imi/" & rsProduct("Pic1")
		parity = - parity
		If parity = 1 Then color = "gainsboro" Else color = "white" End If

		PID					= rsProduct("PID")
		ItemID				= rsProduct("ItemID")
		ProductName			= rsProduct("PName")
		Manufacturer		= rsProduct("Manufa")
		Description			= rsProduct("Descr")
        Deleted				= rsProduct("Deleted")
		NewProductsSubgroup	= rsProduct("NewProductsSubgroup")
		RebatesSubgroup		= rsProduct("RebatesSubgroup")
		'Cost = rsProduct("Cost")

' *******************************
%>
<!-- #include virtual="mobile1/Search/SearchSummary.inc2.asp" -->
<%
' *******************************

		End If
		rsProduct.moveNext
	Next ' row

	Response.Write "</table>"


	NumPages = Celg(TotalNumMembers / NumRows)   ' Using Ceiling function I found on the web; see above.
	'Response.Write "NumPages = " & NumPages & "<br>"
	'Response.Write "ShowPageNum = " & ShowPageNum & "<br>"
	
	MembersOnLastPage = TotalNumMembers Mod NumRows
	If MembersOnLastPage = 0 Then
		MembersOnLastPage = MembersPerPage
	End If
	
	
	   
	If NumPages > 1 Then
	Response.Write "<br>"
	Response.Write "<form name='BottomMenu'>"
	Response.Write "<center>"
	
	Response.Write "<A href='javascript:pPrev(BottomMenu.Page," & MembersPerPage & ");'><img border=0 hspace='5' align='middle' src='../images/NavImages/ArrowLeft.gif'></a>"
	Response.Write "<A href='javascript:pPrev(BottomMenu.Page," & MembersPerPage & ");'>Previous Page</a>"
	Response.Write "&nbsp;&nbsp;"  	
	Response.Write "<select name='Page' onChange='pGo(BottomMenu.Page," & MembersPerPage & ");'>"
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
	Response.Write "<A href='javascript:pNext(BottomMenu.Page," & MembersPerPage & ");'>Next Page</a>"
	Response.Write "<A href='javascript:pNext(BottomMenu.Page," & MembersPerPage & ");'><img border=0 hspace='5' align='middle' src='../images/NavImages/ArrowRight.gif'></a>"
	Response.Write "</center>"
	Response.Write "</form>"
	Else
	Response.Write "<br><br>"
	End If 


' **************************************************************************************************************************

ElseIf ComingFrom = "KeywordSearch" Then

	' 5/6/10: Cannot pre-count the number of products in this case, because the filtering by Keyword is done in the loop below, not via the
	' SQL SELECT statement. As a fresult, also there are no pagination menus in this case.

	ProductSQL 		= Session("ProductSQL")
    'Response.Write "<br>** ProductSQL = " & ProductSQL

	Set Conn 		= Server.CreateObject("ADODB.Connection")
	Conn.Open Session("ConnectionString")
	Set rsProduct 	= Conn.Execute(ProductSQL)	
	Set Conn 		= Nothing
%>
	<table align="center" border="0" cellPadding="1" cellSpacing="1" width="95%">
		<tr><td height='10pt'></td></tr>
		<tr>
			<td align="middle">
			<font size=5><%=Session("SummaryHeading")%></font><br /><br />
			</td>
	</table>

<%
	' This "duplication" of the above loop is unfortunately needed, because I cannot match fields to Keyword adequately using just SQL, because of the 
	' unavailability (discussed above) of Replace in MS Access when accessed via ADO.
	StartRecordNumber = (ShowPageNum - 1) * MembersPerPage + 1				' ShowPageNum is set above using ShowPageNum = Request.QueryString("ShowPageNum").
	'Response.Write "<br>StartRecordNumber = " & StartRecordNumber 
	
	'rsProduct.moveAbsolute(StartRecordNumber)
	' 8/17/05: Iterate over ALL records, but only display the subset needed for this page.
	' This is a kludge because I can't get rsProduct.moveAbsolute(StartRecordNumber) method to work.
	' It is apparently only available when using the recordset DTC, which I
	' have not (yet?) figured out how to use on the MS Access database (not the SQL Server database that I am used to) that this website uses.
	rsProduct.moveFirst
	endRecordNumber = StartRecordNumber + NumRows - 1
	parity = -1
	Response.Write "<table align='center' cellpadding='5' cellspacing='0' border='0' width=" & PageWidth & ">"

	row = 0
	Do While NOT rsProduct.EOF 			
			graphicFile = "../../Imi/" & rsProduct("Pic1")
			PID 			= rsProduct("PID")

			ItemID 			= rsProduct("ItemID")
			ItemID0 		= LCase(Replace(ItemID, "-", ""))		' Remove dashes; to allow for the kind of matching against Keyword that Sani wants.
			ItemID0 		= Replace(ItemID0, " ", "")				' Remove spaces; to allow for the kind of matching against Keyword that Sani wants.
			'Response.Write	"<br>ItemID0 = (" & row & ") " & ItemID0 
			
			ProductName 	= rsProduct("PName")
			ProductName0 	= LCase(Replace(ProductName, "-", ""))
			ProductName0 	= Replace(ProductName0, " ", "")
			'Response.Write	"<br>ProductName0 = " & ProductName0 
			
			Manufacturer 	= rsProduct("Manufa")
			Manufacturer0 	= LCase(Replace(Manufacturer, "-", ""))
			Manufacturer0 	= Replace(Manufacturer0, " ", "")
			'Response.Write	"<br>Manufacturer0 = " & Manufacturer0 
			
			Description 	= rsProduct("Descr")
			Description0 	= LCase(Replace(Description, "-", ""))
			Description0 	= Replace(Description0, " ", "")
			'Response.Write	"<br>Description0 = " & Description0

            		Deleted         = rsProduct("Deleted")
			
			'Cost = rsProduct("Cost")

			NewProductsSubgroup	= rsProduct("NewProductsSubgroup")
			RebatesSubgroup		= rsProduct("RebatesSubgroup")
				
			' A 3rd argument of 1 in Instr below is supposed to make the comparison case-insensitive (I think), but it didn't work. So I am using LCase on fields and Keyword.
			If CBool(Instr(ItemID0, Keyword0)) OR CBool(Instr(ProductName0, Keyword0)) OR CBool(Instr(Manufacturer0, Keyword0)) OR CBool(Instr(Description0, Keyword0)) Then    
				row = row + 1
				parity = - parity
				If parity = 1 Then color = "gainsboro" Else color = "white" End If

' *******************************
%>
<!-- #include virtual="mobile1/Search/SearchSummary.inc2.asp" -->
<%
' *******************************

			End If 
		
		rsProduct.moveNext
	Loop 	' While NOT rsProduct.EOF

	Response.Write "</table>"


End If    


%>




<br>

