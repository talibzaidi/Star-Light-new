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

	If TRUE Then 	' This block is always valid, whether user selected a category or a subcategory.
		Set Conn 	= Server.CreateObject("ADODB.Connection") 
		Conn.Open Session("ConnectionString")
		CatSQL 		= "SELECT AreaName, AreaDesc FROM Area51 WHERE AID=" & CategoryID
		'Response.Write "<br>CatSQL = " & CatSQL 
		Set rsCat 	    = Conn.Execute(CatSQL)
		CatName 	    = rsCat("AreaName")					' i.e. Name of Product Area or Category.
        CatDesc         = rsCat("AreaDesc")                 ' i.e. Description of Product Area or Category.
		'Response.Write "<br>CatName = "	& CatName 
        'Response.Write "<br>CatDesc = "	& CatDesc 
	End If
	
	If SID <> 0 Then	' This occurs when user selected a subcategory, not a category.
		SubCatSQL 	= "SELECT Subname, SubDesc FROM Subarea WHERE AID=" & CategoryID & " AND SID=" & SubCategoryID
		'Response.Write "<br>SubCatSQL = " & SubCatSQL 
		Set rsSubCat    = Conn.Execute(SubCatSQL)
		SubCatName      = rsSubCat("Subname")			    ' i.e. Name of Product SubArea or SubCategory.
        SubCatDesc 	    = rsSubCat("SubDesc")			    ' i.e. Description Product SubArea or SubCategory.
		'Response.Write "<br>SubCatName = "	& SubCatName 
        'Response.Write "<br>SubCatDesc = "	& SubCatDesc
	End If
	
	If SubCategoryID <> 0 Then		' User selected a subcategory, not a category.
		ProductSQL = "SELECT * FROM Product WHERE SID = " & CStr(SubCategoryID) & " AND Cost <> 0 ORDER BY Cost"
		'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
		ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE SID = " & CStr(SubCategoryID) & " AND Cost <> 0"
		SummaryHeading = "<table><tr><td align=right><b>Category:</b></td><td>&nbsp;&nbsp;</td><td align=left>" & CatName & "</td></tr>" 
		SummaryHeading = SummaryHeading & "<tr><td align=right><b>Subcategory:</b></td><td>&nbsp;&nbsp;</td><td align=left>" & SubCatName & "</td></tr></table>"
		Session("SummaryHeading")   = SummaryHeading 
        Session("Description")      = SubCatDesc
	Else								' User selected a category, not a subcategory.
		ProductSQL = "SELECT * FROM Product INNER JOIN SubArea ON Product.SID = SubArea.SID WHERE SubArea.AID = " & CStr(CategoryID) & " AND Cost <> 0 ORDER BY Cost"
        '8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
		ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product INNER JOIN SubArea ON Product.SID = SubArea.SID WHERE SubArea.AID = " & CStr(CategoryID) & " AND Cost <> 0"
		Session("SummaryHeading")   = "<b>Category:</b>&nbsp;&nbsp;" & CatName
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



<HTML>


<HEAD>
	<!-- <meta http-equiv="Content-Type" content="text/html; charset=utf-8"> -->
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
    <meta http-equiv="content-type" content="text/html; charset=UTF-32">
    <meta http-equiv="content-language" content="en">
    <title>GPS sensors | GPS OEM engine boards | Oem GPS | tracking GPS | night vision optics | two-way communication | starlite-intl.com</title>
    <meta NAME="Description" CONTENT="Complete line of GPS sensors, GPS oem boards, OEM GPS, GPS engine, GPS engine Boards, GPS tracking. Also GPS Smartphones, truck GPS, 2-way radios, CB radios, radio scanners, antennas and accessories, night vision optics./">
    <meta NAME="Keywords" CONTENT="GPS sensors,OEM GPS boards,GPS boards,GPS engine,GPS oem engine Boards,GPS tracking,GPSMap,GPS Smartphones,2-way radios,CB radios,radio scanners,antennas and accessories,night vision optics,Garmin,USGlobal,Pharos,Uniden,Cobra,Midland /">
</HEAD>



<body topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">

<%
ComingFrom = Request.QueryString("CF")
'Response.Write "<br>ComingFrom = "	& ComingFrom 




ShowPageNum = Request.QueryString("ShowPageNum")
If ShowPageNum = "" Then ShowPageNum = 1 End If
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
	<XXXiframe name="inlineframe" src="https://www.starlite-intl.com/OEM_GPS_sensors/searchsummary.asp?CF=CSCS&AID=45&SID=173&ShowPageNum=1" frameborder="0" scrolling="auto" width="95%" height="600" marginwidth="5" marginheight="5" ></iframe> 

	<!-- #INCLUDE file="searchsummaryOEM_GPS_sensors.inc.asp" -->

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
