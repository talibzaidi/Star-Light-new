<%@ Language=VBScript %>


<SCRIPT id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</SCRIPT>


<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm  method="POST">




<%

Sub btnFindKeyword_onclick()
Keyword = Trim(Request.Form("txtKeyword"))  
Response.Write "Keyword = " & Keyword & "<br>"
'ProductSQL = "SELECT * FROM Product WHERE PName LIKE '%" & CStr(Keyword) & "%' AND Cost <> 0 ORDER BY Cost"
ProductSQL = "SELECT * FROM Product WHERE (	 PName  LIKE '%" & CStr(Keyword) & "%' OR " & _
											"Descr  LIKE '%" & CStr(Keyword) & "%' OR " & _
											"ITEMID LIKE '%" & CStr(Keyword) & "%' OR " & _
											"Text1  LIKE '%" & CStr(Keyword) & "%') AND " & _
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
ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE " & _
											  "( PName  LIKE '%" & CStr(Keyword) & "%' OR " & _
												"Descr  LIKE '%" & CStr(Keyword) & "%' OR " & _
												"ITEMID LIKE '%" & CStr(Keyword) & "%' OR " & _
												"Text1  LIKE '%" & CStr(Keyword) & "%') AND " & _
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



Sub btnFindProductName_onclick()
ProductName = Trim(Request.Form("txtProductName"))   
Response.Write "ProductName = " & ProductName & "<br>"
ProductSQL = "SELECT * FROM Product WHERE PName LIKE '%" & CStr(ProductName) & "%' AND Cost <> 0 ORDER BY Cost"
Response.Write "ProductSQL = " & ProductSQL & "<br>"
'Response.End
Session("ProductSQL")= ProductSQL

'8/17/05: The following is a kludge because I can't get rsProduct.RecordCount to work in searchsummary.asp (see above).
ProductCountSQL = "SELECT Count(Cost) AS Cnt0, Count(PID) As Cnt FROM Product WHERE PName LIKE '%" & CStr(ProductName) & "%' AND Cost <> 0"
Response.Write "ProductCountSQL = " & ProductCountSQL & "<br>"
Session("ProductCountSQL")= ProductCountSQL
Session("SummaryHeading")= "Product Name: " & CStr(ProductName)
Response.Write Session("SummaryHeading")
'Response.End

Response.Redirect "searchsummary.asp?ShowPageNum=1"     
End Sub		' btnFindKeyword_onclick()



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


<HEAD>
<link rel="stylesheet" type="text/css" href="../../../Misc/StyleSheet1.css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>


<body topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">

<!-- #INCLUDE VIRTUAL = "Misc/Header.short.inc" -->

<br><br><br><br><br><br>


<table align='center' border='0' cellspacing='0' cellpadding=5>		<% ' Start Outer Table %>

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
<OBJECT id=btnFindKeyword style="LEFT: 0px; WIDTH: 46px; TOP: 0px; HEIGHT: 27px" 
	height=27 width=46 classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731">
	<PARAM NAME="_ExtentX" VALUE="1217">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnFindKeyword">
	<PARAM NAME="Caption" VALUE="Find">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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

<!--METADATA TYPE="DesignerControl" endspan-->
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
<OBJECT id=btnFindProductName style="LEFT: 0px; WIDTH: 46px; TOP: 0px; HEIGHT: 27px" 
	height=27 width=46 classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731">
	<PARAM NAME="_ExtentX" VALUE="1217">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnFindProductName">
	<PARAM NAME="Caption" VALUE="Find">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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

<!--METADATA TYPE="DesignerControl" endspan-->
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
<OBJECT id=btnFindManufacturer style="LEFT: 0px; WIDTH: 46px; TOP: 0px; HEIGHT: 27px" 
	height=27 width=46 classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731">
	<PARAM NAME="_ExtentX" VALUE="1217">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnFindManufacturer">
	<PARAM NAME="Caption" VALUE="Find">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
</tr>



<tr>
<td height='50'>     
<b>Site Map - </b><br>Product Category or Subcategory:
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
<OBJECT id=btnFindCatAndSubCat style="LEFT: 0px; WIDTH: 46px; TOP: 0px; HEIGHT: 27px" 
	height=27 width=46 classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731">
	<PARAM NAME="_ExtentX" VALUE="1217">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnFindCatAndSubCat">
	<PARAM NAME="Caption" VALUE="Find">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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

<!--METADATA TYPE="DesignerControl" endspan-->

</td>
</tr>
				
							
</table>			<% ' End Outer Table %>


</BODY>


<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>


</HTML>
