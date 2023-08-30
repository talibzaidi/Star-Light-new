
<% ' 4/20/10: Drop-down menu for Products. Dynamically generated from database. %>


<% If FALSE Then 	' This is here just as a model for how to structure 3-level suckerfish drop-downs. %>
<ul class="sf-menu" id="suckerfishnav">	 
	<li><a>Products</a>
	<ul>
	
		<li><a href="https://www.starlite-intl.com/scart/scartstart.asp?pid=0&sid=11&area=Specials&sar=Specials">1. Cat 1</a>
		<ul>
			<li><a>* 1.1</a></li>
			<li><a>* 1.1</a></li>
		</ul>
		</li>
	
		<li><a href="https://www.starlite-intl.com/scart/scartstart.asp?pid=0&sid=11&area=Specials&sar=Specials">2. Cat 2</a>
		<ul>
			<li><a>* 2.1</a></li>
			<li><a>* 2.1</a></li>
		</ul>
		</li>
		
	</ul>
	</li>
</ul>
<% End If 	' FALSE %>



<%
' 4/20/10: This is based on the loop in search/search.asp.

'MenuSQL = "SELECT * FROM Subarea INNER JOIN Area51 ON Subarea.AID = Area51.AID ORDER BY Subarea.AID ASC, Subarea.SID ASC"
MenuSQL = "SELECT * FROM Subarea INNER JOIN Area51 ON Subarea.AID = Area51.AID ORDER BY AreaName ASC, Subname ASC"		
Set Conn = Server.CreateObject("ADODB.Connection")

'Response.Write "<br>" & Session("ConnectionString")	' 6/13/12: Old method, using DSN-type string. Being phased out.
'Conn.Open Session("ConnectionString")

'Response.Write Server.MapPath("searchsummary.asp")		' Just to help figure out physical path to use in global.asa for Session("ConnectionString").
'Response.Write "<br>" & Session("ConnectionString")	' 6/13/12: New method, using regular connection string. Being phased in.
Conn.Open Session("ConnectionString")

Set rsSubArea = Conn.Execute(MenuSQL)
Set Conn = Nothing
AIDprevious =  -1
'Conn.Close
%>


<!-- <ul class="sf-menu" id="suckerfishnav"> -->
	<!-- <li <%=ProductsStyle%>><a href="https://www.starlite-intl.com/search/search.asp">Products</a> -->
	<li <%=ProductsStyle%>><a>Products</a>

<%	
Do While Not rsSubArea.EOF
	AID 				= rsSubArea("Subarea.AID")					' i.e. ID of Product Area / Category.
	SID 				= rsSubArea("SID")							' i.e. ID of Product SubArea / SubCategory.
	CatName 			= rsSubArea("AreaName")						' i.e. Name of Product Area / Category.
	SubCatName 			= rsSubArea("Subname")   					' i.e. Name of Product SubArea / SubCategory.
	'Response.Write "<br>SubCatName = " & SubCatName 
	
	If SID <> "" AND AID <> 0 AND SubCatName <> "" AND SubCatName <> "test" Then 		
		
		URLSubCat = "https://www.starlite-intl.com/search/searchsummary.asp?CF=CSCS&AID=" & AID & "&SID=" & SID & "&ShowPageNum=1"  	' User selected a Subcategorgy.
		URLCat = "https://www.starlite-intl.com/search/searchsummary.asp?CF=CSCS&AID=" & AID & "&ShowPageNum=1"							' User selected a Category.

		If (AID <> AIDprevious) AND (AIDprevious= -1) Then     ' i.e. just started the 1st new catergory / area.
			' I am assuming that 1st category is NOT Warranties,  Warranties, Gift Certificates or Tools.
			Response.Write "<ul>"
			Response.Write "<li><a href=" & URLCat & ">" & CatName & "</a><ul>"
			Response.Write "<li><a href=" & URLSubCat & ">" & SubCatName & "</a></li>"
		ElseIf (AID <> AIDprevious) Then     ' i.e. just started the 2nd and later new catergory / area.
			' Don't want to list subcategories for Warranties (there is only one), Gift Certificates (there is only one) or Tools (there are too many).
			If (AIDprevious <> 97) AND (AIDprevious <> 66) AND (AIDprevious <> 50) Then   
				Response.Write "</ul></li>" 
			Else 
				' Do nothing
			End If
			
			If (AID <> 97) AND (AID <> 66) AND (AID <> 50) Then  ' Don't want to list subcategories for Warranties, Gift Certificates or Tools.
				Response.Write "<li><a href=" & URLCat & ">" & CatName & "</a><ul>"
				Response.Write "<li><a href=" & URLSubCat & ">" & SubCatName & "</a></li>"
			Else
				Response.Write "<li><a href=" & URLCat & ">" & CatName & "</a></li>"			
			End If
		Else	' i.e. still in same catergory / area as before.
			If (AID <> 97) AND (AID <> 66) AND (AID <> 50) Then
				Response.Write "<li><a href=" & URLSubCat & ">" & SubCatName & "</a></li>"
			End If
		End If
		
		'Response.Redirect URL
		
		AIDprevious = AID
	End If 
	
	rsSubArea.MoveNext
Loop
Response.Write "</ul></li>"
rsSubArea.Close 
%><!-- </ul>  -->