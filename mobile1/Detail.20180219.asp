<%@ LANGUAGE = VBScript %>

<% 'Response.Write "<br>Session('Country') = "	& Session("Country") %>

<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<%response.buffer=true%>

<% PID = ReQuest("PID") %>
<% ar = Request("Area") %>
<% Area = Request("Area") %>

<% 'Response.Write "<br>Session('Country') = " & Session("Country") & "<br>" %>

<% if Area="iii" then
   Area = Request("Manufat")
   ar = Request("Manufat")
   end if

   if Area = "New Products" then
     response.redirect "./scart.asp?pid=0&sid=11&area=New!&sar=New%20Products"
   end if	

   if Area="Choose a catalog area." then
   response.redirect "https://www.starlite-intl.com/mobile1/index.asp"
   end if
      
    If True Then  ' [BN, 2/7/18] ?The following was the original, old code, here in this file. It seems to be obsolete now?
        If (Request("Canada") <> "" OR Request("  USA  ") <> "") then
            If Request("Canada") <> "" then
                Session("Country") = "Canada"
            else
                Session("Country") = "USA"
            end if
        end if
    End If  ' False/True
%>


<% sar = ReQuest("sar") %>
<% 'sar = Replace( sar, " ", "%20") %>
<% SID = ReQuest("SID") %>


<%
' [BN, 12/6/17] Copied from Details.asp of main (non-mobile) site.

' 11/9/15: Using the connection string method, as opposed to the earlier DSN method.
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open Session("ConnectionString")
SQLstring = "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  PID = " + PID  +   " "
Set RS = Server.CreateObject("ADODB.Recordset")
'rsSpecials.Open SQLstring, Conn1, adOpenStatic, adLockOptimistic, adCmdText 
RS.Open SQLstring, Conn1, 3, 3, &H0001 
Set Conn1 = Nothing

'on error resume next
'do while not rs.eof

' ***********************************************************

' The following accesses the meta data, if any, listed in the database for this product.
' This is so we can dynamically specialize the <head> section of the generated html page to be specific to the product being dynamically 
' displayed in the <body> section of the generated page. That will hopefully cause Google to rank the generated (dynamic) html page
' higher in the organic search results. (See comments in my file MyTests > DynamicMetatags > test1.html of the non-mobile site.) 
If NOT (RS("tagTitle") = "") Then
    tagTitle = "<title>" & RS("tagTitle") & "</title>"
Else
    tagTitle = "<title>GPS sensors, OEM GPS, GPS modules, tracking GPS, lidar-lite, night vision optics, CB radios, Star Lite International</title>"
End If

If NOT (RS("tagDescription") = "") Then
    tagDescription = "<meta name='description' content='" & RS("tagDescription") & "'>"
Else
    tagDescription = "<meta name='description' content='Large selection of Garmin GPS,Lidar-Lite,USglobal GPS,OEM GPS,GPS sensors,bluetooth GPS,fish finders,sounders,CB radios and walky-talky,radio scanners,digital cameras,car audio and video'>"
End If

tagTitle1       = "<title>" & RS("tagTitle") & "</title>"
tagDescription1 = "<meta name='description' content='" & RS("tagDescription") & "'>"

tagTitle2       = "<title>" & RS("PName") & " | " & RS("ItemID") & "</title>"
tagDescription2 = "<meta name='description' content='" & RS("Descr") & "'>"

' ***********************************************************
%>



<%
' The following column widths are defined here so they can be used in Subs below to easily keep their columns aligned.
' XXXCol1Width=50 : XXXCol2Width=5 : XXXCol3Width=120 : XXXCol4Width=310 : XXXCol5Width=120 : XXXCol6Width=120
' XXXColnWidth=10  ' i.e. last col.
	
' 2/23/06: Makes the table of a product's accessories or the table of all products that the product is an accessory of.
Sub MakeTable(Direction)
	Select Case Direction
	Case "Accessories"
			' Added on 2/22/06, by Nadel ...
			If NOT IsNULL(rs("HasAccessories")) AND Trim(rs("HasAccessories")) <> "" Then
				MakeTable2("Accessories")
			End If 
	Case "AccessoryOf"
			' Added on 2/23/06, by Nadel ...
			'If NOT IsNULL(rs("IsAccessoryOf")) AND Trim(rs("IsAccessoryOf")) <> "" Then
			'	MakeTable2("AccessoryOf")
			'End If 
	End Select
End Sub ' MakeTable


' Added on 2/23/06, by Nadel ...
Sub MakeTable2(Direction)

	Response.Write "<br><table style='border:0px solid #CCCCCC' align='center' width='100%'>"
	If Direction = "Accessories" Then
		Response.Write "<tr align=left style='background:#1800FF;' ><td colspan=4 ><font color='white' size=4 style='line-height:1.5'><b>&nbsp;Has Accessories ...</b></font></td></tr>"
	Else
		Response.Write "<tr align=left style='background:#1800FF;' ><td colspan=4><font color='white' size=4 style='line-height:1.5'><b>&nbsp;Is an Accessory of ...</b></font></td></tr>"			
	End If
	Response.Write "<tr style='border:0px solid black;'>"
	Response.Write "<td >&nbsp;</td>"
	Response.Write "<td align=right><font><b>List Price</b></font></td>"
	Response.Write "<td align=right><font color='red'><b>Our Price</b></font></td>"
	Response.Write "<td >&nbsp;</td>"
	Response.Write "</tr>"
	'Dim i
	'Dim AccessoryID
	If Direction = "Accessories" Then
        'Response.Write "<br>" & rs("HasAccessories")
		AccessoryID = Split(rs("HasAccessories"), ",")
	Else   ' 2/25/06: No longer used.
		AccessoryID = Split(rs("IsAccessoryOf"), ",")	' So in this case "AccessoryID" really stands for "AccessoryOfID"
	End If 



	parity = -1
	color = "white"
	For i = LBound(AccessoryID) to UBound(AccessoryID)

		parity = - parity
		'If parity = 1 Then color = "gainsboro" Else color = "white" End If
		If parity = 1 Then color = "white" Else color = "white" End If   ' Elliminates alternating colors.

		ItemID = Trim(AccessoryID(i))
		'Response.Write "<br>ItemID = " & ItemID
		If InStr(ItemID,"~~") Then  ' If it starts with "~~" then it is a heading, not an ID. 
		' Using InStr(ItemID,"~") above, instead of Left(ItemID,1), to allow for carriage returns to occur in front of "~".
			ItemID = replace(ItemID, "~~", "")  ' Remove (all) "~~".
			Response.Write "<tr><td colspan=4 align=left style='background:#EEEEEE;'><font color=navy><b>" & ItemID & "</b></font></td></tr>"
		Else
			AccessorySQL = "Select PName, PID, ITEMID, PIC1, MSL, Duty, Cost, GPM, ShowPrice, Deleted from PRODUCT WHERE ITEMID = '" & ItemID & "'"
			'Response.Write "<br>AccessorySQL = " & AccessorySQL
			Set conn = Server.CreateObject("ADODB.Connection")
    		Conn.Open Session("ConnectionString")
    		'If rsAccessory.isOpen() Then rsAccessory.close() End If 
    		Set rsAccessory = Conn.Execute(AccessorySQL)
    			
    		If NOT rsAccessory.EOF Then  ' Sometimes an Accessory listed in HasAccessories column of Product table may not really be in the Product table.
				Response.Write "<tr style='border:0px solid black' bgcolor='" & color & "'>"
					
				If Trim(rsAccessory("PIC1")) <> "notava1t.gif" Then
  					Response.Write "<td XXXwidth=" & Col1Width & ">"
   					'Response.Write "<a href=https://www.starlite-intl.com/mobile1/Detail.asp?PID=" & Trim(rsAccessory("PID")) & ">"
					Response.Write "<img border=0 width='30' src='../imi/" & Trim(rsAccessory("PIC1")) & "'"
    				'Response.Write "</a>"
					Response.Write "</td>"
				Else
					Response.Write "<td width=" & Col1Width & ">"
					'Response.Write "No Picture"
					Response.Write "</td>"
				End If
				
  				Response.Write "<td colspan='3' align=left width=" & Col4Width & ">"
   				'Response.Write "<a href=https://www.starlite-intl.com/mobile1/Detail.asp?PID=" & Trim(rsAccessory("PID")) & ">"
  				Response.Write Trim(rsAccessory("PName"))
    			'Response.Write "</a>"
    			Response.Write "</td>"

				Response.Write "</tr>"
				Response.Write "<tr style='border:0px solid black' bgcolor='" & color & "'>"

				Response.Write "<td align=left width=" & Col3Width & "><font size=1>" & ItemID & "</font></td>"
    				
    			Response.Write "<td align=right width=" & Col5Width & ">"	' List Price ...
                If Not Trim(rsAccessory("Deleted")) Then
  				    If Session("Country") = "USA" Then    
					    ' 6/18/06, commented out, BN: Response.Write formatcurrency(rsAccessory("MSL")*RS("Freight"))
					    Response.Write formatcurrency(Trim(rsAccessory("MSL")))
				    Else
					    ' 6/18/06, commented out, BN: Response.Write formatcurrency(((rsAccessory("MSL")*rsAccessory("Duty"))*RS("Freight"))*RS("Exch")) 
					    Response.Write formatcurrency(Trim(rsAccessory("MSL"))*Trim(rsAccessory("Duty"))*Trim(RS("Exch"))) 
				    End If
                End If
    			Response.Write "</td>"
    				
    			Response.Write "<td align=right width=" & Col6Width & ">"   ' Our Price ...
    			'Response.Write "HI" ' "rsAccessory('ShowPrice') = " & rsAccessory("ShowPrice")
                If Not Trim(rsAccessory("Deleted")) Then
				    If Session("Country") = "USA" Then    
 					    If Trim(rsAccessory("ShowPrice")) = True Then			'  "ShowPrice" really means "Don't Show Price"  
						    Response.Write "<font size=1>Click ORDER<br>to see price<br>or to order</font>"
					    Else
						    Response.Write "<font color=red><i>" & formatcurrency(Trim(rsAccessory("Cost"))*Trim(RS("Freight"))*(1/(1-(Trim(rsAccessory("GPM")))))) & "</i></font>"
					    End If
				    Else
					    If CBool(Trim(rsAccessory("ShowPrice")))  Then			'  "ShowPrice" really means "Don't Show Price" 
						    Response.Write "<font size=1>Click ORDER<br>to see price<br>or to order</font>"
					    Else 
						    Response.Write "<font color=red><i>" & formatcurrency(Trim(rsAccessory("Cost"))*Trim(rsAccessory("Duty"))*Trim(RS("Freight"))*Trim(RS("Exch"))*(1/(1-(Trim(rsAccessory("GPM")))))) & "</i></font>"
					    End If
				    End If
                End If
    			Response.Write "</td>"
  						
                If Not Trim(rsAccessory("Deleted")) Then
                  	Response.Write "<td align=center>" 
                    'Response.Write Trim(rsAccessory("ITEMID"))
 				    Response.Write "<a href=https://www.starlite-intl.com/mobile1/scart/scart.asp?action=add&item=" & Trim(rsAccessory("ITEMID")) & "&count=1&sid=0&Area=&sar=Special" & ">"
                    Response.Write "<img src='../Images/order.gif'  border=0>"	
    			    Response.Write "</a>"
                    Response.Write "</td>"
                Else
                    Response.Write "<td align=center style='border: 0px solid blue;' XXXwidth='100'>" 
				    Response.Write "<font color='#B90000' size='2'>No Longer Available.</font><br>"
				    'Response.Write "<a href=https://www.starlite-intl.com/mobile1/Detail.asp?pid=" & Trim(rsAccessory("PID")) & "><font color='navy' size='1'>Click for possible alternatives and accessories</font></a>"
                    Response.Write "</td>"
                End If
				
  				'Response.Write "<td style='border:1px solid black' width=" & ColnWidth & ">&nbsp;" 
  				'Response.Write "</td>" 
  
   				Response.Write "</tr>"
				Response.Write "<tr><td colspan='4'><hr style='color:#EEEEEE'></td></tr>"
   			End If 
   		End If
	Next
	Response.Write "</table>"
	Response.Write "<br><br>"
End Sub  ' MakeTable2


' Added 2/26/06, to compute parent lists (IsAnAccessoryOf lists) instead of just looking them up. This avoids need for
' Sani to have to double-enter both children and parents.
Sub MakeTable3(ItemID)
	'Response.Write "<br>**ItemID = " & ItemID
	' The following ParentsSQL is not a reliable approach, because it will also find products with children whose ItemIDs 
	' subsume ItemID (rather than just being *equal* to ItemID).
	'ParentsSQL = "Select PName, PID, ITEMID, PIC1, MSL, Duty, Cost, GPM, ShowPrice from PRODUCT " & _
	'			"WHERE HasAccessories LIKE '%" & ItemID & "%'"     
	ParentsSQL = "Select PName, PID, ITEMID, PIC1, MSL, Duty, Cost, GPM, ShowPrice, Deleted, HasAccessories from PRODUCT " & _
				"WHERE HasAccessories LIKE '%_%'"   ' i.e. returns products having at least one child.
	'Response.Write "<br>ParentsSQL = " & ParentsSQL
	'Response.End
	Set connParents = Server.CreateObject("ADODB.Connection")
    connParents.Open Session("ConnectionString")
    Set rsParents = connParents.Execute(ParentsSQL)
    NumberOfParentsOfItemID = 0
	Do while NOT rsParents.EOF  ' Remember, rsParents is the set of parents of some product, not necessarily of ItemID.
		Children = rsParents("HasAccessories")
		If InChildren(Children, ItemID) Then
			NumberOfParentsOfItemID = NumberOfParentsOfItemID + 1
			If NumberOfParentsOfItemID = 1 Then
				Response.Write "<table style='border:0px solid #1800FF' align='center' width='100%'>"			
				Response.Write "<tr align=left style='background:#1800FF;' ><td colspan=8><font color='white' size=4 style='line-height:1.5'><b>&nbsp;Is an Accessory of ...</b></font></td></tr>"			
				Response.Write "<tr><td>&nbsp;</td>"
				Response.Write "<td align=right><font><b>List Price</b></font></td>"
				Response.Write "<td align=right><font color='red'><b>Our Price</b></font></td>"
				Response.Write "<td>&nbsp;</td></tr>"
			End If
			
			Response.Write "<tr>" 
		
			If rsParents("PIC1") <> "notava1t.gif" Then
  				Response.Write "<td width=" & Col1Width & ">"
   				Response.Write "<a href=https://www.starlite-intl.com/mobile1/Detail.asp?PID=" & rsParents("PID") & ">"
				Response.Write "<img border=0 width='30' src='../Imi/" & rsParents("PIC1") & "'"
				Response.Write "</a>"
				Response.Write "</td>"
			Else
				'Response.Write "<td>"
  				Response.Write "<td width=" & Col1Width & ">"
				'Response.Write "**" & PIC1 & "**"
				Response.Write "1</td>"
			End If

  			Response.Write "<td align=left colspan='3' width=" & Col4Width & ">"
   			Response.Write "<a href=https://www.starlite-intl.com/mobile1/Detail.asp?PID=" & rsParents("PID") & ">"
  			Response.Write rsParents("PName")
			Response.Write "</a>"
			Response.Write "</td>"		
  
   			Response.Write "</tr>"
			
			' ***********************************************************

			Response.Write "<tr>" 
						
			Response.Write "<td align=left width=" & Col3Width & "><font size=1>" & rsParents("ItemID") & "</font></td>"	
						
    		Response.Write "<td align=right width=" & Col5Width & ">"	' List Price ...
            If Not rsParents("Deleted") Then
  			    If Session("Country") = "USA" Then    
				    ' 6/18/06, commented out, BN: Response.Write formatcurrency(rsParents("MSL")*RS("Freight"))
											      Response.Write formatcurrency(rsParents("MSL"))
			    Else
				    ' 6/18/06, commented out, BN: Response.Write formatcurrency(((rsParents("MSL")*rsParents("Duty"))*RS("Freight"))*RS("Exch")) 
											      Response.Write formatcurrency(rsParents("MSL")*rsParents("Duty")*RS("Exch"))
			    End If
            End If
			Response.Write "</td>"
						
    		Response.Write "<td align=right width=" & Col6Width & ">"   ' Our Price ...
			'Response.Write "HI" ' "rsParents('ShowPrice') = " & rsParents("ShowPrice")
            If Not rsParents("Deleted") Then
			    If Session("Country") = "USA" Then    
 				    If rsParents("ShowPrice") = true Then			'  "ShowPrice" really means "Don't Show Price"  
					    Response.Write "<font size=1>Click ORDER to see<br>price or to order</font>"
				    Else
					    Response.Write "<font color=red><i>" & formatcurrency(rsParents("Cost")*RS("Freight")*(1/(1-(rsParents("GPM"))))) & "</i></font>"
				    End If
			    Else
				    If CBool(rsParents("ShowPrice")) Then			'  "ShowPrice" really means "Don't Show Price" 
					    Response.Write "<font size=1>Click ORDER<br>to see price<br>or to order</font>"
				    Else 
					    Response.Write "<font color=red><i>" & formatcurrency(rsParents("Cost")*rsParents("Duty")*RS("Freight")*RS("Exch")*(1/(1-(rsParents("GPM"))))) & "</i></font>"
				    End If
			    End If
            End If
			Response.Write "</td>"
  						
            If Not rsParents("Deleted") Then
                Response.Write "<td align=center>" 
                'Response.Write rsParents("ITEMID")
 				Response.Write "<a href=https://www.starlite-intl.com/mobile1/scart/scart.asp?action=add&item=" & Trim(rsParents("ITEMID")) & "&count=1&sid=0&Area=&sar=Special" & ">"
  				Response.Write "<img src='../Images/order.gif'  border=0>"	
    			Response.Write "</a>"
                Response.Write "</td>"
            Else
                Response.Write "<td align=center XXXwidth='100'>" 
				Response.Write "<font color='#B90000' size='2'>No Longer Available.</font><br>"
                Response.Write "<a href=https://www.starlite-intl.com/mobile1/Detail.asp?pid=" & rsParents("PID") & "><font color='navy' size='1'>Click for possible alternatives and accessories</font></a>"
				Response.Write "</td>"
			End If
  
   			Response.Write "</tr>"

			' ***********************************************************
			Response.Write "<tr><td colspan='4'><hr style='color:#EEEEEE'></td></tr>"
   		End If 
	rsParents.Movenext
    loop
	Response.Write "</table>"
	Response.Write "<br>"
	rsParents.close  
	connParents.close
End Sub ' MakeTable3
   

' Tests if product ItemID is in children list Children
Function InChildren(Children, ItemID)
	'Response.Write "<br>ItemID = " & ItemID
	InFlag = False
	AccessoryID = Split(Children, ",")	
	i = 0
	'For i = LBound(AccessoryID) to UBound(AccessoryID)
	Do While (i <= UBound(AccessoryID)) AND (InFlag = False)
		AID = Trim(AccessoryID(i))
		If AID = ItemID Then
			'Response.Write AID & " " 
			InFlag = True
		End If
		i = i + 1
	Loop
	InChildren = InFlag
	'Response.Write "<br>InChildren = " & InChildren
End Function 


'Function GetWarrantyPID(ItemID)
'	Set rsWarrProducts = CreateObject("ADODB.Recordset")
'	rsWarrProducts.Open "SELECT * FROM Product WHERE  ItemID = '" & ItemID & "' ", "DSN=STAREC1" , 1, 4
'	GetWarrantyPID = rsWarrProducts("PID")
'End Function


Sub OutputWarrantyRow(ItemID, Description)
	'Response.Write "<br>ItemID = " & ItemID & ", " & "Description = " & Description
	Set rsWarrProducts = CreateObject("ADODB.Recordset")
	rsWarrProducts.Open "SELECT *, Rates.ExchangeRate1 AS Exch, Rates.Freight FROM Product, Rates WHERE  ItemID = '" & ItemID & "' ", "DSN=STAREC1" , 1, 4
	
    ' 11/18/15: I replaced DSN method of line above with the folllowing Connection String method. 	
    WarrProducts_SQL = "SELECT *, Rates.ExchangeRate1 AS Exch, Rates.Freight FROM Product, Rates WHERE  ItemID = '" & ItemID & "' "
    Set Conn = Server.CreateObject("ADODB.Connection") 
    Conn.Open Session("ConnectionString")
    Set rsWarrProducts = Conn.Execute(WarrProducts_SQL)	
    Set Conn = Nothing
    
    WarrantyPID = Trim(rsWarrProducts("PID"))
	WarrantyPic = Trim(rsWarrProducts("Pic1"))
	WarrantyPrice = Trim(rsWarrProducts("Cost"))*Trim(rsWarrProducts("Freight"))*(1/(1-(Trim(rsWarrProducts("GPM"))))) 
	If Session("Country") = "Canada" Then
		WarrantyPrice = WarrantyPrice * Trim(rsWarrProducts("Exch"))
	End If
	
	Response.Write "<tr>"
		Response.Write "<td width=" & Col1Width & " align=center>"
   		Response.Write "<a href=https://www.starlite-intl.com/mobile1/Detail.asp?PID=" & WarrantyPID & ">"
		Response.Write "<img border=0 width='30' src='../imi/" & WarrantyPic & "'"
    	Response.Write "</a>"
		Response.Write "</td>"
		
		Response.Write "<td width=" & Col2Width & ">"
		Response.Write "</td>"
		
		Response.Write "<td width=" & Col3Width & " align=left>"
		Response.Write "<font size=1>" & ItemID & "</font>"
		Response.Write "</td>"
		
		Response.Write "<td width=" & Col4Width & " align=left>"
		Response.Write Description
		Response.Write "</td>"
		
		Response.Write "<td align=right width=" & Col5Width & ">"
		Response.Write "</td>"
		
		Response.Write "<td align=right width=" & Col6Width & ">"
		Response.Write formatcurrency(WarrantyPrice)
		Response.Write "</td>"
		
  		Response.Write "<td align=right>"  	
 		Response.Write "<a href=https://www.starlite-intl.com/mobile1/scart/scart.asp?action=add&item=" & ItemID & "&count=1&sid=0&Area=&sar=Special" & ">"
  		Response.Write "<img src='../Images/order.gif'  border=0>"	
    	Response.Write "</a>"
    	Response.Write "</td>"
		
  		Response.Write "<td width=" & ColnWidth & ">" 
  		Response.Write "</td>" 
		
	Response.Write "</tr>"
End Sub		' OutputWarrantyRow
%>



<html>

<head>
    <% 'Response.Write vbCrLf & vbTab & tagTitle1 %>
    <% Response.Write vbCrLf & vbTab & tagTitle2 %>
    <% 'Response.Write vbCrLf & vbTab & tagDescription1 %>
    <% Response.Write vbCrLf & vbTab & tagDescription2 & vbCrLf %>
	<% ' [BN, 12/6/17] The description metatag below is way too general and is therefore not relevant to any individual product that would be displayed by this file Display.asp. 
       ' <meta name="description" content="Large selection of Garmin GPS,Lidar-Lite,USglobal GPS,OEM GPS,GPS sensors,bluetooth GPS,fish finders,sounders,CB radios and walky-talky,radio scanners,digital cameras,car audio and video"> 
    %>
	<% ' [BN, 12/6/17] The keywords metatag is no longer used by Google (and other search engines?), and even if it was, 
       ' the list of keywords below is way too general and is therefore not relevant to any individual product that would be displayed by this file Display.asp. 
       ' <meta name="keywords" content="Gps,Gps sensors,Gps sensor,Gps engine,Gps navigation,oem Gps,GpsMap,Nuvi,Gps accessories,Gps system,Lidar-Lite,range finder,tracker Gps,auto Gps,portable Gps,handheld Gps,marine Gps,Gps marine network,Gps receiver,Gps antennas,fish finder,sounders,Gps cartography,Gps equipment,Garmin Gps,bluetooth Gps,global positioning,tracking Gps,fleet tracking Gps,USglobal Gps,Gps on sale,navigation electronics,Cobra,Midland,amateur radios,Galaxy radio,Magnum radio,radio scanner,radio scanners,scanner,digital cameras,power supplies,regulated power supplies,Fujifilm,Nikon,Olympus,Panasonic,Motorola,Canon"> 
    %>
	<meta http-equiv="content-language" content="en">
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
    <% ' The following stylesheet was causing trouble on this (mobile site) page by indenting the left edge of the <body>.
        ' I have not investigated why, but at least for now I am just commenting out that stylesheet.
    %>
    <% ' The following stylesheet was causing trouble on this page (of the mobile site) by indenting the left edge of the <body>.
       ' I have not investigated why, but at least for now I am just commenting out the stylesheet.
       ' 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus (on main, non-mobile, site) to work. 
       ' 12/6/17: But in any case, it is apparently not needed here (on the mobile site).
    %>
	<XXXlink rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> 
	<% ' <meta name="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds"> %>
	
	<meta name="viewport" content="width=device-width; initial-scale=1.0"> 
	<!-- foneFrame.css is the stylesheet with comments, so it is readable.
	     foneFrame-min.css is the minimized version; it is smaller and loads faster. -->
	<link href="foneFrame.css" rel="stylesheet" type="text/css">
	<!-- The following 2 lines are not strict HTML5. -->
	<meta name="HandheldFriendly" content="true"/>
	<meta name="MobileOptimized" content="320"/>
	
	<!-- You can use different style sheets for mobile vs. computer browsers: -->
	<!--  <link href="style-mobile.css" rel="stylesheet" type="text/css" media="handheld"> -->
	<!--  <link href="style-computer.css" rel="stylesheet" type="text/css" media="screen"> -->
	<!-- The favicon & iOS home screen icon are both 57x57 PNG's. Use a full URL file path for Android devices.  -->
	<!--  <link rel="apple-touch-icon-precomposed" href="http://yoursite.com/apple-touch-icon.png">  -->
	<!--  <link rel="icon" type="image/vnd.microsoft.icon" href="http://yoursite.com/favicon.png">  -->
	<!-- Site maps help search spiders where to find your pages.  www.xml-sitemaps.com  -->
	<!--  <link rel="alternate" type="application/rss+xml" title="ROR" href="ror.xml"> -->
	<!-- Your Google Analytics code goes here, just before the </head> tag. -->

    <!-- 11/10/13: For the accordion menu from menucool.com, where its HTML is in a separate file, and does not have to be repeated in each webpage that has the menu. -->
    <link href="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.css" rel="stylesheet" type="text/css" />
    <script src="https://www.starlite-intl.com/mobile1/Misc/MenuCoolAmenuOneForAll/amenu/accordionmenu.js" type="text/javascript"></script>
	<script type="text/javascript">		amenu.close(true);</script>

    <!-- <meta name="viewport" content="initial-scale=1.0"> -->

</head>


<body style="border: solid 0px blue;">
<% 
    ' Load the JavaScript SDK for use in generating Facebook buttons below; see "fb-like" below. 
    ' My code for both was generated at https://developers.facebook.com/docs/plugins/like-button
%>
<div id="fb-root"></div>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = 'https://connect.facebook.net/en_US/sdk.js#xfbml=1&version=v2.11';
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));
</script>


<table style="border: solid 0px green; align: auto;" width="100%">   <% ' Start Table 1 %>
<tr><td>

<% InArea = "Products" %>

<!-- #include virtual="mobile1/Misc/Header.INC" -->


<% ' ************************** %>

<% ' 2/7/18: This form is from the included file "INC/LeftColumn.inc.asp" in Detail.asp of main site. %>
                        <center>
						<form method="get" name="Country">
						<br><font face="Tahoma" size="2">You are currently a 
						<% If Session("Country") = "Canada" Then %>
							<img src="https://www.starlite-intl.com/Images/can1.gif" WIDTH="36" HEIGHT="18" alt="Canada" title="Canada"> 
						<% Else				' Previously: ElseIf Session("Country") = "USA" Then 
							Session("Country") = "USA"
						%>
							<img src="https://www.starlite-intl.com/Images/USA1.gif" WIDTH="34" HEIGHT="18" alt="USA" title="USA"> 
						<% End if %> 
						customer. Click on a button to change countries:
						</font>
								
                        <input type="submit" name="Canada" value="Canada">
                        <input type="submit" name="  USA  " value="USA">

						<input type="hidden" name="pid" value="<%=request("pid")%>">
						<input type="hidden" name="sid" value="<%=request("sid")%>">
						<input type="hidden" name="area" value="<%=request("area")%>">
						<input type="hidden" name="sar" value="<%=request("sar")%>">
						<input type="hidden" name="action" value="<%=request("action")%>">

                        <% ' Added 12/18/17 ... to try to recover from a file corruption that occured on saving this?/some? file. 
                           ' Need If-Thens below to avoid type mismatch error in request(x) = "" cases.
                            If request("SpecialsSubgroup") <> "" Then
                                Response.Write "<input type='hidden' name='SpecialsSubgroup' value='" & request("SpecialsSubgroup") & "'>"
                            End If
                            If request("NewProductsSubgroup") <> "" Then
                                Response.Write "<input type='hidden' name='NewProductsSubgroup' value='" & request("NewProductsSubgroup") & "'>"
                            End If
                            If request("RebatesSubgroup") <> "" Then
                                Response.Write "<input type='hidden' name='RebatesSubgroup' value='" & request("RebatesSubgroup") & "'>"
                            End If
                        %>

						</form>
                        </center>

<% 
'Response.Write "Request('Canada') = "		& Request("Canada")
'Response.Write "<br>Request('  USA  ') = "	& Request("  USA  ")
'Response.Write "<br>Session('Country') = "	& Session("Country")
%>

<% ' ************************** %>


<% '*********************************************************************************************************************** %>


<% ' Start Table 1.1 %>
<table style="border-right:0px solid #84bff1;" width="100%" XXXwidth='1120' cellpadding="0" cellspacing="0" align="center" > <% ' Start Table 1.1 %>
    <tr>

        <!-- <td width="100%" background="Images/bluebackground2.jpg" valign=top> -->	
		<td width="100%" valign=top>
					<br>
					<% ' Start Table 1.1.2 %>
					<table border="0" cellpadding="0" cellspacing="0" align="center" width=100%>  
                    <!--
                    <tr>
                        <td align="right">
                           <% ' My code for Facebook buttons was generated at https://developers.facebook.com/docs/plugins/like-button  
                           ' See there for parameter setting options. %>
                            <div class="fb-like" data-href="https://www.facebook.com/starliteintl" data-width="100" data-layout="button" data-action="like" data-size="large" data-show-faces="false" data-share="true"></div> 
                            <br />
                        </td>
                    </tr>
                    -->
					<tr>
						<td valign="top" align="center">
						<!-- #include file="DETAIL.INC" -->
                		</td>
            		</tr>
            		</table>				
            		<% ' End Table 1.1.2 %>
            		   		
		</td>         
	</tr>
		 
</table>		
<% ' End Table 1.1 %>
       		
       		 
<!--#INCLUDE file="Misc/Footer.INC"-->

</td>
</tr>
	
	
</table>   
<% ' End Table 1 %>

</body>

</html>





