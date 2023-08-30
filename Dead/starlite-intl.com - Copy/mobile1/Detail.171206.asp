<%@ LANGUAGE = VBScript %>


<!--[if IE]>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<![endif]-->
<!-- The above seems to be needed for IE to get the drop-down menubar to work properly. -->



<%response.buffer=true%>

<% PID = ReQuest("PID") %>
<% ar = Request("Area") %>
<% Area = Request("Area") %>

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
      If (Request("Canada") <> "" OR Request("  USA  ") <> "") then

 If Request("Canada") <> "" then

 Session("Country") = "Canada"


 else

 Session("Country") = "USA"


 end if

end if
%>

<% sar = ReQuest("sar") %>
<% 'sar = Replace( sar, " ", "%20") %>
<% SID = ReQuest("SID") %>


<%
Set RS = CreateObject("ADODB.Recordset")
'RS.Open "SELECT *, Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE PID = " + PID + " ", "DSN=STAREC1" , 1, 4

' 11/18/15: I replaced DSN method of line above with the folllowing Connection String method. 	
RS_SQL   = "SELECT *, Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE PID = " & PID & " "
Set Conn = Server.CreateObject("ADODB.Connection") 
Conn.Open Session("ConnectionString")
Set RS 	 = Conn.Execute(RS_SQL)	
Set Conn = Nothing

'on error resume next
'do while not rs.eof
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
    <link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/mobile1/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
	<meta charset=utf-8>
	<title>GPS Best Source: Garmin GPS, USglobalSat GPS, Pharos GPS. Cobra, Uniden, Midland, Motorola and brand name manufacturers of communication and electronic products. Hand tools</title>
	<meta name="keywords" content="Gps,Gps sensors,Gps sensor,Gps engine,Gps navigation,oem Gps,GpsMap,Nuvi,Gps accessories,Gps system,tracker Gps,auto Gps,portable Gps,handheld Gps,marine Gps,Gps marine network,Gps receiver,Gps antennas,fish finder,sounders,Gps cartography,Gps equipment,pda Gps,Garmin Gps,bluetooth Gps,global positioning,tracking Gps,fleet tracking Gps,USglobal Gps,discount Gps,Gps on sale,navigation electronics,Cobra,Midland,amateur radios,Galaxy radio,Magnum radio,radio scanner,radio scanners,scanner,digital cameras,power supplies,regulated power supplies,Fujifilm,Nikon,Olympus,Panasonic,Motorola,Canon">
	<meta name="description" content="Large selection of Garmin GPS,USglobal GPS,OEM GPS,GPS sensors,bluetooth GPS,fish finders,sounders,CB radios and walky-talky,radio scanners,digital cameras,car audio and video,hand tools,mechanics tools">
	<meta http-equiv="content-language" content="en">
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">

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

	<script language="Javascript">
	<!--
		once = new MakeArray(6)
		over = new MakeArray(6)
		under = new MakeArray(6)
		standard = new MakeArray(1)
		once[0].src = "Images/question1.gif"
		once[1].src = "Images/scart1.gif"
		once[2].src = "Images/home1.gif"
		once[3].src = "Images/new1.gif"
					once[4].src = "Images/cat1.gif"
		once[5].src = "Images/ex1.gif"    
		over[0].src = "Images/question2.gif"
		over[1].src = "Images/scart2.gif"
		over[2].src = "Images/home2.gif"
		over[3].src = "Images/new2.gif"
		over[4].src = "Images/cat2.gif"
		over[5].src = "Images/ex2.gif"
		under[0].src = "Images/helpnav.gif"
		under[1].src = "Images/shoppingcartnav.gif"
		under[2].src = "Images/homenav.gif"
		under[3].src = "Images/newproductsnav.gif"
		under[4].src = "Images/onlinecataloguenav.gif"
		under[5].src = "Images/specialsnav.gif"
		standard[0].src = "Images/emptynav.jpg"
	
	
	function MakeArray(n) 

		{

		this.length = n

		for (var i = 1; i<=n; i++) 

			{

			this[i-1] = new Image()

			}

		return this

		}

	function msover(inum,d_inum) 

		{

			if ((over[inum].src != "")) 

				{

				document.images[d_inum].src = over[inum].src
				document.images[7].src = under[inum].src
				}

		}


	function msout(inum,d_inum) 

		{

			if ((once[inum].src != "")) 

				{

				document.images[d_inum].src = once[inum].src
				document.images[7].src = standard[0].src
				}

		}

	// -->
	</script>

	<style>
		td {border: 0px solid black;}
	</style>

</head>


<body>


<table style="border:0px solid blue;" XXXwidth='1100'  XXXbgcolor="" align='center'>		<% ' Start Table 1 %>
<tr><td>

<% InArea = "Products" %>

<!-- #include virtual="mobile1/Misc/Header.INC" -->


<% '*********************************************************************************************************************** %>


<% ' Start Table 1.1 %>
<table style="border-right:0px solid #84bff1;" XXXwidth='1120' cellpadding="0" cellspacing="0" align="center" > <% ' Start Table 1.1 %>
    <tr>

        <!-- <td width="100%" background="Images/bluebackground2.jpg" valign=top> -->	
		<td XXXwidth="100%" valign=top>
					<br>
					<% ' Start Table 1.1.2 %>
					<table border=0 cellpadding="0" cellspacing="0" align="center" XXXwidth=100%>  
					<tr>
						<td valign="top" align="center">
						<!--#include file="DETAIL.INC"-->
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





