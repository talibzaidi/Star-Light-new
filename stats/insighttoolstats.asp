<%currtime = convertTime(now())


thepage = request.servervariables("URL")
thepos=0
found = true
while found
	currpos = instr(thepos+1,thepage,"/")
	if currpos > thepos then
		thepos = currpos
	end if
	if currpos = 0 then
		found = false
	end if
wend
actpage = mid(thepage,thepos+1,len(thepage) - thepos)
rootdir = left(thepage,len(thepage)-len(actpage))

statdbpath = Server.MapPath(".")
if thepos > 1 then
statdbpath = left(statdbpath,(len(statdbpath) - len(rootdir) +1))
end if

statdbpath = statdbpath & "\stats\statsdb.mdb"

theconnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & statdbpath
Set statConn = Server.CreateObject("ADODB.Connection")
Set SRS = Server.CreateObject("ADODB.Recordset")
statConn.Open theconnection

time24 = convertTime(DateAdd("d","-1",now()))
timeweek = convertTime(DateAdd("d","-7",now()))
timemonth = convertTime(DateAdd("m","-1",now()))
timemonth2 = convertTime(DateAdd("m","-2",now()))
timemonth3 = convertTime(DateAdd("m","-3",now()))
timemonth4 = convertTime(DateAdd("m","-4",now()))
timemonth5 = convertTime(DateAdd("m","-5",now()))
timemonth6 = convertTime(DateAdd("m","-6",now()))

ssql = "delete * from SiteCount where Date < '" & timemonth6 & "'"
SRS.Open ssql, statConn, 3,3
ssql = "delete * from PageCount where Date < '" & timemonth6 & "'"
SRS.Open ssql, statConn, 3,3

thedatatype = request.form("databut")
thetimerefer = request.form("timebut")
if thedatatype = "" then
	thedatatype = "hits"
end if
if thetimerefer = "" then
	thetimerefer = "24hr"
end if


	

%>

<html>
<head>
<title>Insight Stats Tool</title>
<script language="javascript">
function SetDataType(thedata)
{

	document.datatype.databut.value=thedata;
	document.datatype.submit();
}
function SetTimeRefer(thetime)
{	document.timerefer.timebut.value=thetime;
	document.timerefer.submit();
}
function showRemote(thepage,thewindow,wid,hei) {


var windowprops = "toolbar=0,location=0,directories=0,status=0, " +
"menubar=0,scrollbars=1,resizable=1,width="+wid+",height="+hei+",screenX=10,screenY=10,top=10,left=10";

OpenWindow = window.open(thepage, thewindow, windowprops);
self.focus();
}
function showVal(theval)
{
document.timerefer.barval.value=theval;
}
function pageNav(thenav)
{

if(thenav=="first")
{
	document.nav.intpage.value=1;
}
if(thenav=="prev")
{
	document.nav.intpage.value=parseFloat(document.nav.intpage.value) - 1;
	if(document.nav.intpage.value == 0)
	{
		document.nav.intpage.value = 1;
	}
}
if(thenav=="next")
{
	document.nav.intpage.value=parseFloat(document.nav.intpage.value)+1;
}
if(thenav=="last")
{
	document.nav.intpage.value = -1
}

document.nav.submit();
}
</script>
<link rel="stylesheet" type="text/css" href="insight.css">
</head>

<BODY BGCOLOR=#FFFFFF link="#000000" vlink="#000000" alink="#808000" leftmargin="0" marginheight="0" marginwidth="0" topmargin="0">
<!-- ImageReady Slices (insighttooltracker.jpg) -->
<TABLE WIDTH=722 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD COLSPAN=6>
			<IMG SRC="images/insighttooltracker_01.gif" WIDTH=722 HEIGHT=1></TD>
	</TR>
	<TR>
		<TD ROWSPAN=3>
			<IMG SRC="images/insighttooltracker_02.gif" WIDTH=3 HEIGHT=527></TD>
		<TD ROWSPAN=3>
			<IMG SRC="images/insighttooltracker_03.gif" WIDTH=20 HEIGHT=527></TD>
		<TD COLSPAN=3>
			<IMG SRC="images/insighttooltracker_04.gif" WIDTH=671 HEIGHT=108></TD>
		<TD ROWSPAN=3>
			<IMG SRC="images/insighttooltracker_05.gif" WIDTH=28 HEIGHT=527></TD>
	</TR>
	<TR>
		<TD COLSPAN=3>
			<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="363">
              <tr>
                <td width="100%">
                	<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
						<tr>
							<td>
								<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
									<tr height="40">
										<td colspan="2" height="40">
											<form name="datatype" method="post">
											<table width="100%" cellpadding="0" cellspacing="0" border="0">
												<tr>
													<td align="left"><font class="menu">Server Time:&nbsp;<%=now()%></font></td>
													<td align="center" width="100">
													<%if thedatatype="hits"then
														thebut = "hitsdown.gif"
													else 
														thebut = "hitsup.gif"
													end if%><input type="image" name="hits" onClick="SetDataType('hits');" src="images/<%=thebut%>"></td>
													<td align="center" width="100">
													<%if thedatatype="Refer"then
														thebut = "referdown.gif"
													else 
														thebut = "referup.gif"
													end if%><input type="image" name="refer" onClick="SetDataType('Refer');" src="images/<%=thebut%>"></td>
													<td align="center" width="100">
													<%if thedatatype="IP"then
														thebut = "ipdown.gif"
													else 
														thebut = "ipup.gif"
													end if%><input type="image" name="ip" onClick="SetDataType('IP');" src="images/<%=thebut%>"></td>
													<td align="center" width="100">
													<%if thedatatype="Page"then
														thebut = "pvdown.gif"
													else 
														thebut = "pvup.gif"
													end if%><input type="image" name="page" onClick="SetDataType('Page');" src="images/<%=thebut%>"></td>
												</tr>
											</table>
											<input type="hidden" name="databut" value="<%=request.form("databut")%>"> 
											<input type="hidden" name="timebut" value="<%=request.form("timebut")%>">
											
											</form>
										</td>
									</tr>
									<tr>
										<td valign="top" width="120">
											<form name="timerefer" method="post">
											<table width="120" cellpadding="0" cellspacing="0" border="0">
												<tr height="40">
													<td align="center" height="40">
													<%if thetimerefer="24hr"then
														thebut = "lst24down.gif"
													else 
														thebut = "lst24up.gif"
													end if%><input type="image" name="last24" onClick="SetTimeRefer('24hr');" src="images/<%=thebut%>"></td>
												</tr>
												<tr height="40">
													<td align="center" height="40">
													<%if thetimerefer="days"then
														thebut = "lstmnthdown.gif"
													else 
														thebut = "lstmnthup.gif"
													end if%><input type="image" name="days" onClick="SetTimeRefer('days');" src="images/<%=thebut%>"></td>
												</tr>
												<tr height="40">
													<td align="center" height="40">
													<%if thetimerefer="months"then
														thebut = "lst6mnthdown.gif"
													else 
														thebut = "lst6mnthup.gif"
													end if%><input type="image" name="months" onClick="SetTimeRefer('months');" src="images/<%=thebut%>"></td>
												</tr>
												<tr>
													<td align="center" height="40">&nbsp;</td>
												</tr>
												<tr>
													<td align="center" height="40"><input type="image" name="export" onClick="showRemote('export.asp?timebut=<%=request.form("timebut")%>&databut=<%=request.form("databut")%>','Export',10,10);" src="images/export.gif"></td>
												</tr>
												<tr><td height="40" align="center">&nbsp;<font class="menu">Max.</font><input type="text" name="maxval" readonly size="10" border="0"></td></tr>
												<tr><td height="40" align="center">&nbsp;<font class="menu">Value</font><input type="text" name="barval" size="10" border="0"></td></tr>
											</table>
											<input type="hidden" name="timebut" value="<%=request.form("timebut")%>">
											<input type="hidden" name="databut" value="<%=request.form("databut")%>">
											
											</form>
								
										</td>
										<td valign="top">
											<table width="100%" cellpadding="0"cellspacing="0" border="0">
										<%
								'DIM drArray
								if thedatatype = "hits" then
								
									ssql1 = "select Date, sum(Count) as Total from SiteCount where Date like '"
									ssql2 = "%' group by Date"
									
									drcount = 0 
									if thetimerefer="24hr" then
									
										for x = 24 to 0 step -1
											REDIM Preserve drArray(2, drcount+1)
											datetosearch = convertTime(DateAdd("h","-" & x,now()))
											hourtodisplay = hour(DateAdd("h","-" & x,now()))
											datetosearch = left(datetosearch,10)
											ssql = ssql1 & datetosearch & ssql2
											SRS.Open ssql,statConn, 1,4
											
											drArray(0,drcount) = hourtodisplay
											
											if not SRS.EOF then
												drArray(1,drcount) = SRS.recordcount
											else
												drArray(1,drcount) = 0
											end if
											SRS.Close
											drcount = drcount + 1
										next
									end if
									if thetimerefer="days" then
										for x = 30 to 0 step -1
											REDIM Preserve drArray(2, drcount+1)
											datetosearch = convertTime(DateAdd("d","-" & x,now()))
											daytodisplay = day(DateAdd("d","-" & x,now()))
											datetosearch = left(datetosearch,8)
											ssql = ssql1 & datetosearch & ssql2
											SRS.Open ssql,statConn, 1,4
											drArray(0,drcount) = daytodisplay
											if not SRS.EOF then
												drArray(1,drcount) = SRS.recordcount
											else
												drArray(1,drcount) = 0
											end if
											SRS.Close
											drcount = drcount + 1
										next
									end if
									if thetimerefer="months" then
										for x = 6 to 0 step -1
											REDIM Preserve drArray(2, drcount+1)
											datetosearch = convertTime(DateAdd("m","-" & x,now()))
											monthtodisplay = month(DateAdd("m","-" & x,now()))
											datetosearch = left(datetosearch,6)
											ssql = ssql1 & datetosearch & ssql2
											SRS.Open ssql,statConn, 1,4
											
											drArray(0,drcount) = monthtodisplay
											if not SRS.EOF then
												drArray(1,drcount) = SRS.recordcount
											else
												drArray(1,drcount) = 0
											end if
											SRS.Close
											drcount = drcount + 1
										next
									end if%>
												<tr><script language="javascript">
									document.timerefer.maxval.value=<%=max(drArray)%>;
									</script><%if max(drArray) > 0 then
									
									for x = 0 to Ubound(drArray,2)-1
										if max(drArray) > 0 then
										totPixels = Int((drArray(1,x)*300)/max(drArray))
										else
										totPixels = 0
										end if%>
													<td width ="15" align="center" valign="bottom"><a href="#" onmouseover="showVal(<%=drArray(1,x)%>);" onmouseout="showVal('');"><img src="images/bar.gif" width="10" height="<%=totPixels%>" border="0"></a></td>
																<%next%>
												</tr>
												<tr><%
									for x = 0 to Ubound(drArray,2)-1%>
													<td width="15" align="center"><font class="statsreport"><%=drArray(0,x)%></font></td>
																<%next%>
												</tr><%
										
									else%>
													<td><font class="statsreport">No Stats</font></td>
												</tr>
									<%end if
								else
									drtable="SiteCount"
									if thedatatype = "Page" then
										drtable="PageCount"
									end if
									if thetimerefer = "24hr" then
										drtime = time24
									end if
									if thetimerefer = "days" then
										drtime = timemonth
									end if
									if thetimerefer = "months" then
										drtime = timemonth6
									end if	
									ssql = "select " & thedatatype & ", sum(Count) as Total from " & drtable & " where Date >= '" & drtime &"' group by " & thedatatype
									
									SRS.Open ssql, statConn, 1,4
									
									
										if request.form("intpage") <> "" then
											intpage = request.form("intpage")
										else 
											intpage=1
										end if
										
									
		
									
									
									
									
									
									if not SRS.EOF then
										'Dim drArray
										drArray = SRS.getrows
										%><script language="javascript">
									document.timerefer.maxval.value=<%=max(drArray)%>;
									</script>
										<%maxpages = int(UBound(drArray,2)/10)
										
										if (UBound(drArray,2) mod 10) > 0 then
											maxpages = maxpages + 1
										end if
										if maxpages = 0 then
											maxpages = 1
										end if
										if CInt(intpage) > maxpages or CInt(intpage) = -1 then
											intpage = maxpages
										end if
										
										
										
										pagestart = (intpage-1) * 10
										pageend = (intpage*10-1)%>
										<form name="nav" method="post">
										<tr>
										<td colspan="2" align="right"><input type="image" src="images/first.gif" name="first" value="<<" onClick="pageNav('first');">
										<input type="image" name="prev" src="images/prev.gif" value="<" onClick="pageNav('prev');">
										<input type="image" name="next" src="images/next.gif" value=">" onClick="pageNav('next');">
										<input type="image" name="last" src="images/last.gif" value=">>" onClick="pageNav('last');">
										</td></tr>
									<input type="hidden" name="timebut" value="<%=request.form("timebut")%>">
											<input type="hidden" name="databut" value="<%=request.form("databut")%>">
											<input type="hidden" name="intpage" value="<%=intpage%>">
										</form>
										<tr>
											<td colspan="2" align="right"><font class="reccount">Page:<%=intpage%>&nbsp;of&nbsp<%=maxpages%></font></td>
										</tr>
										<tr><td colspan="2">&nbsp;</td></tr>
									<%
									
									
										for x = pagestart to pageend
											if x > UBound(drArray,2) then
												exit for
											end if
											if max(drArray) > 0 then
											totPixels = Int((drArray(1,x)*340)/max(drArray))
											else
											totPixels = 0
											end if
											 
											if len(drArray(0,x)) <= 30 then
												thedatalen = len(drArray(0,x))
											else
												thedatalen = 30
											end if
												%>
												<tr height="24">
												<%if thedatatype="IP" then%>
												<td width="210" align="center" height="24"><font class="statsreport"><%=drArray(0,x)%></font></td>
													
												<%else%>
												<td width="210" height="24"><font class="statsreport"><a href="#" onClick="alert('<%=drArray(0,x)%>');"><%=left(drArray(0,x),thedatalen)%></a></font></td>
													
												<%end if%>
													<td height="24"><a href="#" onmouseover="showVal(<%=drArray(1,x)%>);" onmouseout="showVal('');"><img src ="images/bar.gif" height="10" width="<%=totPixels%>"></a></td>
												</tr>
										<%next
									else%>
												<tr height="24">
													<td colspan="2" height="24"><font class="statsreport">No Stats</font></td>
												</tr>
									<%end if
								end if%>
									
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
             	</td>
       		</tr>
    	</table>
  	</TD>
</TR>
	<TR>
		<TD valign="bottom">
			<IMG SRC="images/insighttooltracker_07.gif" WIDTH=166 HEIGHT=52></TD>
		<TD valign="bottom">
			<IMG SRC="images/insighttooltracker_08.gif" WIDTH=257 HEIGHT=52></TD>
		<TD valign="bottom">
			<A HREF="http;/www.internetadvertisingcorp.com">
				<IMG SRC="images/insighttooltracker_09.gif" WIDTH=248 HEIGHT=52 BORDER=0></A></TD>
	</TR>
</TABLE>
<!-- End ImageReady Slices -->
</BODY>
</HTML>
<%


statConn.Close
Set SRS = nothing
set statConn = nothing
Function Max(aNumberArray)
	Dim I
	Dim dblHighestSoFar

	' Y'all really don't want comments on this one too, do you?  It's exactly the
	' same as above except for the > instead of <.  I also changed the variable name
	' from dblLowestSoFar to dblHighestSoFar so it made more sense.
	
	' Notice about the "Y'all"...
	'             we've just moved to Georgia and I'm practicing my accent!  ;)

	dblHighestSoFar = Null

	For I = 0 to UBound(aNumberArray,2)
		' Testing line left in for debugging if needed
		'Response.Write aNumberArray(1,I) & "<BR>"
		If IsNumeric(aNumberArray(1,I)) Then
			If CDbl(aNumberArray(1,I)) > dblHighestSoFar Or IsNull(dblHighestSoFar) Then
				dblHighestSoFar = CDbl(aNumberArray(1,I))
			End If
		End If
	Next 'I
	
	Max = dblHighestSoFar
End Function



Function convertTime(timein)
			timeout = Year(timein)
		if month(timein) < 10 then
			timeout = timeout & "0"
		end if
		timeout = timeout & month(timein)
		if day(timein) < 10 then
			timeout = timeout & "0"
		end if
		timeout = timeout & day(timein)
		if hour(timein) < 10 then
			timeout = timeout & "0"
		end if
		timeout = timeout & hour(timein)
		if minute(timein) < 10 then
			timeout = timeout & "0"
		end if
		timeout = timeout & minute(timein)
		if second(timein) < 10 then
			timeout = timeout & "0"
		end if
		timeout = timeout & second(timein)
	convertTime=timeout
end function%>
