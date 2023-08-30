 <%@ Language=VBScript %>
<%
Response.Buffer = True
%>
<%browserdetect = request.servervariables("HTTP_USER_AGENT")
if inStr(browserdetect,"MSIE") > 0 and inStr(browserdetect,"Mac_PowerPC") > 0 then%>
<script language="javascript">
alert("Log Exports are currently unavailable with Mac IE.  Please use Netscape or a Win version of IE.");
window.close();
</script>
<%
end if
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
dlpath = statdbpath
statdbpath = statdbpath & "\statsdb.mdb"

theconnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & statdbpath

set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set theCurrentFolder = objFSO.GetFolder( dlpath )
Set curFiles = theCurrentFolder.Files 

currentSlot = -1 ' start before first slot
' We collect all the info about each file and put it into one
' "slot" in our "theFiles" array. 
'
For Each fileItem in curFiles
	fname = fileItem.Name
	
	if inStr(fname,".log") > 0 then
		objFSO.DeleteFile(dlpath & "\" & fname)
	end if
Next



'Public Function exportdata(ByVal LogID As String, ByVal UserName As String, ByVal ConStr As String) As String
'LogID is the log we want to get
'UserName is the user name
'ConStr is the connection string


Dim FileName
Dim SelectStr

'determine which log is being called

'create filename for text file
FileName = request.servervariables("HTTP_HOST") & Year(Date) & Month(Date) & Day(Date) & Hour(Time) & Minute(Time) & Second(Time) & ".log"
'Establish data connection

Set statConn = Server.CreateObject("ADODB.Connection")
statConn.Open theconnection

Set SRS = Server.CreateObject("ADODB.Recordset")
time24 = convertTime(DateAdd("d","-1",now()))
timeweek = convertTime(DateAdd("d","-7",now()))
timemonth = convertTime(DateAdd("m","-1",now()))
timemonth2 = convertTime(DateAdd("m","-2",now()))
timemonth3 = convertTime(DateAdd("m","-3",now()))
timemonth4 = convertTime(DateAdd("m","-4",now()))
timemonth5 = convertTime(DateAdd("m","-5",now()))
timemonth6 = convertTime(DateAdd("m","-6",now()))
thedatatype = request.querystring("databut")
thetimerefer = request.querystring("timebut")
if thedatatype = "" then
	thedatatype = "hits"
end if
if thetimerefer = "" then
	thetimerefer = "24hr"
end if

'Create file path
Dim OutString
OutString = dlpath & "\" & FileName


set objSave = objFSO.OpenTextFile(OutString, 8, True)

if thedatatype = "hits" then
								
	ssql1 = "select Date, sum(Count) as Total from SiteCount where Date like '"
	ssql2 = "%' group by Date"
									
	drcount = 0 
	
	
	
	if thetimerefer="24hr" then
		objSave.WriteLine ("Last 24hrs|Current Time " & now() & vbCrLf)
		objSave.WriteLine ("Hour|Count" & vbCrLf)								
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
		objSave.WriteLine ("Last Month|Current Time " & now() & vbCrLf)
		objSave.WriteLine ("Date|Count" & vbCrLf)
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
		objSave.WriteLine ("Last 6 Months|Current Time " & now() & vbCrLf)
		objSave.WriteLine ("Date|Count" & vbCrLf)
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
	end if
	for x = 0 to Ubound(drArray,2)
		thewriteline = drArray(0,x) & "|" & drArray(1,x) & vbCrLf
		
		objSave.WriteLine (thewriteline)
		
		
	next
	objSave.Close
else
	drtable="SiteCount"
	if thedatatype = "Page" then
		drtable="PageCount"
	end if
	if thetimerefer = "24hr" then
		drtime = time24
		objSave.WriteLine ("Last 24hrs|Current Time " & now() & vbCrLf)
	end if
	if thetimerefer = "days" then
		drtime = timemonth
		objSave.WriteLine ("Last Month|Current Time " & now() & vbCrLf)
	end if
	if thetimerefer = "months" then
		drtime = timemonth6
		objSave.WriteLine ("Last 6 Months|Current Time " & now() & vbCrLf)
	end if	
	ssql = "select " & thedatatype & ", sum(Count) as Total from " & drtable & " where Date >= '" & drtime &"' group by " & thedatatype

	objSave.WriteLine (thedatatype &"|Count" & vbCrLf)
	
	SRS.Open ssql, statConn, 1,4
	if not SRS.EOF then
		drArray = SRS.getrows
		for x = 0 to UBound(drArray,2)
			objSave.WriteLine (drArray(0,x) & "|" & drArray(1,x) & vbCrLf)
			
		next
	end if
	SRS.Close
	objSave.Close
end if

Set objSave = nothing


Set objTextStream = objFSO.OpenTextFile(OutString, 1)
response.addheader "Content-Disposition", "attachment; filename="  & FileName
 	response.ContentType = "text"
	Response.binarywrite objTextStream.ReadAll
	objTextStream.Close
	Set objTextStream = Nothing
delresponse = objFSO.DeleteFile(OutString, False)
statConn.Close
Set SRS = nothing
Set statConn = nothing
Set objFSO = nothing



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

