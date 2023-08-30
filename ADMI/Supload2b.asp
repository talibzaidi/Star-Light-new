<html>



<head>
    <title>File Upload Processor</title>
    <link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
</head>



<body>

<% 
'Response.Write "<br>Response.Form('Submit') = " & Request.Form("Submit")

Set Upload = Server.CreateObject("Persits.Upload.1") 
'Set Upload = Server.CreateObject("Persits.Upload") 
' Re "Upload.OverwriteFiles = False" below, see Section 2.5 of http://www.aspupload.com/manual_simple.html. It says there ...
' To prevent name collisions, AspUpload appends the original file name with an integer number in parentheses. 
' For example, if the file MyFile.txt already exists in the upload directory, and another file with the same name is 
' being uploaded, AspUpload will save the new file under the name MyFile(1).txt. If more copies of MyFile.txt are uploaded, 
' they will be saved under the names MyFile(2).txt, MyFile(3).txt, etc. 
Upload.OverwriteFiles = False 

'Count = Upload.Save("C:\websites\4rft4c\imi")  ' Old version, disabled by hosting company. Use the following instead.
Count = Upload.SaveVirtual("\imi") 
'Response.End
%>

<br />
<% = Count %> files uploaded. 

<br /><br />Files (a number in paretheses denotes a duplicate file name):<br /><br />
<%
For Each File in Upload.Files
    Response.Write "&nbsp;&nbsp;&nbsp;" & File.Name & " = " & File.Path & "&nbsp;&nbsp;&nbsp;(" & File.Size &" bytes)"

    'If fso.FileExists(File.Path) Then
    '    Response.Write "&nbsp;&nbsp;" & File.Name & " exists"
    '    'fso.DeleteFile("c:\test.txt")
    'End If

    Response.Write "<br>"
Next

Set Upload  = nothing 
%>

</body>


</html>
