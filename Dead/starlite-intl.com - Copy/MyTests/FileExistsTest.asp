<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">


<head>
    <title></title>
</head>


<body>

<%
Response.Write "Hello!!!"
'Response.End

<!-- From : http://developer.earthskater.net/asp/asp_fso.asp -->

Set fs=Server.CreateObject("Scripting.FileSystemObject")
fname = "../Imi/00002.gif"
If (fs.FileExists(server.mappath(fname)))=true Then
      Response.Write("<br><br>File " & fname & " exists.")
Else
      Response.Write("<br><br>File " & fname & " does not exist.")
End If

set fs=nothing
%>

</body>


</html>
