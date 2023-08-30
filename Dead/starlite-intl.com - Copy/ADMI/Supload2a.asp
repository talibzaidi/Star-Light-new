<html>



<head>
    <title>File Upload Processor</title>
    <link rel="stylesheet" type="text/css" href="https://www.starlite-intl.com/Misc/StyleSheet1.css"> <!-- 3/24/10: Copied from Mit Mazel; was needed to allow drop-down menus to work. -->
</head>



<body>

<% 
'Response.Write "<br>Response.Form('Submit') = " & Request.Form("Submit")


    <!-- From : http://developer.earthskater.net/asp/asp_fso.asp -->
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    For i = 1 To 3
        FN = "FileName" & i
        FileName = Request.Form(FN)
        Response.Write "<br><br>File " & FileName
        FileName = "../Imi/" & FileName
        If (fso.FileExists(server.mappath(FileName))) = True Then
              Response.Write " <font color='red'><b>exists</b></font>."
        Else
              Response.Write " does not exist yet."
        End If
    Next

    Set fso     = nothing
%>

</body>


</html>
