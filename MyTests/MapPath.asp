<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>

<%
Path = Server.mapPath("test1.asp")						' This worked!
Response.Write "<br>Path = " & Path

Path = Server.mapPath("../EDATA/ec-star-001.mdb")		' This worked!	
Response.Write "<br>Path = " & Path
%>

</body>

</html>
