<%
  AdNum = Request("AdNum")
  if AdNum < 1 then
    AdNum = 0
  end if

  
   
  set rstAdvertisement = CreateObject("ADODB.Recordset")
  rstAdvertisement.Open "SELECT * FROM Banner WHERE Advertisement=" & AdNum, "DSN=STAREC1", 1, 2

  'Determine if we are adding or updating
  if rstAdvertisement.EOF then
    rstAdvertisement.AddNew
  end if

  'Load the database buffer
  rstAdvertisement.Fields("ALocation") = Request("Location")
  rstAdvertisement.Fields("AURL") = Request("URL")
  rstAdvertisement.Fields("AImage") = Request("Image")
  rstAdvertisement.Fields("AName") = Request("AName")
  
  if Request("Show") = "T" then
    rstAdvertisement.Fields("AShow") = true
  else
    rstAdvertisement.Fields("AShow") = false
  end if

  'Update the database and retrieve the record id
  rstAdvertisement.Update
  AdUnq = rstAdvertisement("Advertisement")

  rstAdvertisement.Close
  set rstAdvertisement = nothing

  response.redirect "sbg.asp"
%>
<html>

<head>
<title>Sanction - Version Orange</title>
</head>

<body>
</body>
</html>
