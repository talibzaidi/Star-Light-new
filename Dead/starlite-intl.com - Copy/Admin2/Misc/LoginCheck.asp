

<% 
If NOT Session("LoggedIn") = "Yes" Then
	Response.Write "<br><br><br><br><br><br><br><br>"
	Response.Write "<center><font color=red size=4 face=Tahoma>You must be logged in to use this page.</font></center>"
	Response.End
End If
%>

