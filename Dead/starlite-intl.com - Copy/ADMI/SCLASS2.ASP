<%@ LANGUAGE = VBScript %>
<!--#include file="ADOVBS.INC"-->
<% 
    If (Session("Access") < "1") Then 
	Response.Redirect "login.asp"
    End If
%>
<%
msg=""

Action = Left(UCase(Request("Action")),6)
mSubmitted = date & " " & time

id = CInt(Request("Classified")) 

If Action = "SUBMIT" Then
	
	
              
		 SQL = "Delete * From CLASSFD WHERE Index =" & id
                 Set conn = Server.CreateObject("ADODB.Connection")
    		 Conn.Open Session("ConnectionString")
    		 Conn.Execute(SQL)
		 Response.Redirect "sclass.asp" 
	
End If  'Action = "Submit"
%>


<HTML>
<HEAD>
<TITLE>SANCTION: USER ADMINISTRATION TOOL - C.1999 - DESIGNWERX</TITLE>
</HEAD>
<body bgcolor="#000000" TOPMARGIN="5" LEFTMARGIN="5" LINK="#E00000" VLINK="#E00000">



</BODY>
</HTML>