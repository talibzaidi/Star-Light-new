
<% ActiveTabColor = "silver" %>

<table bgcolor='lightblue' align='center' width='100%'>
<tr>
	<td>
	<table align="center" border="0" bordercolor="red" cellpadding="5">   
	<tr>
		<td align=center <% If InSection="Login" Then%> bgcolor=<%=ActiveTabColor%><% End If %>>
		<a href="http://www.starlite-intl.com/Admin2/Login.asp">
		<font size=2>Login</font></a>
		</td>
		
		<td align=center>
		<a href="http://www.starlite-intl.com/Admin2/Logout.asp">
		<font size=2>Logout</font></a>
		</td>

		<td align=center <% If InSection="Products" Then%> bgcolor=<%=ActiveTabColor%><% End If %>>
		<a href="http://www.starlite-intl.com/Admin2/Products.asp">
		<font size=2>Products</font></a>
		</td>
		
		<td align=center <% If InSection="Areas" Then%> bgcolor=<%=ActiveTabColor%><% End If %>>
		<a href="http://www.starlite-intl.com/Admin2/Areas.asp">
		<font size=2>Areas</font></a>
		</td>
		
		<td align=center <% If InSection="SubAreas" Then%> bgcolor=<%=ActiveTabColor%><% End If %>>
		<a href="http://www.starlite-intl.com/Admin2/SubAreas.asp">
		<font size=2>SubAreas</font></a>
		</td>
		
		<td align=center <% If InSection="Orders" Then%> bgcolor=<%=ActiveTabColor%><% End If %>>
		<a href="http://www.starlite-intl.com/Admin2/Orders.asp">
		<font size=2>Orders</font></a>
		</td>
		
		<td align=center <% If InSection="Emails" Then%> bgcolor=<%=ActiveTabColor%><% End If %>>
		<a href="http://www.starlite-intl.com/Admin2/Emails.asp">
		<font size=2>Emails</font></a>
		</td>
		
		<td align=center <% If InSection="EmailSend" Then%> bgcolor=<%=ActiveTabColor%><% End If %>>
		<a href="http://www.starlite-intl.com/Admin2/EmailSend.asp">
		<font size=2>Email Send</font></a>
		</td>
		
	</tr>
	</table>
	</td>
</tr>
</table>
            

