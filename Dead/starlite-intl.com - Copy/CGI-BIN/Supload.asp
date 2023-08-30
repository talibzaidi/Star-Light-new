<html>

<head>

<title>Sanction - Version (Orange)</title>
</head>

<body bgcolor="#000000" text="#FFBD00" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">

<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td bgcolor="#FFBD00"><font face="Arial"><img
        src="Simages/sanction.gif"
        width="330" height="82"></font></td>
        <td bgcolor="#FFBD00"><font face="Arial"></font>&nbsp;</td>
        <td align="right" bgcolor="#FFBD00"><font face="Arial">

<% if Session("Acess") = 1 or  Session("Acess") = 2 or Session("Acess") = 3 Then %>
<a href="../ADMI/sanction.asp">
<% end if %>
<img
        src="Simages/homegif.GIF"
        width="84" height="82" border="0"></a></font></td>
    </tr>
    <tr>
        <td><font face="Arial"><img
        src="Simages/btcurve.gif"
        width="330" height="82"></font></td>
        <td><font face="Arial"></font>&nbsp;</td>
        <td><font size="2" face="Arial" color="#FFBD00">CLIENT IMAGE UPLOAD: Use this to upload client images. Or don't.</font></td>
    </tr>
    <tr>
        <td><font face="Arial"><img
        src="Simages/blcurve.GIF"
        width="102" height="256"></font></td>
        <td colspan="2"><font face="Arial">


<!--#include file=libGlobalFuncs.inc-->







<!-- BEGIN TABLE CONTAINING  -->
<TABLE WIDTH=100% BORDER=0>
	<TR>
		<TD VALIGN=TOP >
			<TABLE WIDTH=100% CELLPADDING=5 CELLSPACING=5 BORDER=0>
				<TR>
					<TD>
						
  					</TD>
					<TD>
						<CENTER>
						<H2>Upload client images.</H2>
						</CENTER>

										

						<CENTER>
  <form enctype="multipart/form-data" action="../cgi-bin/upload.exe" method=post>
<input name="filename" type="file" SIZE="15">
<input type="submit" value=" Upload file ">
</form>
   
						</CENTER>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>



</font>&nbsp;</td>
        
    </tr>
</table>
</body>
</html>
