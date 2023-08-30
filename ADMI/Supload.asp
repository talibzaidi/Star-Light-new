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
    <td align="right" bgcolor="#FFBD00"><font face="Arial"><a href="sanction.asp"><img
    src="Simages/homegif.GIF"
    width="84" height="82" border="0"></a></font></td>
</tr>

<tr>
    <td><font face="Arial"><img
    src="Simages/btcurve.gif"
    width="330" height="82"></font></td>
    <td><font face="Arial"></font>&nbsp;</td>
    <td><font size="2" face="Arial" color="#FFBD00">CLIENT IMAGE UPLOAD: Use this to upload client images. </font></td>
</tr>

<tr>
    <td><font face="Arial"><img
    src="Simages/blcurve.GIF"
    width="102" height="256"></font></td>
    <td colspan="2">

        <font face="Arial">
        <!-- BEGIN TABLE CONTAINING  -->
        <TABLE WIDTH=100% BORDER="0">
	    <TR>
		    <TD VALIGN=TOP >
			    <TABLE WIDTH=100% CELLPADDING=5 CELLSPACING=5 BORDER=0>
				<TR>
					<TD>
						<CENTER>
						    <H2>Check for Duplicate File Names</H2>
                            <!--
                            <form enctype="multipart/form-data" action="../cgi-bin/upload.exe" method=post>
                            <input name="filename" type="file" SIZE="15">
                            <input type="submit" value=" Upload file ">
                            </form>
                            -->
                            <FORM METHOD="POST" ACTION="Supload2a.asp">
                                <INPUT TYPE=text SIZE=60 NAME="FileName1"><BR>
                                <INPUT TYPE=text SIZE=60 NAME="FileName2"><BR>
                                <INPUT TYPE=text SIZE=60 NAME="FileName3"><BR>
                                <INPUT TYPE=SUBMIT Name="Submit0" VALUE="Check">
                            </FORM>
						</CENTER>
  					</TD>

					<TD>
						<CENTER>
						    <H2>Upload Client Images</H2>
                            <!--
                            <form enctype="multipart/form-data" action="../cgi-bin/upload.exe" method=post>
                            <input name="filename" type="file" SIZE="15">
                            <input type="submit" value=" Upload file ">
                            </form>
                            -->
                            <FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="Supload2b.asp">
                                <INPUT TYPE=FILE SIZE=60 NAME="FILE1"><BR>
                                <INPUT TYPE=FILE SIZE=60 NAME="FILE2"><BR>
                                <INPUT TYPE=FILE SIZE=60 NAME="FILE3"><BR>
                                <INPUT TYPE=SUBMIT Name="Submit" VALUE="Upload!">
                            </FORM>
						</CENTER>
					</TD>
				</TR>
			    </TABLE>
		    </TD>
	    </TR>
        </TABLE>

        </font>&nbsp;
    </td>
        
</tr>
</table>

</body>

</html>
