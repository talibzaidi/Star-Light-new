<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<% spec = 2 %>
<% locale = "Classified" %>
<!--#include file="ADOVBS.INC"-->
<% 
   If (Request("Message")="") Then
       Response.Redirect "addlist.asp"
   End If 
   If (Request("Author")="") Then
       Response.Redirect "addlist.asp"
   End If 
 
 
    On Error Resume Next
    SQL = "INSERT INTO CLASSFD ( Author, Message, Email, Area, Contact) "
        SQL = SQL & " VALUES( '"& ReQuest("Author") &"','"& ReQuest("Message") &"','"& ReQuest("Email") &"','"& ReQuest("R1") &"','"& ReQuest("Contact") &"' )"                             
    Set conn = Server.CreateObject("ADODB.Connection")
        Conn.Open Session("ConnectionString")
        Conn.Execute(SQL)
    
        strSQL = "SELECT * FROM CLASSFD WHERE Message = '"& ReQuest("Message") &"' "
    Set rst = Server.CreateObject("ADODB.Recordset")
    rst.Open strSQL, conn, adOpenStatic, _
     adLockOptimistic, adCmdText 
        ' Update record
     rst("Datet") = Now()
    rst.Update
    rst.Close
    Conn.Close

    
        SQL = "DELETE * FROM CLASSFD WHERE Datet <= (Now() - 30) "
        Set conn = Server.CreateObject("ADODB.Connection")
        Conn.Open Session("ConnectionString")
        Conn.Execute(SQL)
        Conn.Close
        SQL = "DELETE * FROM CLASSFD WHERE Datet = NULL "
        Set conn = Server.CreateObject("ADODB.Connection")
        Conn.Open Session("ConnectionString")
        Conn.Execute(SQL)
        Conn.Close
    
        
%>
<html>

<head>
<meta name="keywords" content="gps,cb radios,frs,gmrs,radio scanners,2-way radios,car audio,hand tools">
<meta name="description" content="Online store for GPS, cb radios, frs, gmrs, antennas, car audio, dj, hand tools.  Shopping on a secure SSL line. Accepting Visa, Mastercard, American Express cards.">
<meta name ="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds">

<title>Starlite International LLC - Online Store</title>
<script language="Javascript">
<!--

    once = new MakeArray(6)
    over = new MakeArray(6)
    under = new MakeArray(6)
    standard = new MakeArray(1)
 
    once[0].src = "Images/question1.gif"
    once[1].src = "Images/scart1.gif"
    once[2].src = "Images/home1.gif"
    once[3].src = "Images/new1.gif"
                once[4].src = "Images/cat1.gif"
    once[5].src = "Images/ex1.gif"    
        
    over[0].src = "Images/question2.gif"
    over[1].src = "Images/scart2.gif"
    over[2].src = "Images/home2.gif"
    over[3].src = "Images/new2.gif"
    over[4].src = "Images/cat2.gif"
    over[5].src = "Images/ex2.gif"

    under[0].src = "Images/helpnav.gif"
    under[1].src = "Images/shoppingcartnav.gif"
    under[2].src = "Images/homenav.gif"
    under[3].src = "Images/newproductsnav.gif"
    under[4].src = "Images/onlinecataloguenav.gif"
    under[5].src = "Images/specialsnav.gif"

    standard[0].src = "Images/emptynav.jpg"

function MakeArray(n) 

    {

    this.length = n

    for (var i = 1; i<=n; i++) 

        {

        this[i-1] = new Image()

        }

    return this

    }

function msover(inum,d_inum) 

    {

        if ((over[inum].src != "")) 

            {

            document.images[d_inum].src = over[inum].src
            document.images[7].src = under[inum].src
            }

    }


function msout(inum,d_inum) 

    {

        if ((once[inum].src != "")) 

            {

            document.images[d_inum].src = once[inum].src
            document.images[7].src = standard[0].src
            }

    }

// -->
</script>
</head>

<body background="Images/background.jpg" bgcolor="#FFFFFF"
link="#000000" vlink="#000000" topmargin="0" leftmargin="0"
marginwidth="0" marginheight="0">


<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td><div align="left"><table border="0" cellpadding="0"
        cellspacing="0" width="575">
            <tr>
                <td> <!--#include file="NAV.INC"--><img
                src="Images/toptitle.jpg" width="411" height="29"><br>
                </td>
            </tr>
            <tr>
                <td width="575"><img src="Images/emptynav.jpg"
                width="164" height="14"><img
                src="Images/bottitle.JPG" width="411" height="14"></td>
            </tr>
            <tr>
                <td><img src="Images/leftbar.gif" width="176"
                height="23"><a href="class.asp"><img
                src="Images/classifieds.gif" alt="Classifieds"
                border="0" width="115" height="23"></a><a
                href="links.asp"><img src="Images/links.gif"
                alt="Links" border="0" width="91" height="23"></a><a
                href="contact.asp"><img
                src="Images/ContactUS.gif" alt="Contact Us"
                border="0" width="126" height="23"></a><a
                href="help.asp"><img src="Images/Help.gif"
                alt="Help" border="0" width="67" height="23"></a></td>
            </tr>
        </table>
        </div></td>
        <td width="100%"
        background="Images/topback.gif">&nbsp;</td>
    </tr>
    <tr>
        <td><table border="0" cellpadding="5" cellspacing="0">
            <tr>
                <td align="center" valign="top" width="166"><img
                src="Images/logo.gif" width="140" height="145">

<!--#include file="SPECIAL.INC"-->
<br>

</td>
                <td valign="top" width="382"><div align="center"><center>
                </center></div>

<br>
  <table border="0" width="100%">
  
        <tr background="bg.gif">
            <td><div align="center"><center><table border="0" width="80%"  cellpadding="5">
                <tr>
                    <td><p align="center"><font color="#000000" face="Tahoma"><strong><u>Thank you for posting</u>:</strong></font></p>
                  
                    <p><font color="#000000" face="Arial"><strong>You
                    Posted in: </font><font face="Arial" color="000000"> <%=Request("R1")%></strong></font></p>
                    <p><font color="#000000" face="Arial"><strong>Your
                    contact name: </font><font face="Arial" color="000000"> <%=Request("Author")%></strong></font></p>
                <p><font color="#000000" face="Arial"><strong>Your
                    contact phone: </font><font face="Arial" color="000000"> <%=Request("Contact")%></strong></font></p> 
                    <p><font color="#000000" face="Arial"><strong>Your
                    ad: </font><font face="Arial" color="000000"> <%=Request("Message")%></strong></font></p>
                    <p><font color="#000000" face="Arial"><strong>Your
                    email: </font><font face="Arial" color="000000"> <%=Request("Email")%></strong></font></p>
                    <p><font color="#000000" face="Arial"><strong>Date:&nbsp;<%=Request("Date")%></font><font face="Arial" color="000000"> 
                    </strong></font></p>
                    <p align="center"><font face="Tahoma" size=5 color="000000"> <b>To continue click <a href="class.asp">here. <b></a></font>&nbsp;</p><br><br>
                    </td>
                </tr>
            </table>
            </center></div></td>
        </tr>

    </table>          


</div align="center"></center>
                <div align="center"><center><table border="1"
                cellpadding="3" cellspacing="0" width="95%"
                bgcolor="#0000FF" bordercolor="#000000">
                    <tr>
                        <td align="center"><a href="#top"><font color="#FFFFFF"
                        size="4"><strong>RETURN TO TOP OF PAGE</strong></font></a></td>
                    </tr>
                </table>
                </center></div></td>
            </tr>
        </table>
        </td>
        <td>&nbsp;</td>
    </tr>
    <tr>
        <td><img src="Images/bottompage.GIF" width="575"
        height="52"></td>
        <td
        background="Images/botback.gif">&nbsp;</td>
    </tr>
</table>
</body>
</html>