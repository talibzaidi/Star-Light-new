
<!-- #INCLUDE file="RATES.INC" -->
<!-- #INCLUDE file="Scart.inc.SRs.asp"   4/22/10: This is the file of all Subs and Functions that used to be in the present file, but that I factored out on 4/22/10. -->


<% session("initq") = 1 %>
<%

'on error resume next

sAction 	= CStr(Request.QueryString("action"))
iItemID 	= (Request.QueryString("item"))
iItemCount 	= CInt(Request.QueryString("count"))
' 4/22/10: I don't know why, but on 4/22/10 I found I needed to add the following, or for some crazy reason it would always use SID = 275. 
' But this had not been needed for years before. Weird!
SID 		= Request.QueryString("SID")


If (sar = "New%20Products") OR (sar = "New Products") OR (sar = "New!") Then 
  	Set RS = CreateObject("ADODB.Recordset")
  	RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE  NewProduct = yes ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4    
   	'ar = "New Products"  
   	  
Elseif sar ="Specials" Then 
	response.redirect "scartstart.asp"
	
Elseif sar = "Manufa" Then 
    Set RS = CreateObject("ADODB.Recordset")
  	RS.Open "SELECT *,  Rates.ExchangeRate1 AS Exch,  Rates.Freight AS Freight FROM Product, Rates WHERE  Manufa Like '" & ReQuest("Manufa") & "' ORDER BY MSL ASC", "DSN=STAREC1" , 1, 4    
    ar="Search"
    
Else 
	'Response.Write "Got to here."
   	Set RS = CreateObject("ADODB.Recordset")
   	' 4/22/10: I don't know why, but on 4/22/10 I found I needed to change "+ sid +" in the next line to "+ CStr(sid) +", or I would
   	' get an error message about a type mismatch. But this had not been needed for years before. Weird!
  	RS.Open "SELECT *, Rates.ExchangeRate1 AS Exch, Rates.Freight AS Freight FROM Product, Rates WHERE SID = " + CStr(sid) + " ORDER BY MSL ASC ", "DSN=STAREC1" , 1, 4
End If

    countChoc = RS.RecordCount
    set RXS = CreateObject("ADODB.Recordset")
  	sar = Replace( sar, " ", "%20") 
%>




<% ' ***** Begin the infamous runtime script *****
' Declare our Vars
Dim dictCart ' as dictionary
Dim sAction ' as string
Dim iItemID ' as integer
Dim iItemCount ' as integer

' Get a reference to the cart if it exists otherwise create it
If IsObject(Session("cart")) Then
    Set dictCart = Session("cart")
Else
    ' We use a dictionary so we can name our keys to correspond to our
    ' item numbers and then use their value to hold the quantity.  An
    ' array would also work, but would be a little more complex and
    ' probably not as easy for readers to follow.
    Set dictCart = Server.CreateObject("Scripting.Dictionary")
End If

' Get all the parameters passed to the script

%>




<center>
<table BORDER="0" CELLSPACING="0" CELLPADDING="0">
<tr><td>
<% ' Response.Write "sAction = " & sAction & "<br>" %>
<% ar = Replace( ar, " ", "%20") %>
<% sar = Replace( sar, " ", "%20") %>
<% Area = Replace( Area, " ", "%20") %>
<% sid = Replace( sid, " ", "%20") %>
<% '************************QUANTITY AREA **** QTY AREA *******************************
' Select action based on user input
Select Case sAction

    Case "add"
        AddItemToCart iItemID, iItemCount
        ShowItemsInCart
        %>
        </td></tr>
        
        <tr><td ALIGN="right">
<!--        <a HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><img SRC="https://www.starlite-intl.com/images/shop_checkout.gif" BORDER="0" ALT="Checkout" WIDTH="46" HEIGHT="46"></a><br><font face="tahoma" size="1"><b>CHECKOUT</b></font><br><br><br>
-->
        <%
        
    Case "del"
        RemoveItemFromCart iItemID, iItemCount
        ShowItemsInCart
        %>
        <%
        
  	Case "qty"

        NewItemsCart iItemID, CDbl(Request("Qty"))
        'AddItemToCart iItemID, CDbl(Request("Qty"))
        if CDbl(Request("Qty")) = 0 then
            RemoveItemFromCart iItemID, CDbl(Request("Qty"))
        end if
        ShowItemsInCart
        %>
        </td></tr>
        <tr><td ALIGN="right">
<!--        <a HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><img SRC="https://www.starlite-intl.com/images/shop_checkout.gif" BORDER="0" ALT="Checkout" WIDTH="46" HEIGHT="46"></a><br><font face="tahoma" size="1"><b>CHECKOUT</b></font><br><br><br>
-->
        <%
        
  	Case "viewcart"
        ShowItemsInCart
        %>
        </td></tr>
        <tr><td ALIGN="right">


<!--        <a HREF="./scart.asp?action=checkout&amp;sid=<%=sid%>&amp;Area=<%=Area%>&amp;sar=<%=sar%>"><img SRC="https://www.starlite-intl.com/images/shop_checkout.gif" BORDER="0" ALT="Checkout" WIDTH="46" HEIGHT="46"></a><br><font face="tahoma" size="1"><b>CHECKOUT</b></font><br><br><br>
-->
        <%
        
    Case "checkout"
        ' PlaceOrder  [BN, 9/9/05]: I commeneted this out. It doesn't seem to belong here. Don't know why it was here.
        %>
        </td></tr>
        <tr><td ALIGN="left">
        <br>
        
        <center>
        <font face=Tahoma>
		Please enter the following information to complete your order. <script type="text/javascript" src="https://seal.XRamp.com/seal.asp?type=H"></script>
		<br>
        Your order will be shipped when your payement clears. 
		<br>
		In the event that there is a problem in processing your order, you will be notified via phone or e-mail.
		<br>
		For further inquiries call: 1-800-387-8535 or <a href="mailto:staff@starlite-intl.com">email us</a>.
        </font>
        </center>

<script Language="JavaScript">
<!--
function Validator1(theForm)
{

if (theForm.FName.value == "")
  {
    alert("Please enter your First Name");
    theForm.FName.focus();
    return (false);
  }
if (theForm.LName.value == "")
  {
    alert("Please enter your Last Name");
    theForm.LName.focus();
    return (false);
  }
if (theForm.Email.value == "")
  {
    alert("Please enter your Email Address");
    theForm.Email.focus();
    return (false);
  }
if (theForm.Phone.value == "")
  {
    alert("Please enter your Phone Number");
    theForm.Phone.focus();
    return (false);
  }
if (theForm.Address.value == "")
  {
    alert("Please enter your Street Address");
    theForm.Address.focus();
    return (false);
  }
if (theForm.City.value == "")
  {
    alert("Please enter your City");
    theForm.City.focus();
    return (false);
  }
if (theForm.City.value == "")
  {
    alert("Please enter your City");
    theForm.City.focus();
    return (false);
  }
if (theForm.Country.selectedIndex == 0 || theForm.Country.selectedIndex == 1)
  {
    alert("Please select a Country");
    theForm.Country.focus();
    return (false);
  }  
 
if ((theForm.Country.value == "United States" & theForm.StateProv.value == "") || (theForm.Country.value == "Canada" & theForm.StateProv.value == ""))
  {
    alert("Please enter a State");
    theForm.StateProv.focus();
    return (false);
  }  
  
if ((theForm.Country.value == "United States" & theForm.Postal.value == "") || (theForm.Country.value == "Canada" & theForm.Postal.value == ""))
  {
    alert("Please enter a Postal Code");
    theForm.StateProv.focus();
    return (false);
  }  
 
if (theForm.Payment.selectedIndex == 0 || theForm.Payment.selectedIndex == 1)
  {
    alert("Please select a Payment Method");
    theForm.Payment.focus();
    return (false);
  }  
return (true);
}

//-->
</script>

                
        <form action="EmailBuild.asp" method="post" language='JavaScript' onsubmit='return Validator1(this)' name='CheckoutDataForm'>

		<div align="center"><center>
		
		<table border="0" align=center cellpadding=4>
    
        <tr>
            <td width=150 align=right><font size="2" face="Tahoma"><strong>First Name:</strong></font></td>
            <td>
            <input type="text" size="25" name="FName"></td>
        </tr>
        
        <tr>
            <td width=150 align=right><font size="2" face="Tahoma"><strong>Last Name:</strong></font></td>
            <td>
            <input type="text" size="25" name="LName"></td>
        </tr>
        
		<tr>
            <td align=right><font size="2" face="Tahoma"><strong>Email Address:</strong></font></td>
            <td>
            <input type="text" size="25" name="Email"></td>
        </tr>
        
         <tr>
            <td valign="top"></td>
            <td valign="top">
            <input type="checkbox" name="OptOutOfEmailing">
            <font size="2" face="Tahoma">
            <strong>I do not want to receive information on specials and new products.</strong>
            </font>
            </td>
        </tr>
        
		<tr>
            <td align=right><font size="2" face="Tahoma"><strong>Phone Number:</strong></font></td>
            <td><input type="text" size="25" name="Phone" value></td>
        </tr>
        
        <tr>
            <td align=right><font size="2" face="Tahoma"><strong>Street
            Address:</strong></font></td>
            <td>
            <input type="text" size="25" name="Address"></td>
        </tr>
        
        <tr>
            <td align=right><font size="2" face="Tahoma"><strong>City:</strong></font></td>
            <td>
            <input type="text" size="25" name="City"></td>
        </tr>
        
        <tr>
            <td align=right><font size="2" face="Tahoma"><strong>State/Province:</strong></font></td>
            <td><input type="text" size="25" name="StateProv"></td>
        </tr>
        
         <tr>
            <td valign="top"></td>
            <td valign="top">
            <input type="checkbox" name="BigShip">
            <font size="2" face="Tahoma">
            <strong>I live in Alaska, Hawaii, Nunavut, or Northwest Territories.</strong><br>
            Your Shipping and Handling charge will be recalculated at 3 times the displayed amount.
            </font><br>
            </td>
        </tr>
        
		<tr>
            <td align=right><font size="2" face="Tahoma"><strong>Postal Code:</strong></font></td>
            <td><input type="text" size="25" name="Postal"></td>
        </tr>
        <tr>
            <td align=right><font size="2" face="Tahoma"><strong>Country:</strong></font></td>
            <td>
<select name="Country" id="Country">
<option value="SELECT COUNTRY" SELECTED>Select a Country ...</option>
<option></option>
<option value="Afghanistan">Afghanistan</option>
<option value="Albania">Albania</option>
<option value="Algeria">Algeria</option>
<option value="American Somoa">American Samoa</OPTION>
<option value="Andorra">Andorra</option>
<option value="Angola">Angola</option>
<option value="Anguilla">Anguilla</option>
<option value="Antarctica">Antarctica</option>
<option value="Aantigua and Barbuda">Antigua and Barbuda</option>
<option value="Argentina">Argentina</option>
<option value="Armenia">Armenia</option>
<option value="Aruba">Aruba</option>
<option value="Australia">Australia</option>
<option value="Austria">Austria</option>
<option value="Azerbbaijan">Azerbaijan</option>
<option value="Bahamas">Bahamas</option>
<option value="Bahrain">Bahrain</option>
<option value="Bangedesh">Bangladesh</option>
<option value="Barbados">Barbados</option>
<option value="Belarus">Belarus</option>
<option value="Belguim">Belgium</option>
<option value="Belize">Belize</option>
<option value="Benin">Benin</option>
<option value="Bermuda">Bermuda</option>
<option value="Bhutan">Bhutan</option>
<option value="Bolivia">Bolivia</option>
<option value="Bosnia and Herzegovina">Bosnia and Herzegovina</OPTION>
<option value="Botswana">Botswana</option>
<option value="Bouvet Island">Bouvet Island</option>
<option value="Brazil">Brazil</option>
<option value="British Indian Ocean Territory">British Indian Ocean Territory</option>
<option value="Brunei">Brunei</option>
<option value="Bulgaria">Bulgaria</option>
<option value="Burkina Faso">Burkina Faso</option>
<option value="Burundi">Burundic</option>
<option value="Cambodia">Cambodia</option>
<option value="Cameroon">Cameroon</option>
<option value="Canada">Canada</option>
<option value="Cape Verde">Cape Verde</option>
<option value="Cayman Islands">Cayman Islands</OPTION>
<option value="Central African Republic">Central African Republic</option>
<option value="Chad">Chad</option>
<option value="Chile">Chile</option>
<option value="China">China</option>
<option value="Christmas Island">Christmas Island</option>
<option value="Cocos Islands">Cocos Islands</option>
<option value="Columbia">Colombia</option>
<option value="Comoros">Comoros</option>
<option value="Congo">Congo</option>
<option value="Cook Islands">Cook Islands</option>
<option value="Costa Rica">Costa Rica</option>
<option value="Cote d&#146;Iviore">Côte d'Ivoire</option>
<option value="Croatia">Croatia</option>
<option value="Cuba">Cuba</option>
<option value="Cyrus">Cyprus</option>
<option value="Czech Republic">Czech Republic</option>
<option value="Congo (DRC)">Congo (DRC) </option>
<option value="Denmark">Denmark</option>
<option value="Djibouti">Djibouti</option>
<option value="Dominica">Dominica</option>
<option value="Dominion Republic">Dominican Republic</option>
<option value="East Timor">East Timor</option>
<option value="Ecuador">Ecuador</option>
<option value="Eqypt">Egypt</option>
<option value="El Salvador">El Salvador</option>
<option value="Equatorial Guinea">Equatorial Guinea</option>
<option value="Erutrea">Eritrea</option>
<option value="Estonia">Estonia</option>
<option value="Ethiopia">Ethiopia</option>
<option value="Falkland Islands">Falkland Islands</option>
<option value="Faroe Islands">Faroe Islands</option>
<option value="Fiji Islands">Fiji Islands</option>
<option value="Finland">Finland</option>
<option value="France">France</option>
<option value="French Guiana">French Guiana</option>
<option value="French Polynesia">French Polynesia</option>
<option value=" French Southern and Antarctic Lands ">French Southern and Antarctic Lands</option>
<option value="Gabon">Gabon</option>
<option value="Gambia">Gambia</option>
<option value="Georgia">Georgia</option>
<option value="Germany">Germany</option>
<option value="Ghana">Ghana</option>
<option value="Gibraltar">Gibraltar</option>
<option value="Greece">Greece</option>
<option value="Greenland">Greenland</option>
<option value="Grenada">Grenada</option>
<option value="Guadeloupe">Guadeloupe</option>
<option value="Gua,">Guam</option>
<option value="Guatemala">Guatemala</option>
<option value="Guinea">Guinea</option>
<option value="GuineaBissau">GuineaBissau</option>
<option value="Guyana">Guyana</option>
<option value="Haiti">Haiti</option>
<option value="Heard Island">Heard Island</option>
<option value="Honduras">Honduras</option>
<option value="Hong Kong">Hong Kong</option>
<option value="Hungary">Hungary</option>
<option value="Iceland">Iceland</option>
<option value="India">India</option>
<option value="Indonesia">Indonesia</option>
<option value="Iran">Iran</option>
<option value="Iraq">Iraq</option>
<option value="Ireland">Ireland</option>
<option value="Israel">Israel</option>
<option value="Italy">Italy</option>
<option value="Jamaica">Jamaica</option>
<option value="Japan">Japan</option>
<option value="Jordan">Jordan</option>
<option value="Kazakhstan">Kazakhstan</option>
<option value="Kenya">Kenya</option>
<option value="Kiribati">Kiribati</option>
<option value="Korea">Korea</option>
<option value="Kuwait">Kuwait</option>
<option value="Kyrgyzstan">Kyrgyzstan</option>
<option value="Laos">Laos</option>
<option value="Latvia">Latvia</option>
<option value="Lebanon">Lebanon</option>
<option value="Lesotho">Lesotho</option>
<option value="Liberia">Liberia</option>
<option value="Libya">Libya</option>
<option value="Liechtenstein">Liechtenstein</option>
<option value="Lthuania">Lithuania</option>
<option value="Luxembourg">Luxembourg</option>
<option value="Macau">Macau</option>
<option value="Macedonia">Macedonia</option>
<option value="Madagascar">Madagascar</option>
<option value="Malawi">Malawi</option>
<option value="Malaysia">Malaysia</option>
<option value="Maldives">Maldives</option>
<option value="Mali">Mali</option>
<option value="Malta">Malta</option>
<option value="Marshall Islands">Marshall Islands</option>
<option value="Martinique">Martinique</option>
<option value="Mauritania">Mauritania</option>
<option value="Mauritius">Mauritius</option>
<option value="Mayotte">Mayotte</option>
<option value="Mexico">Mexico</option>
<option value="Micronesia">Micronesia</option>
<option value="Moldova">Moldova</option>
<option value="Monaco">Monaco</option>
<option value="Mongolia">Mongolia</option>
<option value="Montserrat">Montserrat</option>
<option value="Morocco">Morocco</option>
<option value="Mozambique">Mozambique</option>
<option value="Myanmar">Myanmar</option>
<option value="Namibia">Namibia</option>
<option value="Nauru">Nauru</option>
<option value="Nepal">Nepal</option>
<option value="Netherlands">Netherlands</option>
<option value="Netherlands Antilles">Netherlands Antilles</option>
<option value="New Caledonia">New Caledonia</option>
<option value="New Zealand">New Zealand</option>
<option value="Nicaragua">Nicaragua</option>
<option value="Niger">Niger</option>
<option value="Nigeria">Nigeria</option>
<option value="Niue">Niue</option>
<option value="Norfolk Island">Norfolk Island</option>
<option value="North Korea">North Korea</option>
<option value="Northern Mariana Islands">Northern Mariana Islands</option>
<option value="Norway">Norway</option>
<option value="Oman">Oman</option>
<option value="Pakistan">Pakistan</option>
<option value="Palau">Palau</option>
<option value="Panama">Panama</option>
<option value="Papua New Guinea">Papua New Guinea</option>
<option value="Paraguay">Paraguay</option>
<option value="Peru">Peru</option>
<option value="Philippines">Philippines</option>
<option value="Pitcairn Islands">Pitcairn Islands</option>
<option value="Poland">Poland</option>
<option value="Portugual">Portugal</option>
<option value="Puerto Rico">Puerto Rico</option>
<option value="Qatar">Qatar</option>
<option value="Reunion">Reunion</option>
<option value="Romania">Romania</option>
<option value="Russia">Russia</option>
<option value="Rwanda">Rwanda</option>
<option value="St. Kits and Nevis">St. Kitts and Nevis</option>
<option value="St. Lucia">St. Lucia</option>
<option value="Samoa">Samoa</option>
<option value="San Marino">San Marino</option>
<option value="Saudi Arabia">Saudi Arabia</option>
<option value="Senegal">Senegal</option>
<option value="Seychelles">Seychelles</option>
<option value="Sierra Leone">Sierra Leone</option>
<option value="Sinapore">Singapore</option>
<option value="Slovakia">Slovakia</option>
<option value="Slovenia">Slovenia</option>
<option value="Solomon Islands">Solomon Islands</option>
<option value="Somalia">Somalia</option>
<option value="South Africa">South Africa</option>
<option value="South Georgia">South Georgia </option>
<option value="Spain">Spain</option>
<option value="Sri Lanka">Sri Lanka</option>
<option value="St. Helena">St. Helena</option>
<option value="Sudan">Sudan</option>
<option value="Suriname">Suriname</option>
<option value="Swaziland">Swaziland</option>
<option value="Sweden">Sweden</option>
<option value="Switzerland">Switzerland</option>
<option value="Syria">Syria</option>
<option value="Taiwan">Taiwan</option>
<option value="Tajikistan">Tajikistan</option>
<option value="Tanzania">Tanzania</option>
<option value="Thailand">Thailand</option>
<option value="Togo">Togo</option>
<option value="Tokelau">Tokelau</option>
<option value="Tonga">Tonga</option>
<option value="Trinidad and Tobago">Trinidad and Tobago</option>
<option value="Tunisia">Tunisia</option>
<option value="Turkey">Turkey</option>
<option value="Turkmenistan">Turkmenistan</option>
<option value="Tuvalu">Tuvalu</option>
<option value="Uganda">Uganda</option>
<option value="Ukraine">Ukraine</option>
<option value="United Kingdom">United Kingdom</option>
<option value="USA">United States</option>
<option value="Uruguay">Uruguay</option>
<option value="Uzbekistan">Uzbekistan</option>
<option value="Vanuartu">Vanuatu</option>
<option value="Vatican City">Vatican City</option>
<option value="Venezuela">Venezuela</option>
<option value="Viet Name">Viet Nam</option>
<option value="Virgin Islands">Virgin Islands</option>
<option value="Yemen">Yemen</option>
<option value="Yugoslavia">Yugoslavia</option>
<option value="Zambia">Zambia</option>
<option value="Zimbabwe">Zimbabwe</option>
</select></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        <tr>
            <td align=right><font size="2" face="Tahoma"><strong>Payment Method:</strong></font></td>
            <td>
            <select name="Payment" size="1">
                <option selected value="Choose">Select a Method ...</option>
                <option></option>
                <option value="Visa">Visa</option>
                <option value="Master Card">Master Card</option>
                <option value="Discover">Discover</option>
				<option value="American Express">American Express</option>
                <option value="Check">Personal Check, sent in advance (No C.O.D)</option>
                <option value="Money Order">Money Order, sent in advance (No C.O.D)</option>
                <option value="Pay Pal">Pay Pal</option>
            </select></td>
        </tr>
       
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>    
    <tr>
            <td colspan="2"><font size="2" face="Tahoma">Taxes will be applied to the total order charge including Shipping & Handling ...
			<br><U>Unless you are tax exempt</U> in Michigan add 6%, In Ontario add 8% for P.S.T. and 5% G.S.T., any other Canadian Province add only 5% for G.S.T., as applicable. International customers will be responsible for paying taxes/duties if any in their country. </font>&nbsp;</td>
            
        </tr>
        
        <tr>
            <td align=right><font size="2" face="Tahoma"><strong>State / Provincial Tax:</strong></font></td>
            <td><input type="text" size="3" name="Taxx1" value="0">&nbsp;%</td>
        </tr>
        <tr>
            <td align=right><font size="2" face="Tahoma"><strong>Federal Tax:</strong></font></td>
            <td><input type="text" size="3" name="Taxx2" value="0">&nbsp;%</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        <tr>
            <td align="center" colspan="2"><font size="2" face="Tahoma"><strong><input type="submit" name="B3" value="Place My Order"></strong></font></td>
        </tr>
		
    </table>
    </center></div>

        </form>
        
        

        <%
        
    Case Else ' Shop
        ShowFullCatalog
        %>
        </td></tr>
        <tr><td ALIGN="right">
        
        <%
End Select


' Return cart to Session for storage
Set Session("cart") = dictCart
%>
</td></tr>
</table>
</center>
