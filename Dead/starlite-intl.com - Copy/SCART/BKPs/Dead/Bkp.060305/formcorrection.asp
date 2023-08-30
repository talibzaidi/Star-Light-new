<%@ LANGUAGE = VBScript %>
<%response.buffer=true%>
<% ar = Request("Area") %>
<% Area = Request("Area") %>
<% sar = ReQuest("sar") %>
<% SID = ReQuest("SID") %>
<% spec = 6 %>


<html>


<head>
<meta name="keywords" content="gps,navigation electronics,cb radios,frs,gmrs,radio scanners,2-way radios,hand tools">
<meta name="description" content="Online store for GPS, Navigation electronics, cb radios, frs, gmrs, antennas, car audio, hand tools.  Shopping on a secure SSL line. Accepting Visa, Mastercard, American Express cards.">
<!-- <meta name="Author" content=" IAC @ www.ontbiz.com/iac - Designed and Programmed by Anyperson.Com www.anyperson.com/tds"> -->
<title>Starlite International LLC - Online Store</title>

<script language="Javascript">
<!--
	once = new MakeArray(6)
	over = new MakeArray(6)
	under = new MakeArray(6)
	standard = new MakeArray(1)
	once[0].src = "../Images/question1.gif"
	once[1].src = "../Images/scart1.gif"
	once[2].src = "../Images/home1.gif"
	once[3].src = "../Images/new1.gif"
    once[4].src = "../Images/cat1.gif"
	once[5].src = "../Images/ex1.gif"    
	over[0].src = "../Images/question2.gif"
	over[1].src = "../Images/scart2.gif"
	over[2].src = "../Images/home2.gif"
	over[3].src = "../Images/new2.gif"
	over[4].src = "../Images/cat2.gif"
	over[5].src = "../Images/ex2.gif"
	under[0].src = "../Images/helpnav.gif"
	under[1].src = "../Images/shoppingcartnav.gif"
	under[2].src = "../Images/homenav.gif"
	under[3].src = "../Images/newproductsnav.gif"
	under[4].src = "../Images/onlinecataloguenav.gif"
	under[5].src = "../Images/specialsnav.gif"
	standard[0].src = "../Images/emptynav.jpg"
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



<body bgcolor="white" link="black" vlink="black" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">


<!--#include virtual="Misc/Header.INC"-->


	
<br><br><br>
<center>      
<font face="tahoma" size="4"><b>You Must Fill Out All Form Fields To Continue!</b></font>
</center>

<br><br>
<form action="sendmail.asp">

    <table border="0" width="350" align='center'>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Name:</strong></font></td>
            <td><strong><input type="text" size="25" name="Name" value="<%=Request("Name")%>"></strong></td>
        </tr>
         <tr>
            <td><font size="2" face="Tahoma"><strong>E-Mail Address:</strong></font></td>
            <td><input type="text" size="25" name="Email" value="<%=Request("Email")%>"></td>
        </tr>
	 <tr>
            <td><font size="2" face="Tahoma"><strong>Phone Number:</strong></font></td>
            <td><input type="text" size="25" name="Phone" value="<%=Request("Phone")%>"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Street
            Address:</strong></font></td>
            <td><input type="text" size="25" name="Address" value="<%=Request("Address")%>"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>City:</strong></font></td>
            <td><input type="text" size="25" name="City" value="<%=Request("City")%>"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>State/Province:</strong></font></td>
            <td><input type="text" size="25" name="StateProv" value="<%=Request("StateProv")%>"></td>
        </tr>
   <tr>
            <td><font size="2" face="Tahoma"><strong>Postal Code:</strong></font></td>
            <td><input type="text" size="25" name="Postal" value="<%=Request("Postal")%>"></td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Country:</strong></font></td>
            <td>
            
<select name="Country" id="Country">
<option selected value="<%=Request("Country")%>"><%=Request("Country")%></option>
<option value="Afghanistan">Afghanistan</OPTION>
<option value="Albania">Albania</OPTION>
<option value="Algeria">Algeria</OPTION>
<option value="American Somoa">American Samoa</OPTION>
<option value="Andorra">Andorra</OPTION>
<option value="Angola">Angola</OPTION>
<option value="Anguilla">Anguilla</OPTION>
<option value="Antarctica">Antarctica</OPTION>
<option value="Aantigua and Barbuda">Antigua and Barbuda</OPTION>
<option value="Argentina">Argentina</OPTION>
<option value="Armenia">Armenia</OPTION>
<option value="Aruba">Aruba</OPTION>
<option value="Australia">Australia</OPTION>
<option value="Austria">Austria</OPTION>
<option value="Azerbbaijan">Azerbaijan</OPTION>
<option value="Bahamas">Bahamas</OPTION>
<option value="Bahrain">Bahrain</OPTION>
<option value="Bangedesh">Bangladesh</OPTION>
<option value="Barbados">Barbados</OPTION>
<option value="Belarus">Belarus</OPTION>
<option value="Belguim">Belgium</OPTION>
<option value="Belize">Belize</OPTION>
<option value="Benin">Benin</OPTION>
<option value="Bermuda">Bermuda</OPTION>
<option value="Bhutan">Bhutan</OPTION>
<option value="Bolivia">Bolivia</OPTION>
<option value="Bosnia and Herzegovina">Bosnia and Herzegovina</OPTION>
<option value="Botswana">Botswana</OPTION>
<option value="Bouvet Island">Bouvet Island</OPTION>
<option value="Brazil">Brazil</OPTION>
<option value="British Indian Ocean Territory">British Indian Ocean Territory</OPTION>
<option value="Brunei">Brunei</OPTION>
<option value="Bulgaria">Bulgaria</OPTION>
<option value="Burkina Faso">Burkina Faso</OPTION>
<option value="Burundi">Burundic</OPTION>
<option value="Cambodia">Cambodia</OPTION>
<option value="Cameroon">Cameroon</OPTION>
<option value="Canada">Canada</OPTION>
<option value="Cape Verde">Cape Verde</OPTION>
<option value="Cayman Islands">Cayman Islands</OPTION>
<option value="Central African Republic">Central African Republic</OPTION>
<option value="Chad">Chad</OPTION>
<option value="Chile">Chile</OPTION>
<option value="China">China</OPTION>
<option value="Christmas Island">Christmas Island</OPTION>
<option value="Cocos Islands">Cocos Islands</OPTION>
<option value="Columbia">Colombia</OPTION>
<option value="Comoros">Comoros</OPTION>
<option value="Congo">Congo</OPTION>
<option value="Cook Islands">Cook Islands</OPTION>
<option value="Costa Rica">Costa Rica</OPTION>
<option value="Cote d’Iviore">Côte d'Ivoire</OPTION>
<option value="Croatia">Croatia</OPTION>
<option value="Cuba">Cuba</OPTION>
<option value="Cyrus">Cyprus</OPTION>
<option value="Czech Republic">Czech Republic</OPTION>
<option value="Congo (DRC)">Congo (DRC) </OPTION>
<option value="Denmark">Denmark</OPTION>
<option value="Djibouti">Djibouti</OPTION>
<option value="Dominica">Dominica</OPTION>
<option value="Dominion Republic">Dominican Republic</OPTION>
<option value="East Timor">East Timor</OPTION>
<option value="Ecuador">Ecuador</OPTION>
<option value="Eqypt">Egypt</OPTION>
<option value="El Salvador">El Salvador</OPTION>
<option value="Equatorial Guinea">Equatorial Guinea</OPTION>
<option value="Erutrea">Eritrea</OPTION>
<option value="Estonia">Estonia</OPTION>
<option value="Ethiopia">Ethiopia</OPTION>
<option value="Falkland Islands">Falkland Islands</OPTION>
<option value="Faroe Islands">Faroe Islands</OPTION>
<option value="Fiji Islands">Fiji Islands</OPTION>
<option value="Finland">Finland</OPTION>
<option value="France">France</OPTION>
<option value="French Guiana">French Guiana</OPTION>
<option value="French Polynesia">French Polynesia</OPTION>
<option value="French Southern and Antarctic Lands ">French Southern and Antarctic Lands</OPTION>
<option value="Gabon">Gabon</OPTION>
<option value="Gambia">Gambia</OPTION>
<option value="Georgia">Georgia</OPTION>
<option value="Germany">Germany</OPTION>
<option value="Ghana">Ghana</OPTION>
<option value="Gibraltar">Gibraltar</OPTION>
<option value="Greece">Greece</OPTION>
<option value="Greenland">Greenland</OPTION>
<option value="Grenada">Grenada</OPTION>
<option value="Guadeloupe">Guadeloupe</OPTION>
<option value="Gua,">Guam</OPTION>
<option value="Guatemala">Guatemala</OPTION>
<option value="Guinea">Guinea</OPTION>
<option value="GuineaBissau">GuineaBissau</OPTION>
<option value="Guyana">Guyana</OPTION>
<option value="Haiti">Haiti</OPTION>
<option value="Heard Island">Heard Island</OPTION>
<option value="Honduras">Honduras</OPTION>
<option value="Hong Kong">Hong Kong</OPTION>
<option value="Hungary">Hungary</OPTION>
<option value="Iceland">Iceland</OPTION>
<option value="India">India</OPTION>
<option value="Indonesia">Indonesia</OPTION>
<option value="Iran">Iran</OPTION>
<option value="Iraq">Iraq</OPTION>
<option value="Ireland">Ireland</OPTION>
<option value="Israel">Israel</OPTION>
<option value="Italy">Italy</OPTION>
<option value="Jamaica">Jamaica</OPTION>
<option value="Japan">Japan</OPTION>
<option value="Jordan">Jordan</OPTION>
<option value="Kazakhstan">Kazakhstan</OPTION>
<option value="Kenya">Kenya</OPTION>
<option value="Kiribati">Kiribati</OPTION>
<option value="Korea">Korea</OPTION>
<option value="Kuwait">Kuwait</OPTION>
<option value="Kyrgyzstan">Kyrgyzstan</OPTION>
<option value="Laos">Laos</OPTION>
<option value="Latvia">Latvia</OPTION>
<option value="Lebanon">Lebanon</OPTION>
<option value="Lesotho">Lesotho</OPTION>
<option value="Liberia">Liberia</OPTION>
<option value="Libya">Libya</OPTION>
<option value="Liechtenstein">Liechtenstein</OPTION>
<option value="Lthuania">Lithuania</OPTION>
<option value="Luxembourg">Luxembourg</OPTION>
<option value="Macau">Macau</OPTION>
<option value="Macedonia">Macedonia</OPTION>
<option value="Madagascar">Madagascar</OPTION>
<option value="Malawi">Malawi</OPTION>
<option value="Malaysia">Malaysia</OPTION>
<option value="Maldives">Maldives</OPTION>
<option value="Mali">Mali</OPTION>
<option value="Malta">Malta</OPTION>
<option value="Marshall Islands">Marshall Islands</OPTION>
<option value="Martinique">Martinique</OPTION>
<option value="Mauritania">Mauritania</OPTION>
<option value="Mauritius">Mauritius</OPTION>
<option value="Mayotte">Mayotte</OPTION>
<option value="Mexico">Mexico</OPTION>
<option value="Micronesia">Micronesia</OPTION>
<option value="Moldova">Moldova</OPTION>
<option value="Monaco">Monaco</OPTION>
<option value="Mongolia">Mongolia</OPTION>
<option value="Montserrat">Montserrat</OPTION>
<option value="Morocco">Morocco</OPTION>
<option value="Mozambique">Mozambique</OPTION>
<option value="Myanmar">Myanmar</OPTION>
<option value="Namibia">Namibia</OPTION>
<option value="Nauru">Nauru</OPTION>
<option value="Nepal">Nepal</OPTION>
<option value="Netherlands">Netherlands</OPTION>
<option value="Netherlands Antilles">Netherlands Antilles</OPTION>
<option value="New Caledonia">New Caledonia</OPTION>
<option value="New Zealand">New Zealand</OPTION>
<option value="Nicaragua">Nicaragua</OPTION>
<option value="Niger">Niger</OPTION>
<option value="Nigeria">Nigeria</OPTION>
<option value="Niue">Niue</OPTION>
<option value="Norfolk Island">Norfolk Island</OPTION>
<option value="North Korea">North Korea</OPTION>
<option value="Northern Mariana Islands">Northern Mariana Islands</OPTION>
<option value="Norway">Norway</OPTION>
<option value="Oman">Oman</OPTION>
<option value="Pakistan">Pakistan</OPTION>
<option value="Palau">Palau</OPTION>
<option value="Panama">Panama</OPTION>
<option value="Papua New Guinea">Papua New Guinea</OPTION>
<option value="Paraguay">Paraguay</OPTION>
<option value="Peru">Peru</OPTION>
<option value="Philippines">Philippines</OPTION>
<option value="Pitcairn Islands">Pitcairn Islands</OPTION>
<option value="Poland">Poland</OPTION>
<option value="Portugual">Portugal</OPTION>
<option value="Puerto Rico">Puerto Rico</OPTION>
<option value="Qatar">Qatar</OPTION>
<option value="Reunion">Reunion</OPTION>
<option value="Romania">Romania</OPTION>
<option value="Russia">Russia</OPTION>
<option value="Rwanda">Rwanda</OPTION>
<option value="St. Kits and Nevis">St. Kitts and Nevis</OPTION>
<option value="St. Lucia">St. Lucia</OPTION>
<option value="Samoa">Samoa</OPTION>
<option value="San Marino">San Marino</OPTION>
<option value="Saudi Arabia">Saudi Arabia</OPTION>
<option value="Senegal">Senegal</OPTION>
<option value="Seychelles">Seychelles</OPTION>
<option value="Sierra Leone">Sierra Leone</OPTION>
<option value="Sinapore">Singapore</OPTION>
<option value="Slovakia">Slovakia</OPTION>
<option value="Slovenia">Slovenia</OPTION>
<option value="Solomon Islands">Solomon Islands</OPTION>
<option value="Somalia">Somalia</OPTION>
<option value="South Africa">South Africa</OPTION>
<option value="South Georgia">South Georgia </OPTION>
<option value="Spain">Spain</OPTION>
<option value="Sri Lanka">Sri Lanka</OPTION>
<option value="St. Helena">St. Helena</OPTION>
<option value="Sudan">Sudan</OPTION>
<option value="Suriname">Suriname</OPTION>
<option value="Swaziland">Swaziland</OPTION>
<option value="Sweden">Sweden</OPTION>
<option value="Switzerland">Switzerland</OPTION>
<option value="Syria">Syria</OPTION>
<option value="Taiwan">Taiwan</OPTION>
<option value="Tajikistan">Tajikistan</OPTION>
<option value="Tanzania">Tanzania</OPTION>
<option value="Thailand">Thailand</OPTION>
<option value="Togo">Togo</OPTION>
<option value="Tokelau">Tokelau</OPTION>
<option value="Tonga">Tonga</OPTION>
<option value="Trinidad and Tobago">Trinidad and Tobago</OPTION>
<option value="Tunisia">Tunisia</OPTION>
<option value="Turkey">Turkey</OPTION>
<option value="Turkmenistan">Turkmenistan</OPTION>
<option value="Tuvalu">Tuvalu</OPTION>
<option value="Uganda">Uganda</OPTION>
<option value="Ukraine">Ukraine</OPTION>
<option value="United Kingdom">United Kingdom</OPTION>
<option value="United States">United States</OPTION>
<option value="Uruguay">Uruguay</OPTION>
<option value="Uzbekistan">Uzbekistan</OPTION>
<option value="Vanuartu">Vanuatu</OPTION>
<option value="Vatican City">Vatican City</OPTION>
<option value="Venezuela">Venezuela</OPTION>
<option value="Viet Name">Viet Nam</OPTION>
<option value="Virgin Islands">Virgin Islands</OPTION>
<option value="Yemen">Yemen</OPTION>
<option value="Yugoslavia">Yugoslavia</OPTION>
<option value="Zambia">Zambia</OPTION>
<option value="Zimbabwe">Zimbabwe</OPTION>
</select>
            
            
            </td>
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
            <td><font size="2" face="Tahoma"><strong>Payment By:</strong></font></td>
            <td>
            <select name="Payment" size="1">
                <option selected value="<%=Request("Payment")%>"><%=Request("Payment")%></option>
                <option value="Visa">Visa</option>
                <option value="Master Card">Master Card</option>
                <option value="Discover">Discover</option>
				<option value="American Express">American Express</option>
                <option value="Cheque">Personal Cheque</option>
                <option value="Money Order">Money Order</option>
                <option value="Pay Pal">Pay Pal</option>
            </select>
            </td>
        </tr>
       
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        
        <tr>
            <td><font size="2" face="Tahoma"><strong>State / Provincial Tax:</strong></font></td>
            <td><input type="text" size="3" name="Taxx1" value="0">&nbsp;%</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong>Federal Tax:</strong></font></td>
            <td><input type="text" size="3" name="Taxx2" value="0">&nbsp;%</td>
        </tr>
        <tr>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
            <td><font size="2" face="Tahoma"><strong></strong></font>&nbsp;</td>
        </tr>
        <tr>
            <td align="center" colspan="2">
            <input type="submit" name="B3" value="Place My Order">
            </td>
        </tr>
    </table>

</form>


</body>


</html>
