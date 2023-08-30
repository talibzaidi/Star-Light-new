<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">


<html>



<head>
<title>Suckerfish drop-down menu test 1</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="http://www.mitmazel.com/Misc/StyleSheet1.css" media=screen> 
<!-- suckerfishnav.css is no longer needed here, because it is now incorporated inside StyleSheet1.css ...
<link rel="stylesheet" type="text/css" href="suckerfishnav.css" media=screen>  
-->
</head>



<body bottomMargin="0" leftMargin="0" rightMargin="0" topMargin="0">

<!-- #INCLUDE VIRTUAL = "header.htm" -->

<% 
InArea = "Home"				' Parameter of NavBars1.inc file 
InSubArea = "About Us"		' Parameter of NavBars1.inc file 
%>

<!-- #INCLUDE VIRTUAL = "NavBar1.inc" -->



<!-- Start Drop-Down Menu Test ---------------------------------------------->

<!-- <script type="text/javascript" src="http://www.guardyoureyes.org/GUE/WP271/wp-content/plugins/multi-level-navigtion-plugin/scripts/superfish.js"></script> -->
<script type="text/javascript" src="http://www.mitmazel.com/mytests/navbars/dropdowns/suckerfish/superfish.js"></script>


<!--[if lte IE 6]>
<!-- <script type="text/javascript" src="http://www.guardyoureyes.org/GUE/WP271/wp-content/plugins/multi-level-navigtion-plugin/scripts/suckerfish_ie.js"></script> -->
<script type="text/javascript" src="http://www.mitmazel.com/mytests/navbars/dropdowns/suckerfish/suckerfish_ie.js"></script>
<![endif]-->


<br><br>

<!-- 
<table align=center border=0 cellpadding=20>
<tr>

<td align=center>
-->

<!--
<div id="pixopoint_menu_wrapper1">
	<div id="pixopoint_menu1">
-->


<ul class="sf-menu" id="suckerfishnav">
	<li class="current_page_item"><a href="http://www.mitmazel.com">Home</a>
	<ul>
		<li class="page_item page-item-334"><a href="#">Singles <font size=1>>></font></a>
		<ul>
			<li class="page_item page-item-408"><a href="#">How MM Works</a></li>
			<li class="page_item page-item-410"><a href="#">Success Stories</a></li>
			<li class="page_item page-item-414"><a href="#">Member Comments</a></li>
			<li class="page_item page-item-408"><a href="#">Become a Member <font size=1>>></font></a>
			<ul>
				<li class="page_item page-item-351"><a href="http://www.mitmazel.com/singles/memberJoinAssoc1.asp">Associate Member</a></li>
				<li class="page_item page-item-299"><a href="http://www.mitmazel.com/singles/memberJoinFull.asp?PrType=1MFT">Full Member</a></li>
			</ul>
			</li>
			<li class="page_item page-item-410"><a href="#">Find Your Bashert</a></li>
			<li class="page_item page-item-414"><a href="#">Find a Matchmaker</a></li>
			<li class="page_item page-item-414"><a href="#">Find a Sponsor</a></li>
			<li class="page_item page-item-414"><a href="#">Terms & Conditions</a></li>
			<li class="page_item page-item-414"><a href="#">Privacy Policy</a></li>
		</ul>
		</li>
		<li class="page_item page-item-336"><a href="http://www.mitmazel.com/relatives/relatives.asp">Relatives & Friends</a></li>
		<li class="page_item page-item-342"><a href="#">Matchmakers <font size=1>>></font></a>
		<ul>
			<li class="page_item page-item-408"><a href="#">Find a Matchmaker</a></li>
			<li class="page_item page-item-410"><a href="#">Become a Matchmaker</a></li>
		</ul>				
		</li>
		<li class="page_item page-item-340"><a href="#">Sponsors <font size=1>>></font></a>
		<ul>
			<li class="page_item page-item-408"><a href="#">Find a Sponsor</a></li>
			<li class="page_item page-item-410"><a href="#">Become a Sponsor</a></li>
		</ul>
		</li>
		<li class="page_item page-item-344"><a href="http://www.mitmazel.com/singles/membercomments.asp">Member Comments</a></li>
		<li class="page_item page-item-347 haschildren"><a href="#">About Us <font size=1>>></font></a>
		<ul>
			<li class="page_item page-item-408"><a href="#">Contact Us</a></li>
			<li class="page_item page-item-410"><a href="#">Privacy Policy</a></li>
			<li class="page_item page-item-414"><a href="#">Terms & Conditions</a></li>
		</ul>
		</li>
	</ul>
	</li>

	<li class="page_item page-item-332 haschildren"><a href="http://www.mitmazel.com/singles/myMitMazel.asp"><font color=red>New!</font>&nbsp;myMM</a>
	<ul>
		<li class="page_item page-item-334"><a href="http://www.mitmazel.com/singles/MyStatus.asp">My Account</a></li>
		<li class="page_item page-item-336"><a href="http://www.mitmazel.com/singles/ProfileEdit.asp">My Profile</a></li>
		<li class="page_item page-item-340"><a href="http://www.mitmazel.com/singles/MyPicturesAdd1.asp">My Pictures</a></li>
		<li class="page_item page-item-342"><a href="http://www.mitmazel.com/singles/MyMail.asp?SearchType=Inbox&ShowPageNum=1&MemberID=Anyone">My MMmail</a></li>
		<li class="page_item page-item-344"><a href="http://www.mitmazel.com/singles/MemberFavoritesDisplay.asp?ShowPageNum=1">My Favorites</a></li>
		<li class="page_item page-item-347 haschildren"><a href="http://www.mitmazel.com/singles/MyViewers.asp?ShowPageNum=1">Who Viewed Me</a>
		<li class="page_item page-item-347 haschildren"><a href="http://www.mitmazel.com/singles/MyViewees.asp?ShowPageNum=1">Who I Viewed</a>
		<li class="page_item page-item-347 haschildren"><a href="http://www.mitmazel.com/singles/RequestDeletionForm.asp">Change My Status</a>
		<li class="page_item page-item-347 haschildren"><a href="http://www.mitmazel.com/singles/emailMsgToRebbetzinForm.asp">Ask the Webmaster</a></li>
	</ul>
	</li>
	
	<li class="page_item page-item-291 haschildren"><a href="http://www.mitmazel.com/singles/memberJoin.asp">Join</a>
	<ul>
		<li class="page_item page-item-351"><a href="http://www.mitmazel.com/singles/memberJoinAssoc1.asp">Associate Members</a></li>
		<li class="page_item page-item-299"><a href="http://www.mitmazel.com/singles/memberJoinFull.asp?PrType=1MFT">Full Members</a></li>
		<li class="page_item page-item-361"><a href="http://www.mitmazel.com/matchmakers/matchmakeradd.asp">Matchmakers</a></li>
		<li class="page_item page-item-358"><a href="http://www.mitmazel.com/sponsors/sponsoradd.asp">Sponsors</a></li>
	</ul>
	</li>
	
	<li class="page_item page-item-311"><a href="http://www.mitmazel.com/Login/Login.asp">Login</a></li>
	
	<li class="page_item page-item-383"><a href="http://www.mitmazel.com/singles/memberSearchQuick.asp">Search</a>
	<ul>
		<li class="page_item page-item-351"><a href="http://www.mitmazel.com/singles/memberSearchQuick.asp">Quick Search</a></li>
		<li class="page_item page-item-299"><a href="http://www.mitmazel.com/singles/memberSearchAdvanced.asp?PrefsType=1">Advanced Search</a></li>
	</ul>
	</li>
	
	<li class="page_item page-item-2"><a href="http://www.mitmazel.com/AboutUs/howmmworks.asp">How MM Works</a></li>
	
	<li><a href="http://www.mitmazel.com/faqs/faqsindex.asp?ShowPageNum=1">FAQ</a></li>
	
	<li><a href="http://www.mitmazel.com/successstories/success_stories.asp?ShowPageNum=1">Success Stories</a></li>
	
	<li class="cat-item cat-item-34"><a href="#">etc</a>
	<ul>
		<li class="page_item page-item-351"><a href="http://www.halachos.com/Halacha/halacha.asp">Halachos</a></li>
		<li class="page_item page-item-299"><a href="http://www.jewishclassifieds.com">JewishClassifieds.com</a></li>
		<li class="page_item page-item-351"><a href="http://www.mitmazel.com/searchengine/sitesearch.asp">Jewish Web <font size=1>>></font></a>
		<ul>
			<li class="page_item page-item-408"><a href="http://www.mitmazel.com/searchengine/sitesearch.asp">Search Jewish Web</a></li>
			<li class="page_item page-item-410"><a href="http://www.mitmazel.com/searchengine/siteadd.asp">Join Jewish Web</a></li>
		</ul>
		</li>
		<li class="page_item page-item-299"><a href="http://www.mitmazel.com/JewishFun/JewishJokes/JewishJokesdbsearch.asp">Jokes & Stories</a></li>
		<li class="page_item page-item-351"><a href="http://www.mitmazel.com/JewishFun/JewishCartoons/JewishCartoons.asp?ShowPageNum=1">Cartoons</a></li>
		<li class="page_item page-item-299"><a href="http://www.mitmazel.com/JewishFun/JewishCartoonsRabbiSock/JewishCartoonsRabbisock.asp?ShowPageNum=1">Rabbi Sock</a></li>
	</ul>
	</li>

</ul>

<!--
	</div>
</div>
-->

<!--
</td>

</tr>
</table>
-->

<!-- End Drop-Down Menu Test ---------------------------------------------->



<br><br>
<center><font color="#000080" size="5">About&nbsp; Us</font></center>


<table border="0" cellPadding="0" cellSpacing="15" width="50%" align="center">
                <tr>
                    <td valign="middle" align="center">
                    <a href="contactus.asp" title="You never write! You never call!">
                    <img alt src="http://www.mitmazel.com/images/NavImages/contact75.jpg" WIDTH="75" HEIGHT="75" border=0></a>
                    </td>
                    <td align="left">
                        <font size="4">
                        <a href="contactus.asp" title="You never write! You never call!">Contact  Us</a>
                        </font>
                     </td>
                </tr>
                
                <tr>
                    <td valign="middle" align="center">
                    <a href="privacy.asp">
                    <img alt src="http://www.mitmazel.com/images/NavImages/privacy75.jpg" WIDTH="75" HEIGHT="75" border=0>
                    </a>
                    </td>
                    <td align="left">
                        <font size="4">Our <a href="privacy.asp">Privacy Policy</a></font>
                    </td>
                </tr>

                <tr>
                    <td valign="middle" align="center">
                    <a href="terms.asp">
                    <img alt src="http://www.mitmazel.com/images/NavImages/terms75.jpg" WIDTH="75" HEIGHT="75" border=0>
                    </a>
                    </td>
                    <td align="left">
                        <font size="4">Our <a href="terms.asp">Terms and Conditions</a> of Usage</font>
                    </td>
               </tr>
</table>

<br><br>
<!-- #INCLUDE VIRTUAL = "Misc/Footer.inc" -->



</body>
</html>
