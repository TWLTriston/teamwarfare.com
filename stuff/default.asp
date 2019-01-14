<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TeamWarfare Stuff"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strPlayerHandle
strPlayerHandle = Session("uName")
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<% Call ContentStart("TeamWarfare Mouse Pads") %>
	<script language="javascript" type="text/javascript">
	<!--
		function fPopImages() {
			var objImageWin;
			objImageWin = window.open ("inuse/pics.html", "PicWin", "height=650,width=700,toolbar=0,location=0,resizable=0,scrollbars=0");
			 objImageWin.focus();
		}
	//-->
	</script>

<table border="0" cellspacing="0" cellpadding="0" width="97%">
<tr>
	<td>
		<table border="0" cellspacing="0" cellpadding="0" align="right">
		<tr>
			<td><a href="javascript:fPopImages();"><img src="/images/pad/padimage.jpg" height="300" width="350" alt="TeamWarfare Mouse Pad" border="0" /></a></td>
		</tr>
		<tr>
			<td align="center">
				<a href="javascript:fPopImages();">view more images</a>
			</td>
		</tr>
		</table>
		<br />
		TeamWarfare Stuff is now available! For all of you TeamWarfare fans who’ve been yearning for some tangible way to show off your affection for your favorite gaming league, we've now opened the Stuff Shop.  Dedicated to meeting the needs of the serious online gamer, only the best will satisfy us so the hunt for the quality we demand is grueling.  Our discriminating staff discovered the excellence of XTrac mouse pads during a recent tournament. <br />
		<br />
We've combined highly detailed graphics with precision texture to ensure that most optical mice will respond flawlessly to the surface through a full range of movement, from hair-trigger sniping, to breakneck 180-degree turns.  SureGrip backing assures you won’t slip up while you push your gaming to the limit all night long on the enormous surface area of this super thin pad.<br />
<br />
Simply said, if you’re a gamer, this is the mouse pad for you. <br />
<br />
Quick Facts: <br />
·	Large Surface — 8.5" x 11" <br />
·	Super Thin — Only 3/32" Thick! <br />
·	Durable, Flexible Plastic — Ensures durability and portability <br />
·	SureGrip Rubber Backing — Firmly holds to most surfaces <br />
·	Precision Texture & Detailed Graphics — Reliable and consistent mouse response <br />
<br /><br />
These pads are available in a limited quantity at this time for <b>$20.00 USD</b>.<br />
<br />
<b>For a limited time only! We are including FREE USPS shipping to anyone in the continental 48 states.</b>  Additional shipping charges may apply for shipping to addresses outside of the continental US.<br />
<br />
<font color="red"><b>Notice: This mousepad has been tested and may NOT work properly with recently purchased Logitech Opitcal Mice (MX500) or blue optic mice.</b></font><br />
<br />
Click the "Buy Now" button below to order this pad now for $20.00 using Paypal. You will receive a confirmation of your order within 24 hours. If you do not receive confirmation, please contact <a href="mailto:stuff@teamwarfare.com">stuff@teamwarfare.com</a> for assistance.<br />
 

		<br />
		
		<table border="0" cellspacing="0" cellpadding="0" align="center">
<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="stuff@teamwarfare.com">
<input type="hidden" name="undefined_quantity" value="1">
<input type="hidden" name="item_name" value="TeamWarfare Mouse Pad">
<input type="hidden" name="item_number" value="TWLPad">
<input type="hidden" name="amount" value="20.00">
<input type="hidden" name="cn" value="Notes">
<input type="hidden" name="currency_code" value="USD">
<input type="hidden" name="lc" value="US">
<% If Len (strPlayerHandle) > 0 Then %>
<input type="hidden" name="on0" value="TWL Player Name">
<input type="hidden" name="os0" value="<%=Server.HTMLEncode(Session("uName") & "")%>">
<% End If %>
<tr>
	<td><input type="image" src="https://www.paypal.com/en_US/i/btn/x-click-but23.gif" border="0" name="submit" alt="Make payments with PayPal - it's fast, free and secure!"></td>
</tr>
</form>
</table>
		
	</td>
</tr>
</table>

<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>