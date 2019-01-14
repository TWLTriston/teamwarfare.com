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
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<% Call ContentStart("TeamWarfare Mouse Pads") %>

<table border="0" cellspacing="0" cellpadding="0" width="97%">
<tr>
	<td>
		<img src="/images/pad/padimage.jpg" height="300" width="350" alt="TeamWarfare Mouse Pad" border="0"  align="right" />
		<br />
		TeamWarfare Stuff is now available! For all of you TeamWarfare fans who’ve been yearning for some tangible way to show off your affection for your favorite gaming league, we've now opened the Stuff Shop.  Dedicated to meeting the needs of the serious online gamer, only the best will satisfy us so the hunt for the quality we demand is grueling.  Our discriminating staff discovered the excellence of XTrac mouse pads during a recent tournament. <br />
		<br />
We've combined highly detailed graphics with precision texture to ensure that optical mice will respond flawlessly to the surface through a full range of movement, from hair-trigger sniping, to breakneck 180-degree turns.  SureGrip backing assures you won’t slip up while you push your gaming to the limit all night long on the enormous surface area of this super thin pad.<br />
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
Click the "Buy Now" button below to order this pad now for $20.00 using Paypal. You will receive a confirmation of your order within 24 hours. If you do not receive confirmation, please contact <a href="mailto:stuff@teamwarfare.com">stuff@teamwarfare.com</a> for assistance.<br />
 

		<br />
		
		<table border="0" cellspacing="0" cellpadding="0" align="center">
<form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="PadPurchase">
<input type="hidden" name="cmd" value="_s-xclick">
<input type="hidden" name="encrypted" value="-----BEGIN PKCS7-----
MIIHRwYJKoZIhvcNAQcEoIIHODCCBzQCAQExggEwMIIBLAIBADCBlDCBjjELMAkG
A1UEBhMCVVMxCzAJBgNVBAgTAkNBMRYwFAYDVQQHEw1Nb3VudGFpbiBWaWV3MRQw
EgYDVQQKEwtQYXlQYWwgSW5jLjETMBEGA1UECxQKbGl2ZV9jZXJ0czERMA8GA1UE
AxQIbGl2ZV9hcGkxHDAaBgkqhkiG9w0BCQEWDXJlQHBheXBhbC5jb20CAQAwDQYJ
KoZIhvcNAQEBBQAEgYCPOU8904JcDYT/qnzpLWiQ0B6PADhR8Uo1oOp43Wa8Tlrr
e/eO28EqFvUFOaAZ8aEpNyMkjlrBD4FNVyB1SQ6+VoilsGwDfqGbAK5PCYCbSeXV
w53CO/5N6aZj+NDa0WJeEfJ0dwMz5WOed6BlWhkTu0Ovbt4gzqJuXiXDBOWVMDEL
MAkGBSsOAwIaBQAwgcQGCSqGSIb3DQEHATAUBggqhkiG9w0DBwQIxlVz3R0/E2qA
gaDSKq/ar1Fwj2Vy8XYeH6/1ho7JE8M6Jhxc6Ki3BeEnueO0qYWtahiJAVfUANKK
q1Lq/H+dfaJkiSll9nFLOchSKNeCs8OUEZoFN/q3GrNhsMKNM1E+QF6xDmZk5xrq
AnLqOleCwVnIozCkUFikUHv95AmJCF1KN2XoCddnMQ2MJa3KnscfWCjFTb66N43t
PNuy27yIqWcMJbGVbE7drsB7oIIDhzCCA4MwggLsoAMCAQICAQAwDQYJKoZIhvcN
AQEFBQAwgY4xCzAJBgNVBAYTAlVTMQswCQYDVQQIEwJDQTEWMBQGA1UEBxMNTW91
bnRhaW4gVmlldzEUMBIGA1UEChMLUGF5UGFsIEluYy4xEzARBgNVBAsUCmxpdmVf
Y2VydHMxETAPBgNVBAMUCGxpdmVfYXBpMRwwGgYJKoZIhvcNAQkBFg1yZUBwYXlw
YWwuY29tMB4XDTA0MDIxMzEwMTMxNVoXDTM1MDIxMzEwMTMxNVowgY4xCzAJBgNV
BAYTAlVTMQswCQYDVQQIEwJDQTEWMBQGA1UEBxMNTW91bnRhaW4gVmlldzEUMBIG
A1UEChMLUGF5UGFsIEluYy4xEzARBgNVBAsUCmxpdmVfY2VydHMxETAPBgNVBAMU
CGxpdmVfYXBpMRwwGgYJKoZIhvcNAQkBFg1yZUBwYXlwYWwuY29tMIGfMA0GCSqG
SIb3DQEBAQUAA4GNADCBiQKBgQDBR07d/ETMS1ycjtkpkvjXZe9k+6CieLuLsPum
sJ7QC1odNz3sJiCbs2wC0nLE0uLGaEtXynIgRqIddYCHx88pb5HTXv4SZeuv0Rqq
4+axW9PLAAATU8w04qqjaSXgbGLP3NmohqM6bV9kZZwZLR/klDaQGo1u9uDb9lr4
Yn+rBQIDAQABo4HuMIHrMB0GA1UdDgQWBBSWn3y7xm8XvVk/UtcKG+wQ1mSUazCB
uwYDVR0jBIGzMIGwgBSWn3y7xm8XvVk/UtcKG+wQ1mSUa6GBlKSBkTCBjjELMAkG
A1UEBhMCVVMxCzAJBgNVBAgTAkNBMRYwFAYDVQQHEw1Nb3VudGFpbiBWaWV3MRQw
EgYDVQQKEwtQYXlQYWwgSW5jLjETMBEGA1UECxQKbGl2ZV9jZXJ0czERMA8GA1UE
AxQIbGl2ZV9hcGkxHDAaBgkqhkiG9w0BCQEWDXJlQHBheXBhbC5jb22CAQAwDAYD
VR0TBAUwAwEB/zANBgkqhkiG9w0BAQUFAAOBgQCBXzpWmoBa5e9fo6ujionW1hUh
PkOBakTr3YCDjbYfvJEiv/2P+IobhOGJr85+XHhN0v4gUkEDI8r2/rNk1m0GA8HK
ddvTjyGw/XqXa+LSTlDYkqI8OwR8GEYj4efEtcRpRYBxV8KxAW93YDWzFGvruKnn
LbDAF6VR5w/cCMn5hzGCAZowggGWAgEBMIGUMIGOMQswCQYDVQQGEwJVUzELMAkG
A1UECBMCQ0ExFjAUBgNVBAcTDU1vdW50YWluIFZpZXcxFDASBgNVBAoTC1BheVBh
bCBJbmMuMRMwEQYDVQQLFApsaXZlX2NlcnRzMREwDwYDVQQDFAhsaXZlX2FwaTEc
MBoGCSqGSIb3DQEJARYNcmVAcGF5cGFsLmNvbQIBADAJBgUrDgMCGgUAoF0wGAYJ
KoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMDQwMzI1MTYx
MTQ4WjAjBgkqhkiG9w0BCQQxFgQU5v+IXECSOLMyr+i2ylZ6Qg4hCAwwDQYJKoZI
hvcNAQEBBQAEgYAEcTMZAWN6Qm7h9S/iTRLWiduu8C1NLhGhnp29X+yxxV6CeUGq
AbDljCBB8wvr636UhzACoUfvFgi5D55oFq5+L7/NhJ8omVnMIDqDi2INfyAQDU+W
Oz8j7DP6cSLl9fVygGXz7mWVFd3O9NXhQqwrfzQ4t7Oz/dh6cBXZD2llfQ==
-----END PKCS7-----
">
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