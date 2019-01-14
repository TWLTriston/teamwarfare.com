<%
If strTournamentName = "Operation Triple Threat" Then 
	Call ContentStart(strTournamentName & " Tournament Sponsors")
	%>
	<div align="center"> 
	<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" 
		codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" WIDTH="728" HEIGHT="90" id="sc_banner" ALIGN=""> 
	<PARAM NAME="movie" VALUE="/images/allinwonder9700_728x90.swf?clickTag=http://www.ati.com"> 
	<PARAM NAME="quality" VALUE="high"> 
	<PARAM NAME="bgcolor" VALUE="#000000"> 
	<EMBED src="/images/allinwonder9700_728x90.swf?clickTag=http://www.ati.com" quality="high" bgcolor="#FFFFFF"  WIDTH="728" HEIGHT="90" NAME="rvs_banner" ALIGN="" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED> 
	</OBJECT> 
	<br /><br />
	</div>
	<center><a href="http://www.alienware.com/index.asp?from=Ravenshield 3 promo:01_banner_468x60"><img src="http://www.alienware.com/main/affiliate_pages/banners/01_banner_468x60.gif" width="468" height="60" border="0" alt="Alienware - Ultimate Gaming PC"></a><br />
	<br />
	</center>
	<%
	Call ContentEnd()
End If
If strTournamentName = "Answer the Call" Then
	Dim objBannerRotator, strSponsor
	Set objBannerRotator = Server.CreateObject( "MSWC.ContentRotator" )
	strSponsor = objBannerRotator.ChooseContent("cod_sponsors.txt")
	Set objBannerRotator = Nothing
	Call ContentStart("")
	%>
	<center>
	Below is a prize sponsor for Answer the Call, for a full list <a href="default.asp?tournament=Answer+the+Call&page=prizes">click here</a>.<br /><br />
	<% Response.Write strSponsor %>
	</center>
	<%
	Call ContentEnd()
End If
%>