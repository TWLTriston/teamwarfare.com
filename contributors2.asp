<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "teamwarfare.com"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim dictContributors
Set dictContributors = Server.CreateObject("Scripting.Dictionary")
dictContributors.Add "TotalCarnage", "200"
dictContributors.Add "Team Naturally Skilled", "100"
dictContributors.Add "ICFire", "50"
dictContributors.Add "KennyRemotelabs", "50"
dictContributors.Add "NixFix", "50"
dictContributors.Add "vawlk", "50"
dictContributors.Add "biscuit", " 50"
dictContributors.Add "subroutine", "50"
dictContributors.Add "Vendetta", "50"
dictContributors.Add "BlackBeltJones", "50"
dictContributors.Add "Norrin", "50"
dictContributors.Add "Ananasi", "50"
dictContributors.Add "Ireckon", "50"
dictContributors.Add "slurp", "40"
dictContributors.Add "Cryl", "30"
dictContributors.Add "Lessershade", "30"
dictContributors.Add "ricefrog", "30"
dictContributors.Add "Pancho Villa", "30"
dictContributors.Add "Fury X", "25"
dictContributors.Add "b-o-b-c-a-t", "25"
dictContributors.Add "ICE9", "25"
dictContributors.Add "Firewing", "25"
dictContributors.Add "PaveHawk-", "25"
dictContributors.Add "Bullseye", "20"
dictContributors.Add "Badwill", "20"
dictContributors.Add "ICRED", "20"
dictContributors.Add "tao| Death ICFU", "20"
dictContributors.Add "BlindOldMan", "20"
dictContributors.Add "DEATH-BY-FIRE", "20"
dictContributors.Add "Dragon Star", "20"
dictContributors.Add "ManBeef", "20"
dictContributors.Add "Janus", "20"
dictContributors.Add "Uso", "20"
dictContributors.Add "Squidly", "20"
dictContributors.Add "Juzfugen", "20"
dictContributors.Add "SmackDab", "20"
dictContributors.Add "Recluse", "20"
dictContributors.Add "[enraitsu]", "20"
dictContributors.Add "encryptor", "20"
dictContributors.Add "Got Haggis?", "20"
dictContributors.Add "Flyguy", "20"
dictContributors.Add "|AssMan|", "20"
dictContributors.Add "Captain_Kibitz", "20"
dictContributors.Add "-DeVast-", "20"
dictContributors.Add "Don Corleone", "20"
dictContributors.Add "envy", "15"
dictContributors.Add "Hypn0tik", "15"
dictContributors.Add "DoMiNuS", "15"
dictContributors.Add "Legend", "15"
dictContributors.Add "riverrat", "10"
dictContributors.Add "Jericho", "10"
dictContributors.Add "NoMadd", "10"
dictContributors.Add "Qing", "10"
dictContributors.Add "Silicon Demon", "10"
dictContributors.Add "ShinShun", "10"
dictContributors.Add "beLIEve", "10"
dictContributors.Add "AADiC_", "10"
dictContributors.Add "Llamasaurus Rex", "10"
dictContributors.Add "MrWatermonkey", "10"
dictContributors.Add "Buzzed", "10"
dictContributors.Add "^BuGs^", "10"
dictContributors.Add "Respice", "10"
dictContributors.Add "cimat1984", "10"
dictContributors.Add "Barak", "10"
dictContributors.Add "Sky Marshall", "5, 5"
dictContributors.Add "mojo (Team Vibez)", "5"
dictContributors.Add "Ceiba", "5"
dictContributors.Add "OutlawDienekes", "5"
dictContributors.Add "azn_essence", "5"
dictContributors.Add "SL={Dawg_Pack}", "5"
dictContributors.Add "-Forever-", "3.??"
dictContributors.Add "zigma", "1"
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("")
%>
<span style="font-size: 30px;">Thank you!</span>
<table border="0" width="97%">
<tr><td>
We would like to thank the people below for their generous contributions. <br /><br />We appreciate your continued support, <br />Triston, Polaris, and the TWL Staff<br /><br />
<b>TWL Contributors:</b><br />
<table border="0" cellspacing="0" cellpadding="0" width="97%">
<%
Dim strCont, i
i = 0
For each strCont in DictContributors
	i = i + 1
	if i mod 3 = 1 then
		if i > 1 then
			response.write "</tr>"
		end if
		response.write "<tr>"
	end if
	Response.write "<td>" & strCont & "</td>"
Next
Response.write "</tr>"
%>
</table>
<br />
<b>What are TWL Contributors?</b><br />
TWL Contributors are users of teamwarfare.com who felt the need to give back, via a monetary donation through paypal. 
The people above were under no obligation to make a donation. This is a public acknowledgement, and a thank you to them all.
<br /><br />
<b>How do I become a TWL Contributor?</b><br />
We are not soliciting donations, but having a buffer to help upgrade hardware and keep things running smoothly never hurts. 
You may send a donation to us via <a href="https://www.paypal.com/xclick/business=paypal%40teamwarfare.com&item_name=Contributions&item_number=Contributions&cn=Notes+%28include+TWL+login+name%29">paypal</a>.
</td>
</tr>
</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oRS = Nothing
Set oConn = Nothing
%>