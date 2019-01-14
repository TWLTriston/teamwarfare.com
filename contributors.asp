	<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "teamwarfare.com"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgcblack, bgcheader

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()%>
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

strSQL = "SELECT PlayerHandle FROM tbl_players WHERE Contributor = 1 ORDER BY ContributorAmount DESC"
oRs.Open strSQL, oConn
If Not(ors.eof and ors.bof) Then
	i = 0
	Do While Not(ors.EOF)
		i = i + 1
		if i mod 3 = 1 then
			if i > 1 then
				response.write "</tr>"
			end if
			response.write "<tr>"
		end if
		Response.write "<td>" & Server.HTMLENcode(oRs.Fields("PlayerHandle").Value & "") & "</td>"
		oRs.MoveNext
	Loop
End If
oRs.nextRecordSet

Dim arrOtherContributions(1)
arrOtherContributions (0) = "The High End Gaming Clan"
DIm j
For J = 0 to UBound(arrOtherContributions)
	i = i + 1
	if i mod 3 = 1 then
		if i > 1 then
			response.write "</tr>"
		end if
		response.write "<tr>"
	end if
	Response.write "<td>" & Server.HTMLENcode(arrOtherContributions(j) & "") & "</td>"
Next
Response.write "</tr>"


Dim strPlayerHandle
strPlayerHandle = Session("uName")

%>
</table>
<br />
<b>What are TWL Contributors?</b><br />
TWL Contributors are users of teamwarfare.com who felt the need to give back, via a monetary donation through paypal. 
The people above were under no obligation to make a donation. This is a public acknowledgement, and a thank you to them all.
<br /><br />
<b>How do I become a TWL Contributor?</b><br />
As we are going forward with a banner style advertising system, we are currently not accepting any further contributions. Thank you to all who have helped us get to where we are at!
<% If False Then %>
We are not soliciting donations, but having a buffer to help upgrade hardware and keep things running smoothly never hurts. 
You may send a donation to us by clicking the button below and using paypal.
<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="paypal@teamwarfare.com">
<input type="hidden" name="item_name" value="Contributions">
<input type="hidden" name="item_number" value="Contributions">
<input type="hidden" name="cn" value="Notes">
<input type="hidden" name="currency_code" value="USD">
<input type="hidden" name="tax" value="0">
<% If Len (strPlayerHandle) > 0 Then %>
<input type="hidden" name="on0" value="TWL Player Name">
<input type="hidden" name="os0" value="<%=Server.HTMLEncode(Session("uName") & "")%>">
<% Else %>
<input type="hidden" name="on0" value="TWL Player Name">
<input type="hidden" name="os0" value="Please enter your name below.">
<% End If %>
<center>
<input type="image" src="https://www.paypal.com/en_US/i/btn/x-click-but21.gif" border="0" name="submit" alt="Make payments with PayPal - it's fast, free and secure!" style="border: 0;">
</center>
</form>
<% End If %>

<!-- You may send a donation to us via <a href="https://www.paypal.com/xclick/business=paypal%40teamwarfare.com&item_name=Contributions&item_number=Contributions&cn=Notes+%28include+TWL+login+name%29">paypal</a>.
//-->
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