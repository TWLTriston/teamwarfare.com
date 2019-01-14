<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Member Preferences"

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

Dim AllSame, AllDifferent, Ladderpage, SigCheck, Sel, Showsig
Dim showsigs, showsigsword,showsigssel 

if Len(Session("ShowSigs")) = 0 Then
	Session("ShowSigs") = 1
End If
showsigs = Session("ShowSigs")
if showsigs then
	showsigsword = "Show signatures"
	showsigssel = "checked"
Else
	showsigsword = "Do not show signatures"
	showsigssel = ""
End If
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->
<% Call Content33BoxStart("Member Preferences - Ladder Settings") %>
	<form action="saveitem.asp" method="post" id=form2 name=form2>
	<table width="97%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td>
			This section defines how your computer downloads any page with a multipage list. Each option is customizable, but defaults are offered. These are stored inside a cookie on your local computer, so each workstation you
			access teamwarfare.com from, you will need to reconfigure these.
		</td>
	</tr>
	</table>
<% Call Content33BoxMiddle() %>
	<table width=97% align=center border=0 cellspacing=0 cellpadding=0 class="cssBordered">
<%
'--------
' Look up their current settings
'--------
AllSame = ""
AllDifferent = "mchecked"
LadderPage = request.cookies("PerPage")("LadderView")
SigCheck = request.cookies("PerPage")("ShowSig")
if LadderPage = "" then
	LadderPage = 25
end if
if SigCheck = "" then
	SigCheck = "n"
end if
if SigCheck = "y" then
	showsig = "Auto show sig"
	sel = " checked "
else
	showsig = "Do not auto show sig"
	sel = ""
end if
%>
	<tr><td colspan="2"><p class=small><b>View Ladder</b></p></td></tr>
	<tr height=30 bgcolor=<%=bgcone%>><td align=right><p class=small>Current:</p></td><td align=left><p class=small>&nbsp;<b><%=Ladderpage%></b> per page</p></td></tr>
	<tr bgcolor=<%=bgctwo%>><td align=right><p class=small>New:</p></td>
	<td align=left>&nbsp;<SELECT name=LadderView>
		<option value=10 <%if ladderpage = "10" then Response.Write " selected "%>>10 per page<br>
		<option value=15 <%if ladderpage = "15" then Response.Write " selected "%>>15 per page<br>
		<option value=20 <%if ladderpage = "20" then Response.Write " selected "%>>20 per page<br>
		<option value=25 <%if ladderpage = "25" then Response.Write " selected "%>>25 per page<br>
		<option value=30 <%if ladderpage = "30" then Response.Write " selected "%>>30 per page<br>
		<option value=40 <%if ladderpage = "40" then Response.Write " selected "%>>40 per page<br>
		<option value=50 <%if ladderpage = "50" then Response.Write " selected "%>>50 per page<br>
		<option value=75 <%if ladderpage = "75" then Response.Write " selected "%>>75 per page<br>
		<option value=100 <%if ladderpage = "100" then Response.Write " selected "%>>100 per page<br>
	</select></td></tr>
	<tr height=30 bgcolor=<%=bgcone%>><td colspan=2><p class=small>This used when viewing the complete ladder. Default is <b>25</b> ranked teams per page, before another page is generated.</p></td></tr>
	<tr><td colspan=2 align=center><input type=submit class=bright id=submit1 value="Save Preferences"><input type=hidden name=savetype value=SetCookie></td></tr>
</table>

<% Call Content33BoxEnd() %>
<% Call Content33BoxStart("Member Preferences - Forum Settings") %>
	<table width="97%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td>
			This section defines how you interact with the forums.
		</td>
	</tr>
	</table>

	
<% Call Content33BoxMiddle() %>
		<table width=97% align=center border=0 cellspacing=0 cellpadding=0 class="cssBordered">

		<tr><td colspan="2"><p class=small><b>Show Signature:</b></p></td></tr>
		<tr height=30 bgcolor=<%=bgcone%>><td align=right><p class=small>Current:</p></td><td align=left><p class=small>&nbsp;<b><%=showsig%></b></p></td></tr>
		<tr height=30 bgcolor=<%=bgctwo%>><td align=right><p class=small>New Setting:</p></td><td align=left><p class=small>&nbsp;<b><input type=checkbox id=showsig name=showsig value="ShowSigAuto" <%=sel%>></b></p></td></tr>
		<tr height=30 bgcolor=<%=bgcone%>><td colspan=2><p class=small>This setting will automatically 'check' the box in the reply/new thread pages in the forum section. Default is 'un-checked'</p></td></tr>

		<tr><td colspan="2"><p class=small><b>Show Forum Signatures:</b></p></td></tr>
		<tr height=30 bgcolor=<%=bgcone%>><td align=right><p class=small>Current:</p></td><td align=left><p class=small>&nbsp;<b><%=showsigsword%></b></p></td></tr>
		<tr height=30 bgcolor=<%=bgctwo%>><td align=right><p class=small>New Setting:</p></td><td align=left><p class=small>&nbsp;<b><input type=checkbox id=showsigs name=showsigs value="ShowSigs" <%=showsigssel %>></b></p></td></tr>
		<tr height=30 bgcolor=<%=bgcone%>><td colspan=2><p class=small>If you want to "hide" other user sigs, uncheck this box (You must be logged in for this to take affect).</p></td></tr>
		<tr><td colspan=2 align=center><input type=submit class=bright id=submit1 value="Save Preferences"></td></tr>		
		</table>
<% Call Content33BoxEnd() %>
<% Call Content33BoxStart("Member Preferences - Skin Style") %>
	<table width="97%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td>
			Choose how you would like TWL to be displayed by choosing a style option from the right. You must have a TWL account to take advantage of this feature.
		</td>
	</tr>
	</table>

	
<% Call Content33BoxMiddle() %>
		<table width=97% align=center border=0 cellspacing=0 cellpadding=0 class="cssBordered">
		<tr>
			<td colspan="3">Choose a Style:</td>
		</tr>
		<tr>
			<td valign="top" bgcolor="<%=bgctwo%>" width="20"><input class="borderless" type="radio" name="radStyle" id="radStyle" value="7" <% If Session("StyleID") = 7 Then Response.Write " checked=""checked"" " End If %> /></td>
			<td valign="top" bgcolor="<%=bgctwo%>">TeamWarfare - Classic</td>
		</tr>
		<tr>
			<td valign="top" bgcolor="<%=bgcone%>" width="20"><input class="borderless" type="radio" name="radStyle" id="radStyle" value="8" <% If Session("StyleID") = 8 Then Response.Write " checked=""checked"" " End If %> /></td>
			<td valign="top" bgcolor="<%=bgcone%>">TeamWarfare - Blue</td>
		</tr>
		<tr>
			<td valign="top" bgcolor="<%=bgctwo%>" width="20"><input class="borderless" type="radio" name="radStyle" id="radStyle" value="9" <% If Session("StyleID") = 9 Then Response.Write " checked=""checked"" " End If %> /></td>
			<td valign="top" bgcolor="<%=bgctwo%>">TeamWarfare - Green</td>
		</tr>
		<tr>
			<td valign="top" bgcolor="<%=bgctwo%>" width="20"><input class="borderless" type="radio" name="radStyle" id="radStyle" value="10" <% If Session("StyleID") = 10 Then Response.Write " checked=""checked"" " End If %> /></td>
			<td valign="top" bgcolor="<%=bgctwo%>">TeamWarfare - Classic (Larger Font)</td>
		</tr>
		<tr>
			<td valign="top" bgcolor="<%=bgcone%>" width="20"><input class="borderless" type="radio" name="radStyle" id="radStyle" value="11" <% If Session("StyleID") = 11 Then Response.Write " checked=""checked"" " End If %> /></td>
			<td valign="top" bgcolor="<%=bgcone%>">TeamWarfare - Blue (Larger Font)</td>
		</tr>
		<tr>
			<td valign="top" bgcolor="<%=bgctwo%>" width="20"><input class="borderless" type="radio" name="radStyle" id="radStyle" value="12" <% If Session("StyleID") = 12 Then Response.Write " checked=""checked"" " End If %> /></td>
			<td valign="top" bgcolor="<%=bgctwo%>">TeamWarfare - Green (Larger Font)</td>
		</tr>
		<tr><td colspan=3 align=center><input type=submit class=bright id=submit1 value="Save Preferences"></td></tr>		
		</table>
</form>
<% Call Content33BoxEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>
