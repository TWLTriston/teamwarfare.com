<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: View Demo"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim demoID, sec_passed, FileDownloadPath
Dim MatchWinnerID, MatchWinner, MatchLoserID, MatchLoser
Dim LadderName, LadderID, PublicFlag, admiNFlag
Dim PlayerTeamLinkID, MatchWinnerLinkID, MatchLoserLinkID
Dim PlayerName, PlayerID, FileSize, PositionPlayed
Dim MatchDate, MatchWinnerDefending
Dim MatchMap1DefenderScore, MatchMap1AttackerScore
Dim MatchMap2DefenderScore, MatchMap2AttackerScore
Dim MatchMap3DefenderScore, MatchMap3AttackerScore
Dim MatchMap1, MatchMap2, MatchMap3
Dim UploadDate, DownloadCount, Comments, PlayerTeamName
Dim DemoName
Dim PlayerTitle
Dim ex2, ex
Dim times, myID
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" --><%
bgc = bgcone
	
DemoID = Request.QueryString("DemoID")
If DemoID = "" Or Not(IsNumeric(DemoID)) Then 
	Response.Redirect("default.asp")
End If

'security check
strSQL = "SELECT AdminFlag, PublicFlag,TLLinkID FROM tbl_Demos WHERE DemoID=" & DemoID
Set oRS = oConn.Execute(strSQL)
If Not(Ors.Eof And Ors.Bof) Then

	If Not(Session("LoggedIn")) And oRS("AdminFlag") = 0 And oRS("PublicFlag") = 1 Then
		'public demo, user is not logged in
	ElseIf oRS("AdminFlag") = 1 Then
		' demo for admins only
		If Not bSysAdmin and Not bAnyLadderAdmin Then
			oRs.Close
			oConn.Close
			Set oRs = Nothing
			Set oRS2 = Nothing
			Set oConn = Nothing
			Response.Clear
			Response.Redirect("/errorpage.asp?error=22")
		End If
	Elseif oRS("PublicFlag") = 0 Then
		' team only demo
		strSQL = "SELECT TLLinkID FROM vPlayerTeams WHERE PlayerHandle='" & CheckString(session("uName")) & "'"
		oRs2.Open strSQL, oConn
		sec_passed = 0
		If Not(oRS2.EOF AND oRS2.BOF) Then 
			Do While Not oRS2.EOF
				If oRS2("TLLinkID") = oRS("TLLinkID") Then
					sec_passed = 1
				End If
			oRS2.MoveNext
			Loop
		End If
		oRS2.NextRecordset 			
		If bSysAdmin Then 
			sec_passed = 1
		End If
		If sec_passed = 0 Then 
			oRs.Close
			oConn.Close
			Set oRs = Nothing
			Set oRS2 = Nothing
			Set oConn = Nothing
			Response.Clear
			Response.Redirect("/errorpage.asp?error=23")
		End If
	End If
Else
	oRs.Close
	oConn.Close
	Set oRs = Nothing
	Set oRS2 = Nothing
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
oRs.NextRecordset 

'if user is downloading a file
If Request.QueryString("dl") = "1" Then 
	If Not(Session("LoggedIn")) then 
		oConn.Close
		Set oRs = Nothing
		Set oRS2 = Nothing
		Set oConn = Nothing
		Response.Redirect("/errorpage.asp?error=2")
	end if
	
	strSQL = "SELECT h.MatchID, d.FileName FROM tbl_demos d, tbl_history h WHERE d.HistoryID = h.HistoryID AND d.DemoID=" & DemoID
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRS.BOF) Then 
		FileDownloadPath = "/demos/uploads/" & oRS("MatchID") & "/" & oRS("FileName")
		' inc dl #
		oRS.Close
		strSQL = "UPDATE tbl_Demos SET DownloadCount=DownloadCount+1 "
		strSQL = strSQL & " WHERE DemoID=" & DemoID
		oConn.Execute(strSQL)
		oConn.Close
		Set oRS2 = Nothing
		Set oRs = Nothing
		Set oConn = Nothing
		Response.Clear
		Response.Redirect(FileDownloadPath)
	Else
		oRs.Close
		oConn.Close
		Set oRS2 = Nothing
		Set oRs = Nothing
		Set oConn = Nothing
		Response.Clear 
		Response.Redirect("default.asp")
	End If	
End If

'strSQL = "SELECT * FROM tbl_Demos WHERE DemoID=" & DemoID
'Set oRS = oConn.Execute(strSQL)

' get name of demo - Winner vs. Loser

strSQL = "EXECUTE ViewDemo @DemoID = '" & CheckString(DemoID) & "'"
oRS.Open strSQL, oConn ' Set oRS = oConn.Execute(strSQL)
If Not(oRS.EOF and oRS.BOF) Then
		MatchWinnerID  = oRS("MatchWinnerID")
		MatchWinner    = oRS("WinnerName")
		MatchLoserID   = oRS("MatchLoserID")
		MatchLoser     = oRS("LoserName")
		LadderName     = oRS("LadderName")
		LadderID       = oRS("LadderID")

		PublicFlag	   = oRS("PublicFlag")
		AdminFlag      = oRS("AdminFlag")
		
		PlayerTeamLinkID  = oRS("TLLinkID")
		MatchWinnerLinkID = oRS("MatchWinnerID")
		MatchLoserLinkID  = oRS("MatchLoserID")
		PlayerName        = oRS("PlayerHandle")
		PlayerID          = oRS("PlayerID")
		FileSize          = oRS("FileSize")
		PositionPlayed    = oRS("PositionPlayed")

		MatchDate         = oRS("MatchDate")
		MatchWinnerDefending   = oRS("WinnerDefending")
		MatchMap1DefenderScore = oRS("MatchMap1DefenderScore")
		MatchMap1AttackerScore = oRS("MatchMap1AttackerScore")
		MatchMap2DefenderScore = oRS("MatchMap2DefenderScore")
		MatchMap2AttackerScore = oRS("MatchMap2AttackerScore")
		MatchMap3DefenderScore = oRS("MatchMap3DefenderScore")
		MatchMap3AttackerScore = oRS("MatchMap3AttackerScore")
		MatchMap1 = oRS("MatchMap1")
		MatchMap2 = oRS("MatchMap2")
		MatchMap3 = oRS("MatchMap3")
		
		UploadDate = oRS("upload_dtim")
		DownloadCount = oRS("DownloadCount")
		Comments = oRS("Comments")
Else
	oRs.Close
	oConn.Close 
	Set oRs = Nothing
	Set oRS2 = Nothing
	Set oConn = Nothing
	Response.Clear 
	Response.Redirect "default.asp"
End If
oRS.NextRecordset

If PlayerTeamLinkID = MatchWinnerLinkID Then
	PlayerTeamName = MatchWinner
Else
	PlayerTeamName = MatchLoser
End If

'player title
strSQL = "SELECT PlayerTitle FROM tbl_Players WHERE PlayerID=" & PlayerID
oRS.Open strSQL, oConn
If Not(oRS.EOF and ors.BOF) Then
	PlayerTitle = oRS("PlayerTitle")
End If
oRs.NextRecordset 

DemoName = MatchWinner & " vs. " & MatchLoser
FileSize = Round(int(FileSize)/(1024*1024), 2)
If FileSize = 0 Then 
	FileSize = "Less than 1"
End If

Call Content2BoxStart("Demo Info: " & DemoName)
%>
	<table width="780" border="0" cellspacing="0" cellpadding="2">
	<tr>
	<td width="8"><img src="/images/spacer.gif" width="8" height="1"></td>
	<td width="375" valign=top>
	<a href="default.asp">&lt; Back to Demo List</a>
	<br><br>
		<table width="100%" align="center" cellpadding="0" cellspacing="0" border="0" BGCOLOR="#444444">
		<TR><TD>
		<table width="100%" align="center" cellpadding="2" cellspacing="1" border="0">
			<TR BGCOLOR="#000000">
				<TH COLSPAN=2>Demo Details</TH>
			</TR>
			<tr bgcolor="<%=bgctwo%>">
				<td width="30%" align="right">Ladder:</td>
				<td><A HREF="/viewladder.asp?ladder=<%=server.UrlEnCode(LadderName & "")%>"><%=LadderName%></A></td>
			</tr>
			<tr bgcolor="<%=bgcone%>">
				<td align="right">Team 1:</td>
				<td><A HREF="/viewteam.asp?team=<%=server.UrlEnCode(MatchWinner & "")%>"><%=MatchWinner%></A></td>
			</tr>
			<tr bgcolor="<%=bgctwo%>">
				<td align="right">Team 2:</td>
				<td><A HREF="/viewteam.asp?team=<%=server.UrlEnCode(MatchLoser & "")%>"><%=MatchLoser%></A></td>
			</tr>
			<tr bgcolor="<%=bgcone%>">
				<td align="right">Match Date:</td>
				<% If IsDate(MatchDate) Then %>
				<td><%=formatDateTime(MatchDate, vbShortDate)%></td>
				<% Else %>
				<td>&nbsp;</td>
				<% End if %>
			</tr>
			<tr bgcolor="<%=bgctwo%>">
				<td align="right">Upload Date:</td>
				<td><%=formatDateTime(UploadDate, vbShortDate)%></td>
			</tr>
			<tr bgcolor="<%=bgcone%>">
				<td align="right">POV:</td>
				<td><%=PlayerName%> - <%= PlayerTeamName %></td>
			</tr>			
			<tr bgcolor="<%=bgctwo%>">
				<td align="right">Position:</td>
				<td><%=PositionPlayed%></td>
			</tr>			
			<tr bgcolor="<%=bgcone%>">
				<td align="right" valign=top>Match Stats:<br>(Team1 - Team2)</td>
				<td>
					<%
						'map 1
						Response.Write "<b>" & MatchMap1 & ":</b>&nbsp;"
						If MatchWinnerDefending Then 
							Response.Write MatchMap1DefenderScore & "-" & MatchMap1AttackerScore 
						Else
							Response.Write MatchMap1AttackerScore & "-" & MatchMap1DefenderScore
						End If
						Response.Write "<br>"
						
						'map 2
						Response.Write "<b>" & MatchMap2 & ":</b>&nbsp;"
						If MatchWinnerDefending Then 
							Response.Write MatchMap2DefenderScore & "-" & MatchMap2AttackerScore
						Else
							Response.Write MatchMap2AttackerScore & "-" & MatchMap2DefenderScore
						End If						
						Response.Write "<br>"
						
						'map 3
						Response.Write "<b>" & MatchMap3 & ":</b>&nbsp;"
						If MatchWinnerDefending Then 
							Response.Write MatchMap3DefenderScore & "-" & MatchMap3AttackerScore
						Else
							Response.Write MatchMap3AttackerScore & "-" & MatchMap3DefenderScore
						End If						
					%>
				</td>
			</tr>						
		</table>
			</td>
		</tr>						
	</table>
	</td>
	<td width="15"><img src="/images/spacer.gif" width="15" height="1"></td>
	<td width="375" valign="top">
		<table width="100%" cellpadding="2" cellspacing="2" border="0">
			<tr>
				<td width="20%" align="center" valign="middle">
					<a href="viewDemo.asp?DemoID=<%=DemoID%>&dl=1"><img src="winzip_icon.gif" border="0" WIDTH="40" HEIGHT="40"><br>Download!</a>
				</td>
				<td valign="top" width="80%">
					<table width="100%" border="0" align="left">
						<tr bgcolor="<%=bgcone%>">
							<td width="30%" align="right">File Size:</td>
							<td><%=FileSize%> mb</td>
						</tr>
						<tr bgcolor="<%=bgctwo%>">
							<td align="right">Downloads:</td>
							<td><%=DownloadCount%></td>
						</tr>
						<tr bgcolor="<%=bgcone%>">
							<td colspan="2" align="center"><font size="1">All demos are in .ZIP format. Please visit <a href="http://www.winzip.com">www.winzip.com</a> to get Winzip.</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<font size="1">
						<b>Disclaimer:</b> Download files with caution, be sure to run a virus scan over all files download via the internet.
						<br>		
					</font>
						<%If Session("uName") = PlayerName or bSysAdmin Then %>
						<form action="saveflags.asp" method=post>
						<input type=hidden name=DemoID value=<%=DemoID%>>
						<table width=300 align=center border=0 cellpadding=0 cellspacing=0 BGCOLOR="444444">
							<TR><TD>
							<table width=100% align=center border=0 cellpadding=2 cellspacing=1>
							<tr BGCOLOR="#000000"><th colspan=2><b>Edit Demo Flags</b></th></tr>
							<tr bgcolor=<%=bgcone%>>
								<td valign=top ALIGN=RIGHT><b>Public</b></td>
								<td>
								<%If PublicFlag=1 Then ex=" checked"%>
									<input type=checkbox name=PublicFlag<%=ex%>>	
								<br><font size=1>Demos that are not public are viewable only by members of your team.</font>
								</td>
							</tr>
							<tr bgcolor=<%=bgctwo%>>
								<td nowrap ALIGN=RIGHT><b>Admins Only?</b></td>
								<td>
								<%If AdminFlag=1 Then ex2=" checked"%>
									<input type=checkbox name=AdminFlag<%=ex2%>>
								</td>
							</tr>							
							<tr bgcolor="#000000"><td align=center colspan=2><input type=submit name=submitFlags value="  Update  "></td></tr>
						</table>
						</tD></TR></TABLE>
						</form>
						<%End If%>
				</td>
			</tr>
		</table>
	</td>
	<td width="7"><img src="/images/spacer.gif" width="7" height="1"></td>
	</tr>
	</table>
<%
Call Content2BoxEnd() 
Call ContentStart("Demo Comments")
%>
		<table width=97% align=center border=0>
			<tr>
				<td width=70% valign=top>
					<table border=0 cellpadding=0 cellspacing=0 BGCOLOR=#444444 width=100%>
					<TR><TD>
					<table border=0 cellpadding=2 cellspacing=1 width=100%>
						<TR BGCOLOR=#000000><TH COLSPAN=2>Current Comments</TH></TR>
						<tr bgcolor=<%=bgcone%>>
							<td width=30% valign=top height=75>
								<b><%=PlayerName%></b><br>
								<FONT color="#AAAAAA"><%=PlayerTitle%></FONT>
							</td>
							<td width=70% valign=top>
								<table border=0 cellspacing=0 cellpadding=2 width=100%>
								<TR><TD><FONT color="#AAAAAA">&nbsp;Author Comments</FONT><BR><HR CLASS=FORUM></TD></TR>
								<TR><TD><%=Comments%></TD></TR>
								</TABLE>
							</td>
						</tr>
						<%
						
						' get visitors id
						myID = Session("PlayerID")
						strSQL = "SELECT l.*, p.PlayerHandle, p.PlayerTitle FROM lnk_comment_demo l, tbl_players p WHERE p.PlayerID = l.PlayerID AND DemoID=" & DemoID & " ORDER BY LCDID"
						oRS2.Open strSQL, oConn 
						If Not(oRS2.EOF AND oRS2.BOF) Then
							Do While Not oRS2.EOF
									If times mod 2 = 0 Then
										bgc = bgctwo
									Else
										bgc = bgcone
									End If						
								%>
								<tr bgcolor=<%=bgc%>>
									<td valign=top height=75>
										<b><%=oRS2("PlayerHandle")%></b><br>
										<FONT color="#AAAAAA"><%=oRS2("PlayerTitle")%></FONT>
									</td>
									<td valign=top>
										<table border=0 cellspacing=0 cellpadding=2 width=100%>
										<TR><TD><FONT color="#AAAAAA">&nbsp;<%=oRS2("CommentTime")%></FONT><br><hr class=forum></TD></TR>
										<TR><TD VALIGN=TOP><%=oRS2("Comment")%></TD></TR>
										</TABLE>
									</td>
								</tr>
								<%
								oRS2.MoveNext
								times = times + 1
							Loop
						End If
						oRS2.NextRecordset 
						%>
					</table>
					</TD></TR></TABLE>
				</td>
				<td width=30% valign=top ALIGN=RIGHT>
			<%If Not(Session("LoggedIn")) Then%>
				<center><b>You must be logged in to post a comment.</b></center>
			<%Else%>
				<%If Request.QueryString("err") = "1" Then%>
					<center><b>Please fill in the Message field!</b></center>
				<%End If%>
				<form action="savecomment.asp" method=post>
				<input type=hidden name=DemoID value=<%=DemoID%>>
				<input type=hidden name=PlayerID value=<%=myID%>>
					<table border=0 cellpadding=0 cellspacing=0 BGCOLOR=#444444 width=100%>
					<TR><TD>
					<table border=0 cellpadding=2 cellspacing=1 width=100%>
					<TR BGCOLOR=#000000><TH COLSPAN=2>Add Comment</TH></TR>
					<tr bgcolor=<%=bgcone%>>
						<td width=25% align=right>Handle:</td>
						<td><%=Session("uName")%></td>
					</tr>
					<tr bgcolor=<%=bgctwo%>>
						<td valign=top width=25% align=right>Message:</td>
						<td valign=top><textarea name=comment rows=4 cols=20></textarea></td>
					</tr>						
					<tr bgcolor=<%=bgcone%>>
						<td COLSPAN=2 align=CENTER>Use Signature: <input type=checkbox name="sig" checked></td>
					</tr>
					<tr bgcolor=<%=bgctwo%>>
						<td colspan=2 align=center><input type=submit name=submit value=" Reply! "></td>	
					</tr>					
					</table></TD></Tr>
					</table>
				</form>
			<%End If%>
				</td>
			</tr>
		</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS2 = Nothing
Set oRS = Nothing
%>