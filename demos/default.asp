<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Demo Archive"

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

Dim PageTitle, AdditionalSQL, AdditionalFilterSQL
Dim frm_search, frm_teamName, frm_playerName, frm_MapName
Dim frm_position, frm_sort, frm_d1, frm_d2, qry_historyid
Dim OrderBySQL
Dim next_date, start_date, end_date
Dim times, MatchWinnerID, MatchWinner, MatchLoserID, MatchLoser, MatchDate, PlayerName, LadderName
Dim numd, d1, d2, UploadDate
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
frm_search = Request.Form("Search")
frm_teamName = replace(trim(Request.Form("TeamName")), "'", "''")
frm_playerName = replace(trim(Request.Form("PlayerName")), "'", "''")
frm_MapName = replace(trim(Request.Form("MapName")), "'", "''")
'frm_position = replace(trim(Request.Form("Position")), "'", "''")
frm_sort = replace(trim(Request.Form("sort")), "'", "''")
frm_d1 =  replace(trim(Request("d1")), "'", "''")
frm_d2 =  replace(trim(Request("d2")), "'", "''")
qry_historyID = trim(Request.QueryString ("historyID"))

Dim intAdminFlag, intPublicFlag

PageTitle = "TWL: Public Demo Archive"
If Request.QueryString("t") = "team" Then
'	AdditionalSQL = " AND AdminFlag=0 AND (PublicFlag=0 AND d.TLLinkID IN (SELECT TLLinkID FROM vPlayerTeams where PlayerHandle='" & Replace(Session("uName"), "'", "''") & "'))"
'	AdditionalFilterSQL = ""
	PageTitle = "TWL: Private Demo Archive"
	intAdminFlag = 0
	intPublicFlag = 0
ElseIf Request.QueryString ("t") = "me" Then
	frm_playerName = Session("uName")
'	AdditionalSQL = " AND p.PlayerHandle='" & replace(session("uName"), "'", "''") & "'"
'	AdditionalFilterSQL = ""
	PageTitle = "TWL: My Demo Archive"
	intAdminFlag = -1
	intPublicFlag = -1
ElseIf (IsAnyLadderAdmin() Or IsSysAdmin()) AND Request.QueryString ("t") = "all" Then
	AdditionalFilterSQL = ""
	PageTitle = "TWL: Admin Version Demo Archive"
	intAdminFlag = -1
	intPublicFlag = -1
Else
	AdditionalFilterSQL = " AND AdminFlag=0 AND PublicFlag=1"
	intAdminFlag = 0
	intPublicFlag = 1
End If

OrderBySQL = " upload_dtim ASC "

If Len(frm_sort) > 0 Then
	OrderBySQL = frm_sort & " ASC "
Else
	frm_sort = "h.MatchDate"
End If	
If len(frm_search) > 0 Then
	
End If
if bgc=bgcone then
	bgc=bgctwo
else
	bgc=bgcone
end if
	
							
strSQL = "SELECT TOP 1 upload_dtim FROM tbl_Demos WHERE TLLinkID IS NOT NULL " & AdditionalFilterSQL & " ORDER BY upload_dtim ASC"
Set oRS = oConn.Execute(strSQL)
If Not(ors.eof and ors.bof) Then 
	start_date = month(oRS("upload_dtim")) & "/" & day(oRS("upload_dtim")) & "/" & year(oRS("upload_dtim"))
End If
oRS.Close
strSQL = "SELECT TOP 1 upload_dtim FROM tbl_Demos WHERE TLLinkID IS NOT NULL " & AdditionalFilterSQL & " ORDER BY upload_dtim DESC"
Set oRS = oConn.Execute(strSQL)
If Not(ors.eof and ors.bof) Then 
	end_date = month(oRS("upload_dtim")) & "/" & day(oRS("upload_dtim")) & "/" & year(oRS("upload_dtim"))
End If
oRS.Close
						
next_date = start_date
If IsNull(next_date) or Len(next_date) = 0 Then
 next_date = Date()
 start_date = Date()
 end_Date = Date()
End If

	
If len(frm_search) <> 0 Then
	d1 = frm_d1
	d2 = frm_d2
Else
	If Request("d1") = "" Then
		d1 = DateAdd("d", -10, end_date)
		d2 = end_date
	Else
		d1 = Request("d1")
		d2 = Request("d2")
	end If
End If
if len(qry_historyID) > 0 THen
	d1 = "1/1/2000"
	d2 = month(now()) & "/" & day(now()) & "/" & year(now())
End If
'Response.Write Request("D1") & "---" & d1

Call ContentStart(PageTitle)
%>
<%
If Len(Session("PlayerID")) > 0 Then
	%>
		<table width="760" align="center" border="0" cellpadding="0" cellspacing="0" height="100%" BGCOLOR=#444444>
		<TR><TD>
		<table cellpadding=2 cellspacing=1 border=0 ALIGN=CENTER WIDTH=100%>
			<tr HEIGHT=25 VALIGN=CeNTER>
				<td align=center bgcolor=<%=bgcone%> >
					[ <a href="addDemo.asp"><font size=1>Add Demo</font></a> ]
					<% If bSysAdmin or bAnyLadderAdmin Then %>
					[ <a href="default.asp?t=all"><font size="1">All Demos</font></a> ]
					<% End If %>
					[ <a href="default.asp"><font size="1">Public Demos</font></a> ]
					[ <a href="default.asp?t=me"><font size=1>View Only My Demos</font></a> ]
					[ <a href="default.asp?t=team"><font size="1">View Only Team/Private Demos</font></a> ]
				</td>
			</tr>
		</table>
		</TD></TR></TABLE>
		
		<br>
		
		<table width="100%" border="0" cellpadding="0" cellspacing="1">
					<tr>
						<td width="50%" valign="top">
							<table width="97%" align="center" border="0" cellpadding="0" cellspacing="0" BGCOLOR=#444444>
							<TR><TD>
							<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">
								<form ACTION="/demos/default.asp" METHOD="POST">
								<tr>
									<tH colspan=2 bgcolor="000000">Search Demos</tH>
								</tr>
								<tr>
									<td colspan="2" bgcolor="#000000">You must type the exact team/player name for searching.<br />A date range is required for searching.</td>
								</tr>
								<tr bgcolor="<%=bgcone%>">
									<td><b>Team:</b></td>
									<td><input type="text" name="TeamName" VALUE="<%=frm_teamname%>"></td>
								</tr>
								<tr bgcolor="<%=bgctwo%>">
									<td><b>Player:</b></td>
									<td><input type="text" name="PlayerName" VALUE="<%=frm_playername%>"></td>
								</tr>
								<tr bgcolor="<%=bgcone%>">
									<td><b>Map:</b></td>
									<td><input type="text" name="MapName" VALUE="<%=frm_MapName%>"></td>
								</tr>
								<tr bgcolor="<%=bgctwo%>">
									<td><b>Dates Between:</b></td>
									<td><input type="text" name="d1" VALUE="<%=d1%>">-<input type="text" name="d2" VALUE="<%=d2%>"></td>
								</tr>							
								<tr bgcolor="<%=bgcone%>">
									<td><b>Sort Results By:</b></td>
									<td>
										<select name="sort">
											<option value="Winnername" <% If frm_sort = "WinnerName" Then Response.Write " SELECTED "%>>Team One</option>
											<option value="LoserName" <% If frm_sort = "Losername" Then Response.Write " SELECTED "%>>Team Two</option>
											<option value="PlayerHandle" <% If frm_sort = "PlayerHandle" Then Response.Write " SELECTED "%>>Player</option>
											<option value="MapPlayed" <% If frm_sort = "MapPlayed" Then Response.Write " SELECTED "%>>Map</option>
											<option value="PositionPlayed" <% If frm_sort = "PositionPlayed" Then Response.Write " SELECTED "%>>Position</option>
											<option value="upload_dtim" <% If frm_sort = "upload_dtim" Then Response.Write " SELECTED "%>>Match Date</option>
										</select>
									</td>
								</tr>
								<tr BGCOLOR="<%=bgctwo%>">
									<td COLSPAN="2" ALIGN="CENTER"><input TYPE="SUBMIT" VALUE="Search"></td>
								</tr>
									<input TYPE="HIDDEN" Name="Search" VALUE="Search!">
									</td>
								</tr>
								</form>
							</table>
							</td>
							</tr>
							</table>
						</td>
						<td width="50%" valign="top">
							<table width="97%" align="center" border="0" cellpadding="0" cellspacing="0" height="100%" BGCOLOR=#444444>
							<TR><TD>
							<table width="100%" align="center" border="0" cellpadding="0" cellspacing="1" height="100%">
								<tr bgcolor="000000">
									<tH COLSPAN=2>Select Uploaded Date Period</TH>
								</tr>
								<%
								times = 0
								next_date = DateAdd("m", -6, end_date)
								Do While DateValue(next_date) < DateValue(end_date)
									If times mod 2 = 0 Then 
										If bgc = bgcone Then 
											bgc = bgctwo
										Else
											bgc = bgcone
										End If
										If Times <> 0 Then
											Response.Write "</tr>"
										End If
										Response.Write "<tr bgcolor=" & bgc & ">"
									End If
									Response.Write "<td align=center width=50% valign=top><a href=""default.asp?d1=" & next_date & "&d2=" & DateAdd("d", 10, next_date) & """>"
									Response.Write format_date(next_date) & " - " & format_date(DateAdd("d", 10, next_date)) & "</a></td>"
									next_date = DateAdd("d", 10, next_date)
										
									times = times + 1
								Loop
								If times Mod 2 <> 0 Then 
									Response.Write "</TD><td>&nbsp;"
								End If
								%></td></TR>
							</table></td>
							</tr></table>				
						</td>
					</tr>
				</table>
		
		<br><br>
<% End If %>
<table width="760" border="0" cellpadding="0" cellspacing="0" align="center" BGCOLOR=#444444>
<TR><TD>
<table width="100%" border="0" cellpadding="2" cellspacing="1" align="center">
		<% If Len(frm_d1) > 0 Then %>
		<tr>
			<td colspan=9 BGCOLOR=#000000 align="center"><b>Showing Demos Between: <%=d1%> - <%=d2%></b></td>
		</tr>
		<% End If %>
		<tr>
			<tH BGCOLOR=#000000>UL Date</td>
			<tH BGCOLOR=#000000>Ladder</td>
			<tH BGCOLOR=#000000>Team 1</td>
			<tH BGCOLOR=#000000>Team 2</td>
			<tH BGCOLOR=#000000>POV</td>
			<tH BGCOLOR=#000000>Map(s)</td>
			<tH BGCOLOR=#000000>DL</td>
			<tH BGCOLOR=#000000>More</td>
		</tr>
		<%
			strSQL = "EXECUTE ViewDemos "
			If Len(Session("PlayerID")) > 0 Then
				strSQL = strSQL & " @PlayerID = '" & Session("PlayerID") & "'"
			Else
				strSQL = strSQL & " @PlayerID = NULL"
			End If
			If Len(d1) > 0 And IsDate(D1) Then
				strSQL = strSQL & ", @FromDate = '" & d1 & "'"
			Else
				d1 = ""
			End If
			If Len(d2) > 0 And IsDate(D2)  Then
				strSQL = strSQL & ", @ToDate = '" & d2 & "'"
			Else
				d2 = ""
			End If
			If Len(frm_teamname) > 0  Then
				strSQL = strSQL & ", @Teamname = '" & frm_teamname & "'"
			End If
			If Len(frm_MapName) > 0  Then
				strSQL = strSQL & ", @MapName = '" & frm_MapName & "'"
			End If
			If Len(frm_position) > 0  Then
'				strSQL = strSQL & ", @PositionPlayed = '" & frm_position & "'"
			End If
			If Len(frm_playername) > 0  Then
				strSQL = strSQL & ", @PlayerHandle = '" & frm_playername & "'"
			End If
			If Len(qry_historyID) > 0 Then
				strSQL = strSQL & ", @ShowHistoryID = '" & qry_historyID & "'"
			End If
			strSQL = strSQL & ", @PublicFlag = '" & intPublicFlag & "'"
			strSQL = strSQL & ", @AdminFlag = '" & intAdminFlag & "'"
			strSQL = strSQL & ", @SortCriteria = '" & OrderBySQL & "'"
			oRS.Open strSQL, oConn
			If Not(Ors.Eof and Ors.BOF) Then
				Do While Not oRS.EOF
					MatchWinnerID = oRS("MatchWinnerID")
					MatchWinner = oRS("WinnerName")
					MatchLoserID = oRS("MatchLoserID")
					MatchLoser = oRS("LoserName")
					MatchDate = oRS("MatchDate")
					UploadDate = oRS("upload_dtim")
					PlayerName = oRS("PlayerHandle")
					LadderName = oRS("LadderName")
					If numd mod 2 = 0 Then
						bgc = bgcone
					Else
						bgc = bgctwo
					End If
					%>
					<tr bgcolor="<%=bgc%>">
						<td height="30" valign="middle"><%=format_date(DateValue(UploadDate))%></td>
						<td height="30" valign="middle"><a href="/viewladder.asp?ladder=<%=Server.URLEncode(LadderName & "")%>"><%=LadderName%></a></td>
						<td height="30" valign="middle"><a href="/viewteam.asp?team=<%=Server.URLEncode(MatchWinner & "")%>"><%=MatchWinner%></a></td>
						<td height="30" valign="middle"><a href="/viewteam.asp?team=<%=Server.URLEncode(MatchLoser & "")%>"><%=MatchLoser%></a></td>
						<td align="center" height="30" valign="middle"><a href="/viewplayer.asp?player=<%=Server.URLEncode(PlayerName & "")%>"><%=PlayerName%></td>
						<td align="center" height="30" valign="middle"><%=oRS("MapPlayed")%></td>
						<td align="center" height="30" valign="middle"><%=oRS("DownloadCount")%></td>
						<td align="center" valign="middle"><a href="viewDemo.asp?DemoID=<%=oRS("DemoID")%>"><img src="arrow.gif" border="0" WIDTH="16" HEIGHT="15"></a></td>
					</tr>
					<%
					numd = numd + 1
					oRS.MoveNext
				Loop
			Else
				%>
					<tr bgcolor="<%=bgc%>">
						<TD COLSPAN=9><p class=small><I>No demos found.</I></TD>
					</TR>
				<%
			End If
			oRS.Close
		%>
	 </table>
	</td></tr>
	</table>            
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing

Function format_date(d)
	Dim new_date
	Dim new_day, new_month, new_year
		
	new_day = day(d)
	new_month = month(d)
	new_year = year(d)
		
	If len(day(d)) = 1 Then new_day = "0" & day(d)
	If len(month(d)) = 1 Then new_month = "0" & month(d)
	If len(year(d)) = 4 then new_year = right(year(d), 2)	
		
	new_date = new_month & "/" & new_day & "/" & new_year
	format_date = new_date
End Function
%>