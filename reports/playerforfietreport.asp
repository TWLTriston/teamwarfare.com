<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "Player Ladder Forfeit Report"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim strPlayerName, intCount
intCount = 0
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Player Ladder Forfeit Report")

strSQL = "SELECT PlayerLadderName from tbl_playerLadders order by PlayerLadderName"
		
ors.open strsql, oconn
if not(ors.eof and ors.bof) then
	Response.Write "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 ALIGN=CENTER>"
	do while not(ors.eof)
		Response.Write "<TR>"
		Response.Write "<TD ALIGN=CENTER><p class=small><B><a href=""playerforfietreport.asp?ladderName=" & Server.URLEncode(oRS("PlayerLadderName")) & """>" & oRS("PlayerLadderName") & "</A></B></TD>"
		Response.Write "</TR>"
		ors.movenext
	loop
	Response.write "</TABLE>"
end if
ors.close

Call ContentEnd()
If Request("LadderName") <> "" Then 
	Call ContentStart(Request.QueryString("LadderName") & " Forfeit Report")
		strSQL = "EXEC 	sp_PlayerForfiet '" & Replace(Request("LadderName"), "'", "''") & "'"
		oRS.Open strSQL, oConn
		If Not(oRS.eof and oRS.bof) Then
			Do While Not(oRS.EOF)
				intCount = intCount + 1
				If strPlayerName <> oRS("PlayerHandle") Then
					strPlayerName = oRS("PlayerHandle")
					If intCount <> 1 Then
						Response.Write "</TABLE></TD></TR></TABLE><BR><BR>"
					End If
					Response.Write "<TABLE BORDER=0 WIDTH=50% CELLSPACING=0 CELLPADDING=0 BGCOLOR=""#444444"" ALIGN=CENTER><TR><TD>"
					Response.Write "<TABLE BORDER=0 WIDTH=100% CELLSPACING=1 CELLPADDING=2 ALIGN=CENTER>"
					Response.Write "<TR BGCOLOR=""#000000"">"
					Response.Write "<TH WIDTH=""50%""><A HREF=""/viewplayer.asp?player=" & server.URLEncode(strPlayerName) & """>" & strPlayerName & "</A></TH>"
					If oRs("LastLogin") <> "" Then
						Response.Write "<TH ALIGN=RIGHT WIDTH=""50%"">Last Login: " & FormatDateTime(oRS("LastLogin"), vbShortDate) & "</TH></TR>"
					Else
						Response.Write "<TH ALIGN=CENTER WIDTH=""50%""><FONT STYLE=""color:FF0000;""><B>Never Logged In</B></FONT></TH></TR>"
					End If
'					Response.write "<TR BGCOLOR=" & bgcOne & "><TD WIDTH=""50%"">" & oRS("Total_Forfiets") & " Forfeits</TD>"
					Response.Write "<TR BGCOLOR=""#000000""><TH WIDTH=""50%"">Loss To</TH><TH WIDTH=""50%"">Loss Date</TH></TR>"
					bgc = bgctwo
				End If
				Response.write "<TR BGCOLOR=" & bgc & ">"
				Response.write "<TD WIDTH=""50%""><a href=""/viewplayer.asp?player=" & Server.URLEncode(oRS("Opponent")) & """>" & oRS("Opponent") & "</A></TD>"
				Response.Write "<TD ALIGN=RIGHT WIDTH=""50%"">" & FormatDateTime(oRS("MatchDate"), vbShortDate) & "</TD>"
				If bgc = bgcone then
					bgc = bgctwo
				else
					bgc = bgcone
				end if
				oRS.MoveNext
			Loop		
			Response.Write "</TABLE></TD></TR></TABLE>"
		End If
		oRS.Close	
	Call ContentEnd()
End If
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>