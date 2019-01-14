<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Rules"

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

Dim RuleName
rulename = Request.QueryString ("set")

Dim CurrentChapter, quesSQL, intGeneralRuleID
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart(RuleName)

strsql = "select c.rulename as Chapter, c.fldauto, c.GeneralRuleID "
strsql = strsql & " from tbl_chapter c "
strsql = strsql & " where c.rulename = '" & CheckString(RuleName) & "' and c.isactive = 1 "
strsql = strsql & " ORDER BY C.orderingField "
'Response.Write strsql
%>
    <table width="90%" border="0">
<tr><td>
<%
ors.open strsql, oconn
if ors.eof and ors.bof then
	ors.close 
	Response.Write "Unable to find requested rule set."
else
	do while not(ors.eof)
		intGeneralRuleID = ors.fields("GeneralRuleID").value
		CurrentChapter = ors.fields("Chapter").value
		Response.Write "<P class=small><B><center>" & CurrentChapter & "</center></b></p>"
		quessql = "select q.rulename as question, q.answer from tbl_question q "
		quesSQL = quesSQL & "where q.chapter_fldauto = '" & ors.fields("fldauto").value & "' order by q.orderingfield"
'		Response.Write quessql
		ors2.open quesSQL, oconn
		if not(ors2.eof and ors2.bof) then
			do while not(ors2.eof)
				Response.Write "<P class=small><B>" & ors2.fields("Question").value & "</B></P>"
				Response.Write "<p class=small>" & Replace(ors2("answer").Value & "",vbCrlf,"<br>") & "</P>"
				ors2.movenext
			loop
		end if
		ors2.nextrecordset
		ors.movenext
	loop
	ors.nextrecordset

		strsql = "select c.rulename as Chapter, c.fldauto "
		strsql = strsql & " from tbl_chapter c "
		strsql = strsql & " where c.fldauto = '" & intGeneralRuleID & "' and c.isactive = 1"
		strsql = strsql & " ORDER BY C.orderingField "
		ors.open strsql, oconn
		if not(ors.eof and ors.bof) then
			do while not(ors.eof)
				CurrentChapter = ors.fields("Chapter").value
				Response.Write "<P class=small><B><center>" & CurrentChapter & "</center></b></p>"
				quessql = "select q.rulename as question, q.answer from tbl_question q "
				quesSQL = quesSQL & "where q.chapter_fldauto = '" & ors.fields("fldauto").value & "' and q.isactive = 1"
				ors2.open quesSQL, oconn
				if not(ors2.eof and ors2.bof) then
					do while not(ors2.eof)
						Response.Write "<P class=small><B>" & ors2.fields("Question").value & "</B></P>"
						Response.Write "<p class=small>" & Replace(ors2("answer").Value,vbCrlf,"<br>") & "</P>"
						ors2.movenext
					loop
				end if
				ors2.nextrecordset
				ors.movenext
			loop
		end if
end if
%>
</td>
</tr>
</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
%>

