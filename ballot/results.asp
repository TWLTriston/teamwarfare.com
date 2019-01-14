<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Ballot Results"

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

Dim qRS, rRS, vRS, q, BallotName, intBallotID 

intBallotID = Request.QueryString("BallotID")
If Not(IsNumeric(intBallotID)) Then
	Response.Clear
	Response.Redirect "/ballot/"
End If
Dim responses(6)
Dim rvalue(6)
Dim i, pct, totalresponses
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("View Ballot Results")
strSQL = "SELECT ballotid, BName, type, description, qCount FROM tbl_ballot WHERE IsActive=1 ORDER BY bName"
oRS.Open strSQL, oConn
If Not(oRs.EOF AND oRS.BOF) Then
	%>
	<TABLE BORDER=0 CElLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" width="97%">
	<TR><TD>
	<TABLE BORDER=0 CElLSPACING=1 CELLPADDING=2 width="100%">
	<TR BGCOLOR="#000000">
		<TH COLSPAN=3>Choose a ballot</TH>
	</TR>
	<TR BGCOLOR="#000000">
		<TH>Ballot Name</TH>
		<TH>Vote Type</TH>
		<TH>Questions</TH>
	</TR>
	<%
	do while not (ors.EOF)
		If bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		Response.Write "<TR BGCOLOR=" & bgc & ">"
		Response.Write "<TD><A HREF=""results.asp?ballotid=" & oRS("BallotID") & """>" & oRS("bName") & "</A></TD>"
		Response.Write "<TD ALIGN=CENTER>"
		If oRS.Fields("type").Value = 0 Then
			Response.Write "Global"
		Else
			Response.Write "Founder's only"
		End If
		Response.Write "</TD>"
		Response.Write "<TD ALIGN=RIGHT>" & ors.Fields("qCount").Value & "</TD>"
		Response.Write "</TR>"
		Response.Write "<TR BGCOLOR=#000000><TD COLSPAN=3><FONT COLOR=""#888888"">&nbsp;&nbsp;&nbsp;&nbsp;" & oRs.Fields("Description").Value & "</FONT></TD></TR>"
		ors.movenext
	loop
	%>
	</TABLE>
	</TD></TR>
	</TABLE>
	<%
End If
oRS.NextRecordSet
Call ContentEnd()

set qrs = server.createobject("ADODB.RecordSet")
set rrs = server.createobject("ADODB.RecordSet")
set vrs = server.createobject("ADODB.RecordSet")

If Len(intBallotID) > 0 Then
	strsql="select b.bname, q.qid, q.questionnum, q.question "
	strsql = strsql & "from tbl_ballot b, tbl_questions q "
	strsql = strsql & "where b.ballotid = '" & intBallotID  & "' " 
	strsql = strsql & " AND q.ballotid = b.ballotid "
	strsql = strsql & "order by questionnum"
	ors.Open strsql, oconn
	q=0
	if not (ors.EOF and ors.BOF) then
		BallotName = ors.fields("bname").value
	end if
	
	
	Call ContentStart(BallotName & " Results")
	
	if not(ors.eof and ors.bof) then
		Response.Write "<table border=0 cellspacing=0 cellpadding=0 BGCOLOR=""#444444"" width=""760"" ><TR><TD>"
		Response.Write "<table border=0 cellspacing=1 cellpadding=2 width=""100%"" >"
		do while not ors.eof 
			for i = 1 to 5
				responses(i) = 0
				rvalue(i) = ""
			next
			Response.Write "<TR bgcolor=#000000><TD colspan=3><B>Question " & ors.Fields("questionNum").Value & ": " & ors.Fields("question").Value & "</b></TD></TR>"
			response.Write "<TR bgcolor=" & bgctwo & ">"
			response.Write "<TD><B>Response</B></TD>"
			response.Write "<TD><B>Votes PCT</B></TD>"
			response.Write "<TD><B>Total Votes</B></TD></TR>"
			strsql= "select r.rtext, r.rval, cnt = count (v.choice) "
			strsql= strsql & " from tbl_responses r, tbl_votes v "
			strsql= strsql & " where r.qid = " & ors.fields("qid").value
			strsql= strsql & " AND v.qid = " & ors.fields("qid").value
			strsql= strsql & " AND v.choice =* r.rval "
			strsql= strsql & " Group BY r.rtext, r.rval "
			strsql= strsql & " order by r.rval "
			rrs.Open strSQL, oconn
			TotalResponses = 0
			if not (rrs.EOF and rrs.BOF) then
				do while not rrs.EOF
					TotalResponses = TotalResponses + rrs.Fields("cnt").value
					responses(rrs.Fields("rval").value) = rrs.Fields("cnt").value
					RValue(rrs.Fields("rval").value) = rrs.Fields("rtext").value
					rrs.MoveNext
				loop
				for i = 1 to 5
					if rvalue(i) <> "" then
						Response.Write "<TR BGCOLOR=" & bgcOne & "><TD>"
						Response.Write rvalue(i) & "</TD>"
						if TotalResponses <> 0 then
							pct = responses(i) / TotalResponses * 100
						else
							pct = 0
						end if
						Response.Write "<TD BGCOLOR=#000000><img src=""/ballot/images/bar.gif"" height=10 width=" & (fix(pct) + 1) * 4 & "> &nbsp;&nbsp;" & formatnumber(pct,2,-1) & "%</TD>"
						response.Write "<TD>" & responses(i) & "</TD>"
						response.Write "</TR>"
					end if
				next
				response.write "<TR BGCOLOR=#000000><TD colspan=2 align=right><B>Total:</B></TD><TD><B>" & TotalResponses & "</B></TD></TR>"
	'				Response.Write "<li><b>" & vrs.fields(0).value & "</b> Vote(s) for <b>" & rrs.Fields(2).Value & "</b>: " & rrs.Fields(3).Value 
			end if
			rrs.nextrecordset 
					
			ors.MoveNext
		loop
		Response.Write "</table></TD></TR></TABLE><BR><BR>"
	end if
	ors.close
	Call ContentEnd()
End If
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>