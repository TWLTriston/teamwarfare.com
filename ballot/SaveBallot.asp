<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Save Ballot"

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

Dim bID, i , j, qID
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<%
if not(bSysAdmin) then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"
end if 
	
strsql="insert into tbl_ballot(BName, Type, LadderType, QCount, description, IsActive, LadderID) values ('" & CheckString(Request("btitle")) & "','" & CheckString(Request("btype")) & "','" & Left(Request.Form("bLadder"), 1) & "', " & CheckString(request("numqs")) & ",'" & CheckString(Request("bDesc")) & "',1, " & Right(Request.Form("bLadder"), Len(Request.Form("bLadder")) - 1) & ")"
Response.Write strsql & "<BR>"
oconn.Execute strsql

strsql="select top 1 ballotid from tbl_ballot order by ballotid desc"
Response.Write strsql & "<BR>"
ors.Open strsql, oconn

bid=ors.Fields(0).Value 
ors.Close 
for i = 1 to request("numqs")
	strsql="insert into tbl_Questions(BallotID, QuestionNum, Question) values (" & bid & "," & i & ",'" & CheckString(request("q_" & i)) & "')"
	'Response.Write strsql & "<BR>"
	oconn.Execute strsql
	for j=1 to 5
		strsql="select top 1 qid from tbl_questions order by qid desc"
		'Response.Write strsql & "<BR>"
		ors.Open strsql, oconn
		qid=ors.Fields(0).Value 
		ors.Close 
		if trim(request("q_" & i & "_a_" & j))<> "" then
			strsql="insert into tbl_responses(QID, RVal, RText) values (" & qid & "," & j & ",'" & CheckString(request("q_" & i & "_a_" & j)) & "')"
			'Response.Write strsql & "<BR>"
			oconn.Execute strsql
		end if
	next
next
oConn.Close
Set oCOnn = Nothing
Set oRs = Nothing
Response.Clear
Response.Redirect "default.asp"

Response.End 
%>
