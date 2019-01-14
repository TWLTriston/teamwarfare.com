<%' Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & verbage & " Reply"

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
%>
<!-- #INCLUDE virtual="/include/i_funclib.asp" -->
<%
if not(bsysadmin or banyladderadmin) then
	' Require login to perform action.
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=3"
end if

Response.Write "<font color=#ffffff>Querystring: " & Request.QueryString
Response.Write "<br>Form data: " & Request.Form & "</font></br>"

if Request.Querystring("clearDate")="true" then
	strsql="UPDATE tbl_matches SET" &_ 
	" MatchDate='TBD', MatchSelDate1='TBD', MatchSelDate2='TBD', MatchAcceptanceDate='', MatchLockDate=''," &_
	" MatchMap1ID='TBD', MatchMap2ID='TBD', MatchMap3ID='TBD'" &_
	" WHERE MatchID='" & request.Querystring("matchid") & "'"
	oconn.execute (strsql)
	strsql = "delete from tbl_disp_pending where mmID='" & Request.Form("matchid") & "'"
	oconn.execute (strsql)
End if

if Request.Form("saveType") = "changeDate" then
	DateString = Request.form("newMonth") & "/" & Request.form("newDay") & "/" & Request.form("newYear")
	TimeString = Request.Form("newHour") & ":" & Request.Form("newMinute") & ":00 PM" 
	longdate = formatDatetime(cDate(DateString),1)
	datearray = split(longdate,",")
	finaldate = DateArray(0) & ", " & datearray(1)
	timearray = split(cdate(timestring), ":")
	strTimeZone  = Request.Form("timezone")
	If Len(strTimeZone) = 0 Then
		strTimeZone = "EST"
	End If
	finaltime = timearray(0) & ":" & timearray(1) & " PM " & strTimeZone	
	dispTimeString = timearray(0) & ":" & timearray(1) & ":00" 
	strSQL="UPDATE tbl_matches SET" &_
	" MatchDate='" & finaldate & " " & FinalTime & "'" &_
	" WHERE MatchID='" & Request.Form("matchid") & "'"
	'Response.Write strsql &"<BR>"
	oconn.execute (strsql)
	strsql = "update tbl_disp_pending set " &_
	" MDate='" & cDate(DateString) & "', mTime='" & dispTimeString & "'" &_
	" WHERE mmID='" & Request.Form("matchid") & "'"
'	Response.Write strsql
	oconn.execute (strsql)
End If

'-----------------------------------------------
' Housekeeping
'-----------------------------------------------
oConn.Close 
set ors = nothing
set oConn = nothing	
set ors2 = nothing	
Response.Clear
Response.Redirect "/adminmenu.asp"
Response.End
%>	
