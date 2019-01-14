<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Map Edit"

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

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim mName, mLadder, mAbbr, mType, mTerr, mGen, mInv, mTurr, mVPad, mDesc, mImg, isYes, isNo	
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Edit Map") %>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2>
		
<%
	strSQL = "select * from tbl_maps where mapid=" & Request.QueryString("mapid")
	ors.Open strSQL, oconn
	mname=""
	mladder=""
	mabbr=""
	mtype=""
	mterr=""
	mgen=""
	minv=""
	mturr=""
	mvpad=0
	mdesc=""
	mimg=""
	if not (ors.EOF and ors.BOF) then
 		mname=ors.Fields("MapName").Value
		mabbr=ors.Fields("MapAbbreviation").Value
		mtype=ors.Fields("MapType").Value
		mterr=ors.Fields("MapTerrain").Value
		mgen=ors.Fields("Generators").Value
		minv=ors.Fields("Inventories").Value
		mturr=ors.Fields("BaseTurrets").Value
		mvpad=ors.Fields("VehiclePad").Value
		mdesc=ors.Fields("Description").Value
		mimg=ors.Fields("MapImage").Value
	end if
	if mvpad=1 then
		isyes="checked"
		isno=""
	else
		isno="checked"
		isyes=""
	end if
	Response.Write "<form name=newmap action=saveitem.asp method=post>"
	Response.Write "<tr bgcolor=" & bgcone & "><td width=150 align=right><p class=small>Map Name:</p></td><td><input name=mapname class=bright type=text value=""" & Server.HTMLEncode(mname) & """></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & "><td width=150 align=right><p class=small>Abbreviation:</p></td><td><input name=mapabbr class=bright type=text value=""" & Server.HTMLEncode(mabbr) & """></td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & "><td width=150 align=right><p class=small>Type (CTF, C&H, ...):</p></td><td><input name=maptype class=bright type=text value=""" & mtype & """></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & "><td width=150 align=right><p class=small>Terrain Type:</p></td><td><input name=mapterr class=bright type=text value=""" & mterr & """></td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & "><td width=150 align=right><p class=small>Generators (Total):</p></td><td><input name=mapgens class=bright type=text value=""" & mgen & """></td></tr>"		
	Response.Write "<tr bgcolor=" & bgctwo & "><td width=150 align=right><p class=small>Inventories (Total):</p></td><td><input name=mapInv class=bright type=text value=""" & minv & """></td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & "><td width=150 align=right><p class=small>Base Turrets (Total):</p></td><td><input name=mapturr class=bright type=text value=""" & mturr & """></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & "><td rowspan=2 valign=top align=right><p class=small>Vehicle Pad?:</p></td><td><p class=small><input name=mapV type=radio class=borderless value=1 " & isyes & ">Yes</p></td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & "><td><p class=small><p class=small><input name=mapV type=radio class=borderless value=0 " & isno & ">No</p></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & "><td valign=top align=right width=150><p class=small>Description:</p></td><td><textarea name=mapdesc rows=5 cols=35>" & mdesc & "</textarea></td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & "><td width=150 align=right><p class=small>Image File Name:</p></td><td><input name=mapimage class=bright type=text value=""" & Server.HTMLEncode(mimg) & """></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & "><td colspan=2 align=center><input type=hidden name=SaveType value=SaveMap><input type=hidden name=mapid value=" & Request.QueryString("mapid") & "><input type=submit class=bright value='Commit Map Edit' id=submit1 name=submit1></td></tr></form>"
	%>
	</TABLE>
	</TD></TR></TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

