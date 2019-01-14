<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Map Guide"

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

Dim MapName, Ladder, Meth, DisplayName, MapID, lName, lReset
Dim rsPrimary, rsSecondary, intFirstPub
Dim urlIMG, prevmap, nextmap
Dim hasV, mName, mAbbr, mType, mterr, mgen, minv, mdesc, mimage, mturr
Dim mSmImage
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<!-- #Include virtual="/include/dynamicselect.asp" -->
<% 
		Mapname=Request("Mapname")
		Ladder=Request("MapType")
		meth=Request("meth")

		if meth="back" then
			session("map")=request("MapName")
		end if
		displayname=session("map") & " -- "
		lReset=request("lReset")
		if lReset="true" or (ladder="" and mapid="") then
			session("map")= mapname
			displayname=""
			mapid=""
			
		end if
		if ladder="" then ladder=0
		displayname=""
		lName=""
		strSQL="select laddername from tbl_ladders where ladderid=" & ladder
		ors.Open strSQL, oconn
		if not (ors.EOF and ors.BOF) then
			lname=ors.Fields(0).Value 
		end if
		ors.Close
		if Request.Form("submit1") <> "" then
			strSQL="Select top 1 MapName from tbl_maps inner join lnk_L_M on tbl_maps.mapid=lnk_L_M.mapid where ladderid=" & ladder & " order by mapname"
			ors.Open strSQL, oconn
			if not (ors.EOF and ors.BOF) then
				displayname=ors.Fields(0).Value & " -- "
			end if
			ors.Close
		end if
		Call ContentStart(Server.HTMLEncode(DisplayName) & " Map Guide")
		
	SET rsPrimary 	= Server.CreateObject("ADODB.RecordSet")
	SET rsSecondary	= Server.CreateObject("ADODB.RecordSet")
	strSQL= "select LadderName, ladderid from tbl_Ladders where ladderactive = 1 order by LadderName"
	rsPrimary.Open strSQL, oconn
	IF NOT (rsPrimary.BOF AND rsPrimary.EOF) THEN
	   intFirstPub = ladder
	ELSE
	intFirstPub = 0
	END IF
	rsSecondary.Open "select distinct MapName, tbl_maps.mapid, ladderid from tbl_maps inner join lnk_L_M on tbl_maps.MapID=lnk_L_M.mapid order by MapName", oConn 

	Call FillArray ("arrMaps", rsSecondary, "Ladderid", "mapname", "mapname" )
	%>
	<table width=90% border=0 align=center>
		<form name=25 action=mapguide.asp id=25 method=get>
		<tr><td valign=center align=right><p class=small><b>Ladder: </b></p></td><td align=left>
			<%
				Call SelectBox (rsPrimary, "maptype", "Ladderid", "laddername", "MapName", "arrMaps", lName)
			%>
			</td>
		<td valign=center align=right><p class=small><b>Map: </b></p></td><td align=left>
			<%
				rsSecondary.Filter = "Ladderid = '" & intFirstPub & "'"
				Call SelectBox (rsSecondary, "MapName", "MapName", "MapName", "", " " , mapname)
				rsSecondary.Filter = 0
			%>
			</select>
			</td>
			<td align=center><input name=lReset value=true type=hidden><input type=submit name=submit1 value="Go" align=bottom class=bright></td></tr>
		</form> 
	</table>
    <%
	if Ladder <> 0 and mapname <> "0" and mapname <> "" then
		'strSQL="Select * from tbl_maps inner join lnk_L_M on tbl_maps.mapid=lnk_L_M.mapid where Ladderid=" & ladder & " order by MapName"
		strsql = "select MapName, MapAbbreviation, MapType, MapTerrain, Generators, Inventories, BaseTurrets, VehiclePad, Description, MapImage, tbl_maps.MapID "
		strsql = strsql & "from tbl_maps inner join lnk_L_M on tbl_maps.mapid=lnk_L_M.mapid where Ladderid=" & ladder & " order by MapName"
		ors2.Open strSQL, oconn
		if not (ors2.EOF and ors2.BOF) then
			if session("map")="" then
				session("map")=ors2.Fields(0).Value
			end if
			prevmap=""
			nextmap=""
			do while not ors2.EOF
				
				if session("map")=ors2.Fields(0).Value then
					if ors2.Fields(7).Value = 1 then
						HasV="Yes"
					else
						HasV="No"
					end if
					mname=ors2.Fields(0).Value 
					mabbr=ors2.Fields(1).Value
					mtype=ors2.Fields(2).Value
					mterr=ors2.Fields(3).Value
					mgen=ors2.Fields(4).Value
					minv=ors2.Fields(5).Value
					mturr=ors2.Fields(6).Value
					mDesc=ors2.Fields(8).Value
					mapid=ors2.Fields(10).Value 
					mimage=Trim(lCase(ors2.Fields("MapImage").Value ))
					ors2.Movenext 
					If Not(oRs2.EOF) Then
						nextmap=ors2.Fields(0).Value
					End If
					If Len(trim(mImage)) = 0 Then 
						mImage = "none.gif"					
					End If
					If instr(1, mimage, "_") > 0 Then 
						mSmImage = left(mimage, instr(1, mimage, "_")-1)
					End If
				end if
				if nextmap="" then
					If Not(oRS2.EOF) Then
						prevmap=ors2.Fields(0).Value 
					End If
				end if
				If Not(oRS2.EOF) Then 
					ors2.MoveNext 
				End If
				
			loop
		end if
'		Response.write mImage & "--"
'		Response.End
		if nextmap <> "" then
			session("map")=nextmap
		else
			session("map")=prevmap
		end if
		displayname=mname
		%>
		<table width=90% border=0><tr><td align=left>
		<%
		if prevmap <> "" then
			%>
			<a href=mapguide.asp?maptype=<%=server.urlencode(ladder)%>&mapname=<%=server.urlencode(prevmap)%>&meth=back><< Previous (<%=Server.HTMLEncode(prevmap)%>)</a>
			<%
		end if
		if nextmap <> "" then
			%>
			</td><td align=right><a href=mapguide.asp?maptype=<%=server.urlencode(ladder)%>&mapname=<%=server.urlencode(nextmap)%>&meth=next>(<%=Server.HTMLEncode(nextmap)%>) Next >></a></td>
			<%
		end if
	%>
		</tr></table>
	<table width=580 border=0 cellspacing=0 cellpadding=0>
		<tr><td><img src="/images/spacer.gif" width="1" height="7"></td></tr>
		<tr height=20><td align=center colspan=3 CLASS="headline"><b><%=Server.HTMLEncode(mname)%></b></td></tr>
		<tr><td><img src="/images/spacer.gif" width="1" height="7"></td></tr>
		<%
		if mimage = "none.gif" then
			Response.Write "<tr height=1><td><img src=/images/spacer.gif width=100 height=1></td><td width=100><img src=/images/spacer.gif width=80 height=1></td><td rowspan=15 width=370 bgcolor=" & bgcone & "><img src=/Maps/" & mimage & "></a></td></tr>"
		else
				%>
				
				<tr height=1><td><img src=/images/spacer.gif width=100 height=1></td><td width=100><img src=/images/spacer.gif width=80 height=1></td>
				<td rowspan=15 width=370 bgcolor=<%=bgcone%>>
<script language=javascript>
	var urlimg = 'no where';
	urlimg = 'mapdisplay.asp?mapname=<%=server.urlencode(mname)%>&imgurl=maps/<%=mSmImage%>_640x480.jpg';
</script>

				<p class=small><a href="javascript:popup(urlimg, 'MapPicture', 490, 650, 'no')"><img src="/Maps/<%=mimage%>" border=0></a></td></tr>
				<%
		end if	
		%>
		<tr><td align=center colspan=2><p class=headline><b><font color=#ffcf3f>Abbreviation</font></b></p></td></tr>
		<tr><td align=center colspan=2><p class=small><b><%=Server.HTMLEncode(mabbr)%><b></p></td></tr>
		<tr><td align=center colspan=2><p class=headline><b><font color=#ffcf3f>Type</font><b></p></td></tr>
		<tr><td align=center colspan=2><p class=small><b><%=mtype%></b></p></td></tr>
		<tr><td align=center colspan=2><p class=headline><b><font color=#ffcf3f>Terrain</font></b></p></td></tr>
		<tr><td align=center colspan=2><p class=small><b><%=mterr%></b></p></td></tr>
		<tr><td align=center colspan=2><p class=headline><font color=#ffcf3f><b>Generators</b></font></p></td></tr>
		<tr><td align=center colspan=2><p class=small><b><%=mgen%></b></p></td></tr>
		<tr><td align=center colspan=2><p class=headline><b><font color=#ffcf3f>Inventories</font></b></p></td></tr>
		<tr><td align=center colspan=2><p class=small><b><%=minv%></b></p></td></tr>
		<tr><td align=center colspan=2><p class=headline><font color=#ffcf3f><b>Base Turrets</b></font></p></td></tr>
		<tr><td align=center colspan=2><p class=small><b><%=mturr%></b></p></td></tr>
		<tr><td align=center colspan=2><p class=headline><font color=#ffcf3f><b>Vehicle Pad</b></font></p></td></tr>
		<tr><td align=center colspan=2><p class=small><b><%=HasV%></b></p></td></tr>
		<tr><td colspan=3><img src="/images/spacer.gif" width="1" height="5"></td></tr>
		<tr height=22><td align=left colspan=2><p class=headline><font color=#ffcf3f>Description</font></p></td></tr>
		<tr><td colspan=3><img src="/images/spacer.gif" width="1" height="5"></td></tr>
		<tr><td align=left colspan=3><p class=small>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=mdesc%></p></td></tr>

	</table>
	<%
	end if
	if mname <> "" then
		if bSysadmin or bAnyLadderAdmin then
			%>
			<p align=center><a href=MapEdit.asp?MapID=<%=mapid%>>Edit <%=Server.htmlencode(displayname)%></a></p>
			<%
		end if
	end if
	%>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
Set rsPrimary = Nothing
Set rsSecondary = Nothing
%>

