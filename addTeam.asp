<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add a Team"

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

If Not(Session("LoggedIn")) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"
End If

Dim bIsEdit, strTeamName, strMethod, strVerbage
Dim ErrorCode
Dim tOwnerID, tName, tID, tTags, tURL, tEmail, tJoinPass
Dim tConfirmJoinPass, tDesc, tIRC, tIRCServer, tLogoURL

errorcode = request.querystring("error")

bIsEdit = cBool(Request.QueryString("IsEdit"))
If bIsEdit And Session("LoggedIn") Then
	strTeamName = Request.QueryString ("Team")
	If Not(bSysAdmin Or IsTeamFounder(strTeamName)) Then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = "select * from tbl_teams where TeamName='" & CheckString(strTeamName) & "'"
	oRS.Open strSQL, oConn
	If Not(oRs.EOF AND oRS.BOF) Then
			tOwnerID = ors.Fields(5).Value 
			tName = ors.Fields(1).Value
			tID = ors.Fields(0).Value 
			ttags =  ors.Fields(2).Value
			turl = ors.Fields(3).Value
			temail = ors.Fields(4).Value 
			tjoinpass= ors.fields(6).value
			tconfirmjoinpass= ors.fields(6).value
			tDesc=ors.fields(9).value
			tirc=ors.fields("TeamIRC").value
			tircserver=ors.fields("TeamIRCServer").value
			tlogourl=ors.fields("TeamLogoURL").value
	End If
	oRS.Close
	strMethod = "Edit"
	strVerbage = "Edit Team Profile"	
Else 
	strMethod = "New"
	strVerbage = "Create a Team"
End If

strPageTitle = "TWL: " & strVerbage
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart(strVerbage) %>
<form name=frmAddTeam action=saveItem.asp method=post> 
	<table border=0 align=center width="760" cellspacing=0 CELLPADDING=0 BGCOLOR="#444444">
	<TR><TD>
	<table border=0 align=center width="100%" cellspacing=1 cellpadding=2>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	if errorcode=1 then
		response.write "<tr height=20 bgcolor=" & bgc & "><td colspan=2 align=center><b><font color=red>Error: That team name is already in use.</font></b></td></tr>"
		if bgc=bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
	end if
	if errorcode=4 then
		response.write "<tr height=20 bgcolor=" & bgc & "><td colspan=2 align=center><b><font color=red>Error: Can't change team name, that team name is already in use.</font></b></td></tr>"
		if bgc=bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
	end if
	%>
    <tr bgcolor=<%=bgc%> valign=center height="30">
		<td align="right">Team Name:</td>
		<td> &nbsp;
		<% 
		if bIsEdit AND Not(bSysAdmin) then
			response.write tname
		   %>
		   <input type=hidden id=text1 name=TeamName value="<%=Server.HTMLEncode(tname)%>">
		   <%
		else
			 %>
			<input id=text1 name=TeamName maxlength="50" style=" WIDTH: 300px" class=text value="<%=Server.HTMLEncode(tname)%>">
		  	<% If bSysAdmin Then %>
		  	<input type="hidden" name="teamid" id="teamid" value="<%=tid%>" />
		  	<input type="hidden" id=OldTeamName name=OldTeamName value="<%=Server.HTMLEncode(tname)%>" />
		  	<% End If %>
			 <%
		end if
		%>
		</td>
	</tr>
	<%	
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	if errorcode=3 then
		response.write "<tr height=20 bgcolor=" & bgc & "><td colspan=2 align=center><b><font color=#FFD142>Error: Team names cannot contain special characters: ? & = % &gt &lt</font></b></td></tr>"
		if bgc=bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>Team names should be original and unique. Maximum of 30 characters.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
	<tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
    <tr bgcolor=<%=bgc%> valign=center height="30">
      <td align="right">Tags</td>
        <td>&nbsp;
          <input id=text5 name=TeamTags style=" WIDTH: 70px" maxlength="10" class=text value="<%=Server.HTMLEncode(ttags)%>">
        </td>
    </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>Team tags are what you use in game to identify yourself as a member of this team. Up to 10 characters, and any combination of alphanumerics and special characters. This should also be unique and original.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
	<tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
	<tr bgcolor=<%=bgc%> valign=center height="30">
	   <td align="right">Home Page:</td>
	   <td>&nbsp;
	     <input id=text4 name=TeamURL maxlength=40 style=" WIDTH: 300px" class=text value="<%=Server.HTMLEncode(turl)%>">
	   </td>
	 </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>Not neccessary, but definately a good way to advertise your team and bring in potential recruits.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
		<tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
    <tr bgcolor=<%=bgc%> valign=center height="30">
       <td align="right">Team Logo URL: </td>
       <td>&nbsp;
         <input id=text4 name=TeamLOGO maxlength=100 style=" WIDTH: 300px" class=text value="<%=Server.HTMLEncode(tlogourl & "")%>">
       </td>
     </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>A link 200X200 pixel picture (jpg, gif, bmp, png) that represents your team. Should be in the form of: 'http://www.teamwarfare.com/image.jpg'</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
		<tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
    <tr bgcolor=<%=bgc%> valign=center height="30">
       <td align="right" NOWRAP>Team IRC Channel: </td>
       <td>&nbsp;
         <input id=text4 name=TeamIRC maxlength=25 style=" WIDTH: 300px" class=text value="<%=Server.HTMLEncode(tirc & "")%>">
       </td>
     </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>IRC (Internet Relay Chat) is used for a more instant communication. If applicable, please enter it here.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
    <tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
    <tr bgcolor=<%=bgc%> valign=center height="30">
       <td align="right">IRC Server:</td>
       <td>&nbsp;
         <input id=text4 name=TeamIRCServer maxlength=40 style=" WIDTH: 300px" class=text value="<%=Server.HTMLEncode(tircserver & "")%>">
       </td>
     </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>Server that hosts your IRC channel.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
	<tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
	<tr bgcolor=<%=bgc%> valign=center height="30">
	   <td align="right">Team E-mail: </td>
	   <td>&nbsp;
	     <input id=text4 name=TeamEmail  maxlength=40 style=" WIDTH: 300px" class=text value="<%=Server.HTMLEncode(temail)%>">
	   </td>
	 </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>This is an e-mail address used to send out all pertenant information about the team. This includes, players joining and quiting your team, and major ladder activity, especially time critical activities.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
  <tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
    <tr bgcolor=<%=bgc%> valign=center height="30">
       <td align="right">Join Password:</td>
       <td>&nbsp;
         <input id=text4 name=TeamJoinPassword maxlength=10 style=" WIDTH: 80px" class=text value="<%=Server.HTMLEncode(tjoinpass)%>">
       </td>
     </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
	<tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
    <tr bgcolor=<%=bgc%> valign=center height="30">
       <td align="right">Confirm Password:</td>
       <td>&nbsp;
         <input id=text4 name=TeamConfirmJoinPassword maxlength=10 style=" WIDTH: 80px" class=text value=<%=Server.HTMLEncode(tconfirmjoinpass)%>>
       </td>
     </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>Up to 10 characters, the join password is used for all players wishing to join your team. The join password can be changed later, and is not neccessary to administrate your team.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>   
	<tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
    <tr bgcolor=<%=bgc%> valign=center height="30">
       <td align="right" valign=top>Team Description<br /><font color="#ff0000">Do <b><u>not</u></b> put any sound in this description or your team will be removed.</font></td>
       <td>&nbsp;
         <textarea name=TeamDesc rows=5 cols=70><%=Server.HTMLEncode(tDesc)%></textarea>
       </td>
     </tr>
	<%
	If bIsEdit then
		if bgc=bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
		response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#FFD142>A brief description of your team.</font></td></tr></table></td></tr>"
		if bgc=bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
		%>   
		<tr><td colspan="2"bgcolor="000000"><img src="images/spacer.gif" height="8" width="1" alt="" border="0" /></td></tr>
        <tr bgcolor=<%=bgc%> valign=center height="30">
           <td align="right" valign=top>Change Founder</td>
           <td>&nbsp;
		   	<select name=newfounderid>
		   	<%
		   	strSQL = "select distinct p.playerid, p.playerhandle, t.teamname, t.teamid, tpl.isadmin "
		   	strSQL = strSQL & "from lnk_t_p_l tpl, lnk_t_l tl, TBL_players p , tbl_teams t "
		   	strSQL = strSQL & "where tl.teamid = '" & tid & "'"
		   	strSQL = strSQL & "AND tpl.tllinkid=tl.tllinkid "
		   	strSQL = strSQL & "AND p.playerid = tpl.playerid "
		   	strSQL = strSQL & "AND t.teamid = tl.teamid "
		   	strSQL = strSQL & "AND tpl.isadmin = 1 AND tl.isactive = 1 "
		   	strSQL = strSQL & "union all select distinct p.playerid, p.playerhandle, t.teamname, t.teamid, ltp.isadmin "
		   	strSQL = strSQL & "from lnk_league_team_player ltp, lnk_league_team lt, TBL_players p , tbl_teams t, tbl_leagues lg "
		   	strSQL = strSQL & "where lt.teamid = '" & tid & "'"
		   	strSQL = strSQL & "AND ltp.lnkLeagueTeamID=lt.lnkLeagueTeamID "
		   	strSQL = strSQL & "AND p.playerid = ltp.playerid "
		   	strSQL = strSQL & "AND lg.LeagueID = lt.LeagueID "
		   	strSQL = strSQL & "AND t.teamid = lt.teamid "
		   	strSQL = strSQL & "AND ltp.isadmin = 1 AND lt.active = 1 AND lg.LeagueActive = 1 "
		   	strSQL = strSQL & "ORDER BY playerhandle"
		   	ors.open strsql, oconn
		   	if not(ors.eof and ors.bof) then
		   		do while not(ors.eof)
		   			Response.Write "<option value=""" & ors("playerid") & """ "
		   			if ors("playerid") = townerid then
		   				Response.Write " selected "
		   			end if
		   			Response.Write ">" & ors("playerhandle") & "</option>" & vbcrlf
		   			ors.movenext
		   		loop
		   	end if
		   	ors.close
		   	%>
		   	</select>
		   	<input type=hidden name=oldfounderid value="<%=townerid%>">
           </td>
         </tr>
		<% 
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right>"
	Response.Write "<table width=98% ><tr><td><font color=#FFD142>"
	Response.write "In order to bestow ownership upon another player on your team, the player must be on all the ladder you participate on and must be a captain on all ladders."
	Response.Write "</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
End if

if errorcode=2 then
	response.write "<tr bgcolor=" & bgc & " valign=center height=18><td colspan=2 align=center><font color=red>Passwords do not match, please try again</font></td></tr>"
end if
%>
	<tr bgcolor=<%=bgc%> valign=center height="30">
	 <td colspan=2 align=center>
	 	<input id=method name=SaveMethod value="<%=strMethod%>" type=hidden>
	<% if bIsEdit then %>
		<input id=submit1 name=submit1 type=submit value='Edit Team' class=bright>
    <% else %>	
        <input id=submit1 name=submit1 type=submit value='Create Team' class=bright>
    <% end if %>
    </td>
  </tr>
</table>
    </td>
  </tr>
</table>
<input type=hidden name=SaveType value=team>
</form>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>