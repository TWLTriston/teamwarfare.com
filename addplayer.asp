<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add a Player"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
Dim chrPlayerActive, intForumAccess
Dim bSysAdmin2
bSysAdmin2 = IsSysAdminLevel2()
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim bIsEdit, strPlayerName, strMethod, strVerbage
Dim uName, uEmail, uICQ, HideEmail, PlayerTitle, PlayerSignature, intContributor, ContributorAmount, iPlayerID
Dim intPlayerID, intCanActivate, intSuspension
Dim intPlayercommentID,strComment,strForumBanLiftDate, intAccessID
Dim strSiteBanLiftDate, strSuspensionLiftDate

bIsEdit = cBool(Request.QueryString("IsEdit"))
If bIsEdit And Session("LoggedIn") Then
	strPlayerName = Request.QueryString ("PlayerName")
	If strPlayerName = "" Then
		strPlayerName = Session("uName")
	End If
	Dim sUpper
	sUpper = UCase(strPlayerName)
	If ( strPlayerName <> Session("uName") AND ( sUpper = "TRISTON" Or sUpper = "POLARIS" Or sUpper = "TOTALCARNAGE"  ) ) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
		
	If strPlayerName <> Session("uName") AND Not(bSysAdmin) Then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	iPlayerID = 0
	strSQL = "select * from tbl_Players where PlayerHandle='" & CheckString(strPlayerName) & "'"
	oRS.Open strSQL, oConn
	If Not(oRs.EOF AND oRS.BOF) Then
		iPlayerID=oRs.Fields("PlayerID").Value
		uEmail=oRs.Fields("PlayerEmail").Value
		uICQ=oRs.Fields("PlayerICQ").Value
		hideemail=oRs.Fields("PlayerHideEmail").value
		PlayerTitle = oRs.Fields("PlayerTitle").Value
		PlayerSignature = oRs.Fields("PlayerSignature").Value
		intPlayerID = oRs.Fields("PlayerID").Value
		intContributor = oRs.Fields("Contributor").Value
		ContributorAmount = oRs.Fields("ContributorAmount").Value
		chrPlayerActive = oRs.Fields("PlayerActive").Value
		intForumAccess = oRs.Fields("ForumAccess").Value 
		intCanActivate = oRs.Fields("PlayerCanActivate").Value
		intSuspension = oRs.Fields("Suspension").Value
		strForumBanLiftDate = oRs.Fields("ForumBanLiftDate").Value
		strSiteBanLiftDate = oRs.Fields("SiteBanLiftDate").Value
		strSuspensionLiftDate = oRs.Fields("SuspensionLiftDate").Value

	Else
		Response.Redirect "viewplayer.asp?player=" & Server.URLEncode(strPlayerName)
	End If
	oRs.Close
	strSQL = "SELECT * FROM sysadmins WHERE AdminID = '" & iPlayerID & "' AND AdminLevel = 2"
	oRS.Open strSQL, oConn
	If Not(oRs.EOF AND oRS.BOF) Then
		If Not(bSysAdmin2) AND strPlayerName <> Session("uName") Then
			Response.Redirect "viewplayer.asp?player=" & Server.URLEncode(strPlayerName)
		End If
	End If
	oRs.Close
	
	strMethod = "Edit"
	strVerbage = "Edit Player Profile"	
ElseIf bIsEdit AND Not(Session("LoggedIn")) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"
ElseIf Not(bIsEdit) And Session("LoggedIn") Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
Else 
	strMethod = "New"
	strVerbage = "Create a Player Profile"
End If

If Len(intForumAccess) = 0 Or IsNull(intForumAccess) Then
	intForumAccess = 1	
End If
If Len(intCanActivate) = 0 Or IsNull(intCanActivate) Then
	intCanActivate = 1
End If
If Len(intSuspension) = 0 Or IsNull(intSuspension) Then
	intSuspension = 0
End If

strPageTitle = "TWL: " & strVerbage
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart(strVerbage) %>
<script language="javascript" type="text/javascript">
<!--
	function fValidateForm() {
		var oForm = document.frmAddPlayer;
		var errFlag = 0;
		var errStr = "Error:\n";
		
		<% If bSysAdmin Then %>
		var iForumAccess = '<%=intForumAccess%>';
		var iCanActivate = '<%=intCanActivate%>';
		var iSuspension = '<%=intSuspension%>';
		var blnRequireReason = false;
		
		if (iForumAccess != oForm.ForumAccess[oForm.ForumAccess.selectedIndex].value) {
			blnRequireReason = true;
		}
		if (iCanActivate != oForm.CanActivate[oForm.CanActivate.selectedIndex].value) {
			blnRequireReason = true;
		}
		if (iSuspension != oForm.Suspension[oForm.Suspension.selectedIndex].value) {
			blnRequireReason = true;
		}
		if (blnRequireReason && oForm.comment.value.length == 0) {
			errFlag = 1;
			errStr = errStr + "You must enter a comment when changing access";
		}
		<% End If %>
		
		if (errFlag == 0) {
			return true;
		} else {
			alert(errStr);
			return false;
		}
		
	}

//-->
</script>
    <form name=frmAddPlayer id=frmAddPlayer action=saveItem.asp method=post onSubmit="return fValidateForm();"> 
    <table border="0" align="center" width="97%" cellspacing="0" cellpadding="0" class="cssBordered">
		<%	
		if bgc=bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
		if request("error")=2 then
		      response.write "<tr bgcolor=" & bgc & " valign=center height=18><td colspan=2 align=center><font color=red>Error: Player Name or Email Address already exists, please try again.</font></td></tr>"
		end if
		if request("error")=4 then
		      response.write "<tr bgcolor=" & bgc & " valign=center height=18><td colspan=2 align=center><font color=red>Error: Cannot change Player Name. New name already exists.	</font></td></tr>"
		end if
		if bgc=bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
		if request("error")=3 then
		      response.write "<tr bgcolor=" & bgc & " valign=center height=18><td colspan=2 align=center><font color=red>Error: Email Address already exists, please try again.</font></td></tr>"
		end if
		%>
    <tr bgcolor=<%=bgc%> valign=center>
      <td align="right">Player Name: </td>
      <%
      if strMethod<>"Edit" OR bSysAdmin then
	  	%>
	  	<td>&nbsp;&nbsp;<input id=text1 name=PlayerName style=" WIDTH: 150px" class=text maxlength="30" value="<%=Server.HTMLEncode(strPlayerName)%>"></td>
	  	<% If bSysAdmin Then %>
	  	<input type="hidden" name="playerID" id="playerID" value="<%=intPlayerID%>" />
	  	<input type="hidden" id=OldPlayerName name=OldPlayerName value="<%=Server.HTMLEncode(strPlayerName)%>" />
	  	<% End If %>
	  	<%
	  else
	  	%>
	  	<td><%=strPlayerName%></td>
	  	<input name=PlayerName type=hidden value="<%=Server.HTMLEncode(strPlayerName)%>">
	  <%
	  end if
	  %>
    </tr>
	<%	
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#ffcf3f>Member names should be unique. You are discouraged from including your team's tag when registering your name, because you may change teams/join another team. Maximum of 30 characters. <font color=red>NOTE: Please do NOT use non-standard ASCII characters in your username. Certain non-critical site features will not work for member names with ASCII characters.</font></font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
    <tr bgcolor=<%=bgc%> valign=center>
      <td align="right">Password: </td>
      <td>&nbsp;<input id=text5 type=password name=PlayerPassword maxlength="15" style=" WIDTH: 150px" class=text></td>
    </tr>
	<%
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
	<tr bgcolor=<%=bgc%> valign=center>
      <td align="right" NOWRAP>Confirm Password: </td>
      <td>&nbsp;<input type=password id=text5 name=PlayerConfirmPassword maxlength="15" style=" WIDTH: 150px" class=text></td>
    </tr>
	<%	
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#ffcf3f>Choose a password between 4-10 characters, this should be kept private. Your name and password will be stored in your browser, so you should not need to use it often, unless your cookies are deleted, or you are at a remote location. If you are changing your profile, you do not need to enter the password, unless you are changing it.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
    <tr bgcolor=<%=bgc%> valign=center>
      <td align="right">Email: </td>
      <td>&nbsp;<input id=text4 name=PlayerEmail style=" WIDTH: 300px" class=text value="<%=Server.HTMLEncode(uemail)%>"></td>                  
    </tr>
	<%	
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#ffcf3f>Please enter a valid e-mail address. This will be used for all mailings if you are a captain of any team. Your e-mail address will not be used to solicit information, nor sold to other organizations.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	if hideemail=1 then 
		hideemail="checked"
	else
		hideemail="unchecked"
	end if
	%>
    <tr bgcolor=<%=bgc%> valign=center>
      <td align="right">Hide Email: </td>
      <td>&nbsp;<input type=checkbox class=bright name=HideEmail <%=hideemail%>></td>                  
    </tr>
	<%	
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#ffcf3f>Check this box if you wish to not have your e-mail displayed when other members view your profile.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
    <tr bgcolor=<%=bgc%> valign=center>
      <td align="right">Xfire: </td>
      <td>&nbsp;<input id=text3 name=PlayerICQ style="WIDTH: 80px" class=text value="<%=Server.HTMLEncode(uICQ)%>"></td>
    </tr>
	<%	
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#ffcf3f>Xfire - Game-friendly instant messaging available from <a href=http://www.xfire.com target=_new>www.xfire.com</a></font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	%>
    <tr bgcolor=<%=bgc%> valign=center>
      <td align="right" valign=top>Signature: </td>
      <td><textarea name=Signature id=signature cols=70 rows=7><%=PlayerSignature%></textarea></td>
    </tr>
	<%	
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	response.write "<tr bgcolor=" & bgc & "><td colspan=2 align=right><table width=98% ><tr><td><font color=#ffcf3f>Signature's are used on the forums as a way to automatically <i>sign</i> your posts.<br /><b>Note:</b> Any images inside your signature should be <i>work-safe</i>, meaning no nudity, drugs, ect. This is at our discretion, and abuse will result in the loss of your ability to have a signature.</font></td></tr></table></td></tr>"
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc=bgcone
	end if
	if request("error")=1 then
          response.write "<tr bgcolor=" & bgc & " valign=center height=18><td colspan=2 align=center><font color=red>Passwords do not match</font></td></tr>"
	end if
	%>                   
	<tr bgcolor=<%=bgc%> valign=center>
		<td colspan=2 align=center>
		<input type=hidden name=SaveType value="player">
		<input type=hidden name=OldEmail value="<%=uemail%>">
		
		<input type=hidden name=SaveMethod value="<%=strMethod%>">
		<%
		if strMethod="Edit" then
			%>
			<input id=submit1 name=submit1 type=submit value='Save Changes' class=bright>
			<%                          
        else
			%>
			<p>By creating an account on TeamWarfare, you must agree to our <a href="/terms.asp" target="_blank">Terms and Conditions</a>.<br />Do NOT click the button below until you have read and agreed to the terms.</p>
			<input id=submit1 name=submit1 type=submit value='Create Player' class=bright>
			<%
		end if
		%>
        </td>
      </tr>
    </table>
    
<%
'--------------------------------------
' SysAdmin Only Options
'--------------------------------------
If bSysAdmin Then %>
	<br /><br />
	<table border="0" cellspacing="0" cellpadding="0" width="97%" class="cssBordered">
	<tr>
		<th colspan="2">Admin Only Options</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right">Player Title:</td>
		<td bgcolor="<%=bgcone%>">&nbsp;<input id=text3 name=PlayerTitle style="WIDTH: 150px" maxlength="500" class=text value="<%=Server.HTMLEncode(PlayerTitle)%>"></td>
		<input type="hidden" id="OldPlayerTitle" name=OldPlayerTitle value="<%=Server.HTMLEncode(PlayerTitle)%>" />
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right">Player Active<br /> (Email Verified): </td>
		<td bgcolor="<%=bgctwo%>">&nbsp;<select name="PlayerActive" id="PlayerActive">
			<option value="N" <% If uCase(Trim(chrPlayerActive)) = "N" THen Response.Write " SELECTED " End If %>>No</option>
			<option value="Y" <% If uCase(Trim(chrPlayerActive)) = "Y" THen Response.Write " SELECTED " End If %>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
   <td bgcolor="<%=bgcone%>" align="right">TWL Contributor: </td>
   <td bgcolor="<%=bgcone%>">&nbsp;<select name="Contributor">
			<OPTION VALUE="0" <% If cStr("" & intContributor) = "0" THen Response.Write " SELECTED " End If %>>No</OPTION>
			<OPTION VALUE="1" <% If cStr("" & intContributor) = "1" THen Response.Write " SELECTED " End If %>>Yes</OPTION>
		</SELECT>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right">Amount: </td>
		<td bgcolor="<%=bgctwo%>">&nbsp;$ <input type="text" id="ContributorAmount" name="ContributorAmount" style="WIDTH: 50px" maxlength="6" class=text value="<%=Server.HTMLEncode(ContributorAmount & "")%>"></td>
	</tr>
	<tr>
		<th colspan="2">Ban &amp; Suspension Options</th>
	</tr>
	<tr>
		<td colspan="2"><i>If you change any flag to implement a ban, you must fill in an end date in MM/DD/YYYY format, as well as a comment explaining the reason for the ban.</i></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right">Forum Access: </td>
		<td bgcolor="<%=bgcone%>">&nbsp;<select name="ForumAccess" id="ForumAccess">
			<OPTION VALUE="0" <% If cStr("" & intForumAccess) = "0" THen Response.Write " SELECTED " End If %>>No</OPTION>
			<OPTION VALUE="1" <% If cStr("" & intForumAccess) = "1" THen Response.Write " SELECTED " End If %>>Yes</OPTION>
		</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right">Forum Ban Lift Date</td>
		<td bgcolor="<%=bgctwo%>">&nbsp;<input id=date1 name=ForumBanLiftDate style="WIDTH: 60px" maxlength="10" value="<%=Server.HTMLEncode("" & strForumBanLiftDate)%>"></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right">Site Access: </td>
		<td bgcolor="<%=bgctwo%>">&nbsp;<select name="CanActivate" id="CanActivate">
		<OPTION VALUE="0" <% If cStr("" & intCanActivate) = "0" THen Response.Write " SELECTED " End If %>>No</OPTION>
		<OPTION VALUE="1" <% If cStr("" & intCanActivate) = "1" THen Response.Write " SELECTED " End If %>>Yes</OPTION>
		</SELECT></TD>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right">Site Ban Lift Date</td>
		<td bgcolor="<%=bgcone%>">&nbsp;<input id=sitedate name=SiteBanLiftDate style="WIDTH: 60px" maxlength="10" value="<%=Server.HTMLEncode("" & strSiteBanLiftDate)%>"></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right">Is Suspended: </td>
		<td bgcolor="<%=bgctwo%>">&nbsp;<select name="Suspension" id="Suspension">
			<OPTION VALUE="0" <% If cStr("" & intSuspension) = "0" THen Response.Write " SELECTED " End If %>>No</OPTION>
			<OPTION VALUE="1" <% If cStr("" & intSuspension) = "1" THen Response.Write " SELECTED " End If %>>Yes</OPTION>
			</SELECT>
		</TD>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right">Suspension Lift Date</td>
		<td bgcolor="<%=bgcone%>">&nbsp;<input id=suspenddate name=SuspensionLiftDate style="WIDTH: 60px" maxlength="10" value="<%=Server.HTMLEncode("" & strSuspensionLiftDate)%>"></td>
	</tr>
	<tr>
	 <td bgcolor="<%=bgctwo%>" align="right">Comments</td>
	 <td bgcolor="<%=bgctwo%>" >&nbsp;<textarea name="Comment" id="comment" cols="40" rows="4"><%=Server.HTMLEncode(strComment & "")%></textarea></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="center" colspan="2"><input type="submit" value="Save Changes" class="bright" /></td>
	</tr>
	</table>
	<br />
	<table border="0" cellspacing="0" cellpadding="0" width="97%" class="cssBordered">
	<tr>
		<th colspan="2">Admin Comment History</th>
	</tr>
	<%
	// Ban comment retrieval
	strSQL = "SELECT Comment, CommentDate, tbl_players.PlayerHandle FROM tbl_player_comments INNER JOIN tbl_players ON tbl_player_comments.AdminID=tbl_players.PlayerID WHERE tbl_player_comments.PlayerID='" & intPlayerID & "' ORDER BY PlayerCommentID DESC"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		While Not oRs.EOF
			%>
				<tr>
					<td bgcolor="<%=bgc%>" width="15%" valign=top><b><%=server.htmlencode("" & oRs.FieldS("PlayerHandle").Value)%></b><br /><%Response.Write "<span class=""smalldate"">" & FormatDateTime(oRs.FIelds("CommentDate").Value) & "</span>"%></td>
					<td bgcolor="<%=bgc%>" valign="top"><%=ForumEncode2(oRs.Fields("Comment").Value)%></td>
				</tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			oRs.MoveNext
		Wend
	End If
	oRs.Close
	%>
	</table>
<%
End If
%>
    </form>

<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>