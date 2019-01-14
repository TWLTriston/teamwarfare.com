<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Forums Administration"

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

Dim AdminType
Dim CategoryName, CategoryDesc, CategoryOrder, Edit,verbage
Dim ForumName, ForumPass, ForumDesc, ForumOrder, CategoryID, counter, playerhandle, ForumMatchDispute
Dim ReplaceSearch, ReplaceFiller, ReplaceID, ReplaceCategory, ForumID, MakeID, strPlayerHandle
	dim f1, f2

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% 
if not IsSysAdminLevel2() then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Set oRS2 = Nothing
	Response.Clear
	Response.Redirect "/default.asp"
end if
AdminType = request("AdminType")
If Len(AdminType) = 0 Then AdminType = "forum" End IF 
Call ContentStart("Forum Administration")%>
    <table width="90%" border="0">
      <tr>
        <td><p class=text>
<% ' Put links here %>
			<a href="/forums/admin/default.asp?AdminType=category">Category Admin</a><BR>
			<a href="/forums/admin/default.asp?AdminType=forum">Forum Admin</a><BR>
			<a href="/forums/admin/default.asp?AdminType=moderator">Moderator Admin - By Name</a><BR>
			<a href="/forums/admin/default.asp?AdminType=moderatorf">Moderator Admin - By Forum</a><BR>
			<a href="/forums/admin/default.asp?AdminType=access">Who has what forum Access - By name</a><BR>
			<a href="/forums/admin/default.asp?AdminType=accessforums">Who has what forum Access - By Forum</a><BR>
<% ' End Links Section %>
		</td>
	  </tr>
	</table>
<% Call ContentEnd() %>
<%
if AdminType = "category" then 

	if request("CategoryID") <> "" then
		strsql = "select * from tbl_category where CategoryID = " & request("CategoryID")
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then 
			CategoryName = ors.fields("categoryname").value
			CategoryDesc = ors.fields("CategoryDesc").value
			CategoryOrder = ors.fields("CategoryOrder").value
			Edit = true
		end if
		ors.close
	end if
	
	Call ContentStart("Category Admin")
%>
    <table width="90%" border="0">
      <tr>
        <td>
<%
	strsql = "select * from tbl_category order by CategoryOrder"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		response.write "<table border=0 cellspacing=0 width=50% cellpadding=2>"
        do while not (ors.eof)
			Response.Write "<TR><TD><p class=small><a href='/forums/admin/default.asp?AdminType=category&CategoryID=" & ors.fields("CategoryID") & "'>" & server.HTMLEncode(ors.fields("CategoryName").value) & "</a> ---- </p></td></tr>"
			ors.movenext
		loop
		Response.Write "</table>"
	else
		Response.Write "<font color=red><B>No categories found.</b></font>"
	end if
	ors.close
	%>
        <form id=newcategory name=newcategory method=post action="/forums/admin/adminengine.asp">
        <table border=0 cellspacing=0 cellpadding=2 width=97%>
        <TR valign=center height=30 bgcolor=<%=bgcone%>><TD>Category Name: </td><TD><input type=text value="<%=categoryname%>" name="catname"></td></tr>
        <TR valign=top bgcolor=<%=bgctwo%>><TD>Category Description: </td><TD><textarea name="description" cols=50 rows=5><%=categorydesc%></textarea></td></tr>
        <TR valign=center height=30 bgcolor=<%=bgcone%>><TD>Category Order: </td><TD><input type=text value="<%=categoryorder%>" name="catorder"></td></tr>
        <TR bgcolor=<%=bgctwo%>><TD colspan=2 align=center>
        <input type=hidden name=savetype id=savetype value="Category">
        <% if edit then 
			verbage = "Edit Category"
		%>
        <input type=hidden name=catid id=catid value="<%=request("CategoryID")%>">
        <input type=hidden name=edit id=edit value="True">
        <% else 
			verbage = "Make New Category"
		%>
        <input type=hidden name=edit id=edit value="False">
        <% end if %>
        <input type=submit name=submitcat id=submitcat value="<%=verbage%>">
        </TD></TR>
        </table>
        </form>
		</td>
	  </tr>
	</table>
	<% 
	Call ContentEnd()
end if 
' end category stuff
if AdminType = "forum" then 
	if request("ForumID") <> "" then
		strsql = "select * from tbl_forums where ForumID = " & request("ForumID")
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then 
			ForumName = ors.fields("ForumName").value
			ForumPass = ors.fields("ForumPassword").value
			ForumDesc = ors.fields("ForumDescription").value
			ForumOrder = ors.fields("ForumOrder").value
			CategoryID = ors.fields("CategoryID").value
			ForumMatchDispute = ors.fields("ForumMatchDispute").value
			Edit = true
		end if
		ors.close
	end if
	Call ContentStart("Forum Admin")
%>
    <table width="90%" border="0">
      <tr>
        <td>
<%
	Dim strOldCategory, strThisCategory
	strSQL = "SELECT tbl_category.CategoryName, tbl_category.CategoryID, tbl_forums.ForumID, tbl_forums.ForumName, tbl_forums.ForumDescription, "
	strSQL = strSQL & " tbl_forums.ForumThreadCount, tbl_forums.ForumPostCount, tbl_forums.ForumLocked, tbl_forums.ForumLastPostTime, tbl_forums.ForumLastPosterName "
	strSQL = strSQL & " FROM tbl_category, tbl_forums "
	strSQL = strSQL & " WHERE tbl_forums.CategoryID = tbl_category.CategoryID AND ForumID <> 33 "
	strSQL = strSQL & " ORDER BY tbl_category.CategoryOrder ASC, tbl_forums.ForumOrder ASC, tbl_forums.ForumName ASC "
	ors.open strsql, oconn
	strOldCategory = ""
	if not (ors.eof and ors.bof) then
		response.write "<table border=0 cellspacing=0 width=50% cellpadding=2>"
        do while not (ors.eof)
			strThisCategory = ors.fields("CategoryName").value
			If strThisCategory <> strOldCategory Then
				Response.Write "<TR BGCOLOR=" & BGCOne & " style=""cursor:default""><TD colspan=5><a CLASS=""category"" name=""Category" & oRS.Fields("CategoryID").Value & """ href=""default.asp#Category" & oRS.Fields("CategoryID").Value & """>"
				Response.Write server.HTMLEncode(ors.fields("CategoryName").value & "") & "</A>"
				Response.Write "</td></tr>" & vbCrLf
				strOldCategory = strThisCategory
			End If

			Response.Write "<TR><TD><a href='/forums/admin/default.asp?AdminType=forum&forumid=" & ors.fields("ForumID") & "'>" & server.HTMLEncode(ors.fields("ForumName").value) & "</a></td></tr>"
			ors.movenext
		loop
		Response.Write "</table>"
	else
		Response.Write "<font color=red><B>No forums found.</b></font>"
	end if
	ors.close
%>
        <form id=newcategory name=newcategory method=post action="/forums/admin/adminengine.asp">
        <table border=0 cellspacing=0 cellpadding=2 width=97%>
        <TR valign=center height=30 bgcolor=<%=bgcone%>><TD>Forum Name: </td><TD><input type=text value="<%=ForumName%>" name="ForumName"></td></tr>
        <TR valign=center height=30 bgcolor=<%=bgctwo%>><TD>Forum Password: </td><TD><input type=text value="<%=ForumPass%>" name="ForumPass"></td></tr>
        <TR valign=top bgcolor=<%=bgcone%>><TD>Forum Description: </td><TD><textarea name="ForumDesc" cols=50 rows=5><%=ForumDesc%></textarea></td></tr>
        <TR valign=center height=30 bgcolor=<%=bgctwo%>><TD>Forum Order: </td><TD><input type=text value="<%=ForumOrder%>" name="ForumOrder"></td></tr>
        <TR valign=center height=30 bgcolor=<%=bgcone%>><TD>Disputes Forum?</td><TD><select name="ForumMatchDispute"><option value="0">No</option><option value="1"<% If ForumMatchDispute = 1 Then Response.Write " selected=""selected""" End If %>>Yes</option></select></td></tr>
        <TR valign=center height=30 bgcolor=<%=bgctwo%>><TD>Category: </td><TD>
        <select name=CategoryID><option>-- &lt;Select a Category&gt; --</option>
        <%
        strsql = "select CategoryName, CategoryID from tbl_category order by CategoryName asc"
        ors.open strsql, oconn
        if not(Ors.eof and ors.bof) then
			do while not (ors.eof)
				Response.Write "<option value='" & ors.fields("CategoryID").value & "' "
				if ors.fields("CategoryID").value = CategoryID then
					Response.Write " selected "
				end if
				Response.Write ">" & server.HTMLEncode(ors.fields("CategoryName").value) & "</option>"
				ors.movenext
			loop
		end if
		%>
		</select></td></tr>
        <TR bgcolor=<%=bgcone%>><TD colspan=2 align=center>
        <input type=hidden name=savetype id=savetype value="forum">
        <% if edit then 
			verbage = "Edit Forum"
		%>
        <input type=hidden name=forumid id=forumid value="<%=request("ForumID")%>">
        <input type=hidden name=edit id=edit value="True">
        <% else 
			verbage = "Make New Forum"
		%>
        <input type=hidden name=edit id=edit value="False">
        <% end if %>
        <input type=submit name=submitforum id=submitforum value="<%=verbage%>">
        </TD></TR>
        </table>
        </form>
		</td>
	  </tr>
	</table>
	<% 
	Call ContentEnd()
end if 
' end forum stuff
' Start Replace
if AdminType = "replace" then 
	Call ContentStart("Replace Admin")
%>
    <table width="90%" border="0">
      <tr>
        <td>
		<%
	strsql = "select * from tbl_replace order by ReplaceCategory asc"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then 
		response.write "<table border=0 cellspacing=0 cellpadding=2>"
		counter=-1
		Response.Write "<TR>"
		do while not (ors.eof)
			counter=counter+1
			if counter= 4 then
				counter=0
				Response.Write "</TR><TR>" & vbcrlf
			end if
			if ors.fields("ReplaceCategory").value = 2 then
				Response.Write "<TD width=125 align=center><p class=small><a href=/forums/admin/default.asp?AdminType=replace&id=" & ors.fields("ReplaceID").value & ">" & Server.htmlencode(ors.fields("ReplaceSearch").value) & "</a> = " & ors.fields("ReplaceFiller").value & " <BR><a href='adminengine.asp?delete=true&table=replace&ID=" & ors.fields("ReplaceID").value & "'>Delete</a></p></TD>" & vbcrlf
			else		
				Response.Write "<TD width=125 align=center><p class=small><a href=/forums/admin/default.asp?AdminType=replace&id=" & ors.fields("ReplaceID").value & ">" & Server.HTMLEncode (ors.fields("ReplaceSearch").value) & "</a> = " & Server.HTMLEncode (ors.fields("ReplaceFiller").value) & " <BR><a href='adminengine.asp?delete=true&table=replace&ID=" & ors.fields("ReplaceID").value & "'>Delete</a></p></TD>"  & vbcrlf
			end if
			ors.movenext
		loop
		Response.Write "</table>"
	else
		Response.Write "<p class=small><font color=red>No replace text found</font></p>"
	end if
	ors.close
	edit = false
	if Request.QueryString("id") <> "" then
		strsql = "select * from tbl_replace where ReplaceID='" & Request.QueryString("ID") & "'"
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then
			edit = true
			ReplaceSearch = ors.fields("ReplaceSearch").value
			ReplaceFiller = ors.fields("ReplaceFiller").value
			ReplaceID = ors.fields("ReplaceID").value
			ReplaceCategory = ors.fields("ReplaceCategory").value
		end if
		ors.close
	end if
	
%>
        <form id=replace name=replace method=post action="/forums/admin/adminengine.asp">
        <table border=0 cellspacing=0 cellpadding=2 align=center>
        <TR valign=center height=30 bgcolor=<%=bgcone%>><TD>Replace Search Test: </td><TD><input type=text value="<%=Server.HTMLEncode (ReplaceSearch)%>" name="ReplaceSearch"></td></tr>
        <TR valign=center height=30 bgcolor=<%=bgctwo%>><TD>Replace Filler Text: </td><TD><input type=text value="<%=Server.HTMLEncode (ReplaceFiller)%>" name="ReplaceFiller"></td></tr>
        <TR valign=center height=30 bgcolor=<%=bgctwo%>><TD>Replace Category (1 = text, 2=smiley): </td><TD><input type=text value="<%=Server.HTMLEncode (ReplaceCategory)%>" name="ReplaceCategory"></td></tr>
        <TR valign=center bgcolor=<%=bgcone%>><TD Colspan=2 align="center">
        <input type=hidden name=savetype id=savetype value="Replace">
        <% if edit then 
			verbage = "Edit Replace Text"
		%>
        <input type=hidden name=ReplaceID id=ReplaceID value="<%=ReplaceID%>">
        <input type=hidden name=edit id=edit value="True">
        <% else 
			verbage = "Make Replace Text"
		%>
        <input type=hidden name=edit id=edit value="False">
        <% end if %>
        <input type=submit name=submitforum id=submitforum value="<%=verbage%>">
        </TD></TR>
        </table>
        </form>
		</td>
	  </tr>
	</table>
	<%
	call ContentEnd()
end if 
' end Replace
if AdminType = "moderator" then 
	Call ContentStart("Moderator Admin")
	%>
    <table width="90%" border="0">
      <tr>
        <td>
	<%
	strsql = "select FPLinkID, PlayerHandle, ForumName from lnk_f_p inner join tbl_forums on tbl_forums.ForumID=lnk_f_p.forumid inner join tbl_players on tbl_players.playerid = lnk_f_p.playerid ORDER BY PlayerHandle ASC"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then 
		response.write "<table border=0 cellspacing=0 cellpadding=2>"
		do while not (ors.eof)
			PlayerHandle = ors.fields("PlayerHandle").value
			Response.Write "<TR><TD><p class=small><a href=/viewplayer.asp?player=" & Server.URLEncode("" & PlayerHandle) & ">" & Server.HTMLEncode (PlayerHandle) & "</a> = " & Server.HTMLEncode (ors.fields("ForumName").value) & " - <a href='adminengine.asp?delete=true&table=moderator&ID=" & ors.fields("FPLinkID").value & "'>Delete</a></p></TD></TR>" & vbcrlf
			ors.movenext
		loop
		Response.Write "</table>"
	end if
	ors.close
	edit = false
	strPlayerHandle = Request.QueryString("Player")
	If Len(strPlayerHandle) > 0 Then
	%>
        <form id=moderator name=moderator method=post action="/forums/admin/adminengine.asp">
        <table border=0 cellspacing=0 cellpadding=2 align=center>
        <TR valign=center height=30 bgcolor=<%=bgcone%>><TD>Forum: </td><TD><select name=forumid id=forumid><%
        strsql = "Select ForumName, ForumID from tbl_forums order by ForumName asc"
        ors.open strsql, oconn
        if not (ors.eof and ors.bof) then
			do while not(ors.eof)
				Response.Write "<option value='" & ors.fields("ForumID").value & "'"
				Response.Write ">" & ors.fields("ForumName").value & "</option>" & vbcrlf
				ors.movenext
			loop
		end if
		ors.close
		%>
		</select>
        <TR valign=center height=30 bgcolor=<%=bgctwo%>><TD>Member: </td><TD><select name=playerid id=playerid><%
        strsql = "Select PlayerHandle, PlayerID from tbl_players WHERE PlayerHandle LIKE '%" & CheckString(strPlayerHandle) & "%' order by PlayerHandle asc"
        ors.open strsql, oconn
        if not (ors.eof and ors.bof) then
			do while not(ors.eof)
				Response.Write "<option value='" & ors.fields("PlayerID").value & "'"
				Response.Write ">" & ors.fields("PlayerHandle").value & "</option>" & vbcrlf
				ors.movenext
			loop
		end if
		ors.close
		%>
        <TR valign=center bgcolor=<%=bgcone%>><TD Colspan=2 align="center">
        <input type=hidden name=savetype id=savetype value="Moderator">
        <% if edit then 
			verbage = "Edit Moderator"
		%>
        <input type=hidden name=FPLinkID id=FPLinkID value="<%=Request.QueryString ("ID")%>">
        <input type=hidden name=edit id=edit value="True">
        <% else 
			verbage = "New Moderator"
		%>
        <input type=hidden name=edit id=edit value="False">
        <% end if %>
        <input type=submit name=submitforum id=submitforum value="<%=verbage%>">
        </TD></TR>
        </table>
        </form>
        <% 
    Else
		%>
        <table border=0 cellspacing=1 cellpadding=2 align=center>
		<FORM NAME="FrmPlayerSearch" METHOD="GET" ACTION="default.asp">
		<INPUT TYPE=HIDDEN NAME="AdminType" VALUE="moderator">
		<TR BGCOLOR=<%=bgcone%>>
			<TD>Member Name</TD>
			<TD><INPUT TYPE=TEXT NAME="Player"></TD>
		</TR>
		<TR bgcolor="#000000">
			<TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE=SUBMIT VALUE="Search For Player"></TD>
		</TR>		
		</FORM>
		</TABLE>
		<%
	End If 
	%>
		</td>
	  </tr>
	</table>
	<%
	Call ContentEnd()
end if 
if AdminType = "moderatorf" then 
	Call ContentStart("Moderator Admin")
	%>
    <table width="90%" border="0">
      <tr>
        <td>
	<%
	strsql = "select FPLinkID, PlayerHandle, ForumName from lnk_f_p inner join tbl_forums on tbl_forums.ForumID=lnk_f_p.forumid inner join tbl_players on tbl_players.playerid = lnk_f_p.playerid ORDER BY ForumName ASC"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then 
		response.write "<table border=0 cellspacing=0 cellpadding=2>"
		do while not (ors.eof)
			PlayerHandle = ors.fields("PlayerHandle").value
			if f1 <> orS.FIelds("ForumName").Value Then
				f1 = orS.FIelds("ForumName").Value
				Response.Write "<tr><td>&nbsp;</td></tr><tr><td bgcolor=" & bgcone & "><B>" & orS.FIelds("ForumName").Value & "</b></td></tr>"
			End If
			Response.Write "<TR><TD><p class=small><a href=/viewplayer.asp?player=" & Server.URLEncode("" & PlayerHandle) & ">" & Server.HTMLEncode (PlayerHandle) & "</a> = " & Server.HTMLEncode (ors.fields("ForumName").value) & " - <a href='adminengine.asp?delete=true&table=moderator&ID=" & ors.fields("FPLinkID").value & "'>Delete</a></p></TD></TR>" & vbcrlf
			ors.movenext
		loop
		Response.Write "</table>"
	end if
	ors.close
	edit = false
	strPlayerHandle = Request.QueryString("Player")
	If Len(strPlayerHandle) > 0 Then
	%>
        <form id=moderator name=moderator method=post action="/forums/admin/adminengine.asp">
        <table border=0 cellspacing=0 cellpadding=2 align=center>
        <TR valign=center height=30 bgcolor=<%=bgcone%>><TD>Forum: </td><TD><select name=forumid id=forumid><%
        strsql = "Select ForumName, ForumID from tbl_forums order by ForumName asc"
        ors.open strsql, oconn
        if not (ors.eof and ors.bof) then
			do while not(ors.eof)
				Response.Write "<option value='" & ors.fields("ForumID").value & "'"
				Response.Write ">" & ors.fields("ForumName").value & "</option>" & vbcrlf
				ors.movenext
			loop
		end if
		ors.close
		%>
		</select>
        <TR valign=center height=30 bgcolor=<%=bgctwo%>><TD>Member: </td><TD><select name=playerid id=playerid><%
        strsql = "Select PlayerHandle, PlayerID from tbl_players WHERE PlayerHandle LIKE '%" & CheckString(strPlayerHandle) & "%' order by PlayerHandle asc"
        ors.open strsql, oconn
        if not (ors.eof and ors.bof) then
			do while not(ors.eof)
				Response.Write "<option value='" & ors.fields("PlayerID").value & "'"
				Response.Write ">" & ors.fields("PlayerHandle").value & "</option>" & vbcrlf
				ors.movenext
			loop
		end if
		ors.close
		%>
        <TR valign=center bgcolor=<%=bgcone%>><TD Colspan=2 align="center">
        <input type=hidden name=savetype id=savetype value="Moderator">
        <% if edit then 
			verbage = "Edit Moderator"
		%>
        <input type=hidden name=FPLinkID id=FPLinkID value="<%=Request.QueryString ("ID")%>">
        <input type=hidden name=edit id=edit value="True">
        <% else 
			verbage = "New Moderator"
		%>
        <input type=hidden name=edit id=edit value="False">
        <% end if %>
        <input type=submit name=submitforum id=submitforum value="<%=verbage%>">
        </TD></TR>
        </table>
        </form>
        <% 
    Else
		%>
        <table border=0 cellspacing=1 cellpadding=2 align=center>
		<FORM NAME="FrmPlayerSearch" METHOD="GET" ACTION="default.asp">
		<INPUT TYPE=HIDDEN NAME="AdminType" VALUE="moderator">
		<TR BGCOLOR=<%=bgcone%>>
			<TD>Member Name</TD>
			<TD><INPUT TYPE=TEXT NAME="Player"></TD>
		</TR>
		<TR bgcolor="#000000">
			<TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE=SUBMIT VALUE="Search For Player"></TD>
		</TR>		
		</FORM>
		</TABLE>
		<%
	End If 
	%>
		</td>
	  </tr>
	</table>
	<%
	Call ContentEnd()
end if 
if AdminType = "access" then 
	Call ContentStart("Forum Access")
	%>
	<table align="center">
	<tr><td>
	Click a user to remove their access.<br />
	<%
	strSQL = "SELECT ForumAccessID, PlayerHandle, l.ForumPassword, ForumName FROM lnk_f_p_a l INNER JOIN tbl_players p ON p.PlayerID = l.PlayerID INNER JOIN tbl_forums f ON f.ForumID = l.ForumID ORDER BY PlayerHandle ASC "
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then 
		Do While Not(oRs.EOF)
			%>
			<a href="AdminEngine.asp?SaveType=ForumAccess&FAID=<%=oRs.Fields("ForumAccessID")%>"><%=Server.HTMLEncode(oRs.Fields("PlayerHandle") & " - " & oRs.Fields("ForumName").Value)%></a><br />
			
			<%		
			oRs.MoveNext
		Loop
	End If
	oRs.NextRecordSet
	%>
	</td></tr>
	</table>
	<%	
	Call ContentEnd()
end if
if AdminType = "accessforums" then 
	Call ContentStart("Forum Access")
	%>
	<table align="center">
	<tr><td>
	Click a user to remove their access.<br />
	<%
	strSQL = "SELECT ForumAccessID, PlayerHandle, l.ForumPassword, ForumName FROM lnk_f_p_a l INNER JOIN tbl_players p ON p.PlayerID = l.PlayerID INNER JOIN tbl_forums f ON f.ForumID = l.ForumID ORDER BY ForumName, PlayerHandle ASC "
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then 
		Do While Not(oRs.EOF)
			if f1 <> orS.FIelds("ForumName").Value Then
				f1 = orS.FIelds("ForumName").Value
				Response.Write "<br /><B>" & orS.FIelds("ForumName").Value & "</b><hr />"
			End If
			%>
			<a href="AdminEngine.asp?SaveType=ForumAccess&FAID=<%=oRs.Fields("ForumAccessID")%>"><%=Server.HTMLEncode(oRs.Fields("PlayerHandle") & " - " & oRs.Fields("ForumName").Value)%></a><br />
			
			<%		
			oRs.MoveNext
		Loop
	End If
	oRs.NextRecordSet
	%>
	</td></tr>
	</table>
	<%	
	Call ContentEnd()
end if
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS2 = Nothing
Set oRS = Nothing
%>