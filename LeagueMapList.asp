<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle

strPageTitle = "TWL: League Map List"

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

Dim strLeagueName, intMapCount, intHalf, intCurrent, intLeagueID
strLeagueName = Request.QueryString("League")

If Len(strLeagueName) = 0 Then
	oConn.Close 
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/LeagueAdmin.asp"
End If

strSQL = "SELECT LeagueID FROM tbl_leagues WHERE LeagueName ='" & CheckString(strLeagueName) & "'"
oRS.Open strSQL, oConn
If Not(oRs.EOF and ors.BOF) Then
	intLeagueID = oRS.Fields("LeagueID").Value 
End If
oRs.NextRecordset 

If Not(bSysAdmin Or IsLeagueAdminByID(intLeagueID)) Then
	oConn.Close 
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
	
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<SCRIPT LANGUAGE="javascript">
<!--

		var sortitems = 1;

		function move(fbox,tbox) {
			for(var i=0; i<fbox.options.length; i++) {
				if(fbox.options[i].selected && fbox.options[i].value != "") {
					var no = new Option();
					no.value = fbox.options[i].value;
					no.text = fbox.options[i].text;
					tbox.options[tbox.options.length] = no;
					fbox.options[i].value = "";
					fbox.options[i].text = "";
			   }
			}
			BumpUp(fbox);
			if (sortitems) SortD(tbox);
		}
		
		function BumpUp(box)  {
			for(var i=0; i<box.options.length; i++) {
				if(box.options[i].value == "")  {
					for(var j=i; j<box.options.length-1; j++)  {
						box.options[j].value = box.options[j+1].value;
						box.options[j].text = box.options[j+1].text;
					}
				var ln = i;
				break;
			   }
			}
			if(ln < box.options.length)  {
				box.options.length -= 1;
				BumpUp(box);
		   }
		}

		function SortD(box)  {
			var temp_opts = new Array();
			var temp = new Object();
			for(var i=0; i<box.options.length; i++)  {
				temp_opts[i] = box.options[i];
			}
			for(var x=0; x<temp_opts.length-1; x++)  {
				for(var y=(x+1); y<temp_opts.length; y++)  {
					if(temp_opts[x].text > temp_opts[y].text)  {
						temp = temp_opts[x].text;
						temp_opts[x].text = temp_opts[y].text;
						temp_opts[y].text = temp;
						temp = temp_opts[x].value;
						temp_opts[x].value = temp_opts[y].value;
						temp_opts[y].value = temp;
					}
				}
			}
			for(var i=0; i<box.options.length; i++)  {
				box.options[i].value = temp_opts[i].value;
				box.options[i].text = temp_opts[i].text;
		   }
		}
		
		function highlite(box) {
			for (var i=0;i<box.options.length;i++) {
				box.options[i].selected = 1;
			}
		
		}
		
//-->
</SCRIPT>

<form action=saveitem.asp method=post id=form1 name=form1>
<input type=hidden name=LeagueID value="<%=intLeagueID%>">
<input type=hidden name=SaveType value="LeagueMapList">
<%
Call Content2BoxStart("Map Assignments for the " & Server.HTMLEncode(strLeagueName) & " League") 
%>
	<table width=780 border="0" cellspacing="0" cellpadding="0" BACKGROUND="">
	<tr>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	<td width=380 VALIGN=TOP>
		<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER WIDTH="375">
		<TR><TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<TR BGCOLOR="#000000">
			<TH>Maps on TWL</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>">
			<TD ALIGN=CENTER><SELECT NAME=frm_maplist_map0 MULTIPLE SIZE=10 STYLE="width:300px;">
			<%
	
			strSQL = " SELECT m.MapName, m.MapID FROM tbl_maps m "
			strSQL = strSQL & " WHERE m.MapID NOT IN (SELECT MapID FROM lnk_league_maps WHERE leagueid = '" & intLeagueID & "') "
			strSQL = strSQL & " ORDER BY m.MapName "
			oRs.Open strSQL, oConn
			if not (oRS.EOF and oRS.bof) then
				bgc=bgcone
				do while not oRs.EOF
					Response.Write "<OPTION VALUE=""" & oRS.Fields("MapID").Value & """>" & Server.HTMLEncode(oRS.Fields("MapName").Value & "") & "</OPTION>" & vbCrLf
					oRS.MoveNext 
				Loop
			End If
			oRS.NextRecordset 
			%>
			</SELECT>
			</TD>
		</TR>
		<TR BGCOLOR="<%=bgctwo%>">
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="Add to map list ---&gt;" NAME="ADD" ONCLICK="javascript:move(this.form.frm_maplist_map0,this.form.frm_current_maplist_map0);"></TD>
		</TR>
		</TABLE>
		</TD></TR>
		</TABLE>
		</td>
		<td><img src="/images/spacer.gif" width="10" height="1"></td>
		<td width=379 VALIGN=TOP>
		
		<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER WIDTH="375">
		<TR><TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<TR BGCOLOR="#000000">
			<TH>Maps Currently On League</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>">
			<TD ALIGN=CENTER><SELECT NAME=frm_current_maplist_map0 MULTIPLE SIZE=10 STYLE="width:300px;">
			<%
	
			strSQL = "SELECT m.MapName, m.MapID FROM tbl_maps m, lnk_league_maps lnk "
			strSQL = strSQL & " WHERE m.MapID = lnk.MapID AND LeagueID = '" & intLeagueID & "'"
			strSQL = strSQL & " ORDER BY m.MapName "
			oRs.Open strSQL, oConn
			if not (oRS.EOF and oRS.bof) then
				bgc=bgcone
				do while not oRs.EOF
					Response.Write "<OPTION VALUE=""" & oRS.Fields("MapID").Value & """>" & Server.HTMLEncode(oRS.Fields("MapName").Value & "") & "</OPTION>" & vbCrLf
					oRS.MoveNext 
				Loop
			End If
			oRS.NextRecordset 
			%>
			</SELECT>
			</TD>
		</TR>
		<TR BGCOLOR="<%=bgctwo%>">
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="&lt;-- Remove from map list" NAME="ADD" ONCLICK="javascript:move(this.form.frm_current_maplist_map0,this.form.frm_maplist_map0);"></TD>
		</TR>
		<tr>
			<td align="center" bgcolor="#000000"><input name=button type=button value='Save Map Assignments' class=bright ONCLICK="highlite(this.form.frm_current_maplist_map0);this.form.submit();"></td>
		<tr>
		</TABLE>
		</TD></TR>
		</TABLE>
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
		</tr>
	</table>
	</form>
<% 
Call Content2BoxEnd()
Call Contentstart("Add a New Map")

	Response.Write "<form name=newmap action=saveitem.asp method=post>"
	%>
	<table align=center width=760 cellspacing=0 cellpadding=0 border=0 BGCOLOR="#444444">
	<TR><TD>
	<table align=center width=100% cellspacing=1 cellpadding=2 border=0>
	<%
	Response.Write "<tr bgcolor=" & bgcone & " ><td align=right><p class=small>Map Name:</p></td><td>&nbsp;<input name=mapname type=text class=bright></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & " ><td align=right><p class=small>Abbreviation:</p></td><td>&nbsp;<input name=mapabbr type=text class=bright></td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & " ><td align=right><p class=small>Type (CTF, C&H, ...):</p></td><td>&nbsp;<input name=maptype type=text class=bright></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & " ><td align=right><p class=small>Terrain Type:</p></td><td>&nbsp;<input name=mapterr type=text class=bright></td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & " ><td align=right><p class=small>Generators (Total):</p></td><td>&nbsp;<input name=mapgens type=text class=bright></td></tr>"		
	Response.Write "<tr bgcolor=" & bgctwo & " ><td align=right><p class=small>Inventories (Total):</p></td><td>&nbsp;<input name=maptInv type=text class=bright></td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & " ><td align=right><p class=small>Base Turrets (Total):</p></td><td>&nbsp;<input name=mapturr type=text class=bright></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & " ><td align=right rowspan=2 valign=top><p class=small>Vehicle Pad?:</p></td><td>&nbsp;<input name=mapV type=radio class=borderless value=1 checked>Yes</td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & " ><td>&nbsp;<input name=mapV type=radio class=borderless value=0>No</td></tr>"
	Response.Write "<tr bgcolor=" & bgcone & " ><td align=right valign=top><p class=small>Description:</p></td><td>&nbsp;<textarea name=mapdesc rows=5 cols=40></textarea></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & " ><td align=right><p class=small>Image File Name:</p></td><td>&nbsp;<input name=mapimage type=text value='None.gif' class=bright></td></tr>"
	%>
	<%
	Response.Write "<tr bgcolor=" & bgcone & " height=30><td colspan=2 align=center><input type=hidden name=SaveType value=NewMap><input type=hidden name=league value='" & Server.htmlencode(strLeagueName) & "'><input type=submit value='Save New Map' id=submit1 name=submit1 class=bright></td></tr></table></td></tr></table></form>"
	Call ContentEnd()
	
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

