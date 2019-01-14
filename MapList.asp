<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle

strPageTitle = "TWL: Map List"

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

Dim strLadderName, intMapCount, intHalf, intCurrent, intLadderID
strLadderName = Request.QueryString("Ladder")

If Not(bSysAdmin Or IsLadderAdmin(strLadderName)) Then
	oConn.Close 
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

If Len(strLadderName) = 0 Then
	oConn.Close 
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/adminops.asp?rAdmin=Ladder"
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
			//if (sortitems) SortD(tbox);
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
<%

strSQL = "SELECT LadderID FROM tbl_ladders WHERE LadderName ='" & CheckString(strLadderName) & "'"
oRS.Open strSQL, oConn
If Not(oRs.EOF and ors.BOF) Then
	intLadderID = oRS.Fields("LadderID").Value 
End If
oRs.NextRecordset 
%>
<form action=saveitem.asp method=post id=form1 name=form1>
<input type=hidden name=ladder value=<%=intLadderID%>>
<input type=hidden name=SaveType value=MapList>
<% 
Call ContentStart("Definition of the map system")
%>
<TABLE WIDTH="97%" BORDER=0 CELLSPACING=0 CELLPADDING=4>
<TR>
	<TD>
	<B>Before you go editing the map list you need to fully understand the way it works.</B>
	<BR>
	What we have is two different ways to define map lists, basically two pools in which maps will be pulled from.
	<BR>
	One is the <B>All Maps</B> pool. Maps listed in this area will be available for selection for maps 1, map 2, and map 3.
	<BR><BR>
	We also have map # specific map pools. These allow us to have different map lists for all maps. <BR>
	This will allow for composite ladders in which map 1 should be C&H, map 2 should be D&D, and map 3 being whatever the defender chooses.<BR>
	<BR>
	If you have any question on how this works, contact Triston, Polaris or Vlasic before modifying the map lists, since 
	this will affect the way map selection works.
	</TD>
</TR>
</TABLE>
<%
Call ContentEnd()
Call Content2BoxStart("ALL MAP POOL: Map Assignments for the " & Server.HTMLEncode(strLadderName) & " Ladder") 
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
			<TD ALIGN=CENTER><SELECT NAME=frm_maplist_map0 MULTIPLE SIZE=10 STYLE="width:300px;background-color: #000000;">
			<%
	
			strSQL = " SELECT m.MapName, m.MapID FROM tbl_maps m "
			strSQL = strSQL & " WHERE m.MapID NOT IN (SELECT MapID FROM lnk_l_m WHERE LadderID = '" & intLadderID & "' AND MapNumber = 0) "
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
			<TH>Maps Currently On Ladder</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>">
			<TD ALIGN=CENTER><SELECT NAME=frm_current_maplist_map0 MULTIPLE SIZE=10 STYLE="width:300px;">
			<%
	
			strSQL = "SELECT m.MapName, m.MapID FROM tbl_maps m, lnk_l_m lnk "
			strSQL = strSQL & " WHERE m.MapID = lnk.MapID AND LadderID = '" & intLadderID & "' AND lnk.MapNumber = 0"
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
		</TABLE>
		</TD></TR>
		</TABLE>
	
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
<% 
Call Content2BoxEnd()
Call Content2BoxStart("Map 1 Pool: Map Assignments for the " & Server.HTMLEncode(strLadderName) & " Ladder") 
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
			<TD ALIGN=CENTER><SELECT NAME=frm_maplist_map1 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = " SELECT m.MapName, m.MapID FROM tbl_maps m "
			strSQL = strSQL & " WHERE m.MapID NOT IN (SELECT MapID FROM lnk_l_m WHERE LadderID = '" & intLadderID & "' AND MapNumber = 1) "
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="Add to map list ---&gt;" NAME="ADD" ONCLICK="javascript:move(this.form.frm_maplist_map1,this.form.frm_current_maplist_map1);"></TD>
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
			<TH>Maps Currently On Ladder</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>">
			<TD ALIGN=CENTER><SELECT NAME=frm_current_maplist_map1 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = "SELECT m.MapName, m.MapID FROM tbl_maps m, lnk_l_m lnk "
			strSQL = strSQL & " WHERE m.MapID = lnk.MapID AND LadderID = '" & intLadderID & "' AND lnk.MapNumber = 1"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="&lt;-- Remove from map list" NAME="ADD" ONCLICK="javascript:move(this.form.frm_current_maplist_map1,this.form.frm_maplist_map1);"></TD>
		</TR>
		</TABLE>
		</TD></TR>
		</TABLE>
	
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
<% 
Call Content2BoxEnd()
Call Content2BoxStart("Map 2 Pool: Map Assignments for the " & Server.HTMLEncode(strLadderName) & " Ladder") 
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
			<TD ALIGN=CENTER><SELECT NAME=frm_maplist_map2 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = " SELECT m.MapName, m.MapID FROM tbl_maps m "
			strSQL = strSQL & " WHERE m.MapID NOT IN (SELECT MapID FROM lnk_l_m WHERE LadderID = '" & intLadderID & "' AND MapNumber = 2) "
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="Add to map list ---&gt;" NAME="ADD" ONCLICK="javascript:move(this.form.frm_maplist_map2,this.form.frm_current_maplist_map2);"></TD>
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
			<TH>Maps Currently On Ladder</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>">
			<TD ALIGN=CENTER><SELECT NAME=frm_current_maplist_map2 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = "SELECT m.MapName, m.MapID FROM tbl_maps m, lnk_l_m lnk "
			strSQL = strSQL & " WHERE m.MapID = lnk.MapID AND LadderID = '" & intLadderID & "' AND lnk.MapNumber = 2"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="&lt;-- Remove from map list" NAME="ADD" ONCLICK="javascript:move(this.form.frm_current_maplist_map2,this.form.frm_maplist_map2);"></TD>
		</TR>
		</TABLE>
		</TD></TR>
		</TABLE>
	
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
<% 
Call Content2BoxEnd() 
Call Content2BoxStart("Map 3 Pool: Map Assignments for the " & Server.HTMLEncode(strLadderName) & " Ladder") 
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
			<TD ALIGN=CENTER><SELECT NAME=frm_maplist_map3 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = " SELECT m.MapName, m.MapID FROM tbl_maps m "
			strSQL = strSQL & " WHERE m.MapID NOT IN (SELECT MapID FROM lnk_l_m WHERE LadderID = '" & intLadderID & "' AND MapNumber = 3)"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="Add to map list ---&gt;" NAME="ADD" ONCLICK="javascript:move(this.form.frm_maplist_map3,this.form.frm_current_maplist_map3);"></TD>
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
			<TH>Maps Currently On Ladder</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>">
			<TD ALIGN=CENTER><SELECT NAME=frm_current_maplist_map3 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = "SELECT m.MapName, m.MapID FROM tbl_maps m, lnk_l_m lnk "
			strSQL = strSQL & " WHERE m.MapID = lnk.MapID AND LadderID = '" & intLadderID & "' AND lnk.MapNumber = 3"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="&lt;-- Remove from map list" NAME="ADD" ONCLICK="javascript:move(this.form.frm_current_maplist_map3,this.form.frm_maplist_map3);"></TD>
		</TR>
		</TABLE>
		</TD></TR>
		</TABLE>
	
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
<% Call Content2BoxEnd()
Call Content2BoxStart("Map 4 Pool: Map Assignments for the " & Server.HTMLEncode(strLadderName) & " Ladder") 
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
			<TD ALIGN=CENTER><SELECT NAME=frm_maplist_map4 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = " SELECT m.MapName, m.MapID FROM tbl_maps m "
			strSQL = strSQL & " WHERE m.MapID NOT IN (SELECT MapID FROM lnk_l_m WHERE LadderID = '" & intLadderID & "' AND MapNumber = 4)"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="Add to map list ---&gt;" NAME="ADD" ONCLICK="javascript:move(this.form.frm_maplist_map4,this.form.frm_current_maplist_map4);"></TD>
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
			<TH>Maps Currently On Ladder</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>">
			<TD ALIGN=CENTER><SELECT NAME=frm_current_maplist_map4 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = "SELECT m.MapName, m.MapID FROM tbl_maps m, lnk_l_m lnk "
			strSQL = strSQL & " WHERE m.MapID = lnk.MapID AND LadderID = '" & intLadderID & "' AND lnk.MapNumber = 4"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="&lt;-- Remove from map list" NAME="ADD" ONCLICK="javascript:move(this.form.frm_current_maplist_map4,this.form.frm_maplist_map4);"></TD>
		</TR>
		</TABLE>
		</TD></TR>
		</TABLE>
	
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
<% Call Content2BoxEnd() 
Call Content2BoxStart("Map 5 Pool: Map Assignments for the " & Server.HTMLEncode(strLadderName) & " Ladder") 
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
			<TD ALIGN=CENTER><SELECT NAME=frm_maplist_map5 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = " SELECT m.MapName, m.MapID FROM tbl_maps m "
			strSQL = strSQL & " WHERE m.MapID NOT IN (SELECT MapID FROM lnk_l_m WHERE LadderID = '" & intLadderID & "' AND MapNumber = 5)"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="Add to map list ---&gt;" NAME="ADD" ONCLICK="javascript:move(this.form.frm_maplist_map5,this.form.frm_current_maplist_map5);"></TD>
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
			<TH>Maps Currently On Ladder</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>">
			<TD ALIGN=CENTER><SELECT NAME=frm_current_maplist_map5 MULTIPLE SIZE=5 STYLE="width:300px;">
			<%
	
			strSQL = "SELECT m.MapName, m.MapID FROM tbl_maps m, lnk_l_m lnk "
			strSQL = strSQL & " WHERE m.MapID = lnk.MapID AND LadderID = '" & intLadderID & "' AND lnk.MapNumber = 5"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="&lt;-- Remove from map list" NAME="ADD" ONCLICK="javascript:move(this.form.frm_current_maplist_map5,this.form.frm_maplist_map5);"></TD>
		</TR>
		<tr BGCOLOR="#000000"><td colspan=3 align=center>
		<input name=button type=button value='Save Map Assignments' class=bright ONCLICK="highlite(this.form.frm_current_maplist_map0);highlite(this.form.frm_current_maplist_map1);highlite(this.form.frm_current_maplist_map2);highlite(this.form.frm_current_maplist_map3);highlite(this.form.frm_current_maplist_map4);highlite(this.form.frm_current_maplist_map5);this.form.submit();">
		</TD></TR>
		</TABLE>
		</TD></TR>
		</TABLE>
	
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
	</form>
<% Call Content2BoxEnd() %>
<% Call Contentstart("Add a New Map")

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
	Response.Write "<tr bgcolor=" & bgcone & " height=30><td colspan=2 align=center><input type=hidden name=SaveType value=NewMap><input type=hidden name=ladder value='" & Server.htmlencode(strLadderName) & "'><input type=submit value='Save New Map' id=submit1 name=submit1 class=bright></td></tr></table></td></tr></table></form>"
	Call ContentEnd()
	
if bSysAdmin then
	Call ContentStart("Delete a map, this affects all ladders and is not reverable")
	Response.Write "<form name=killmap action=saveitem.asp method=post><table align=center width=45% border=0>"
	Response.Write "<tr bgcolor=" & bgctwo & " height=22><td align=center><p>Select a map to remove</p></td></tr><tr bgcolor=" & bgcone & "><td align=center><p class=small><font color=red>This operation cannot be reversed</font></p></td></tr>"
	Response.Write "<tr bgcolor=" & bgctwo & " height=30><td align=center><select name=mapname class=brightred>"
	strSQL="select distinct mapname from tbl_maps order by mapname"
	ors.Open strSQL, oconn
	if not (ors.EOF  and ors.BOF) then
		do while not ors.EOF
			Response.Write "<option>" & Server.HTMLEncode(ors.Fields(0).Value) & "</option>"
			ors.MoveNext
		loop
	end if
	Response.Write "</select></td></tr><tr bgcolor=" & bgcone & " height=25><td align=center><input type=hidden name=SaveType value=KillMap><input type=submit name=sumbit3 class=bright value='Delete Map'></td></tr></table></form>"
	Call ContentEnd()
end if
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

