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

Dim strLadderName, intLadderID
strLadderName = Request.QueryString("ladder")
If Len(strLadderName) = 0 Then
	oConn.Close 
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "GeneralAdmin.asp"
End If

strSQL = "SELECT EloLadderID FROM tbl_elo_ladders WHERE EloLadderName ='" & CheckString(strLadderName) & "'"
oRS.Open strSQL, oConn
If Not(oRs.EOF and ors.BOF) Then
	intLadderID = oRS.Fields("EloLadderID").Value 
End If
oRs.NextRecordset 

If Not(bSysAdmin Or IsEloLadderAdminByID(intLadderID)) Then
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
//					fbox.options[i].value = "";
//					fbox.options[i].text = "";
			   }
			}
//			BumpUp(fbox);
			if (sortitems) SortD(tbox);
		}
		
		function drop(box) {
			for(var i=0; i<box.options.length; i++) {
				if(box.options[i].selected) {
					box.options[i].value = "";
					box.options[i].text = "";
			   }
			}
			BumpUp(box);
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
<input type=hidden name=LadderID value="<%=intLadderID%>">
<input type=hidden name=SaveType value="LadderMapList">
<%
Call Content2BoxStart("Map Assignments for the " & Server.HTMLEncode(strLadderName) & " Ladder") 
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
			strSQL = strSQL & " "
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
	
			strSQL = "SELECT m.MapName, m.MapID FROM tbl_maps m, lnk_elo_maps lnk "
			strSQL = strSQL & " WHERE m.MapID = lnk.MapID AND EloLadderID = '" & intLadderID & "'"
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
			<TD ALIGN=CENTER><INPUT TYPE=BUTTON VALUE="&lt;-- Remove from map list" NAME="ADD" ONCLICK="javascript:drop(this.form.frm_current_maplist_map0);"></TD>
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
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

