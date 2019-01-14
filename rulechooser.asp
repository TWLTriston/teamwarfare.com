<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Rules"

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

Dim RuleName
rulename = Request.QueryString ("set")

Dim CurrentChapter, quesSQL, intGeneralRuleID
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("TeamWarfare Rules") %>

<form name="frmRuleSet" id="frmRuleSet" action="rulechooser.asp" method="get">
<table class="cssBordered" align="center" width="60%">
<tr>
	<th colspan="2">Choose a ruleset</th>
</tr>
<tr>
	<td bgcolor="<%=bgcone%>" align="right" width="75">Game:</td>
	<td bgcolor="<%=bgcone%>">
		<select name="selGame" id="selGame" onChange="fPopulateTypeBox()" onKeyUp="fPopulateTypeBox()">
			<option value="">Select a game</option>
		</select>
	</td>
</tr>
</table>

<div id="divType" style="visibility:hidden; display: none;">
<br />
<table class="cssBordered" align="center" width="60%">
<tr>
	<td bgcolor="<%=bgcone%>" align="right" width="75">Type:</td>
	<td bgcolor="<%=bgcone%>">
		<select name="selType" id="selType" onChange="fPopulateRulesBox();" onKeyUp="fPopulateRulesBox();">
			<option value="">Select a type</option>
		</select>
	</td>
</tr>
</table>
</div>

<div id="divRules" style="visibility:hidden; display: none;">
<br />
<table class="cssBordered" align="center" width="60%">
<tr>
	<td bgcolor="<%=bgctwo%>" align="right" width="75">Rule Set:</td>
	<td bgcolor="<%=bgctwo%>">
		<select name="selRuleSet" id="selRuleSet" onChange="fShowSubmit();" onKeyUp="fShowSubmit();">
			<option value="">Select a rule set</option>
		</select>
	</td>
</tr>
</table>
</div>

<div id="divGoRules" style="visibility:hidden; display: none;">
<br />
<table class="cssBordered" align="center" width="60%">
<tr>
	<td bgcolor="#000000" align="center" colspan="2">
		<input type="checkbox" value="1" name="chkNewWin" id="chkNewWin" checked="checked" class="borderless" /> Open in new window?<br />
		<br />
		<input type="button" value="View Rule Set" onClick="fViewRuleSet();" />
	</td>
</tr>
</table>
</div>

</form>
<script language="javascript" type="text/javascript">
var arrRuleSets = new Array();
<%
strSQL = ""
strSQL = strSQL & "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, HasChildren = (SELECT COUNT(MenuID) FROM tbl_menus m2 WHERE m2.ParentMenuID = m.MenuID) FROM tbl_menus m WHERE m.ParentMenuID = 2 "
strSQL = strSQL & "	OR m.ParentMenuID IN (SELECT MenuID FROM tbl_menus WHERE ParentMenuID = 2) "
strSQL = strSQL & "	OR m.ParentMenuID IN (SELECT MenuID FROM tbl_menus WHERE ParentMenuID IN (SELECT MenuID FROM tbl_menus WHERE ParentMenuID = 2)) "
strSQL = strSQL & " ORDER BY ParentMenuID, ShowMenuName ASC "

' SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL FROM tbl_menus WHERE ParentMenuID = 2 ORDER BY ShowMenuName ASC "
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF)
		%>
		arrRuleSets.push(Array(<%=oRs.Fields("ParentMenuID").Value%>, <%=oRs.Fields("MenuID").Value%>, "<%=Replace(oRs.Fields("ShowMenuName").Value & "", """", "\""")%>", "<%=Server.HTMLEncode(oRs.Fields("LinkURL").Value & "")%>", <%=oRs.Fields("HasChildren").Value%>));<%
		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet
%>

function fPopulateBox(objSelectBox, intParentID) {
	var objOption = null;
	objSelectBox.length = 1;
	for (var i=0;i<arrRuleSets.length;i++) {
		if (arrRuleSets[i][0] == intParentID) {
			objOption = new Option();
			objOption.value = i;
			objOption.text = arrRuleSets[i][2];
			objSelectBox.options[objSelectBox.options.length] = objOption;
		}
	}
}

fPopulateBox(document.frmRuleSet.selGame, 2);

function fPopulateTypeBox() {
	var blnHasChildren = false;
	var objSelectBox = document.frmRuleSet.selGame;
	document.getElementById("divRules").style.visibility = "hidden";
	document.getElementById("divRules").style.display = "none";
	if (objSelectBox.options[objSelectBox.selectedIndex].value == "") {
		document.getElementById("divType").style.visibility = "hidden";
		document.getElementById("divType").style.display = "none";
	} else {
//		alert(arrRuleSets[objSelectBox.options[objSelectBox.selectedIndex].value][4]);
		for (i=0;i<arrRuleSets.length;i++) {
			if ((arrRuleSets[i][0] == arrRuleSets[objSelectBox.options[objSelectBox.selectedIndex].value][1]) && (arrRuleSets[i][4] > 0)) {
				blnHasChildren = true;
			}
		}
		if (blnHasChildren) {
			document.getElementById("divType").style.visibility = "visible";
			document.getElementById("divType").style.display = "inline";
			fPopulateBox(document.frmRuleSet.selType, arrRuleSets[objSelectBox.options[objSelectBox.selectedIndex].value][1]);
			document.frmRuleSet.selType.focus();
		} else {
			document.getElementById("divType").style.visibility = "hidden";
			document.getElementById("divType").style.display = "none";

			document.getElementById("divRules").style.visibility = "visible";
			document.getElementById("divRules").style.display = "inline";
			fPopulateBox(document.frmRuleSet.selRuleSet, arrRuleSets[objSelectBox.options[objSelectBox.selectedIndex].value][1]);
			document.frmRuleSet.selRuleSet.focus();
		}
	}
}

function fPopulateRulesBox() {
	blnHasChildren = false;
	var objSelectBox = document.frmRuleSet.selType;
	//alert(objSelectBox.options[objSelectBox.selectedIndex].value);
	if (objSelectBox.options[objSelectBox.selectedIndex].value == "") {
		document.getElementById("divRules").style.visibility = "hidden";
		document.getElementById("divRules").style.display = "none";
	} else {
		document.getElementById("divRules").style.visibility = "visible";
		document.getElementById("divRules").style.display = "inline";
		fPopulateBox(document.frmRuleSet.selRuleSet, arrRuleSets[objSelectBox.options[objSelectBox.selectedIndex].value][1]);
		document.frmRuleSet.selRuleSet.focus();
	}
}

function fShowSubmit() {
	var objSelectBox = document.frmRuleSet.selRuleSet;
	if (objSelectBox.options[objSelectBox.selectedIndex].value != "" && arrRuleSets[objSelectBox.options[objSelectBox.selectedIndex].value][3].length > 0) {
		document.getElementById("divGoRules").style.visibility = "visible";
		document.getElementById("divGoRules").style.display = "inline";
	} else {
		document.getElementById("divGoRules").style.visibility = "hidden";
		document.getElementById("divGoRules").style.display = "none";
	}
}

function fViewRuleSet() {
	var objRulesWin = null;
	var objSelectBox = document.frmRuleSet.selRuleSet;
	if (objSelectBox.options[objSelectBox.selectedIndex].value != "") {
		if (document.frmRuleSet.chkNewWin.checked) {
			objRulesWin = window.open(arrRuleSets[objSelectBox.options[objSelectBox.selectedIndex].value][3]);
			objRulesWin.focus();
		} else {
			window.location = arrRuleSets[objSelectBox.options[objSelectBox.selectedIndex].value][3];	
		}
	}
}
</script>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
%>

