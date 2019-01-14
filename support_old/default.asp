<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: IRC Support Page"

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

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Teamwarfare League IRC Support Page") %> 
<!-- START OF CONTENT -->

<script language="javascript" type="text/javascript">
   // Validate all form input before sending the request on its way
   function validate ( frm ) {
      if (frm.username.value == '') {
         alert('Please enter a user name.');
         return false;
      }

      if (frm.server.value == '') {
         alert('Please enter an IRC server.');
         return false;
      }

      // All is well - submit the form
      return true;
   }
</script>

<p align="center">
	<strong>Welcome to Teamwarfare.com IRC Chat Client</strong><br />
	<br />
	This page provides a free java-based irc client to aid gamers trying to contact members of staff.<br />
	<br />
	You can also contact members of staff via the <a href="/staff.asp" target="_blank">Contact Us</a> page.<br />
</p>

<form name="login" method="post" onsubmit="return validate(this);" action="irc.php" target="twlsupport">
<table border="0" align="center" width="50%" cellspacing="0" cellpadding="0" class="cssBordered">
<tr bgcolor="<%=bgcone%>">
	<td align="right">User Name:</td>
	<td><input name="username" id="username" value="" type="text" size="30" /></td>
</tr>
<tr bgcolor="<%=bgctwo%>">
	<td align="right">Select Channel:</td>
	<td>
		<select name="channel" id="channel">
      <option>----- General Channels -----</option>
      <option value="/join #teamwarfare" selected>- TeamWarfare Channel</option>
      <option value="/join #twl_support">- Site Support Channel</option>
      <option>----- Game Channels -----</option>
      <option value="/join #twl_aa">- America's Army</option>
      <option value="/join #twl_bf">- Battlefield Series</option>
      <option value="/join #twl_cod">- Call of Duty Series</option>
      <option value="/join #twl_coh">- Company of Heroes</option>
      <option value="/join #twl_cs">- Counter-Strike Series</option>
      <option value="/join #twl_dm">- Dark Messiah</option>
      <option value="/join #twl_dod">- Day of Defeat: Source</option>
      <option value="/join #twl_et">- Enemy Territory</option>
      <option value="/join #twl_etqw">- Enemy Territory: Quake Wars</option>
      <option value="/join #twl_fear">- F.E.A.R.</option>
      <option value="/join #twl_gr">- Ghost Recon Series</option>
      <option value="/join #twl_moh">- Medal of Honor Series</option>
      <option value="/join #twl_q4">- Quake 4</option>
      <option value="/join #twl_ro">- Red Orchestra</option>
      <option value="/join #twl_rtcw">- Return to Caste Wolfenstein</option>
      <option value="/join #twl_rvs">- Rainbow Six: RavenShield</option>
      <option value="/join #twl_ut">- Unreal Tournament Series</option>
      <option value="/join #twl_q3ut">- Urban Terror</option>
      <option value="/join #twl_vegas">- Rainbow Six: Vegas</option>
      <option value="/join #twl_xbox">- XBox Live</option>
      </select>
    </td>
  </tr>
</tr>
<tr>
	<td colspan="2" align="center">
		<input type="submit" value="Log in" onclick="window.open('irc.php', 'twlsupport', 'resizable=no,status=no,location=no,directories=no,menubar=no,copyhistory=no,toolbar=no,scrollbars=no,width=730,height=620')" />
	</td>
</tr>
</table>
</form>

<p align="center">
	<a href="/forums/forumdisplay.asp?ForumID=19">Site Support Forum</a> | 
	<a href="mailto:sitesupport@teamwarfare.com?body=Please see my support request from the Support Page below:%0A%0A">Email Site Support</a>
</p>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>

