<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Age Verification"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open
Session("LoggedIn") = ""
Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("Age Verification - Change")
%>
<script language="javascript" type="text/javascript">
<!--
//function will check to see if a year is a LeapYear
//It will return true if the year is a leap year and false if it is not.
function isLeapYear(year){
	return ((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0);
}

//function to see if a character is a digit.
function isDigit (c) {
	return ((c >= "0") && (c <= "9"))
}//function to see if a string is a number.
function isFloat (s) {
	var i;
	var seenDecimalPoint = false;
	var decimalpointdelimiter = ".";

	if (s == "") {
		return true;
	}

	if (s == decimalpointdelimiter) return false;

	// Search through string's characters one by one
	// until we find a non-numeric character.
	// When we do, return false; if we don't, return true.

	for (i = 0; i < s.length; i++) {   
		// Check that current character is number.
		var c = s.charAt(i);

		if ((c == decimalpointdelimiter) && !seenDecimalPoint)
                   {seenDecimalPoint = true;}
		else 
                   {if (!isDigit(c)) return false;}
	}

	// All characters are numbers.
	return true;
}

//function will check to see if the date is valid
//It will return true if the date is valid and false if it is not.
function isDate(year, mnth, day)
{
	if (!isFloat(day) || !isFloat(mnth) || !isFloat(year)) return false;   // all values must be numeric 	
	if (!day || !mnth || !year) return false;   // no values can be 0 
	if (mnth > 12 || mnth < 1)  return false;   // there's only 12 month
	if (day  > 31 || day  < 1)  return false;   // and 31 day

	if ((day == 31)                             // not each month has 31 days
	&& (mnth != 1 && mnth != 3 && mnth != 5
	&&  mnth != 7 && mnth != 8 && mnth != 10
	&&  mnth != 12)) return false;

	if (mnth == 2){                               // february 
		if (day > 29) return false;           // has maximum 29 days
		if (day == 29 && !isLeapYear(year))   // and only if it's a leap year
			return false;
	}
	return true; 
}

function fSubmit(objForm) {
	var errFlag = 0;
	var errStr = "Error: \n";
	
	if (objForm.txtName.value.length == 0) {
		errFlag = 1;
		errStr = errStr + "You must enter your name.\n";
	}
	if (objForm.txtDOB.value.length < 8) {
		errFlag = 1;
		errStr = errStr + "You must enter your DOB in MM/DD/YYYY format.\n";
	} else {
		var sDOB = objForm.txtDOB.value;
		if (sDOB.indexOf("/", 0) == 1) {
			sDOB = "0" + sDOB;
		}
		if (sDOB.indexOf("/", 3) == 4) {
			sDOB = sDOB.substr(0,3) + "0" + sDOB.substr(3);
		}
		
		var sDay = Number(sDOB.substr(3,2));
		var sMon = Number(sDOB.substr(0,2));
		var sYear = Number(sDOB.substr(6,4));

		if (!isDate(sYear, sMon, sDay)) {
			errFlag = 1;
			errStr = errStr + "You must enter a valid DOB.\n";
		} else {
			var oDate = new Date();
			var sCDay = oDate.getDay();
			var sCMon = oDate.getMonth();
			var sCYear = oDate.getYear() - 13;
			var oDOB = new Date(sYear, sMon, sDay);
			var oCD = new Date(sCYear, sCMon, sCDay);
			
			if (oDOB > oCD) {
				alert("Under 13");
				errFlag = 1;
				errStr = errStr + "You are under the age of 13. You may not use TeamWarfare.\n";
			} 
		}
			
	}

	if (errFlag == 0) {
		if (confirm("Are you certain the information you are submitting is accurate?")) {
			objForm.submit();
		}
	} else {
		alert(errStr);
	}
}

//-->
</script>
Oops! You've clicked the wrong link. Before you can continue to use TeamWarfare, you 
must fill out the form below.
<br /><br />
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
<tr>
	<td>
	<table border="0" cellspacing="1" cellpadding="4">
	<form name="frmCoppa" id="frmCoppa" action="saveItem.asp" method="post">
	<input type="hidden" name="SaveType" id="SaveType" value="CoppaChange" />
	<tr>
		<th colspan="2" bgcolor="#000000">Change COPPA Status</th>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=bgcone%>"><b>Your Full Name</b></td>
		<td bgcolor="<%=bgcone%>"><input type="text" name="txtName" id="txtName" size="30" style="width: 200px;" maxlength="200" /></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right"><b>Your Date of Birth<br />(MM/DD/YYYY)</b></td>
		<td bgcolor="<%=bgctwo%>"><input type="text" name="txtDOB" id="txtDOB" size="30" style="width: 100px;" maxlength="200" /></td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#000000" align="center"><input type="button" onclick="fSubmit(this.form)" value="Submit Change" /></td>
	</tr>
	</form>
	</table>
	</td>
</tr>
</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>