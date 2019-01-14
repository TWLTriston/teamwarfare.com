<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Error"

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
Dim ErrorCode, errormsg
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% 
Call ContentStart("Error")
errorcode=request("error")
if errorcode = "" then
	errorcode="NO CODE!!!"
end if
%>
    <table width="90%" border="0">
      <tr>
        <td><p class=small>
        	<% if errorcode = "1" then
        		errormsg = "Error produced because the team founder is a member of another team for the same ladder"
        	   elseif errorcode="2" then
        	   	errormsg = "You must be logged in to perform that action/access that page."
        	   elseif errorcode="3" then
        	   	errormsg = "Proper authorization was not found to access that page/perform that action. Please re-login and retry.<br><br>"
        	   	errormsg = errormsg & "If you feel you should have authorization please e-mail the administrators immediately."
        	   elseif errorcode="4" then
        	   	errormsg = "The team you have attempted to challenge is no longer available to challenge, please choose another team to challenge."
        	   elseif errorcode="5" then
        	   	errormsg = "No shoutcast network specified, please check the linking URL and notify an administrator."
        	   elseif errorcode="6" then
        	   	errormsg = "Your team is unavailable to challenge at this time, this can be caused by a current challenge, or you may have been challenged in the time you attempted to challenge another team."
        	   elseif errorcode="7" then
        	   	errormsg = "The page you have tried to access was passed invalid data, please check your URL and try again."
        	   elseif errorcode="8" then
        	   	errormsg = "No search criteria found, please go back and enter some data!"
        	   elseif errorcode="9" then
        	   	errormsg = "Cannot join competition, team founder found on another team for same competion."
        	   elseif errorcode="10" then
        	   	errormsg = "You must choose two (2) different dates and two (2) different times for a match acceptance. Please press the back button and submit new dates to your opponent."
        	   elseif errorcode="11" then
        	   	errormsg = "Unable to Accept challenge.  No maps linked to ladder. Please be patient, the ladder admin will be adding maps asap."
        	   elseif errorcode="12" then
        	   	errormsg = "You must choose two (2) different times for a match acceptance. Please press the back button and submit new times to your opponent."
        	   elseif errorcode="13" then
        		errormsg = "Specified player does not exist.  Press the Back button and try again."
        	   elseif errorcode="14" then
        		errormsg = "Invalid Player Handle and/or Activation Code."
        	   elseif errorcode="15" then
        		errormsg = "Invalid team name, cannot access history without a valid team name."
        	   elseif errorcode="16" then
        		errormsg = "Your upload did not succeed, most likely because your browser does not support Upload via this mechanism.<br>For Internet Explorer Users:<UL><LI>For Windows 95 or Windows NT 4.0:<UL><LI><A HREF=""http://www.microsoft.com/ie/"">Download</A> V3.02 or later of Internet Explorer<LI><A HREF=""http://www.microsoft.com/ie/download"">Download</A> the File Upload Add-on<LI>For further information, See Knowledge Base Article <A HREF=""http://www.microsoft.com/kb/articles/Q165/2/87.htm"">Q165287</A></UL><LI>For Windows 3.1, WFW 3.11 (Windows 16-bit), or Windows NT 3.51:	<UL><A HREF=""http://www.microsoft.com/ie/"">Download</A> V3.02A or later of Internet Explorer for 16-bit Windows	</UL></UL>For Netscape Users:<UL><LI><A HREF=""http://home.netscape.com"">Download</A> a version of Netscape Navigator or Communicator of 2.x or later</UL>For users of other browsers:<UL><LI>Your browser must support a standard called RFC 1867. Please check with your browser vendor forsupport of this standard.</UL>"        		
						elseif errorcode="17" then
        				errormsg = "The file that you uploaded was empty. Most likely, you did not specify a valid filename to your browser or you left the filename field blank. Please try again."
						elseif errorcode="18" then
        				errormsg = "An error occured in your upload.  Please contact a TWL Admin to resolve this issue."        		
        	   elseif errorcode="19" then
        		errormsg = "Server Error - File1 object not found.  Please contact a TWL Admin."	
        	   elseif errorcode="20" then
        		errormsg = "Unable to accept challenge, no maps are linked to the ladder, please notify an admin immediately."	
        	   elseif errorcode="21" then
        	    errormsg = "Sorry, TWL only accepts .ZIP files for demo uploads.  To get Winzip, visit: <a href=""http://www.winzip.com"">www.winzip.com</a>"
        	   elseif errorcode="22" then
        		errormsg = "Sorry, this demo is intended for admins only."
        	   elseif errorcode="23" then
        	    errormsg = "Sorry, this demo is private (participating team only)."
        	   elseif errorcode="24" then
        	    errormsg = "Sorry, TWL does not accept demos over 20 megabytes in size."        	    
        	   elseif errorcode="25" then
        	    errormsg = "Your selected dates/times overlap with a current reservation."        	    
        	   elseif errorcode="26" then
        	    If Request.QueryString("Msg") <> "" Then errormsg = "You left the following fields blank: " & Request.QueryString("Msg") & "<br><br>"
        	    If Request.QueryString("vMsg") <> "" Then errormsg = errormsg & "The following error(s) occured with your entries:" & Request.QueryString("vMsg")
        	   elseif errorcode="27" then
        	    errormsg = "You must be registered and logged in to use the TWL Instant Messenger. Please click <a href=""/addplayer.asp"">here to register</a>."        	    
        	   elseif errorcode="28" then
        	    errormsg = "This thread has been locked. You may not reply to it."        	    
        	   elseif errorcode="29" then
        	    errormsg = "Team founder was found to be on the roster of another team in the league."        	    
        	   elseif errorcode="30" then
        	    errormsg = "Invalid ladder was passed to map statistics. Please check linking url and try again."        	    
        	   elseif errorcode="31" then
        	    errormsg = "There are no maps assigned to this ladder. Please notify your administrator immediately."        	    
        	   elseif errorcode="32" then
        	    errormsg = "Unable to dispute match. Invalid parameters passed from linking page. Please try again."        	    
         	   elseif errorcode="33" then
        	    errormsg = "Invalid lan id passed. Check the linking URL and try again."        	    
         	   elseif errorcode="34" then
        	    errormsg = "You've earned a forum ban. This ban also applies to rants. You are not allowed to post rants or in the forums."        	    
        	   elseif errorcode="NO CODE!!!" then
        	   	errormsg = "No Code? How the heck did you get here? Press back and have a nice day!"
        	   end if
		  response.write errormsg
        	  %>

        	  </p>
        </td>
      </tr>
      </table><br><br>
      <center>
      <table width=60% cellspacing=0 cellpadding=3 border=0>
      <tr align=center bgcolor=<%=bgcone%>>
       <td>
        <p class=small><font color="#FFD142">If you feel as though this is a valid request, ie you shouldnt get an error message, please use the form below to send an e-mail to the developers. 
        <br>Include an e-mail address for the admins to respond to you.</font></p>
        <form action=errormail.asp method=post id=2>
        <textarea name=MailBody cols=40 rows=10></textarea>
        </td></tr>
        <tr align=center bgcolor=<%=bgcone%>>
        <td>
        <input type=submit name=submit value="Send Mail">
        <input type=hidden name=ErrorCode value='<%=errorcode%>'>
        <input type=hidden name=errormessage value='<%=Server.HTMLEncode(errormsg)%>'>
        </td></tr>
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