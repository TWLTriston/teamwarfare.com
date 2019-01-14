<%' Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & verbage & " Reply"

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
<!-- #INCLUDE virtual="/include/i_funclib.asp" -->
<%
if not (bSysAdmin) then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
end if

if Request.form("SaveType") = "Category" then
	if Request.Form("Edit") = "True" then
		categoryID = Request.Form("CatID")
	end if
	CategoryName = replace(Request.form("catname"), "'", "''")
	CategoryDesc = replace(Request.form("description"), "'", "''")
	CategoryOrder = replace(Request.form("catorder"), "'", "''")
	if Request.Form("Edit") = "True" then
		strsql = "update tbl_category set CategoryName ='" & CategoryName 
		strsql = strsql & "', CategoryDesc='" & CategoryDesc & "', CategoryOrder='" & CategoryOrder & "' "
		strsql = strsql & " where CategoryID='" & CategoryID & "'"
	else
		strsql = "insert into tbl_category (CategoryName, CategoryDesc, CategoryOrder) values " 
		strsql = strsql & " ('" & CategoryName & "', '" & CategoryDesc & "', '" & CategoryOrder & "')"
	end if
	Response.Write strsql 
	ors.Open strsql, oconn
end if

if Request.form("SaveType") = "forum" then
	if Request.Form("Edit") = "True" then
		ForumID = Request.Form("ForumID")
	end if

	ForumName = replace(Request.form("ForumName"), "'", "''")
	ForumPass = replace(Request.form("ForumPass"), "'", "''")
	ForumDesc = replace(Request.form("ForumDesc"), "'", "''")
	ForumOrder = replace(Request.form("ForumOrder"), "'", "''")
	CategoryID = replace(Request.form("CategoryID"), "'", "''")
	ForumMatchDispute = replace(Request.form("ForumMatchDispute"), "'", "''")
	if Request.Form("Edit") = "True" then
		strsql = "update tbl_forums set ForumName ='" & ForumName 
		strsql = strsql & "', ForumPassword='" & ForumPass & "', ForumOrder='" & ForumOrder & "', CategoryID='" & CategoryID
		strsql = strsql & "', ForumDescription='" & ForumDesc
		strsql = strsql & "', ForumMatchDispute='" & ForumMatchDispute & "' "
		strsql = strsql & " where ForumID='" & ForumID & "'"
	else
		strsql = "insert into tbl_forums (ForumName, ForumPassword, ForumOrder, CategoryID, ForumDescription, ForumMatchDispute) values " 
		strsql = strsql & " ('" & ForumName & "', '" & ForumPass & "', '" & ForumOrder & "', '" & CategoryID & "', '" & ForumDesc & "', '" & ForumMatchDispute & "')"
	end if
	Response.Write strsql 
	ors.Open strsql, oconn
end if

if Request.QueryString("Delete") = "true" then
	if Request.QueryString("table") = "category" then
		strsql = "delete from tbl_category where CategoryID=" & Request.QueryString("ID")
		Response.Write strsql
		ors.Open strsql, oconn
	end if
	if Request.QueryString("table") = "forum" then
		strsql = "delete from tbl_forums where ForumID=" & Request.QueryString("ID")
		Response.Write strsql
		ors.Open strsql, oconn
	end if
	if Request.QueryString("table") = "replace" then
		strsql = "delete from tbl_replace where ReplaceID=" & Request.QueryString("ID")
		Response.Write strsql
		ors.Open strsql, oconn
	end if
	if Request.QueryString("table") = "moderator" then
		strsql = "delete from lnk_f_p where FPLinkID=" & Request.QueryString("ID")
		Response.Write strsql
		ors.Open strsql, oconn
	end if
end if

if Request.form("SaveType") = "Replace" then
	if Request.Form("Edit") = "True" then
		ReplaceID = Request.Form("ReplaceID")
	end if
	ReplaceCategory = replace(Request.form("ReplaceCategory"), "'", "''")
	ReplaceSearch = replace(Request.form("ReplaceSearch"), "'", "''")
	ReplaceFiller = replace(Request.form("ReplaceFiller"), "'", "''")
	if Request.Form("Edit") = "True" then
		strsql = "update tbl_replace set ReplaceSearch='" & ReplaceSearch & "', ReplaceFiller='" & replaceFiller & "', ReplaceCategory='" & ReplaceCategory & "' "
		strsql = strsql & "where ReplaceID = '" & ReplaceID & "'"
	else
		strsql = "insert into tbl_replace (ReplaceSearch, replaceFiller, ReplaceCategory) values " 
		strsql = strsql & " ('" & ReplaceSearch & "', '" & ReplaceFiller & "', '" & ReplaceCategory & "')"
	end if
	Response.Write strsql 
	ors.Open strsql, oconn
end if

if Request.form("SaveType") = "Moderator" then
	if Request.Form("Edit") = "True" then
		FPlinkID = Request.Form("FPlinkID")
	end if
	ForumID = replace(Request.form("ForumID"), "'", "''")
	PlayerID = replace(Request.form("PlayerID"), "'", "''")
	if Request.Form("Edit") = "True" then
		strsql = "update lnk_f_P set ForumID='" & ForumID & "', PlayerID='" & PlayerID & "' "
		strsql = strsql & "where FPLInkID = '" & FPLInkID & "'"
	else
		strsql = "insert into lnk_f_P (ForumID, PlayerID) values " 
		strsql = strsql & " ('" & ForumID & "', '" & PlayerID & "')"
	end if
	Response.Write strsql 
	ors.Open strsql, oconn
end if

if Request.QueryString("SaveType") = "ForumAccess" then
	iForumAccessID = Request.QueryString("FAID")
	strSQL = "DELETE FROM lnk_f_p_a WHERE ForumAccessID = '" & iForumAccessID & "'"
	oConn.Execute(strSQL)
End if

oConn.Close
set oConn = nothing
set oRs = nothing
set oRs2 = nothing
set oRs3 = nothing
Response.Clear
Response.Redirect "/forums/admin/"
%>

