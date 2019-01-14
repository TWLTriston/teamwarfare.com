<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Privacy Policy"

Dim strSQL, oConn, oRS
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

<%
Call ContentStart("Privacy Policy ")
%>
<table border="0" width="97%" align="center">
<tr><td>
TeamWarfare is committed to respecting the privacy rights of all visitors to our website. We also recognize that our visitors, and the parents of children visiting our site, need protection of any personally identifying information they choose to share with us. This privacy policy is intended to provide you with the information on how we collect, use and store the information that you provide, so that you can make appropriate choices for sharing information with us. <br />
<br />
If you have any questions, complaints, or comments regarding our privacy statement or policies, please contact us by email at <a href="mailto:triston@teamwarfare.com">triston@teamwarfare.com</a>. Additional mail contact information is provided below.
<br />

<b>Links</b><br />
This policy applies only to site(s) maintained by TeamWarfare. We may create links to the websites of our promotional partners and companies for games that are supported through our competitions. Although we will make every effort to link only to sites that meet similar standards for maintaining each individual's right to privacy, this policy does not apply to third party sites. <br />
<br />
Additionally, many other sites that we are not associated with and have not authorized may have links leading to our site. We cannot control these links and are not responsible for any content appearing on these sites. Furthermore, we do not support links to third party sites in our forums, and we are not responsible for any content appearing on those sites. <br />
<br />

<b>Collection of Personal Information through our Website</b><br />
We do not require that visitors reveal any personally identifying information in order to access our website. However, visitors who do not wish to, or are not allowed by law to share personally identifying information, may not be able to access certain areas or create a player account, which requires registration.<br />
<br />
Although information is required to participate in certain competitions offered through our website, participants provide information on a voluntary basis only. <br />
<br />
Collection of personal information for registration requires your email address only. Additional information may be required for specific competitions, including your name, address, telephone number, date of birth, and/or credit card number. <br />
<br />

<b>Collection of Demographic Information through www.teamwarfare.com</b><br />
Currently, we do not collect any personally identifying demographic information from our members.<br />
<br />

<b>How We Use Your Information </b><br />
Personally identifying information will be saved and used only for user validation, and any prize fulfillment that is necessary. <br />
<br />
When you have provided personally identifying information for a particular purpose, we may disclose your information to certain third parties that we have engaged to assist us in fulfilling your request and in maintaining our database records. This includes, but is not limited to, vendors, affiliates, and prize providers. These companies have agreed to maintain strict security over all personal information received from us, and will not re-distribute your information. TeamWarfare may share information with appropriate third parties, including law enforcement or other similar entity, in connection with a criminal investigation, investigation of fraud, infringement of intellectual property rights, or other activity, which is suspected to be illegal or may expose you or TeamWarfare to legal liability. <br />
<br />

<b>Policies for Children 12 Years Old and Under</b><br />
We will not collect personally identifying information through our website from children 12 and under. You must be 13 years or older to register and participate on TeamWarfare.<br />
<br />

<b>Cookies</b><br />
Cookies are bits of electronic information which a website can transfer to a visitor's hard drive to help tailor and keep records of his or her visit at the site. Cookies allow website operators to better tailor visits to the site to visitor's individual preferences. The use of cookies is standard on the Internet and many major web sites use them. You can choose to set your browser to notify you whenever you are sent a cookie. This gives you the chance to decide whether or not to accept it.<br />
<br />
You must have cookies enabled to register and login to TeamWarfare. We use cookies in certain areas to improve your experience when visiting our websites. For instance, cookies will be used to manage each session you visit our site to make moving around our site more efficient and enjoyable for you. Information collected through cookies or similar techniques will remain anonymous and will not be connected to any user's personally identifying information without his or her consent.<br />
<br />

<b>Third Party Cookies</b><br />
In the course of serving advertisements to this site, our third-party advertiser may place or recognize a unique "cookie" on your browser.<br />
<br />

<b>Log Files</b><br />
We use IP addresses to analyze trends, administer the site, track user's movement, and gather broad demographic information for aggregate use. IP addresses are not linked to personally identifiable information. <br />
<br />

<b>Site and Service Updates</b><br />
We periodically send site and service announcement updates. Members are not able to unsubscribe from service announcements, which contain important information about the service. <br />
<br />

<b>Message Boards / Forums</b><br />
Our message boards are a place where users can go to freely share their thoughts and ideas about TeamWarfare and the respective games it supports. We ask our users to respect the privacy of others. Posting of phone numbers, addresses, or other personal information in the Forums section, which may violate someone's privacy is prohibited. Additionally, we prohibit the posting by any user of his or her own personally identifying information. While we try to moderate our forums, removing malicious and pornographic links / images, some inappropriate items may remain. Our staff can be reached via email <a href="/staff.asp">here</a>. <br />
<br />
We reserve the right to remove any postings that violate this rule or any other term or condition posted on this site. Additionally, we reserve the right to ban any user who has violated this requirement. <br />
<br />

<b>Accuracy & Security</b><br />
We will take appropriate steps to protect all information our visitors share with us. This includes setting up processes to avoid any unauthorized access or disclosure of this information. We will also use our best efforts to maintain accurate personal information collect from our website visitors.<br />
<br />
We collect personally identifying information to the extent deemed reasonably necessary to serve our legitimate business purposes. We maintain safeguards such as: (i) providing consumer access to data for purposes of verification and correction; and (ii) providing technical security measures to ensure the security, integrity, accuracy and privacy of the information you have provided. We also take reasonable steps to assure that third parties to which we transfer any data will provide sufficient protection of that personal information.<br />
<br />

<b>Third Party Advertising</b><br />
We use MaxOnline and other third-party advertising companies to serve ads when you visit our Web site. 
These companies may use information (not including your name, address, email address or telephone number) about your visits to this and other Web sites in order 
to provide advertisements on this site and other sites about goods and services that may be of interest to you. 
If you would like more information about this practice and to know your choices about not having this information used 
by these companies, please <a href="http://www.maxonline.com/privacy_policy/index.php" target="_blank">click here</a>.<br />
<br />

<b>Contact Us</b><br />
For further information on our privacy policy, or for questions on information that we may have collected from you or your children, or should you wish to have your name removed from our records, please contact us by either of the following methods and we will be happy to review, update or remove information as appropriate:<br />
<br />
By Email at: <a href="mailto:triston@teamwarfare.com">triston@teamwarfare.com</a><br />
<br />
or by Mail at: <br />
Myles Angell<br />
17 Spring St<br />
Collinsville, CT 06019<br />
<br />
Updated May 6th, 2005
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