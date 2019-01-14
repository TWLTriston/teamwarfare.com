<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Terms and Conditions"

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

<% Call ContentStart("TERMS AND CONDITIONS FOR USE OF THE TEAMWARFARE WEBSITE") %>
<p>By accessing this Website, you agree to be bound by the terms and conditions appearing in this document and you accept our Privacy Policy which is available <a href="/privacy.asp">here</a>. If there is anything you do not understand please email any enquiry to <a href="mailto:qualitycontrol@teamwarfare.com">qualitycontrol@teamwarfare.com</a>.</p>

<p>In these terms and conditions "We/us/our/Teamwarfare" means TeamWarfare League, "Website" means the website located at <a href="http://www.teamwarfare.com">http://www.teamwarfare.com</a> (or any subsequent URL which may replace it) and all associated websites and micro sites of the TeamWarfare League and "You/your" means you as a user of the Website.</p>

<p>You shall not use the Website for any illegal purposes, and you will use it in compliance with all applicable laws and regulations. You agree not to use the Website in a way that may cause the Website to be interrupted, damaged, rendered less efficient or such that the effectiveness or functionality of the Website is in any way impaired;</p>

<p>You agree not to attempt any unauthorised access to any part or component of the Website; and You agree that in the event that you have any right, claim or action against any Users arising out of that User's use of the Website, then you will pursue such right, claim or action independently of and without recourse to us.</p>

<p>YOU AGREE TO BE FULLY RESPONSIBLE FOR (AND FULLY INDEMNIFY US AGAINST) ALL CLAIMS, LIABILITY, DAMAGES, LOSSES, COSTS AND EXPENSES, INCLUDING LEGAL FEES, SUFFERED BY US AND ARISING OUT OF ANY BREACH OF THE CONDITIONS BY YOU OR ANY OTHER LIABILITIES ARISING OUT OF YOUR USE OF THE WEBSITE, OR THE USE BY ANY OTHER PERSON ACCESSING THE WEBSITE USING YOUR PC OR INTERNET ACCESS ACCOUNT.</p>

<p>We reserve the right to modify or withdraw, temporarily or permanently, the Website (or any part of) with or without notice to you and you confirm that we shall not be liable to you or any third party for any modification to or withdrawal of the Website.</p>

<p>We may alter these terms and conditions from time to time, and your use of the Website (or any part of) following such change shall be deemed to be your acceptance of such change. It is your responsibility to check regularly to determine whether the Conditions have been changed. If you do not agree to any change to the Conditions then you must immediately stop using the Website.</p>

<p>The Website is subject to constant change. You will not be eligible for any compensation because you cannot use any part of the Website or because of a failure, suspension or withdrawal of all or part of the Website.</p>

<p>We are not responsible for the availability of any external sites or resources, and do not endorse and are not responsible or liable, directly or indirectly, for the privacy practices or the content (including misrepresentative or defamatory content) of any third party websites, including (without limitation) any advertising, products or other materials or services on or available from such websites or resources, nor for any damage, loss or offence caused or alleged to be caused by, or in connection with, the use of or reliance on any such content, goods or services available on such external sites or resources.</p>

<p>We have the right, but not the obligation, to monitor any activity and content associated with the Website. We may investigate any reported violation of these Conditions or complaints and take any action that we deem appropriate (which may include, but is not limited to, issuing warnings, suspending, terminating or attaching conditions to your access and/or removing any materials from the Website).</p>

<p>You acknowledge and agree that all copyright, trademarks and all other intellectual property rights in all material or content supplied as part of the Website shall remain at all times vested in us or our licensors. You are permitted to use this material only as expressly authorised by us.</p>

<p>You acknowledge and agree that the material and content contained within the Website is made available for your personal non-commercial use only and that you may download such material and content onto only one computer hard drive for such purpose. Any other use of the material and content of the Website is strictly prohibited. You agree not to (and agree not to assist or facilitate any third party to) copy, reproduce, transmit, publish, display, distribute, commercially exploit or create derivative works of such material and content.</p>

<p>The Website is Copyright, TeamWarfare League. All rights reserved.</p>

<p>WE MAKE NO WARRANTIES, WHETHER EXPRESS OR IMPLIED IN RELATION TO THE ACCURACY OF ANY INFORMATION WE PLACE ON THE WEBSITE. THE WEBSITE IS PROVIDED ON AN "AS IS" AND "AS AVAILABLE" BASIS WITHOUT ANY REPRESENTATION OR ENDORSEMENT. UNLESS SPECIFIED IN SEPARATE TERMS AND CONDITIONS RELATED TO A PARTICULAR PRODUCT OR SERVICE, WE MAKE NO WARRANTIES OF ANY KIND, WHETHER EXPRESS OR IMPLIED, IN RELATION TO THE WEBSITE, OR PRODUCTS OR SERVICES OFFERED ON THE WEBSITE WHETHER BY US OR ON OUR BEHALF (INCLUDING SOFTWARE DOWNLOADS) INCLUDING BUT NOT LIMITED TO, IMPLIED WARRANTIES OF SATISFACTORY QUALITY, FITNESS FOR A PARTICULAR PURPOSE, NON-INFRINGEMENT, COMPATIBILITY, SECURITY, ACCURACY, CONDITION OR COMPLETENESS, OR ANY IMPLIED WARRANTY ARISING FROM COURSE OF DEALING OR USAGE OR TRADE.</p>

<p>UNLESS SPECIFIED IN SEPARATE TERMS AND CONDITIONS RELATED TO A PARTICULAR PRODUCT OR SERVICE, WE MAKE NO WARRANTY THAT THE WEBSITE OR PRODUCTS OR SERVICES OFFERED ON THE WEBSITE WHETHER BY US OR ON OUR BEHALF (INCLUDING SOFTWARE DOWNLOADS) WILL MEET YOUR REQUIREMENTS OR WILL BE UNINTERRUPTED, TIMELY, SECURE OR ERROR-FREE, THAT DEFECTS WILL BE CORRECTED, OR THAT THE WEBSITE OR THE SERVER THAT MAKES IT AVAILABLE OR PRODUCTS OR SERVICES OFFERED ON THE WEBSITE WHETHER BY US OR ON OUR BEHALF (INCLUDING SOFTWARE DOWNLOADS) ARE FREE OF VIRUSES OR BUGS OR ARE FULLY FUNCTIONAL, ACCURATE, OR RELIABLE. WE WILL NOT BE RESPONSIBLE OR LIABLE TO YOU FOR ANY LOSS OF CONTENT OR MATERIAL AS A RESULT OF UPLOADING TO OR DOWNLOADING FROM THE WEBSITE.</p>

<p>YOU ACKNOWLEDGE THAT WE CANNOT GUARANTEE AND THEREFORE SHALL NOT BE IN ANY WAY RESPONSIBLE FOR THE SECURITY OR PRIVACY OF THE WEBSITE AND ANY INFORMATION PROVIDED TO OR TAKEN FROM THE WEBSITE BY YOU.</p>

<p>WE WILL NOT BE LIABLE, IN CONTRACT, TORT (INCLUDING, WITHOUT LIMITATION, NEGLIGENCE), PRE-CONTRACT OR OTHER REPRESENTATIONS (OTHER THAN FRAUDULENT MISREPRESENTATIONS) OR OTHERWISE OUT OF OR IN CONNECTION WITH THE WEBSITE OR PRODUCTS OR SERVICES OFFERED ON THE WEBSITE WHETHER BY US OR ON OUR BEHALF (INCLUDING FREE SOFTWARE DOWNLOADS) FOR ANY ECONOMIC LOSSES (INCLUDING WITHOUT LIMITATION LOSS OF REVENUES, PROFITS, CONTRACTS, BUSINESS OR ANTICIPATED SAVINGS) OR ANY LOSS OF GOODWILL OR REPUTATION, OR ANY LOSS OR CORRUPTION OF DATA, OR ANY SPECIAL OR INDIRECT OR CONSEQUENTIAL LOSSES; IN ANY CASE WHETHER OR NOT SUCH LOSSES WERE WITHIN THE CONTEMPLATION OF EITHER OF US AT THE DATE ON WHICH THE EVENT GIVING RISE TO THE LOSS OCCURRED.</p>

<p>We will not be liable in contract, tort or otherwise if you incur loss or damage connecting to the Website through a third party's hypertext link.</p>

<p>Notwithstanding any other provision in the Conditions, nothing shall limit your rights as a consumer under United States law where or insofar as such rights cannot be derogated from by contract.</p>

<p>Nothing in the Conditions shall exclude or limit our liability for death or personal injury resulting from our negligence or that of our servants, agents or employees.</p>

<p>If any part of the Conditions shall be deemed unlawful, void or for any reason unenforceable, then that provision shall be deemed to be severable from these Conditions and shall not effect the validity and enforceability of any of the remaining provisions of the Conditions.</p>

<p>Nothing shall be construed as a waiver by us of any preceding or succeeding breach of any provision.</p>

<p>These Conditions and documents referred to herein (as amended from time to time) contain the entire agreement between you and us relating to the subject matter covered and supersede any previous agreements, arrangements, undertakings or proposals, written or oral, between you and us in relation to such matters. No oral explanation or oral information given by either of us shall alter the interpretation of these Conditions. You confirm that, in agreeing to accept these Conditions, you have not relied on any representation save insofar as the same has expressly been made a representation in these Conditions and you agree that you shall have no remedy in respect of any misrepresentation which has not become a term of these Conditions save that your agreement contained in this Clause shall not apply in respect of any fraudulent misrepresentation whether or not such has become a term of these Conditions.</p>

<p>You agree to the terms of our privacy policy which is available here.</p>

<p>These Conditions will be exclusively governed by and construed in accordance with the laws of the United States whose Courts will have exclusive jurisdiction in any dispute, save that we have the right, at our sole discretion, to commence and pursue proceedings in alternative jurisdictions.</p>

<p>You may send us notices under or in connection with these Conditions by postal mail to 22 High St, Collinsville, CT 06019 , United States; or by email to triston@teamwarfare.com. As proof of sending does not guarantee our receipt of your notice, you must ensure that you have received an acknowledgement from us, which should be retained by you.</p>
<% Call ContentEnd() %>

<% Call ContentStart("RULES FOR THE USE OF TEAMWARFARE FORUMS") %>

<p>As a user of TeamWarfare forum, chat room or message board ("Online Services") you confirm that you have read, understand and agree to these Rules, TeamWarfare's Terms and Conditions Of Use of its website and you agree to the terms of TeamWarfare's Privacy Policy. If you do not agree, do not use TeamWarfare's Online Services, or if you are already using TeamWarfare's Online Services, please stop immediately.</p>

<p>TeamWarfare's Online Services are not actively monitored and TeamWarfare is not responsible for the content of any messages that are posted. TeamWarfare has the right, but not the obligation, to monitor any activity and content associated with the Online Services. TeamWarfare does not vouch for or warrant the accuracy, completeness, or usefulness of any message. The messages express the views of the author of the message, not the views of TeamWarfare. TeamWarfare reserves the right, in its sole discretion, to edit, delete, or refuse to post any message or thread for any reason whatsoever.</p>

<p>Any user who feels that a posted message is objectionable should contact TeamWarfare by sending an email to qualitycontrol@teamwarfare.com</p>

<p>You agree that you will not upload, post, email or otherwise transmit any material that may infringe any person or entity's intellectual property rights (including patents, trademarks, trade secrets, copyright or other intellectual or other property right).</p>

<p>You agree not to upload, post, email or otherwise transmit any material which (a) is defamatory, libellous, disruptive, threatening, invasive of a person's privacy, harmful, abusive, harassing, obscene, hateful, or racially, ethnically or otherwise objectionable; or that otherwise violates any law, (b) contains software viruses or any other computer codes, files or programs designed to interrupt, destroy, or limit the functionality of any computer software or hardware or telecommunications equipment, (c) may infringe any person or entity's intellectual property rights (including patents, trademarks, trade secrets, copyright or other intellectual or other property right). By posting any material, you confirm, represent and warrant that you have the lawful right to distribute and reproduce such material.</p>

<p>By posting any material, you confirm, represent and warrant that you have the lawful right to distribute and reproduce such material.</p>

<p>You will not impersonate any person or entity or otherwise misrepresent your affiliation with a person or entity.</p>

<p>You agree not to repeatedly post the same or similar message ("flooding") or post excessively large or inappropriate images.</p>

<p>Without TeamWarfare's written permission you will not distribute or publish unsolicited promotions, advertising or solicitations for funds, goods or services, including but not limited to, junk mail, spam and chain letters.</p>

<p>Postings to Online Services become public information. You should be very careful about posting personally identifiable information such as your name, address, telephone number or email address. If you post personal information online, you may receive unsolicited messages from other users in return.</p>

<p>TeamWarfare reserves the right to issue warnings, suspend or terminate the registration of users who refuse to comply with these rules. TeamWarfare may modify these rules from time to time and such modifications will be effective and binding on you when posted online.</p>

<p>You will remain solely liable for the content of any messages or other information you upload or transmit.</p>

<p>By using TeamWarfare's Online Services, you grant TeamWarfare the royalty-free, perpetual, irrevocable, non-exclusive and fully sub licensable right and license to use, reproduce, modify, edit, adapt, publish, translate, create derivative works from, distribute, perform and display the content you post to publicly accessible areas of the TeamWarfare.com site.</p>

<p>TeamWarfare may investigate any reported violation of its policies or complaints and take any action that it deems appropriate. Such action may include, but is not limited to, issuing warnings, removing posted content and/or reporting any activity that is suspects violates any law or regulation to appropriate law enforcement officials, regulators, or other third parties.</p>

<p>Recognising the global reach of the internet, you agree to comply with all local rules regarding online conduct and acceptable content.</p>

<p>WE TAKE PRECAUTIONS WITH CUSTOMER INFORMATION AND PERSONAL DATA, HOWEVER, EXCEPT IN THE CASE OF DEATH OR PERSONAL INJURY RESULTING FROM OUR NEGLIGENCE, WE DO NOT ACCEPT RESPONSIBILITY FOR ANY LOSS OR DAMAGE RESULTING FROM ANY SECURITY BREACHES THAT MAY OCCUR.</p>
	
Updated June 5th, 2006
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>