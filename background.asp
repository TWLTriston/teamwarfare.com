<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Development Background"

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

<% Call ContentStart("Development History") %> 
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH=760>   
<TR><TD>
<p>The development of this ladder has been an extraordinary learning experience for both Myles and myself.  It began around Halloween when Myles was looking for 
some solid PHP development references.  During the discussion he told me of his plan to develop a ladder designed to fit within the Tribalwar.com framework.
After some discussions we concluded that we would be able to turn around a solid product using ASP and SQL Server in a much more condensed time frame.  Development
began almost immediately.
</p>
<p>The development process was very interesting.  I would code the engine in the evening and email the database to Chuck who installed it on his SQL server late at night (my time).
I emailed the ASP files to both Myles and Chuck.  Myles would make adjustments and add features that night.  Chuck was running my latest version at www.clanfade.org 
while Myles was running his latest updates.  Myles would publish the latest ladder files on an FTP site and I would download them the following day and the 
cycle would continue.
</p>
<p>The initial challenge was to develop a flexible system that would allow the team captains the ability to custom tailor their roster to the ladder based on game type.
This sounded a lot easier than it was.  The premise of having unique rosters for every ladder on which a team is participating makes for a messy layout.  Once
we were able to map out the database needs for teams, players, and ladders we began the production of the site in earnest.
</p>
<p>The next hurdle was to retrofit the new ASP system into the rather brilliant framework that was created by Cowboy for Tribalwar.  The templates in use were developed for use with Perl,
PHP, and HTML.  The introduction of ASP into the mix made the initial process rather sketchy.  Once we identified the key elements in the template we were able to 
recreate it in a format that better suited ASP.  With this process under our belt we proceeded to develop the team, player, and ladder tools.
</p>
<P>We spent some time hashing out the rules for the ladder so we would be able to develop the automated code to handle things like random map assignments, match communications,
and the challenges and defense processes.  At this point the principal elements of the ladder engine was complete.  We recruited some important administration personnel
 to help facilitate the final phases of development.  This is when we contacted the key players at Tribalwar to push for this to be released as a principle sector of the Tribalwar community.
</p>
<p>
Since that point it's all been about features and stability.  We have put a lot of thought into each element of this system and we hope it operates in a manor that will best
suit the needs of this great community.  We appreciate the support and feedback we have received from so many people.  This ladder belongs tot he Tribes community.  I hope you all
enjoy it as much as we enjoyed writing it.</P>
<div align=right> - Will&nbsp;</div>
</TD></TR>
</TABLE>
<% 
Call ContentEnd()
Call ContentStart("")
%>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH=760>   
<TR><TD>
<P>The entire TWL, what you see, was done over a two-month period in a constant 
update between Will and myself.  With the help and encouragement from the 
Tribalwar Staff, Chuck, Yankee, and other members of the tribes community, we 
made a ladder from the ground up, and are fairly successful, in my opinion.</P>
<P>This is my first large-scale site/project/program.  Will on the other hand 
has had many years of experience in coding dynamic web pages (PHP, ASP), and 
developing Structure Query Language (SQL).  With his guidance and his father like 
patience, I have learned the ladder inside and out, and helped code some of the 
different bells and whistles.  ASP (Active Server Pages) is my friend now, and I 
feel a personal achievement in doing this for not only my passion of tribes, but 
for myself as a Computer Science Major in a state university.</P>
<P><STRONG>Ladder:</STRONG> </P>
<P>A simple idea on paper, but there are many hurdles that are put in your way.  
First off, how do you make a ladder that isn't like the others that exist?  Also, 
what features of the current ladders make them successful?  Finally, what can you 
do to improve on them all?</P>
<P>Between the Online Gaming League (OGL) and Teamplay.net (TP.net), there is a 
lot of gray space.  These two ladders are the yin and the yang of the Tribes 
gaming community - the OGL being the center point for competition; Teamplay for a 
more lackadaisical attitude towards game play, but higher on the feature scale.  OGL 
has been around for nearly 3 years and was the starting point for tribal 
competition.  After a short time, the TP system was released.  Their loosely 
applied rule system and server reservation system made Teamplay a prime 
target for tribes in addition to the OGL.  Many teams felt TP was a scrim 
league, and used it as such for their upcoming matches on the OGL.  With TP, no 
one looks at rank or any other statistic, considering the possibility of 
falsified numbers (i.e. Thunderwalkers was nearly undefeated after approx 20 
matches, but none of their wins were considered powerful).  OGL, on the 
other hand, rank is everything.  Record is everything.  All rests on match night 
for OGL, and everyone knows it.  From this suspense came the creation of community 
features; forums, demos, and finally <STRONG>shoutcasting.</STRONG></P>
<P>With the TWL, we are the first to bring shoutcasting inside the ladder.  We 
promote it, we encourage it, and we make it easy to shoutcast/be shoutcasted.  
TWL has the option right in the administrative screens to accept/deny a shoutcast.  The shoutcaster then only 
needs to login to the system to find out if there are any available games to 
shoutcast, and then offer their services back to the teams playing the match.  They 
finalize their decision for shoutcasting, and all is done.  No more hunting for 
e-mails, irc infinite idles, or mixed up messaging.</P>
<P>Utilizing the features of other the other ladder systems, TWL combines them 
into a familiar Tribalwar style, with a Tribalwar twist.  Bringing OGL's tight 
rule system and admins, and TP's on server communications, we eliminate the wait 
time for e-mails to be processed, and the chance of a missed delivery.  E-mail is 
used as notification by the system to inform the captains of updates, but captains 
are still expected to check on their team regularly.  Team administration is 
ultra simple.  Bookmark 1 page per ladder, and you are golden, or just bookmark 
your team's main page, and link from there.  Features include, current ladder 
status, roster management, and captain management.  No more clicking through link 
after link trying to find anything.  We have made every effort to put all 
features/page within 3 links of the main page and 4 at the most.  If you have trouble 
finding something, let us know, and we will discover an easier/smarter way to do 
it (give us a suggestion, we just might use it :) ).</P>
<P>Finally, we put much time devising a more fair system.  Designed for Tribes 2 
(where maps are symmetrical) map choices are random for map 1 and 2.  This allows 
attackers no advantage, and defenders no advantage.  This also encourages new and 
ingenious strategies, while keeping the game exciting for all.  In time, some 
maps will prove to not qualify as match quality, and will be removed.  This 
includes maps that are prone to 0-0 ties, and multiple overtime periods.  Also 
encouraging the balanced game play, we are going to bring back a long lost 
piece of the Teamplay system: <STRONG>server reservation</STRONG>.  Using a 
server mod designed by [PoE]BigBrown, we have adapted his method to work with 
the ladder for game time reservations.  With the help and cooperation of the many 
server admins that exist, we can soon compile a large number of servers for 
teams to use on match night, insuring impartial admins, and reliable 
bandwidth.</P>
<P>That is about it for the TWL: a simple system, easy to use, easy to navigate, 
and easy to understand.  Hopefully in time, TWL will be the competitive ladder 
for T2, so that when you join a server, players ask you: "What is your rank on 
TWL CTF?" A pipe dream?  Maybe, but nevertheless, we as a development team are 
satisfied at the accomplishment.  Working on the next piece of the puzzle, and 
taking all feedback into consideration, TWL will get better and better, while 
other systems will just get older and older.</P>
<P>Please e-mail me with any questions <A 
href="mailto:triston@teamwarfare.com">triston@teamwarfare.com</A></P>
<div align=right>- Myles&nbsp;</div>
</TD></TR>
</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>