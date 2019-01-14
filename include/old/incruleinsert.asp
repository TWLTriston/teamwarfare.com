<table cellspacing=0 cellpadding=0 width=97% border=0>
<tr valign=center><td align=center>
<table cellspacing=0 cellpadding=2 width=100% border=0>
<tr><td><p>&nbsp;</td></tr>
<tr bgcolor=<%=bgctwo%>><td align=center>Brief rule overview</td></tr>
<% if mstatus="Attacking" then %>
<tr height=22 bgcolor=<%=bgcone%>><td><li>Defender: <%=Server.HTMLEncode(enemyname)%></td></tr>
<tr height=22 bgcolor=<%=bgctwo%>><td><li>Attacker: <% = Server.HTMLEncode(request.querystring("team")) %></td></tr>
<tr height=22 bgcolor=<%=bgcone%>><td><li><%=Server.HTMLEncode(enemyname)%> will choose two (2) dates for a match.</td></tr>
<% 
 if Request("ladder")="MW4 Team Attrition" or Request("Ladder")= "MW4 TA" then %>
<tr height=22 bgcolor=<%=bgctwo%>><td><li>Team Captains will select drop zones after auto-generated map settings are finalized.</td></tr>
 
<% end if
 else %>
<tr height=22 bgcolor=<%=bgcone%>><td><li>Defender: <% = Server.HTMLEncode(request.querystring("team")) %></td></tr>
<tr height=22 bgcolor=<%=bgctwo%>><td><li>Attacker: <%=Server.HTMLEncode(enemyname)%></td></tr>
<tr height=22 bgcolor=<%=bgcone%>><td><li><% = Server.HTMLEncode(request.querystring("team")) %> will choose two (2) dates for a match.</td></tr>
<%  if Request("ladder")="MW4 Team Attrition" or Request("Ladder")= "MW4 TA" then %>
<tr height=22 bgcolor=<%=bgctwo%>><td><li>Team Captains will select drop zones after auto-generated map settings are finalized.</td></tr>
 
<% end if
end if%>
<tr><td><p>&nbsp;</td></tr>
</table>
</td></tr></table>