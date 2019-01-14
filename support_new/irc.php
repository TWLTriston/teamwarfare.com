<?php
$username = $_POST["username"];
$channel = $_POST["channel"];
?>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>TeamWarfare IRC Support</title>
<style type="text/css">
* {
	font-family: Verdana;
}
</style>
</head>

<body text="gold" bgcolor="#000000">
<p align="center">TeamWarfare IRC Support</p>
<p align="center">
  <applet code=IRCApplet.class archive="irc.jar,pixx.jar" width=707 height=551>
 <param name="CABINETS" value="irc.cab,securedirc.cab,pixx.cab">
<param name="nick" value="<?php print $username; ?>">
<param name="alternatenick" value="<?php print $username; ?>??">
<param name="name" value="<?php print $username; ?>">
<param name="host" value="irc.gamesurge.net">
<param name="port" value="6667">
<param name="command1" value="<?php print $channel; ?>">
<param name="gui" value="pixx">
<param name="quitmessage" value="TWL Support!!!">
<param name="language" value="english">
<param name="pixx:language" value="pixx-english">
<param name="pixx:timestamp" value="true">
<param name="pixx:highlight" value="true">
<param name="pixx:highlightnick" value="true">
<param name="pixx:nickfield" value="true">
<param name="soundbeep" value="snd/bell2.au">
<param name="soundquery" value="snd/ding.au">
  </applet> 
</p>
</body>
</html>