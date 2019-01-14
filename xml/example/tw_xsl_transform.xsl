<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet
xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<xsl:template match="/">
<html>
<body bgcolor="#FFFFFF" text="#000000">
<table width="50%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td colspan="3"> 
      <div align="center"><b><font size="4">XSL Example</font></b></div>
    </td>
  </tr>
  <tr> 
    <td colspan="3"><hr></hr></td>
  </tr>
  <xsl:for-each select="TeamInformation/LadderInformation">
  <tr> 
    <td> 
      <div align="left"><b>Ladder</b></div>
    </td>
    <td> 
      <div align="left"><b>Rung</b></div>
    </td>
    <td> 
      <div align="left"><b>Record</b></div>
    </td>
  </tr>
  <tr> 
    <td>
      <div align="left">
		<a>
		  <xsl:attribute name="href">
			<xsl:value-of select="LadderLink" />
		  </xsl:attribute>
		  <xsl:value-of select="LadderName" />
		</a>
	  </div>
    </td>
    <td > 
      <div align="left"><xsl:value-of select="Rank" /></div>
    </td>
    <td > 
      <div align="left"><xsl:value-of select="Wins" />/<xsl:value-of select="Losses" />/<xsl:value-of select="Forfeits" /></div>
    </td>
  </tr>
  <tr> 
    <td colspan="3"> 
      <div align="left"></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="left"><b>Current Status</b></div>
    </td>
    <td> 
      <div align="left"><b>Maps</b></div>
    </td>
    <td> 
      <div align="left"><b>Match Date</b></div>
    </td>
  </tr>
  <tr> 
    <td>
	  <xsl:value-of select="CurrentStatus" />: 
	  <a>
		<xsl:attribute name="href">
		  <xsl:value-of select="OpponentLink" />
		</xsl:attribute>
	    <xsl:value-of select="DetailedStatus/Opponent" />
	  </a>
	</td>
    <td><xsl:value-of select="DetailedStatus/Map1" />, <xsl:value-of select="DetailedStatus/Map2" />, <xsl:value-of select="DetailedStatus/Map3" /></td>
    <td><xsl:value-of select="DetailedStatus/MatchDate" /></td>
  </tr>
  <tr> 
    <td colspan="3"> 
      <hr></hr>
    </td>
  </tr>
</xsl:for-each>
</table>
</body>
</html>
</xsl:template>
</xsl:stylesheet>

