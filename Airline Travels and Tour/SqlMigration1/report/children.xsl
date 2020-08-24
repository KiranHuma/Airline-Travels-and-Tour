<?xml version="1.0" ?>

<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

<xsl:template match="/">

<HTML><HEAD>
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE9" />
  <script>
function goChild(s) 
{
	top.getNode(s);
}
</script>
<LINK REL='stylesheet' HREF="../textstyles.css" type="text/css"></LINK>


</HEAD><TITLE></TITLE>
<BODY bgcolor="white">



<table align='center' STYLE='table-layout:fixed;border-collapse:collapse' WIDTH='340'>
<caption> <div HEIGHT='50' align='center' valign='middle'><b> Children</b> </div></caption>
<col WIDTH='260'></col><col WIDTH='80'></col>
<tr HEIGHT='30'>
	<th> Name</th>
	<th> Converted</th>
</tr>
<xsl:for-each select="children/child">
<xsl:sort select="@statement" /> 
   <tr HEIGHT='30'>
      <td> <span STYLE='text-decoration:underline;cursor:hand;position:relative;left:6px;'  onclick='goChild({@id});  ' >
	<xsl:value-of select="@name" />

      </span></td>
      <xsl:if test='@converted = 100'>
    	      <td align='right'><span STYLE='position:relative;right:6px;'> <xsl:value-of select="@converted"/>% </span></td>
      </xsl:if>
      <xsl:if test='@converted &lt; 100'>
    	      <td align='right' style="font-size: 14px;color:red;font-weight:bold;"><span STYLE='position:relative;right:6px;'> <xsl:value-of select="@converted"/>%</span></td>
      </xsl:if>

    </tr>
</xsl:for-each>
</table>
</BODY>
</HTML>
</xsl:template>
</xsl:stylesheet>
