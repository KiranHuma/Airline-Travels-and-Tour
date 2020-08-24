<?xml version="1.0" ?>

<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html"/>

<xsl:template match="/">

<HTML><HEAD>
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE9" />
  <script language="javascript">
function goChild(s)
{
  if(top.frames["tree"].inProgress)
    return;
  top.getNode(s);
  top.startNavigate(s);
}
</script>
<style>
table {font-family:  Arial,Helvetica, sans-serif; font-size: 13px;}
</style>

</HEAD><TITLE></TITLE>
<BODY style='filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFFF, endColorstr=#FFBBBBBB, gradientType=0 )'>

<div >

<div STYLE='position:absolute;top:6px;left:20px;'>

<table>
<tr>
<xsl:for-each select="path/node">
<xsl:sort select="position()" order="descending"/> 
  <xsl:if test="position()!=last()">  
      <xsl:if test='@id != "00000000-0000-0000-0000-000000000000"'>
    	<td STYLE='text-decoration:underline;cursor:hand;' >
		<div>
		<xsl:attribute name="onclick"><xsl:value-of select="concat('goChild(',&quot;'&quot;,@id,&quot;'&quot;,')')"/></xsl:attribute>
		<xsl:value-of select="@name" /> 
		</div>
      	</td>
	<td STYLE='color:firebrick;text-decoration:none;'> <xsl:text >&#32;&#62;</xsl:text></td>
      </xsl:if>

      <xsl:if test='@id = "00000000-0000-0000-0000-000000000000"'>
    	<td STYLE='color:firebrick;'     >
		<xsl:value-of select="@name" /> >
      	</td>
      </xsl:if>      
   </xsl:if>
  <xsl:if test="position() =last()">  
     <xsl:if test='@id != "00000000-0000-0000-0000-000000000000"'>
    	<td STYLE='color:blue;font-weight:bold;'   >
		<xsl:value-of select="@name" />
      	</td>
      </xsl:if>
      <xsl:if test='@id = "00000000-0000-0000-0000-000000000000"'>
    	<td STYLE='color:silver;'     >
		<xsl:value-of select="@name" />
      	</td>
      </xsl:if>      
   </xsl:if> 
</xsl:for-each>

</tr>
</table>
</div>
</div>
</BODY>
</HTML>
</xsl:template>
</xsl:stylesheet>
