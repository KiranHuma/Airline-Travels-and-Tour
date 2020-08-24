<?xml version="1.0" ?>

<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template match="/">

    <HTML>
      <HEAD>

        <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE9" />
        <LINK REL='stylesheet' HREF="../textstyles.css" type="text/css"></LINK>

        <script>

          function init()
          {

          try
          {

          var height = top.getCatHeight();
          h = height - 350;
          var eDiv = document.getElementById("innerDiv");

          var str = window.parent.document.location.toString();
          if (str.lastIndexOf("srcIndex")>0)
          h-=35;
          eDiv.style.height = h;
          }
          catch(e)
          {
          var eDiv = document.getElementById("innerDiv");
          eDiv.style.height = "100%";
          eDiv.style.overflowY='visible';
          }
          }

          var lastTimeoutId = null;

          function doResize()
          {
          if (lastTimeoutId != null)
          clearTimeout(lastTimeoutId);
          lastTimeoutId = setTimeout("checkSize()",50);
          }

          function checkSize()
          {
          try
          {
          var height = top.getCatHeight();
          h = height - 185;
          var eDiv = document.getElementById("innerDiv");

          var str = window.parent.document.location.toString();
          if (str.lastIndexOf("srcIndex")>0)
          h-=35;

          var mh = top.getMessagesHeight();
          h -= mh;
          eDiv.style.height = h;
          }
          catch(e)
          {
          //just ignore
          }
          }


        </script>

      </HEAD>
      <TITLE></TITLE>
      <BODY bgcolor="white" style="height:75%;width:100%" onload='javascropt:init();' onresize='javascript:doResize();'>

        <center style="height:75%;width:360">
          <div align='center' style='height:100%;width:360;'>
            <table align='center' STYLE='table-layout:fixed;border-collapse:collapse;width:360'>
              <caption>
                <div HEIGHT='50' align='center' valign='middle'>
                  <b> Objects by categories</b>
                </div>
              </caption>
              <col style='width:200'></col>
              <col style='width:80'></col>
              <col style='width:80'></col>
              <tr HEIGHT='35' >
                <th> Object type</th>
                <th> Total</th>
                <th> With errors</th>
              </tr>
            </table>
            <div id='innerDiv' style='height:75%;width:360;overflow-y:auto'>
              <table align='center' STYLE='table-layout:fixed;border-collapse:collapse;width:360'>
                <xsl:for-each select="statistics/statistic">
                  <xsl:sort select="@category" />
                  <tr style="height:35">
                    <td style='width:200'>
                      <span STYLE='position:relative;left:6px;'>
                        <xsl:variable name='suffix'>
                          <xsl:choose>
                            <xsl:when test='@category="index"'>es</xsl:when>
                            <xsl:when test='@category="query"'>queries</xsl:when>
                            <xsl:otherwise>s</xsl:otherwise>
                          </xsl:choose>
                        </xsl:variable>
                        <xsl:choose>
                          <xsl:when test='@category="query"'>
                            <xsl:value-of select="normalize-space($suffix)"/>
                          </xsl:when>
                          <xsl:when test='@category!="query"'>
                            <xsl:value-of select="concat(@category, normalize-space($suffix))"/>
                          </xsl:when>

                        </xsl:choose>

                      </span>
                    </td>

                    <td align='right' style='width:80'>
                      <span STYLE='position:relative;right:6px;'>
                        <xsl:value-of select="@total"/>
                      </span>
                    </td>
                    
                    <td align='right' style='width:80'>
                      <span STYLE='position:relative;right:6px;'>
                        <xsl:value-of select="@not-converted-count"/>
                      </span>
                    </td>
                  </tr>
                </xsl:for-each>
              </table>
            </div>
          </div>
        </center>
      </BODY>
    </HTML>
  </xsl:template>
</xsl:stylesheet>
