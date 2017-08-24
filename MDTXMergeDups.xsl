<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
    <xsl:output method="xml" indent="yes"/>

  <xsl:key name="tagnames" match="/ServerDescription/*" use="name()"/>
  <xsl:key name="datapertag" match="/ServerDescription/*/*" use="name(..)"/>

  <xsl:template match="/ServerDescription">
    <ServerDescription>   
    <xsl:for-each select ="*">
  <!--    I'm <xsl:value-of select="name()"/>=[<xsl:value-of select="."/>] -->
      <xsl:if test="generate-id() = generate-id(key('tagnames', name(.))[1])">
  <!--    and I'm in the if -->
        <xsl:copy>          
        <xsl:copy-of select="text() | key('datapertag',name())[node()]"/>
        </xsl:copy>
      </xsl:if>
    </xsl:for-each>
    </ServerDescription>
        </xsl:template>
</xsl:stylesheet>
