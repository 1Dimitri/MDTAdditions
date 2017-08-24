<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl xs"   xmlns:xs="http://www.w3.org/2001/XMLSchema"
>
  <xsl:variable name="up" select="'ABCDEFGHIJKLMNOPQRSTUVWXYZ'"/>
  <xsl:variable name="lo" select="'abcdefghijklmnopqrstuvwxyz'"/>

  <!--
  $Id: MDTXEmitVars.xsl 313 2011-09-29 15:37:22Z jq467 $
  $Rev: 313 $
  $URL: svn://lupvs802/MDTRepo/trunk/MDTShare/Scripts/MDTXEmitVars.xsl $
  -->
  <xsl:output method="xml" indent="yes"/>

  <xsl:variable name="s" select ="document('../Control/MDTXemitvars.xsd')/xs:schema//xs:element" />
  <xsl:template match ="/ServerDescription">
    <MediaVarList>
      <xsl:apply-templates />
    </MediaVarList>
  </xsl:template>
  <xsl:template match="comment()">
    <xsl:copy/>
  </xsl:template>
  <xsl:template match="@* | node()">
    <xsl:variable name="varname" select="name(.)"/>
    <xsl:variable name="em" select ="$s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@type" />
    <xsl:variable name="varup" select="translate($varname,$lo,$up)"/>

    <xsl:choose>
      <xsl:when test="$varname=''"></xsl:when>
      <xsl:when test="$em='single'">
        <var name="{$varup}">
          <xsl:copy-of select="text()"/>
        </var>
      </xsl:when>
      <xsl:when test="$em='list'">
        <xsl:for-each select ="*">
          <var name="{$varup}{format-number(position(),'000')}">
            <xsl:copy-of select="text()"/>
          </var>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="$em='csvstring'">
        <var name="{$varup}">
          <xsl:for-each select ="*">
            <xsl:copy-of select="text()"/>
            <xsl:if test="position()!=last()">,</xsl:if>
          </xsl:for-each>
        </var>
      </xsl:when>
	  <xsl:when test="$em='semicolonstring'">
        <var name="{$varup}">
          <xsl:for-each select ="*">
            <xsl:copy-of select="text()"/>
            <xsl:if test="position()!=last()">;</xsl:if>
          </xsl:for-each>
        </var>
      </xsl:when>
      <xsl:when test="$em='nested'">
        <xsl:for-each select ="*">
          <xsl:variable name="prefix" select="translate(name(.),$lo,$up)"/>
          <xsl:variable name="pos" select ="position()-1"/>
          <xsl:for-each select ="*">
            <var name="{$prefix}{$pos}{translate(name(.),$lo,$up)}">
              <xsl:copy-of select="text()"/>
            </var>
          </xsl:for-each>
        </xsl:for-each>
        <var name="{translate($s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@count,$lo,$up)}">
          <xsl:value-of select="count(*)"/>
        </var>
      </xsl:when>
      <xsl:when test="$em='aggregated'">
        <xsl:apply-templates/>
      </xsl:when>
      <xsl:otherwise>
        <var name="UNTYPED_{$varname}">
          <xsl:copy-of select="text()"/>
        </var>
      </xsl:otherwise>
    </xsl:choose>


  </xsl:template>
</xsl:stylesheet>
