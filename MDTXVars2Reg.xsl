<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl xs"   xmlns:xs="http://www.w3.org/2001/XMLSchema">

<!-- $Id: MDTXVars2Reg.xsl 224 2010-08-17 06:39:54Z jq467 $ -->
<!-- $URL $-->
<!-- $Rev: 224 $ -->
<!-- $Date: 2010-08-17 08:39:54 +0200 (Tue, 17 Aug 2010) $ -->
  <xsl:output method="text" indent="no"/>
 <!-- <xsl:strip-space elements="*" /> -->
  
  <xsl:variable name ="keyprefix">HKLM\Software\MDTX2010\Source</xsl:variable>
  <xsl:variable name="s" select ="document('../Control/MDTXemitvars.xsd')/xs:schema//xs:element" />
  <xsl:template match ="/ServerDescription" xml:space="default">
REM AutoGenerated Script to create Keys from xml file
REM See MDTXVars2Reg.xsl $Rev: 224 $ for more details

reg add <xsl:value-of select="$keyprefix"/> /t reg_sz /v SvnRev /d "$Rev: 224 $" /f
      <xsl:apply-templates />
    
  </xsl:template>
  <xsl:template match="comment()">
  </xsl:template>
  <xsl:template match="@* | node()">
    <xsl:variable name="varname" select="name(.)"/>
    <xsl:variable name="em" select ="$s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@type" />    

    <xsl:choose>
      <xsl:when test="$varname=''"></xsl:when>
  <!--    <xsl:when test="$em='aggregated'"><xsl:apply-templates select="*"/></xsl:when>-->
      <xsl:when test="$em='single'">
reg add <xsl:value-of select="$keyprefix"/> /t reg_sz /v <xsl:value-of select="$varname"/> /d "<xsl:value-of select="text()"/>" /f
      </xsl:when>
      <xsl:when test="$em='list' or $em='csvstring'">
reg add <xsl:value-of select="$keyprefix"/>\<xsl:value-of select="$varname"/> /f
        <xsl:for-each select ="*">
reg add <xsl:value-of select="$keyprefix"/>\<xsl:value-of select="$varname"/> /v <xsl:value-of select="format-number(position(),'000')"/> /t reg_sz /d "<xsl:value-of select="text()"/>" /f
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="$em='nested'">
reg add <xsl:value-of select="$keyprefix"/>\<xsl:value-of select="$varname"/> /f
        <xsl:for-each select ="*">
          <xsl:variable name="prefix" select="name(.)"/><xsl:variable name="pos" select ="format-number(position(),'000')"/>
reg add <xsl:value-of select="$keyprefix"/>\<xsl:value-of select="$varname"/>\<xsl:value-of select ="$pos"/> /f
          <xsl:for-each select ="*">
reg add <xsl:value-of select="$keyprefix"/>\<xsl:value-of select="$varname"/>\<xsl:value-of select ="$pos"/> /v <xsl:value-of select="name(.)"/> /t reg_sz /d "<xsl:copy-of select="text()"/>" /f
          </xsl:for-each>
        </xsl:for-each>
reg add <xsl:value-of select="$keyprefix"/>\<xsl:value-of select="$varname"/> /t reg_dword /v <xsl:value-of select="$s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@count"/> /d "<xsl:value-of select="count(*)"/>" /f
      </xsl:when>
      
      <xsl:otherwise>
reg add <xsl:value-of select="$keyprefix"/> /v UNTYPED_<xsl:value-of select="$varname"/> /t reg_sz /d "<xsl:value-of select="text()"/>" /f        
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
