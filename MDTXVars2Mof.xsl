<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl xs"   xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xsl:output method="text" indent="no" />
  <xsl:strip-space elements="*" />

  <!-- $Rev: 353 $ $URL: svn://lupvs802/MDTRepo/trunk/MDTShare/Scripts/MDTXVars2Mof.xsl $ $Id: MDTXVars2Mof.xsl 353 2013-01-08 08:26:58Z jq467 $ -->
  
  <!-- USER VARIABLES -->
  <!-- WMI Class Name -->
  <xsl:variable name ="classprefix">MDTX2010</xsl:variable>
  <!-- Key created by the result of MDTXVars2Reg.xsl -->
  <xsl:variable name ="keyprefix">HKEY_LOCAL_MACHINE\\Software\\MDTX2010\\Source</xsl:variable>
  
  <!-- If you change something below, you may break something -->
  
  <xsl:variable name="s" select ="document('../Control/MDTXemitvars.xsd')/xs:schema//xs:element" />
  <xsl:template match ="/ServerDescription" xml:space="default">
    //==================================================================
    // Register Registry property provider (shipped with WMI)
    //==================================================================
    #pragma namespace("\\\\.\\root\\cimv2")

    instance of __Win32Provider as $InstProv
    {
    Name    ="RegProv" ;
    ClsID   = "{fe9af5c0-d3b6-11ce-a5b6-00aa00680c3f}" ;
    ImpersonationLevel = 1;
    PerUserInitialization = "False";
    };

    instance of __InstanceProviderRegistration
    {
    Provider    = $InstProv;
    SupportsPut = True;
    SupportsGet = True;
    SupportsDelete = False;
    SupportsEnumeration = True;
    };

    instance of __Win32Provider as $PropProv
    {
    Name    ="RegPropProv" ;
    ClsID   = "{72967901-68EC-11d0-B729-00AA0062CBB7}";
    ImpersonationLevel = 1;
    PerUserInitialization = "False";
    };

    instance of __PropertyProviderRegistration
    {
    Provider     = $PropProv;
    SupportsPut  = True;
    SupportsGet  = True;
    };

    //==================================================================
    // Our <xsl:value-of select ="$classprefix"/> class
    //==================================================================

    #pragma namespace ("\\\\.\\root\\cimv2")

    #pragma deleteclass("<xsl:value-of select ="$classprefix"/>",nofail)
    [DYNPROPS]
    class <xsl:value-of select ="$classprefix"/>
    {
    [key]
    string InstanceKey;
    string SvnRev;
    <xsl:apply-templates mode="header"/>
    };

    // Instance definition
    [DYNPROPS]
    instance of <xsl:value-of select ="$classprefix"/>
    {
    InstanceKey = "@";
    [PropertyContext("local|<xsl:value-of select="$keyprefix"/>|SvnRev"), Dynamic, Provider("RegPropProv")]
    SvnRev;

    <xsl:apply-templates mode="contents"/>
    };
    
    // End OF MOF File    
  </xsl:template>
  <xsl:template match="comment()">
  </xsl:template>
  <xsl:template match="@* | node()" mode="header">
    <xsl:variable name="varname" select="name(.)"/>
    <xsl:variable name="em" select ="$s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@type" />

    <xsl:choose>
     <!-- <xsl:when test="'$em='aggregated'"><xsl:apply-templates select="*"/></xsl:when> -->
      <xsl:when test="$varname=''"></xsl:when>
      <xsl:when test="$em='single'">
        string <xsl:value-of select ="$varname"/>;
      </xsl:when>
      <xsl:when test="$em='list' or $em='csvstring'">
        <xsl:for-each select ="*">
        string <xsl:value-of select="$varname"/><xsl:value-of select="format-number(position(),'000')"/>;
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="$em='nested'">
        <xsl:for-each select ="*">
          <xsl:variable name="prefix" select="name(.)"/><xsl:variable name="pos" select ="format-number(position(),'000')"/>
          <xsl:for-each select ="*">
            string <xsl:value-of select="$varname"/><xsl:value-of select ="$pos"/><xsl:value-of select="name(.)"/>;
          </xsl:for-each>
        </xsl:for-each>
        int <xsl:value-of select="$s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@count"/>;
      </xsl:when>
      <xsl:otherwise>
        string UNTYPED_<xsl:value-of select="$varname"/>;
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="@* | node()" mode="contents">
    <xsl:variable name="varname" select="name(.)"/>
    <xsl:variable name="em" select ="$s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@type" />
    <xsl:choose>
   <!--   <xsl:when test="$em='aggregated'"><xsl:apply-templates select="*"/></xsl:when> -->
      <xsl:when test="$varname=''"></xsl:when>
      <xsl:when test="$em='single'">
        [PropertyContext("local|<xsl:value-of select="$keyprefix"/>|<xsl:value-of select ="$varname"/>"), Dynamic, Provider("RegPropProv")]
        <xsl:value-of select ="$varname"/>;
      </xsl:when>
      <xsl:when test="$em='list' or $em='csvstring'">
        <xsl:for-each select ="*">
         [PropertyContext("local|<xsl:value-of select="$keyprefix"/>\\<xsl:value-of select ="$varname"/>|<xsl:value-of select="format-number(position(),'000')"/>"), Dynamic, Provider("RegPropProv")]
         <xsl:value-of select="$varname"/><xsl:value-of select="format-number(position(),'000')"/>;
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="$em='nested'">
        <xsl:for-each select ="*">
          <xsl:variable name="prefix" select="name(.)"/><xsl:variable name="pos" select ="format-number(position(),'000')"/>
          <xsl:for-each select ="*">
            [PropertyContext("local|<xsl:value-of select="$keyprefix"/>\\<xsl:value-of select ="$varname"/>|<xsl:value-of select ="$pos"/>"), Dynamic, Provider("RegPropProv")]
            <xsl:value-of select="$varname"/><xsl:value-of select ="$pos"/><xsl:value-of select="name(.)"/>;
          </xsl:for-each>
        </xsl:for-each>
        [PropertyContext("local|<xsl:value-of select="$keyprefix"/>\\<xsl:value-of select ="$varname"/>|<xsl:value-of select="$s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@count"/>"), Dynamic, Provider("RegPropProv")]
      <xsl:value-of select="$s[@name=$varname]/xs:annotation/xs:appinfo/xs:emit/@count"/>;

      </xsl:when>
      <xsl:otherwise>
        reg add <xsl:value-of select="$classprefix"/> /v UNTYPED_<xsl:value-of select="$varname"/> /t reg_sz /d "<xsl:value-of select="text()"/>" /f
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>
