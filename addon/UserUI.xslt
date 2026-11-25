<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="xml" encoding="utf-8" indent="yes"/>

  <!-- Identity transform -->
  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <!-- Ensure Saulo Tools appears in Window > Dockers even in customized user workspaces -->
  <xsl:template match="commandBarData[@guid='3eaa9bbe-28fd-4672-9128-02974ee96332']/menu">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <xsl:if test="not(item[@guidRef='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03'])">
        <item guidRef="7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>
