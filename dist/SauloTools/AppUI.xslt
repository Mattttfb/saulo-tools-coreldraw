<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:frmwrk="Corel Framework Data">
  <xsl:output method="xml" encoding="utf-8" indent="yes"/>

  <!-- Identity transform -->
  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <!-- Register WPF Host control and Window > Dockers toggle command -->
  <xsl:template match="uiConfig/items">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <xsl:if test="not(itemData[@guid='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d02'])">
        <itemData guid="7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d02"
                  type="wpfhost"
                  hostedType="Addons\SauloTools\SauloTools.dll,SauloTools.DockerUI"
                  enable="true"/>
      </xsl:if>
      <xsl:if test="not(itemData[@guid='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03'])">
        <itemData guid="7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03"
                  noBmpOnMenu="true"
                  type="checkButton"
                  check="*Docker('7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d01')"
                  enable="true"
                  userCaption="Saulo Tools"
                  caption="Saulo Tools"
                  nonLocalizableName="SauloToolsDockerToggle"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>

  <!-- Register the native Docker panel -->
  <xsl:template match="uiConfig/dockers">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <xsl:if test="not(dockerData[@guid='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d01'])">
        <dockerData guid="7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d01"
                    userCaption="Saulo Tools"
                    caption="Saulo Tools"
                    nonLocalizableName="SauloToolsDocker"
                    wantReturn="true"
                    noPadding="true">
          <container>
            <item dock="fill" margin="0,0,0,0" guidRef="7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d02"/>
          </container>
        </dockerData>
      </xsl:if>
    </xsl:copy>
  </xsl:template>

  <!-- Inject into Window > Dockers menu (guid verified in DrawUI.xml) -->
  <xsl:template match="commandBarData[@guid='3eaa9bbe-28fd-4672-9128-02974ee96332']/menu">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <xsl:if test="not(item[@guidRef='7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03'])">
        <item guidRef="7f4c9a12-6d38-4b8e-a519-2e8c4b7f1d03"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>
