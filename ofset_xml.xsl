<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/keys">
<xsl:for-each select="./row">
	<row>
		<xsl:attribute name="name"><xsl:value-of select="./@name"/>"></xsl:attribute>
		<xsl:attribute name="label"><xsl:value-of select="./@label"/>"></xsl:attribute>
	</row>
</xsl:for-each>
</xsl:template>
</xsl:stylesheet>