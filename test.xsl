<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">

<xsl:template match="/">
<table border="2" bgcolor="yellow">
<tr>
<th>title</th>
<th>artist</th>
</tr>
<xsl:for-each select="CATALOG/CD">
<tr>

<td >
<xsl:attribute name="title">
	<xsl:value-of select="./TITLE/@aa"/>
</xsl:attribute>
<xsl:value-of select="TITLE"/>

</td>
<td>
<xsl:value-of select="ARTIST"/>
</td>
</tr>
</xsl:for-each>
</table>
</xsl:template>
</xsl:stylesheet>