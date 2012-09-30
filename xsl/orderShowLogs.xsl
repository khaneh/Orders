<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/logs">
<table align="center">
	<thead>
		<tr>
			<th>تاريخ</th>
			<th>ساعت</th>
			<th>مرحله</th>
			<th>وضعيت</th>
			<th>ثبت كننده</th>
		</tr>
	</thead>
	<tbody>
<xsl:for-each select="./order">
		<tr>
			<td class="date"><xsl:value-of select="./insertedDate"/></td>
			<td class="time"><xsl:value-of select="./insertedDate"/></td>
			<td><xsl:value-of select="./stepName"/></td>
			<td><xsl:value-of select="./statusName"/></td>
			<td><xsl:value-of select="./realName"/></td>
		</tr>
</xsl:for-each>
	</tbody>
</table>
</xsl:template>
</xsl:stylesheet>