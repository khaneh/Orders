<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/camp">
	<table class="list" border="1" cellspacing="0" cellpadding="2" dir="RTL"  borderColor="#555588">
		<tr valign="top" class="head">
			<td>#</td>
			<td>نام</td>
			<td>وضعيت</td>
			<td>اعضاء</td>
			<td>شماره تلفن</td>
			<td>تعداد تلفن</td>
			<td>تلفن‌هاي پاسخ داده شده</td>
		</tr>
		<xsl:for-each select="./row">
			<tr>
				<td>
					<xsl:value-of select="./id"/>
				</td>
				<td>
					<xsl:value-of select="./name"/>
				</td>
				<td>
					<xsl:value-of select="./status"/>
				</td>
				<td>
					<xsl:value-of select="./member"/>
				</td>
				<td class="exten">
					<xsl:value-of select="./exten"/>
				</td>
				<td class="totalCallied"></td>
				<td class="answerdCallid"></td>
			</tr>
		</xsl:for-each>
	</table>
</xsl:template>
</xsl:stylesheet>