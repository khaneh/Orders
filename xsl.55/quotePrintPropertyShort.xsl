<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/keys">
	<xsl:for-each select="./row">
		<li>
		<xsl:variable name="sumPrice" select="0"/>
		<xsl:for-each select="./group">
			<xsl:variable name="gname" select="./@name"/>
				<span>
				<xsl:if test="./@isPrice='no'">
					<b>
						<xsl:value-of select="./key[@name='description']"/> 
						<xsl:if test="./key[@name='pages']!=''"> (<xsl:value-of select="./key[@name='pages']"/> صفحه)</xsl:if>
					</b>
				</xsl:if>
				<xsl:value-of select="./@label"/> 
				<xsl:if test="./@isPrice!='no'"></xsl:if>
				</span>
				<span>|</span>
		</xsl:for-each>
		<!--span>(<xsl:value-of select="sum(replace(./group/key[substring(@name,string-length(@name)-5,6)='-price']/text(),',',''))"/>) ريال	</span-->
		</li>
	</xsl:for-each>
</xsl:template>
</xsl:stylesheet>