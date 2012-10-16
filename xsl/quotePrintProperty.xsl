<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/keys">
	<xsl:for-each select="./row">
		<xsl:for-each select="./group">
			<xsl:variable name="gname" select="./@name"/>
			<div class='rspan3'>
				<li>
					<xsl:if test="./@isPrice='no'">
						<b>
							<xsl:value-of select="./key[@name='description']"/> 
							<xsl:if test="./key[@name='pages']!=''"> (<xsl:value-of select="./key[@name='pages']"/> صفحه)</xsl:if>
						</b>
					</xsl:if>
					<xsl:value-of select="./@label"/> <xsl:if test="./@isPrice!='no'">(<xsl:value-of select="./key[@name=concat($gname,'-price')]"/>) ريال</xsl:if>
				</li>
			</div>
		</xsl:for-each>
		<div class="clear"></div>
	</xsl:for-each>
</xsl:template>
</xsl:stylesheet>