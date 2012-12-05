<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/stock">
<div class="well well-small">
	<center class="label label-inverse">درخواستهاي كالا از انبار</center>
	<xsl:for-each select="./req">
		<div class="customerHasInventory">
			<xsl:attribute name="title">
				<xsl:value-of select="./Comment/."/>
			</xsl:attribute>
			<a>
				<xsl:attribute name="href">
					<xsl:if test="concat(./link/.,'')!=''">
						<xsl:value-of select="concat('../inventory/default.asp?show=',./link/.)"/>
					</xsl:if>
				</xsl:attribute>
				<span>
					<xsl:attribute name="class">
						<xsl:value-of select="./StatusClass/."/>
					</xsl:attribute>
					<xsl:value-of select="./Status/."/>
				</span>
				<span><xsl:value-of select="./ItemName/."/></span>
				<span>(تعداد </span>
				<xsl:value-of select="./Qtty/."/> 
				<xsl:value-of select="./unit/."/>
				<span>)</span>
			</a>
			<xsl:if test="./customer=-1">
				<div class="label label-info">ارسالي</div>
			</xsl:if>
		</div>
	</xsl:for-each>
</div>
</xsl:template>
</xsl:stylesheet>