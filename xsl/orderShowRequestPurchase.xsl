<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/purchase">
<div class="well well-small">
	<center class="label label-inverse">درخواستهاي خريد سرويس و كالا</center>
	<xsl:for-each select="./req">
		<div>
			<xsl:attribute name="title">
				<xsl:value-of select="./Comment/."/>
			</xsl:attribute>
			<xsl:if test="./Status='new'">
				<span class="label label-info">جديد</span>
			</xsl:if>
			<xsl:if test="./Status='ord'">
				<span class="label label-success">سفارش</span>
			</xsl:if>
			<xsl:if test="./Status='del'">
				<span class="label label-important">حذف شده</span>
			</xsl:if>
			<span><xsl:value-of select="./TypeName/."/></span>
			<span>(تعداد </span>
			<xsl:value-of select="./Qtty/."/> 
			<span> قيمت </span>
			<xsl:value-of select="./Price/."/>
			<span>)</span>
		</div>
	</xsl:for-each>
</div>
</xsl:template>
</xsl:stylesheet>