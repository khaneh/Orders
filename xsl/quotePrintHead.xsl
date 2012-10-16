<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/head">
	<div class="logo"><img src="/images/logo-bw.jpg"/></div>
	<p class="inthename">به نام خدا</p>
	<p class="date"><xsl:value-of select="./today/shamsiToday"/></p>
	<xsl:if test="concat(./companyName,'')!=''">
		<p class="company">شركت محترم <xsl:value-of select="./companyName"/></p>
	</xsl:if>
	
	<p class="name">
		<xsl:if test="concat(./dear,'')='آقاي'"><span>جناب</span></xsl:if>
		<xsl:if test="concat(./dear,'')='خانم'"><span>سركار</span></xsl:if> <xsl:value-of select="./dear"/> <xsl:value-of select="./customerName"/></p>
	<p>با سلام،</p>
	<p>احتراما استعلام درخواستي از طرف آن مجموعه محترم به شرح ذيل اعلام مي‌گردد:</p>
	<p>هزينه توليد <xsl:value-of select="./qtty"/> عدد <xsl:value-of select="./orderTitle"/> در ابعاد <xsl:value-of select="./paperSize"/> در مجموع  <xsl:value-of select="./totalPrice"/> ريال مي‌باشد. كه به شرح ذيل مي‌باشد:</p>
</xsl:template>
</xsl:stylesheet>