<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/head">
	<div class="logo"><img src="/images/logo-bw.jpg"/></div>
	<p class="inthename">�� ��� ���</p>
	<p class="date"><xsl:value-of select="./today/shamsiToday"/></p>
	<xsl:if test="concat(./companyName,'')!=''">
		<p class="company">���� ����� <xsl:value-of select="./companyName"/></p>
	</xsl:if>
	
	<p class="name">
		<xsl:if test="concat(./dear,'')='����'"><span>����</span></xsl:if>
		<xsl:if test="concat(./dear,'')='����'"><span>�����</span></xsl:if> <xsl:value-of select="./dear"/> <xsl:value-of select="./customerName"/></p>
	<p>�� ����</p>
	<p>������� ������� �������� �� ��� �� ������ ����� �� ��� ��� ����� �흐���:</p>
	<p>����� ����� <xsl:value-of select="./qtty"/> ��� <xsl:value-of select="./orderTitle"/> �� ����� <xsl:value-of select="./paperSize"/> �� �����  <xsl:value-of select="./totalPrice"/> ���� ������. �� �� ��� ��� ������:</p>
</xsl:template>
</xsl:stylesheet>