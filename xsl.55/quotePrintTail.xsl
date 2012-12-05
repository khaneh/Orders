<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/head">
	<p>����� ��� ��� ������ ���� ��� ���� ������ ������ �� ���� ������ ������</p>
	<p>������ ������� ��� �� ����� <xsl:value-of select="./today/date"/> ������</p>
	<p>���� ����� ����� �� �� ����� ����� � ������ <xsl:value-of select="./deposit"/> ���� �� ����� ��� ������ <xsl:value-of select="./productionDuration"/> ��� ���� ������.</p>
	<xsl:if test="./dueDate=0">
		<p>���� ������ ����� ������ �� ���� ����� ����� ���</p>
	</xsl:if>
	<xsl:if test="./dueDate>0">
		<p>���� ������ ����� ������ <xsl:value-of select="./dueDate"/> ��� �� �� ����� ����� ���</p>
	</xsl:if>
	<xsl:if test="./chequeDueDate=0">
		<p>���� ������ �� ������ �� ��� ����� ����� ���</p>
	</xsl:if>
	<xsl:if test="./chequeDueDate>0">
		<p>���� �� ���� �� <xsl:value-of select="./chequeDueDate"/> ��� �� �� ����� ����� ���</p>
	</xsl:if>	
	<p>����� �����: <xsl:value-of select="./salesPerson"/> ����� <xsl:value-of select="./extention"/></p>
	<div class="sign">
		<p>�� ����<br/><xsl:value-of select="./approvedByName"/><br/>���� �ǁ � ���</p>
	</div>
	<div class="tail">
		<p>����� | ������ ����� | ����� 545 | ���� 66042700 | ��� 66042704 | ����� �� www.pdhco.com</p>
	</div>
</xsl:template>
</xsl:stylesheet>