<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/rows">
	<table class="RepTable" width="90%" align="center">
		<tr>
			<td>#</td>
			<td>تاريخ</td>
			<td>نام حساب</td>
			<td>توضيحات</td>
			<td>بدهكار</td>
			<td>بستانكار</td>
		</tr>
		<xsl:for-each select="./row">
		<tr>
			<xsl:if test="./Voided=-1">
				<xsl:attribute name="class">voided</xsl:attribute>
			</xsl:if>
			<td>#</td>
			<td><xsl:value-of select="./EffectiveDate"/></td>
			<td>
				<a>
					<xsl:attribute name="href">../CRM/AccountInfo.asp?act=show&amp;SelectedCustomer=<xsl:value-of select="./Account"/></xsl:attribute>
					<xsl:value-of select="./AccountTitle"/>
				</a>
			</td>
			<td>
				<a>
					<xsl:attribute name="href">AccountReport.asp?act=showMemo&amp;sys=AR&amp;memo=<xsl:value-of select="./Link"/></xsl:attribute>
					<xsl:value-of select="./Description"/>
				</a>
			</td>
			<td class="debit">
				<xsl:if test="./IsCredit=0">
					<xsl:value-of select="./AmountOriginal"/>
				</xsl:if>
				<xsl:if test="./IsCredit=-1">0</xsl:if>
			</td>
			<td class="credit">
				<xsl:if test="./IsCredit=-1">
					<xsl:value-of select="./AmountOriginal"/>
				</xsl:if>
				<xsl:if test="./IsCredit=0">0</xsl:if>
			</td>

		</tr>
		</xsl:for-each>
		<tr>
			<td colspan="4"/>
			<td id="sumDebit"></td>
			<td id="sumCredit"></td>
		</tr>
	</table>
</xsl:template>
</xsl:stylesheet>