<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/costs">
	<TABLE border="0" cellspacing="0" cellpadding="2" align="center" style=" color:black; direction:RTL; width:700; ">
		<tr bgcolor="black">
			<td align="center" colspan="8" style="color:yellow;padding:3px 0 8px 0;">عمليات‌هاي انجام شده براي اين سفارش:</td>
		</tr>
		<tr bgcolor="black" style="color:yellow;">
			<td>مركز</td>
			<td>درايور</td>
			<td>عمليات</td>
			<td>تعداد</td>
			<td>زمان</td>
			<td>شرح</td>
			<td>ثبت كننده</td>
			<td>شماره سفارش</td>
		</tr>
		<xsl:for-each select="./cost">
			<tr bgcolor="#CCCCCC">
				<td class="delCost">
					<xsl:value-of select="./centerName"/>
					<xsl:if test="./canEdit='yes'">
						<span class="delCostBtn label label-important">
							<xsl:attribute name="costID"><xsl:value-of select="./id"/></xsl:attribute>حذف</span>
					</xsl:if>
				</td>
				<td><xsl:value-of select="./driverName"/></td>
				<td><xsl:value-of select="./operationName"/></td>
				<td><xsl:value-of select="./theCount"/></td>
				<td><xsl:value-of select="./theTime"/></td>
				<td><xsl:value-of select="./description"/></td>
				<td>
					<xsl:attribute name="title">تاريخ ثبت: <xsl:value-of select="./insertDate"/></xsl:attribute>
					<xsl:value-of select="./realName"/>
				</td>
				<td>
					<a>
						<xsl:attribute name="href">../order/order.asp?act=show&amp;id=<xsl:value-of select="./order"/></xsl:attribute>
						<xsl:value-of select="./order"/>
					</a>
				</td>
				
			</tr>
		</xsl:for-each>
	</TABLE>
</xsl:template>
</xsl:stylesheet>
