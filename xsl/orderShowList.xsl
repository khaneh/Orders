<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/rows">
	<table class="list" border="1" cellspacing="0" cellpadding="2" dir="RTL"  borderColor="#555588" >
		<TR valign="top" class="head">
			<TD width="40">#</TD>
			<TD width="65">تاريخ ثبت<br/>تاريخ تحويل/اعتبار</TD>
			<TD width="130">نام شركت - مشتري</TD>
			<TD width="80">عنوان كار</TD>
			<TD width="36">نوع</TD>
			<TD width="45">مرحله</TD>
			<TD width="38">سفارش گيرنده</TD>
			<TD width="18">وضع</TD>
			<TD width="30">صورتحساب</TD>
			<td width="50">قيمت</td>
		</TR>
		<xsl:for-each select="./row">
			<tr>
				<td>
					<a>
						<xsl:attribute name="href">order.asp?act=show&amp;id=<xsl:value-of select="./id"/></xsl:attribute>
						<xsl:value-of select="./id"/>
					</a>
				</td>
				<td class="orderDates">
					<span class="createdDate">
						<xsl:value-of select="./createdDate"/>
					</span>
					<span class="createdTime">
						<xsl:value-of select="./createdDate"/>
					</span>
					<br/>
					<span class="returnDate">
						<xsl:value-of select="./returnDate"/>
					</span>
					<span class="returnTime">
						<xsl:value-of select="./returnDate"/>
					</span>
				</td>
				<td>
					<xsl:attribute name="title"><xsl:if test="string-length(./Tel1)>3"><xsl:value-of select="./Tel1"/>,	</xsl:if><xsl:if test="string-length(./Tel2)>3"><xsl:value-of select="./Tel2"/>,	</xsl:if><xsl:if test="string-length(./Mobile1)>3"><xsl:value-of select="./Mobile1"/>,	</xsl:if><xsl:if test="string-length(./Mobile2)>3"><xsl:value-of select="./Mobile2"/></xsl:if>
					</xsl:attribute>
					<a>
						<xsl:attribute name="href">../CRM/AccountInfo.asp?act=show&amp;selectedCustomer=<xsl:value-of select="./Customer"/></xsl:attribute>
						<xsl:value-of select="./AccountTitle"/>					
						<br/>
						<xsl:if test="string-length(./Tel1)>3">
							<xsl:value-of select="./Tel1"/>
						</xsl:if>
						<xsl:if test="3>string-length(./Tel1)">
							<xsl:value-of select="./Tel2"/>
						</xsl:if>
					</a>
				</td>
				<td><xsl:value-of select="./orderTitle"/></td>
				<td><xsl:value-of select="./typeName"/></td>
				<td><xsl:value-of select="./stepName"/></td>
				<td><xsl:value-of select="./RealName"/></td>
				<td>
					<xsl:attribute name="title">
						<xsl:value-of select="./statusName"/>
					</xsl:attribute>
					<img>
						<xsl:attribute name="src">
							<xsl:value-of select="./statusIcon"/>
						</xsl:attribute>
					</img>
				</td>
				<td><xsl:value-of select="./invStatus"/></td>
				<td><xsl:value-of select="./Price"/></td>
			</tr>
		</xsl:for-each>
	</table>
</xsl:template>
</xsl:stylesheet>