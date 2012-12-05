<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/rows">
	<table class="list" border="1" cellspacing="0" cellpadding="2" dir="RTL"  borderColor="#555588" >
		<TR valign="top" class="head">
			<TD>#</TD>
			<TD width="40">�����</TD>
			<TD width="55"><small>����� ���<br/>����� ����<br/>����� �����</small></TD>
			<TD width="55"><small>����� ����� ������<br/>����� ����� ����</small></TD>
			<TD width="130">��� ���� - �����</TD>
			<TD width="80">����� ���</TD>
			<TD width="36">���</TD>
			<TD width="45">�����</TD>
			<TD width="38">����� ������</TD>
			<TD width="18">���</TD>
			<TD width="30"><small>��������</small></TD>
			<td width="50">����</td>
		</TR>
		<xsl:for-each select="./row">
			<tr>
				<td/>
				<td>
					<a>
						<xsl:attribute name="href">order.asp?act=show&amp;id=<xsl:value-of select="./id"/></xsl:attribute>
						<xsl:value-of select="./id"/>
					</a>
				</td>
				
				<td class="orderDates">
					<span class="faDate">
						<xsl:value-of select="./createdDate"/>
					</span>
					<br/>
					<span class="faDate">
						<xsl:value-of select="./approvedDate"/>
					</span>
					<br/>
					<span class="faDate">
						<xsl:value-of select="./orderDate"/>
					</span>
				</td>
				<td class="orderDates">
					<span class="faDate">
						<xsl:value-of select="./contractDate"/>
					</span>
					<br/>
					<span class="faDate">
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
					<div>
						<xsl:attribute name="title">
							<xsl:value-of select="./statusName"/>
						</xsl:attribute>
						<xsl:if test="./status=1">
							<i class="icon-play"></i>
						</xsl:if>
						<xsl:if test="./status=2">
							<i class="icon-trash"></i>
						</xsl:if>
						<xsl:if test="./status=3">
							<i class="icon-pause"></i>
						</xsl:if>
						<xsl:if test="./status=4">
							<i class="icon-question-sign"></i>
						</xsl:if>
					</div>
					<div>
						<xsl:attribute name="title">
							<xsl:if test="./isPrinted!=0">�ǁ ���</xsl:if>
						</xsl:attribute>
						<xsl:if test="./isPrinted!=0"><i class="icon-print"></i></xsl:if>
					</div>
					<div>
						<xsl:attribute name="title">
							<xsl:if test="./isOrder!=0">�����</xsl:if>
							<xsl:if test="./isOrder=0">�������</xsl:if>
						</xsl:attribute>
						<xsl:if test="./isOrder!=0"><i class="icon-book"></i></xsl:if>
						<xsl:if test="./isOrder=0"><i class="icon-list-alt"></i></xsl:if>
					</div>
					<div>
						<xsl:attribute name="title">
							<xsl:if test="./isApproved!=0">���� ���</xsl:if>
						</xsl:attribute>
						<xsl:if test="./isApproved!=0"><i class="icon-ok-sign"></i></xsl:if>
					</div>
					<div>
						<xsl:attribute name="title">
							<xsl:if test="./customerApprovedType!=0">
								<xsl:value-of select="./customerApprovedTypeName"/>
							</xsl:if>
						</xsl:attribute>
						<xsl:if test="./customerApprovedType!=0"><i class="icon-thumbs-up"></i></xsl:if>
					</div>
				</td>
				<td>
					<xsl:if test="./invoiceID=0">
						<xsl:value-of select="./invStatus"/>
					</xsl:if>
					<xsl:if test="./invoiceID>0">
						<a>
							<xsl:attribute name="href">../AR/AccountReport.asp?act=showInvoice&amp;invoice=<xsl:value-of select="./invoiceID"/></xsl:attribute>
							<xsl:value-of select="./invStatus"/>
						</a>
					</xsl:if>
				</td>
				<td><xsl:value-of select="./Price"/></td>
			</tr>
		</xsl:for-each>
		<tr class="sumTotal">
			<td colspan="11">���</td>
			<td/>
		</tr>
	</table>
</xsl:template>
</xsl:stylesheet>