<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/logs">
<table align="center"  Cellspacing="0" Cellpadding="2">
	<thead>
		<tr title="جهت نمايش/پنهان كردن تاريخچه كليك كنيد">
			<th>تاريخ</th>
			<th>ساعت</th>
			<th>مرحله</th>
			<th></th>
			<th></th>
			<th></th>
			<th></th>
			<th></th>
			<th>بها</th>
			<th>ثبت كننده</th>
		</tr>
	</thead>
	<tbody>
<xsl:for-each select="./order">
		<tr>
			<xsl:attribute name="class">
				<xsl:if test="./status!=1">pause</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="logID">
				<xsl:value-of select="./id"/>
			</xsl:attribute>
			<td class="date showLogedOrder" title="جهت نمايش كليك كنيد"><xsl:value-of select="./insertedDate"/></td>
			<td class="time"><xsl:value-of select="./insertedDate"/></td>
			<td><xsl:value-of select="./stepName"/></td>
			<td>
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
			</td>
			<td>
				<xsl:attribute name="title">
					<xsl:if test="./isPrinted!=0">چاپ شده</xsl:if>
				</xsl:attribute>
				<xsl:if test="./isPrinted!=0"><i class="icon-print"></i></xsl:if>
			</td>
			<td>
				<xsl:attribute name="title">
					<xsl:if test="./isOrder!=0">سفارش</xsl:if>
					<xsl:if test="./isOrder=0">استعلام</xsl:if>
				</xsl:attribute>
				<xsl:if test="./isOrder!=0"><i class="icon-book"></i></xsl:if>
				<xsl:if test="./isOrder=0"><i class="icon-list-alt"></i></xsl:if>
			</td>
			<td>
				<xsl:attribute name="title">
					<xsl:if test="./isApproved!=0">تاييد شده</xsl:if>
				</xsl:attribute>
				<xsl:if test="./isApproved!=0"><i class="icon-ok-sign"></i></xsl:if>
			</td>
			<td>
				<xsl:attribute name="title">
					<xsl:if test="./customerApprovedType!=0">
						<xsl:value-of select="./customerApprovedTypeName"/>
					</xsl:if>
				</xsl:attribute>
				<xsl:if test="./customerApprovedType!=0"><i class="icon-thumbs-up"></i></xsl:if>
			</td>
			<td><xsl:value-of select="./price"/></td>
			<td><xsl:value-of select="./realName"/></td>
		</tr>
</xsl:for-each>
	</tbody>
</table>
</xsl:template>
</xsl:stylesheet>