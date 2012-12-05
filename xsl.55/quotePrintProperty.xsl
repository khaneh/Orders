<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/keys">
	<xsl:for-each select="./row">
		<xsl:for-each select="./group">
			<xsl:variable name="gname" select="./@name"/>
			<div>
				<xsl:attribute name="class">
					<xsl:if test="./@name!='desc'">rspan4</xsl:if>
				</xsl:attribute>
				<li>
					<xsl:choose>
						<xsl:when test="./@name='desc'">
							<b>
								<xsl:value-of select="./key[@name='description']"/> 
								<xsl:if test="./key[@name='pages']!=''"> (<xsl:value-of select="./key[@name='pages']"/> ����)</xsl:if>
							</b>
							<br/>
						</xsl:when>
						<xsl:when test="./@name='proof'">
							<span><xsl:value-of select="./key[@name='proof-form']"/></span>
							<span> ��� </span>
							<span><xsl:value-of select="./@label"/>(</span>								
							<xsl:for-each select="./key[@type='check']">
								<span><xsl:value-of select="./@label"/> </span>
							</xsl:for-each>
							<span>)</span>
							<span> �� ����� </span>
							<span><xsl:value-of select="./key[@name='proof-size']"/></span>
							<span> ����� �� </span>
							<span><xsl:value-of select="./key[@name='proof-date']"/></span>
						</xsl:when>
						<xsl:when test="./@name='plate'">
							<span><xsl:value-of select="./key[@name='plate-qtty']"/> ��� </span>
							<span><xsl:value-of select="./@label"/><xsl:value-of select="./key[@name='plate-color-count']"/>&#160;�� &#160;<xsl:value-of select="./key[@name='plate-type']"/></span>
						</xsl:when>
						<xsl:when test="./@name='print'">
							<span><xsl:value-of select="./key[@name='print-form']"/> ��� �� ����� </span>
							<span><xsl:value-of select="./key[@name='print-qtty']"/> ��� </span>
							<span><xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='print-type']"/> </span>
							<span><xsl:value-of select="./key[@name='print-color']"/> �� </span>
							<xsl:if test="./key[@name='print-sp-color']>0">
								<span><xsl:value-of select="./key[@name='print-sp-color']"/> �� ��� </span>
							</xsl:if>
							<span><xsl:value-of select="./key[@name='print-d-type']"/> ��� </span>
							<span><xsl:value-of select="./key[@name='print-control']"/></span>
						</xsl:when>
						<xsl:when test="./@name='paper'">
							<span><xsl:value-of select="./key[@name='paper-qtty']"/> ���</span>
							<span><xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='paper-weight']"/> ���� <xsl:value-of select="./key[@name='paper-type']"/> </span>
							<span><xsl:value-of select="./key[@name='paper-size']"/> </span>
						</xsl:when>
						<xsl:when test="./@name='binding'">
							<span><xsl:value-of select="./key[@name='binding-form']"/> ��� </span>
							<span><xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='binding-type']"/> </span>
							<span><xsl:value-of select="./key[@name='binding-qtty']"/> ��� </span>
							<xsl:if test="./key[@name='binding-orient']!='��'">
								<span><xsl:value-of select="./key[@name='binding-orient']"/></span>
							</xsl:if>
						</xsl:when>
						<xsl:when test="./@name='packing'">
							<span><xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='packing-type']"/> </span>
							<span><xsl:value-of select="./key[@name='packing-qtty']"/> ���</span>
						</xsl:when>
						<xsl:when test="./@name='delivery'">
							<span><xsl:value-of select="./@label"/> �� <xsl:value-of select="./key[@name='delivery-address']"/> </span>
						</xsl:when>
						<xsl:when test="./@name='verni'">
							<span>&#160;<xsl:value-of select="./key[@name='verni-form']"/> ��� </span>
							<span><xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='verni-mat']"/></span>
							<span>&#160;<xsl:value-of select="./key[@name='verni-wat']"/></span>
							<span>&#160;<xsl:value-of select="./key[@name='verni-qtty']"/> ���</span>
							<span>&#160;<xsl:value-of select="./key[@name='verni-type']"/></span>
						</xsl:when>
						<xsl:when test="./@name='selefon'">
							<span>&#160;<xsl:value-of select="./key[@name='selefon-form']"/> ���</span>
							<span>&#160;<xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='selefon-face']"/></span>
							<span>&#160;<xsl:value-of select="./key[@name='selefon-mat']"/></span>
							<span>&#160;<xsl:value-of select="./key[@name='selefon-hot']"/></span>
							<span>&#160;<xsl:value-of select="./key[@name='selefon-qtty']"/> ���</span>
							<span>&#160;<xsl:value-of select="./key[@name='selefon-size']"/></span>
						</xsl:when>
						<xsl:when test="./@name='uv'">
							<span>&#160;<xsl:value-of select="./key[@name='uv-form']"/> ���</span>
							<span>&#160;<xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='uv-spot']"/></span>
							<span>&#160;<xsl:value-of select="./key[@name='uv-mat']"/></span>
							<span>&#160;<xsl:value-of select="./key[@name='uv-qtty']"/> ���</span>
							<span>&#160;<xsl:value-of select="./key[@name='uv-size']"/></span>
						</xsl:when>
						<xsl:when test="./@name='fold'">
							<span><xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='fold-count']"/> ���</span>
						</xsl:when>
						<xsl:when test="./@name='service'">
							<span><xsl:value-of select="./key[@name='service-item']"/>&#160;<xsl:value-of select="./key[@name='service-qtty']"/> ���</span>
						</xsl:when>
						<xsl:when test="./@name='serviceIN'">
							<span><xsl:value-of select="./key[@name='serviceIN-item']"/>&#160;<xsl:value-of select="./key[@name='serviceIN-qtty']"/> ���</span>
						</xsl:when>
						<xsl:when test="./@name='cutting'">
							<span><xsl:value-of select="./@label"/>&#160;<xsl:value-of select="./key[@name='cutting-count']"/> ���</span>
						</xsl:when>
						<xsl:when test="./@name='snap'">
							<span><xsl:value-of select="./key[@name='snap-qtty']"/>&#160;<xsl:value-of select="./@label"/></span>
							<span>&#160; �� ����� <xsl:value-of select="./key[@name='snap-size']"/></span>
						</xsl:when>
						<xsl:when test="./@name='invRequest'">
							<span><xsl:value-of select="./key[@name='invRequest-qtty']"/></span>
							<span>&#160;<xsl:value-of select="./key[@name='invRequest-name']"/></span>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="./@label"/>
						</xsl:otherwise>
					</xsl:choose>
				</li>
			</div>
		</xsl:for-each>
		<div class="clear"></div>
	</xsl:for-each>
</xsl:template>
</xsl:stylesheet>