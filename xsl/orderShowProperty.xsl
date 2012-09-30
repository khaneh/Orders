<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/keys">
<xsl:for-each select="./row">
	<div class='myRow'>
	<div class='exteraArea'>
		<xsl:attribute name="type"><xsl:value-of select="./@name"/></xsl:attribute>
		<xsl:if test="./@label!=''">
			<div class="rowTitle">
				<h3><xsl:value-of select="./@label"/></h3>
			</div>
		</xsl:if>
	<xsl:for-each select="./group">
		<div class='group'>
		<xsl:attribute name="groupName"><xsl:value-of select="./@name"/></xsl:attribute>
		<fieldset class="group">
			<xsl:if test="./@label!=''">
				<legend>
					<span class="groupName"><xsl:value-of select="./@label"/></span>
				</legend>
			</xsl:if>
			<xsl:if test="./@isPrice!='no'">
				<div class="priceGroup">
					<div class="priceGroupTitle">
						<span class="sp1">بهاي كل</span>
						<span class="sp2">تخفيف</span>
						<span class="sp2">مازاد</span>
					</div>
					<div class="clear"></div>
					<div class="priceGroup">
					<xsl:variable name="gname" select="./@name" />
						<span class="price sp1"><xsl:value-of select="./key[@name=concat($gname,'-price')]"/></span>
						<span class="dis sp2"><xsl:value-of select="./key[@name=concat($gname,'-dis')]"/></span>
						<span class="over sp2"><xsl:value-of select="./key[@name=concat($gname,'-over')]"/></span>
					</div>
				</div>
			</xsl:if>
			
			<xsl:for-each select="./key[@type!='price']">
				<div class="rspan-f">
					<div class="rspan1">
						<label class='key-label'>
							<xsl:value-of select="./@label"/>
						</label>
					</div>
					<div class="rspan2">
						<xsl:choose>
							<xsl:when test="./@type='text'">
								<span class="data"><xsl:value-of select="."/></span>
							</xsl:when>
							<xsl:when test="./@type='textarea'">
								<span class="data"><xsl:value-of select="."/></span>
							</xsl:when>
							<xsl:when test="./@type='check'">
								<span class="data"><xsl:value-of select="."/></span>
							</xsl:when>
							<xsl:when test="./@type='radio'">
								<span class="data"><xsl:value-of select="."/></span>
							</xsl:when>
							<xsl:when test="./@type='option'">
								<span class="data"><xsl:value-of select="."/></span>
							</xsl:when>
							<xsl:when test="./@type='option-other'">
								<span class="data"><xsl:value-of select="."/></span>
							</xsl:when>
							<xsl:when test="./@type='option-fromTable'">
								<span class="data"><xsl:value-of select="."/></span>
							</xsl:when>
						</xsl:choose>
					</div>
				</div>
			</xsl:for-each>
		</fieldset>
		</div>
	</xsl:for-each>
	</div>
	<div>
		<xsl:attribute name="id">extreArea-<xsl:value-of select="./@name"/></xsl:attribute>
	</div>
	</div>
	<hr/>
</xsl:for-each>
</xsl:template>
</xsl:stylesheet>