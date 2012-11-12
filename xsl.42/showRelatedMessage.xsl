<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/rows">
<xsl:variable name="class" select="CusTD3" />
<fieldset class="msg">
	<legend>
		<span>يادداشت‌ها</span>
		<input class="btn btn-info" type="button" value="نوشتن يادداشت" name="addNewMessage" id="addNewMessage"/>
	</legend>
	<ul class="msg">
		<xsl:for-each select="./row">
			<li>
				<xsl:attribute name="class">
					<xsl:if test="./Urgent='0'">msg urgent0</xsl:if>
					<xsl:if test="./Urgent='1'">msg urgent1</xsl:if>
					<xsl:if test="./Urgent='2'">msg urgent2</xsl:if>
				</xsl:attribute>
				<div class="clearfix">
					<xsl:if test="./toName!=''">
						<div class="rfloat label label-inverse absBottomLeft">
							<xsl:value-of select="./toName"/>
						</div>
					</xsl:if>
					<xsl:if test="./typeName!=''">
						<div class="rfloat label label-info absTopLeft">
							<xsl:value-of select="./typeName"/>
						</div>
					</xsl:if>
					<div class="clearfix">
						<span class="rpad badge">
							<xsl:value-of select="./fromName"/>
						</span>
						<span class="rpad msgBody">
							<xsl:value-of select="./MsgBody"/>
						</span>
					</div>
					<div class="msgDate">
						<span>
							<xsl:value-of select="./MsgDate"/>
						</span>
						<span class="lpad">
							<xsl:value-of select="./MsgTime"/>
						</span>
					</div>
				</div>
				
			</li>
		</xsl:for-each>	
	</ul>
	<xsl:if test="count(//row)=0">	
		<div>هيچ</div>
	</xsl:if>
</fieldset>
</xsl:template>
</xsl:stylesheet>
