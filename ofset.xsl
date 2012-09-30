<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/keys">
<xsl:for-each select="./row">
	<div class='myRow'>
	<input type='hidden' value='0'>
		<xsl:attribute name="id"><xsl:value-of select="./@name"/>-maxID</xsl:attribute>
	</input>
	<div class='exteraArea'>
		<xsl:attribute name="id"><xsl:value-of select="./@name"/>-0</xsl:attribute>
		<xsl:if test="./@label!=''">
			<div class="rowTitle">
				<i class="icon-chevron-left"/>
				<xsl:value-of select="./@label"/>
			</div>
		</xsl:if>
	<xsl:for-each select="./group">
		<div class='group'>
		<xsl:attribute name="groupName"><xsl:value-of select="./@name"/></xsl:attribute>
		<xsl:if test="./@isPrice!='no'">
			<div class="priceGroupTitle">
				<span class="sp1">بهاي كل</span>
				<span class="sp2">تخفيف</span>
				<span class="sp2">مازاد</span>
			</div>
			<div class="clear"></div>
			<div class="priceGroup">
				<input type="text" size="12" class="myInput" value="0" dir="ltr">
					<xsl:attribute name="name"><xsl:value-of select="./@name"/>-price</xsl:attribute>
					<xsl:attribute name="onblur">calc_<xsl:value-of select="./@name"/>(this)</xsl:attribute>
					<xsl:if test="./@disable='yes'"><xsl:attribute name="disabled">true</xsl:attribute></xsl:if>
					<xsl:if test="./@isPrice='yes'"><xsl:attribute name="readonly">true</xsl:attribute></xsl:if>
				</input>
				<input type="text" size="2" class="myInput" value="0%" onclick="clearThis(this);" dir="ltr">
					<xsl:attribute name="name"><xsl:value-of select="./@name"/>-dis</xsl:attribute>
					<xsl:attribute name="onblur">calc_<xsl:value-of select="./@name"/>(this)</xsl:attribute>
					<xsl:if test="./@disable='yes'"><xsl:attribute name="disabled">true</xsl:attribute></xsl:if>
				</input>
				<input type="text" size="2" class="myInput" value="0%" onclick="clearThis(this);" dir="ltr">
					<xsl:attribute name="name"><xsl:value-of select="./@name"/>-over</xsl:attribute>
					<xsl:attribute name="onblur">calc_<xsl:value-of select="./@name"/>(this)</xsl:attribute>
					<xsl:if test="./@disable='yes'"><xsl:attribute name="disabled">true</xsl:attribute></xsl:if>
				</input>
			</div>
		</xsl:if>
		<xsl:if test="./@label!=''">
			<div class="groupTitle">
				<i class="icon-chevron-left icon-white"/>
				<xsl:if test="./@disable='yes'">
					<input type='checkbox' value='0'>
						<xsl:attribute name="name"><xsl:value-of select="./@name"/>-disBtn</xsl:attribute>
						<xsl:attribute name="onclick">disGroup(this);<xsl:if test="./@blur='yes'">calc_<xsl:value-of select="./@name"/>($(this).parent());</xsl:if></xsl:attribute>
					</input>
				</xsl:if>
				<xsl:value-of select="./@label"/>
			</div>
		</xsl:if>
		<xsl:for-each select="./key">
			<label class='key-label'>
				<xsl:value-of select="./@label"/>
			</label>
			<xsl:choose>
				<xsl:when test="./@type='text'">
					<input type='text'>
						<xsl:attribute name="name">
							<xsl:value-of select="./@name"/>
						</xsl:attribute>
						<xsl:attribute name="size">
							<xsl:value-of select="."/>
						</xsl:attribute>
						<xsl:if test="./@readonly='yes'">
							<xsl:attribute name="readonly">true</xsl:attribute>
						</xsl:if>
						<xsl:if test="./@default!=''">
							<xsl:attribute name="value">
								<xsl:value-of select="./@default"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:attribute name="group">
							<xsl:value-of select="../../@name"/>
						</xsl:attribute>
						<xsl:if test="../@disable='yes'">
							<xsl:attribute name="disabled">true</xsl:attribute>
						</xsl:if>
						<xsl:if test="./@blur='yes'">
							<xsl:attribute name="onblur">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
						</xsl:if>
						<xsl:if test="./@force='yes'">
							<xsl:attribute name="force">yes</xsl:attribute>
						</xsl:if>
					</input>
				</xsl:when>
				<xsl:when test="./@type='textarea'">
					<textarea class='myDesc'>
						<xsl:attribute name="name">
							<xsl:value-of select="./@name"/>
						</xsl:attribute>
						<xsl:attribute name="cols">
							<xsl:value-of select="."/>
						</xsl:attribute>
						<xsl:if test="../@disable='yes'">
							<xsl:attribute name="disabled">true</xsl:attribute>
						</xsl:if>
						<xsl:if test="./@force='yes'">
							<xsl:attribute name="force">yes</xsl:attribute>
						</xsl:if>
						<xsl:if test="./@blur='yes'">
							<xsl:attribute name="onblur">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
						</xsl:if>
						<xsl:if test="./@default!=''"><xsl:value-of select="./@default"/></xsl:if>
					</textarea>
				</xsl:when>
				<xsl:when test="./@type='check'">
					<input type='checkbox' value='on-0'>
						<xsl:attribute name="name">
							<xsl:value-of select="./@name"/>
						</xsl:attribute>
						<xsl:if test=".='checked'">
							<xsl:attribute name="checked">true</xsl:attribute>
						</xsl:if>
						<xsl:if test="../@disable='yes'">
							<xsl:attribute name="disabled">true</xsl:attribute>
						</xsl:if>
						<xsl:if test="./@blur='yes'">
							<xsl:attribute name="onclick">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
						</xsl:if>
						<xsl:if test="./@price!=''">
							<xsl:attribute name="price">
								<xsl:value-of select="./@price"/>
							</xsl:attribute>
						</xsl:if>
					</input>
				</xsl:when>
				<xsl:when test="./@type='radio'">
					<xsl:for-each select="./option">
						<label>
							<xsl:value-of select="./@label"/>
						</label>
						<input type='radio'>
							<xsl:attribute name="name">
								<xsl:value-of select="../@name"/>
							</xsl:attribute>
							<xsl:attribute name="value">
								<xsl:value-of select="."/>
							</xsl:attribute>
							<xsl:if test="../../@disable='yes'">
								<xsl:attribute name="disabled">true</xsl:attribute>
							</xsl:if>
							<xsl:if test="./@default='yes'">
								<xsl:attribute name="checked">true</xsl:attribute>
							</xsl:if>
							<xsl:if test="./@blur='yes'">
								<xsl:attribute name="onchange">calc_<xsl:value-of select="../../@name"/>(this);</xsl:attribute>
							</xsl:if>
						</input>
					</xsl:for-each>
				</xsl:when>
				<xsl:when test="./@type='option'">
					<select>
						<xsl:attribute name="name">
							<xsl:value-of select="./@name"/>
						</xsl:attribute>
						<xsl:if test="../@disable='yes'">
							<xsl:attribute name="disabled">true</xsl:attribute>
						</xsl:if>
						<xsl:if test="../@blur='yes'">
							<xsl:attribute name="onchange">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
						</xsl:if>
						<xsl:for-each select="./option">
							<option>
								<xsl:attribute name="value">
									<xsl:value-of select="."/>
								</xsl:attribute>
								<xsl:if test="./@price!=''">
									<xsl:attribute name="price">
										<xsl:value-of select="./@price"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test="./@addr!=''">
									<xsl:attribute name="addr">
										<xsl:value-of select="./@addr"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:value-of select="./@label"/>
							</option>
						</xsl:for-each>
					</select>
				</xsl:when>
				<xsl:when test="./@type='option-other'">
					<select>
						<xsl:attribute name="name">
							<xsl:value-of select="./@name"/>
						</xsl:attribute>
						<xsl:if test="../@disable='yes'">
							<xsl:attribute name="disabled">true</xsl:attribute>
						</xsl:if>
						<xsl:if test="../@blur='yes'">
							<xsl:attribute name="onchange">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
						</xsl:if>
						<xsl:for-each select="./option">
							<option>
								<xsl:attribute name="value">
									<xsl:value-of select="."/>
								</xsl:attribute>
								<xsl:if test="./@price!=''">
									<xsl:attribute name="price">
										<xsl:value-of select="./@price"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:value-of select="./@label"/>
							</option>
						</xsl:for-each>
						<option value='-1'>ساير</option>
					</select>
					<input type='text' onblur='addOther(this);'>
						<xsl:attribute name="name"><xsl:value-of select="./@name"/>-addValue</xsl:attribute>
					</input>
				</xsl:when>
				<xsl:when test="./@type='option-fromTable'">
					<select class='fromTable'>
						<xsl:attribute name="name">
							<xsl:value-of select="./@name"/>
						</xsl:attribute>
						<xsl:attribute name="table">
							<xsl:value-of select="./@table"/>
						</xsl:attribute>
						<xsl:if test="../@disable='yes'">
							<xsl:attribute name="disabled">true</xsl:attribute>
						</xsl:if>
						<xsl:if test="../@blur='yes'">
							<xsl:attribute name="onchange">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
						</xsl:if>
					</select>
				</xsl:when>
			</xsl:choose>
			<xsl:if test="./@br='yes'">
				<br/>
			</xsl:if>
		</xsl:for-each>
		</div>
	</xsl:for-each>
	</div>
	<div>
		<xsl:attribute name="id">extreArea-<xsl:value-of select="./@name"/></xsl:attribute>
	</div>
	<div class="funcBtn">
	<p>
		<a class="btn btn-danger" title="حذف آخرين خط" >
			<xsl:attribute name="onclick">removeRow('<xsl:value-of select="./@name"/>');</xsl:attribute>
			<i class="icon-trash icon-white" />
		</a>
		<a class="btn btn-success" title="اضافه" >
			<xsl:attribute name="onclick">cloneRow('<xsl:value-of select="./@name"/>');</xsl:attribute>
			<i class="icon-plus-sign icon-white" />
		</a>
	</p>
	</div>
	</div>
</xsl:for-each>
</xsl:template>
</xsl:stylesheet>