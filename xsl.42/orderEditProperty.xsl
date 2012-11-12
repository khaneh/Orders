<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/keys">
<xsl:for-each select="./rows">
	<xsl:variable name="rname" select="./@name" />
	<div class='myRow'>
	<xsl:for-each select="./row">
		<div class='exteraArea'>
			<xsl:attribute name="type"><xsl:value-of select="$rname"/></xsl:attribute>
			<xsl:if test="../@label!=''">
				<div class="rowTitle">
					<h3><xsl:value-of select="../@label"/></h3>
				</div>
			</xsl:if>
				<div class="funcBtn">
				<p>
					<a class="btn btn-danger" rel="tooltip" data-original-title="جهت حذف آخرين خط كليك كنيد" onclick="removeRow(this);">
						<i class="icon-trash icon-white" />
					</a>
					<a class="btn btn-success" rel="tooltip" data-original-title="جهت اضافه شدن سطري ديگر كليك كنيد" onclick="cloneRow(this)">
						<i class="icon-plus-sign icon-white" />
					</a>
				</p>
				</div>

		<xsl:for-each select="./group">
			<div class='group'>
			<xsl:attribute name="groupName"><xsl:value-of select="./@name"/></xsl:attribute>
			<fieldset class="group">
			<xsl:if test="./@label!=''">
				<legend>
					<span class="groupName"><xsl:value-of select="./@label"/></span>
					<xsl:if test="./@disable='yes'">
						<input type='checkbox' value='0' class="disBtn">
							<xsl:attribute name="name"><xsl:value-of select="./@name"/>-disBtn</xsl:attribute>
							<xsl:attribute name="onclick">disGroup(this);<xsl:if test="./@blur='yes'">calc_<xsl:value-of select="./@name"/>($(this).parent());</xsl:if></xsl:attribute>
							<xsl:if test="./@hasValue='yes'">
								<xsl:attribute name="checked">true</xsl:attribute>
							</xsl:if>
						</input>
					</xsl:if>
				</legend>
			</xsl:if>
			<xsl:if test="./@isPrice!='no'">
				<div class="priceGroup">
					<div class="priceGroupValue">
						<xsl:variable name="gname" select="./@name" />
						<div>
							<input type="text" size="12" class="myInput" out='yes' dir="ltr" rel="tooltip" data-placement="right" data-original-title="بهاي كل">
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-price</xsl:attribute>
								<xsl:attribute name="onblur">calc_<xsl:value-of select="./@name"/>(this)</xsl:attribute>
								<xsl:if test="./@disable='yes' and ./@hasValue!='yes'"><xsl:attribute name="disabled">true</xsl:attribute></xsl:if>
								<xsl:if test="./@isPrice='yes'"><xsl:attribute name="readonly">true</xsl:attribute></xsl:if>
								<xsl:if test="concat(./@price,'')=''">
									<xsl:attribute name="value">0</xsl:attribute>
								</xsl:if>
								<xsl:if test="./@price!=''">
									<xsl:attribute name="value">
										<xsl:value-of select="./@price"/>
									</xsl:attribute>
								</xsl:if>
							</input>
						</div>
						<div>
							<input type="text" size="12" class="myInput" onclick="clearThis(this);" out='yes' dir="ltr" rel="tooltip" data-placement="right" data-original-title="تخفيف">
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-dis</xsl:attribute>
								<xsl:attribute name="onblur">calc_<xsl:value-of select="./@name"/>(this)</xsl:attribute>
								<xsl:if test="concat(./@disable,'')='yes' and concat(./@hasValue,'')!='yes'"><xsl:attribute name="disabled">true</xsl:attribute></xsl:if>
								<xsl:if test="concat(./@dis,'')=''">
									<xsl:attribute name="value">0</xsl:attribute>
								</xsl:if>
								<xsl:if test="./@dis!=''">
									<xsl:attribute name="value">
										<xsl:value-of select="./@dis"/>
									</xsl:attribute>
								</xsl:if>
							</input>
						</div>
						<div>
							<input type="text" size="12" class="myInput" onclick="clearThis(this);" out='yes' dir="ltr" rel="tooltip" data-placement="right" data-original-title="مازاد">
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-over</xsl:attribute>
								<xsl:attribute name="onblur">calc_<xsl:value-of select="./@name"/>(this)</xsl:attribute>
								<xsl:if test="./@disable='yes' and ./@hasValue!='yes'"><xsl:attribute name="disabled">true</xsl:attribute></xsl:if>
								<xsl:if test="concat(./@over,'')=''">
									<xsl:attribute name="value">0</xsl:attribute>
								</xsl:if>
								<xsl:if test="./@over!=''">
									<xsl:attribute name="value">
										<xsl:value-of select="./@over"/>
									</xsl:attribute>
								</xsl:if>
							</input>
						</div>
						<div>
							<input type="text" size="12" class="myInput" onclick="clearThis(this);" out='yes' dir="ltr" rel="tooltip" data-placement="right" data-original-title="برگشت">
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-reverse</xsl:attribute>
								<xsl:attribute name="onblur">calc_<xsl:value-of select="./@name"/>(this)</xsl:attribute>
								<xsl:if test="./@disable='yes' and ./@hasValue!='yes'"><xsl:attribute name="disabled">true</xsl:attribute></xsl:if>
								<xsl:if test="concat(./@reverse,'')=''">
									<xsl:attribute name="value">0</xsl:attribute>
								</xsl:if>
								<xsl:if test="./@reverse!=''">
									<xsl:attribute name="value">
										<xsl:value-of select="./@reverse"/>
									</xsl:attribute>
								</xsl:if>
							</input>
						</div>
						<input type="hidden" out='yes'>
							<xsl:attribute name="name"><xsl:value-of select="./@name"/>-w</xsl:attribute>
							<xsl:attribute name="value">
								<xsl:value-of select="./@w"/>
							</xsl:attribute>
						</input>
						<input type="hidden" out='yes'>
							<xsl:attribute name="name"><xsl:value-of select="./@name"/>-l</xsl:attribute>
							<xsl:attribute name="value">
								<xsl:value-of select="./@l"/>
							</xsl:attribute>
						</input>
						<xsl:if test="./@hasStock='yes' or ./@hasStock='option'">
							<input type="hidden" out='yes'>
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-stockDesc</xsl:attribute>
								<xsl:attribute name="value">
									<xsl:value-of select="./@stockDesc"/>
								</xsl:attribute>
							</input>
							<input type="hidden" out='yes'>
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-stockName</xsl:attribute>
								<xsl:attribute name="value">
									<xsl:value-of select="./@stockName"/>
								</xsl:attribute>
							</input>
						</xsl:if>
						<xsl:if test="./@hasPurchase='yes'">
							<input type="hidden" name="purchaseName" out='yes'>
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-purchaseName</xsl:attribute>
								<xsl:attribute name="value">
									<xsl:value-of select="./@purchaseName"/>
								</xsl:attribute>
							</input>
							<input type="hidden" name="purchaseDesc" out='yes'>
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-purchaseDesc</xsl:attribute>
								<xsl:attribute name="value">
									<xsl:value-of select="./@purchaseDesc"/>
								</xsl:attribute>
							</input>
							<input type="hidden" name="purchaseTypeID" out='yes'>
								<xsl:attribute name="name"><xsl:value-of select="./@name"/>-purchaseTypeID</xsl:attribute>
								<xsl:attribute name="value">
									<xsl:value-of select="./@purchaseTypeID"/>
								</xsl:attribute>
							</input>
						</xsl:if>
					</div>
				</div>
			</xsl:if>
			<div class="contentGroup">
				<xsl:for-each select="./key">
					<div class="rspan-f">
						<div class="rspan1 key-label" rel="tooltip">
							<xsl:attribute name="data-original-title">
								<xsl:value-of select="./@label"/>	
							</xsl:attribute>
							<xsl:value-of select="./@label"/>
						</div>
						<div class="rspan2">
							<xsl:choose>
								<xsl:when test="./@type='text'">
									<input type='text' out='yes'>
										<xsl:attribute name="name">
											<xsl:value-of select="./@name"/>
										</xsl:attribute>
										<xsl:attribute name="size">
											<xsl:value-of select="."/>
										</xsl:attribute>
										<xsl:if test="./@readonly='yes'">
											<xsl:attribute name="readonly">true</xsl:attribute>
										</xsl:if>
										<xsl:if test="./@value!=''">
											<xsl:attribute name="value">
													<xsl:value-of select="./@value"/>
												</xsl:attribute>
										</xsl:if>
										<xsl:if test="concat(./@value,'')=''">
											<xsl:if test="./@default!=''">
												<xsl:attribute name="value">
													<xsl:value-of select="./@default"/>
												</xsl:attribute>
											</xsl:if>
										</xsl:if>
										<xsl:if test="./@price!=''">
											<xsl:attribute name="price">
												<xsl:value-of select="./@price"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:attribute name="group">
											<xsl:value-of select="../../@name"/>
										</xsl:attribute>
										<xsl:if test="../@disable='yes' and concat(../@hasValue,'')!='yes'">
											<xsl:attribute name="disabled">true</xsl:attribute>
										</xsl:if>
										<xsl:if test="./@blur='yes'">
											<xsl:attribute name="onblur">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
										</xsl:if>
										<xsl:if test="./@force='yes'">
											<xsl:attribute name="force">yes</xsl:attribute>
										</xsl:if>
										<xsl:if test="./@ltr='yes'">
											<xsl:attribute name="dir">ltr</xsl:attribute>
										</xsl:if>
									</input>
								</xsl:when>
								<xsl:when test="./@type='textarea'">
									<textarea class='myDesc' out='yes'>
										<xsl:attribute name="name">
											<xsl:value-of select="./@name"/>
										</xsl:attribute>
										<xsl:attribute name="cols">
											<xsl:value-of select="."/>
										</xsl:attribute>
										<xsl:if test="../@disable='yes' and concat(../@hasValue,'')!='yes'">
											<xsl:attribute name="disabled">true</xsl:attribute>
										</xsl:if>
										<xsl:if test="./@force='yes'">
											<xsl:attribute name="force">yes</xsl:attribute>
										</xsl:if>
										<xsl:if test="./@blur='yes'">
											<xsl:attribute name="onblur">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
										</xsl:if>
										<xsl:if test="concat(./@value,'')=''">
											<xsl:if test="./@default!=''"><xsl:value-of select="./@default"/></xsl:if>
										</xsl:if>
										<xsl:if test="./@value!=''">
											<xsl:value-of select="./@value"/>
										</xsl:if>
									</textarea>
								</xsl:when>
								<xsl:when test="./@type='check'">
									<input type='checkbox' value='on' out='yes'>
										<xsl:attribute name="name">
											<xsl:value-of select="./@name"/>
										</xsl:attribute>
										<xsl:if test="concat(./@value,'')=''">
											<xsl:if test=".='checked'">
												<xsl:attribute name="checked">true</xsl:attribute>
											</xsl:if>
										</xsl:if>
										<xsl:if test="./@value='on'">
											<xsl:attribute name="checked">true</xsl:attribute>
										</xsl:if>
										<xsl:if test="../@disable='yes' and concat(../@hasValue,'')!='yes'">
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
										<input type='radio' out='yes'>
											<xsl:attribute name="name">
												<xsl:value-of select="concat(../@name,'*',../../../@id)"/>
											</xsl:attribute>
											<xsl:attribute name="value">
												<xsl:value-of select="."/>
											</xsl:attribute>
											<xsl:if test="../../@disable='yes' and concat(../../@hasValue,'')!='yes'">
												<xsl:attribute name="disabled">true</xsl:attribute>
											</xsl:if>
											<xsl:if test="concat(../@value,'')=''">
												<xsl:if test="./@default='yes'">
													<xsl:attribute name="checked">true</xsl:attribute>
												</xsl:if>
											</xsl:if>
											<xsl:if test="../@value=.">
												<xsl:attribute name="checked">true</xsl:attribute>
											</xsl:if>
											<xsl:if test="../@blur='yes'">
												<xsl:attribute name="onchange">calc_<xsl:value-of select="../../@name"/>(this);</xsl:attribute>
											</xsl:if>
											<xsl:if test="./@price!=''">
												<xsl:attribute name="price">
													<xsl:value-of select="./@price"/>
												</xsl:attribute>
											</xsl:if>
										</input>
										<xsl:if test="./@br='yes'">
											<br/>
										</xsl:if>
									</xsl:for-each>
								</xsl:when>
								<xsl:when test="./@type='option'">
									<select out='yes'>
										<xsl:attribute name="name">
											<xsl:value-of select="./@name"/>
										</xsl:attribute>
										<xsl:if test="../@disable='yes' and concat(../@hasValue,'')!='yes'">
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
												<xsl:if test="./@hasStock!=''">
													<xsl:attribute name="hasStock">
														<xsl:value-of select="./@hasStock"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="../@value=.">
													<xsl:attribute name="selected">true</xsl:attribute>
												</xsl:if>
												<xsl:value-of select="./@label"/>
											</option>
										</xsl:for-each>
									</select>
								</xsl:when>
								<xsl:when test="./@type='option-other'">
									<select out='yes'>
										<xsl:attribute name="name">
											<xsl:value-of select="./@name"/>
										</xsl:attribute>
										<xsl:attribute name="checkOther">
											<xsl:value-of select="./@checkOther"/>
										</xsl:attribute>
										<xsl:if test="../@disable='yes' and concat(../@hasValue,'')!='yes'">
											<xsl:attribute name="disabled">true</xsl:attribute>
										</xsl:if>
										<xsl:if test="../@blur='yes'">
											<xsl:attribute name="onchange">calc_<xsl:value-of select="../@name"/>(this);checkOther(this);</xsl:attribute>
										</xsl:if>
										<xsl:if test="../@blur!='yes'">
											<xsl:attribute name="onchange">checkOther(this);</xsl:attribute>
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
												<xsl:if test="../@value=.">
													<xsl:attribute name="selected">true</xsl:attribute>
												</xsl:if>
												<xsl:value-of select="./@label"/>
											</option>
										</xsl:for-each>
										<option value='-1'>
											<xsl:if test="concat(substring(./@value,1,6),'')='other:'">
												<xsl:attribute name="selected">true</xsl:attribute>
												<xsl:value-of select="substring(./@value,7)"/>
											</xsl:if>
											<xsl:if test="concat(substring(./@value,1,6),'')!='other:'">ساير</xsl:if>
										</option>
									</select>
									<input type='text' onblur='addOther(this);'>
										<xsl:attribute name="name"><xsl:value-of select="./@name"/>-addValue</xsl:attribute>
									</input>
								</xsl:when>
								<xsl:when test="./@type='option-fromTable'">
									<select class='fromTable' out='yes'>
										<xsl:attribute name="name">
											<xsl:value-of select="./@name"/>
										</xsl:attribute>
										<xsl:attribute name="thisValue">
											<xsl:value-of select="./@value"/>
										</xsl:attribute>
										<xsl:attribute name="table">
											<xsl:value-of select="./@table"/>
										</xsl:attribute>
										<xsl:if test="../@disable='yes' and concat(../@hasValue,'')!='yes'">
											<xsl:attribute name="disabled">true</xsl:attribute>
										</xsl:if>
										<xsl:if test="../@blur='yes'">
											<xsl:attribute name="onchange">calc_<xsl:value-of select="../@name"/>(this);</xsl:attribute>
										</xsl:if>
									</select>
								</xsl:when>
							</xsl:choose>
						</div>
					</div>
					<!--xsl:if test="./@br='yes'">
						<br/>
					</xsl:if-->
				</xsl:for-each>
			</div>
			</fieldset>
			</div>
		</xsl:for-each>
		<div class="unusedGroup"/>
		</div>
	</xsl:for-each>
	</div>
	<hr/>
</xsl:for-each>
</xsl:template>
</xsl:stylesheet>