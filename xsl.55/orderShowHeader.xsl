<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/head">
	<input type="hidden" name="customerID">
		<xsl:attribute name="value"><xsl:value-of select="./customer/id"/></xsl:attribute>
	</input>
	<input type="hidden" name="orderType">
		<xsl:attribute name="value"><xsl:value-of select="./orderTypeID"/></xsl:attribute>
	</input>
	<input type="hidden" name="returnDateTime" id="returnDateTime"/>
	<table cellspacing="0" cellpadding="2" dir="RTL" width="700px" align="center">
		<tr>
			<xsl:if test="./isOrder!='yes'">
				<xsl:attribute name="class">quoteColor</xsl:attribute>
<!-- 				<xsl:attribute name="bgcolor">#555599</xsl:attribute> -->
			</xsl:if>
			<xsl:if test="./isOrder='yes'">
				<xsl:attribute name="class">orderColor</xsl:attribute>
<!-- 				<xsl:attribute name="bgcolor">black</xsl:attribute> -->
			</xsl:if>
			<td align="left">حساب:</td>
			<td align="right" height="25px">
				<a id="customerID" href="#">
					<xsl:attribute name="myID"><xsl:value-of select="./customer/id"/></xsl:attribute>
					<xsl:value-of select="./customer/id"/> - <xsl:value-of select="./customer/accountTitle"/>
				</a>
			</td>
			<td align="left">
				<xsl:if test="./isOrder!='yes'">استعلام گيرنده:</xsl:if>
				<xsl:if test="./isOrder='yes'">سفارش گيرنده:</xsl:if>
			</td>
			<td><xsl:value-of select="./salesPerson"/></td>
			<td align="left">
				<xsl:if test="./isOrder!='yes'">نوع استعلام:</xsl:if>
				<xsl:if test="./isOrder='yes'">نوع سفارش:</xsl:if>
			</td>
			<td class="isRed">
				<b><xsl:value-of select="./orderTypeName"/></b>
			</td>
		</tr>
		<tr>
			<xsl:if test="./isOrder!='yes'">
				<xsl:attribute name="class">quoteColor</xsl:attribute>
			</xsl:if>
			<xsl:if test="./isOrder='yes'">
				<xsl:attribute name="class">orderColor</xsl:attribute>
			</xsl:if>
			<td align="left">
				<xsl:if test="./isOrder!='yes'">شماره استعلام:</xsl:if>
				<xsl:if test="./isOrder='yes'">شماره سفارش:</xsl:if>
			</td>
			<td align="right"><xsl:value-of select="./orderID"/></td>
			<td align="left">تاريخ:</td>
			<td>
				<span><xsl:value-of select="./today/date"/></span>
				<span><xsl:value-of select="./today/wName"/></span>
			</td>
			<td align="left">ساعت:</td>
			<td align="right"><xsl:value-of select="./today/time"/></td>
		</tr>
			<tr class="grayColor">
				<xsl:if test="./isOrder!='yes'">
					<td align="left">زمان توليد:</td>
					<td>
						<h6 class="isBlack">
							<span><xsl:value-of select="./productionDuration"/></span>
							<span>روز</span>
						</h6>
					</td>
					<td align="left">موعد اعتبار:</td>
					<td>
						<h6 class="isBlack">
							<xsl:value-of select="./today/retDate"/>
						</h6>
					</td>
					<td align="left">ساعت اعتبار:</td>
					<td align="right">
						<h6 class="isBlack">
							<xsl:value-of select="./today/retTime"/>
						</h6>
					</td>
				</xsl:if>
				<xsl:if test="./isOrder='yes'">
					<xsl:if test="./today/contractIsNull='yes'">
						<td align="right" colspan="6">
							<h6 class="isBlack">
								<center>تاريخ تحويل قرارداد، فعلا معلوم نيست!</center>
							</h6>
						</td>
					</xsl:if>
					<xsl:if test="./today/contractIsNull='no'">
						<td align="left">تاريخ تحويل قرارداد:</td>
						<td align="right">
							<h4>
								<xsl:attribute name="class">
									<xsl:value-of select="./today/contractClass"/>
								</xsl:attribute>
								<xsl:value-of select="./today/contractDate"/> &#160;<xsl:value-of select="./today/contractTime"/>
							</h4>
						</td>
						<td align="left">تاريخ تحويل عملي:</td>
						<td align="right" colspan="3">
							<h4>
								<xsl:attribute name="class">
									<xsl:value-of select="./today/retClass"/>
								</xsl:attribute>
								<xsl:value-of select="./today/retDate"/> &#160;<xsl:value-of select="./today/retTime"/>
							</h4>
						</td>
					</xsl:if>
						
				</xsl:if>
				
			</tr>
			<tr class="grayColor">
				<td align="left">سايز:</td>
				<td>
					<h6 class="isBlack">
						<xsl:value-of select="./paperSize"/>
					</h6>
				</td>
				<td align="left">عنوان كار داخل فايل:</td>
				<td align="right" colspan="3">
					<h6 class="isBlack">
						<xsl:value-of select="./orderTitle"/>
					</h6>
				</td>
			</tr>
			<tr class="grayColor">
				<td align="left">نام شركت:</td>
				<td align="right"><h6 class="isBlack"><xsl:value-of select="./customer/companyName"/></h6></td>
				<td align="left">نام مشتري:</td>
				<td align="right"><h6 class="isBlack"><xsl:value-of select="./customer/customerName"/></h6></td>
				<td align="left">تلفن:</td>
				<td align="right"><h6 class="isBlack"><xsl:value-of select="./customer/tel"/></h6></td>
			</tr>
			<tr class="grayColor">
				<td align="left">تيراژ:</td>
				<td>
					<h6 class="isBlack">
						<xsl:value-of select="./qtty"/>
					</h6>
				</td>
				<td align="left">توضيحات بيشتر:</td>
				<td align="right" colspan="3" rowspan="3">
					<h6 class="isBlack">
						<xsl:value-of select="./notes"/>
					</h6>
				</td>
			</tr>
			<tr class="grayColor">
				<td align="left">قيمت كل:</td>
				<td colspan="2" class="isPrice">
					<span class="isBlack price" id="sumTotal" rel="tooltip" data-placement="top" data-original-title="بهاي كل سفارش"><xsl:value-of select="./totalPrice"/></span>
					<span class="isBlack dis" id="sumDis" rel="tooltip" data-placement="top" data-original-title="بهاي كل تخفيف"></span>
					<span class="isBlack over" id="sumOver" rel="tooltip" data-placement="top" data-original-title="بهاي كل مازاد"></span>
					<span class="isBlack reverse" id="sumReverse" rel="tooltip" data-placement="top" data-original-title="بهاي كل برگشت"></span>
				</td>
			</tr>
			<tr class="grayColor">
				<td colspan="2" align="right">
					<span>وضعيت:</span>
					<span class="stName">
						<xsl:value-of select="./status/statusName"/>
					</span>
					<span>مرحله:</span>
					<span class="stName">
						<xsl:value-of select="./status/stepName"/>
					</span>
				</td>
				<td>
					<span align="center">نسخه </span>
					<span class="badge badge-important">
						<xsl:value-of select="./ver"/>
					</span>
				</td>
			</tr>
			<tr class="grayColor">
				<td align="left">پيش پرداخت:</td>
				<td align="right"><xsl:value-of select="./deposit"/> ريال</td>
				<td align="left">سررسيد:</td>
				<td align="right"><xsl:value-of select="./dueDate"/> روز</td>
				<td align="left">مهلت دريافت چك:</td>
				<td align="right"><xsl:value-of select="./chequeDueDate"/> روز</td>
			</tr>
	</table>
</xsl:template>
</xsl:stylesheet>