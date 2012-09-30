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
			<td align="left">استعلام گيرنده:</td>
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
				</xsl:if>
				<xsl:if test="./isOrder='yes'">
					<td align="right" colspan="2">
						<xsl:if test="concat(./today/retDate,'')=''">
							<h6 class="isBlack">
								<span style="margin-right: 18px;">تاريخ تحويل قرارداد، فعلا معلوم نيست!</span>
								<!--input name="returnDateNull" type="checkbox"/-->
							</h6>
						</xsl:if>
					</td>
				</xsl:if>
				<td align="left">
					<xsl:if test="./isOrder!='yes'">موعد اعتبار:</xsl:if>
					<xsl:if test="./isOrder='yes'">تاريخ تحويل قرارداد:</xsl:if>
				</td>
				<td>
					<h6 class="isBlack">
						<xsl:value-of select="./today/retDate"/>
					</h6>
				</td>
				<td align="left">
					<xsl:if test="./isOrder!='yes'">ساعت اعتبار:</xsl:if>
					<xsl:if test="./isOrder='yes'">ساعت تحويل:</xsl:if>
				</td>
				<td align="right">
					<h6 class="isBlack">
						<xsl:value-of select="./today/retTime"/>
					</h6>
				</td>
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
				<td align="right" colspan="3" rowspan="2">
					<h6 class="isBlack">
						<xsl:value-of select="./notes"/>
					</h6>
				</td>
			</tr>
			<tr class="grayColor">
				<td align="left">قيمت كل:</td>
				<td colspan="2">
					<h6 class="isBlack">
						<xsl:value-of select="./totalPrice"/>
					</h6>
				</td>
			</tr>
	</table>
</xsl:template>
</xsl:stylesheet>