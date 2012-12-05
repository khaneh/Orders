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
			</xsl:if>
			<xsl:if test="./isOrder='yes'">
				<xsl:attribute name="class">orderColor</xsl:attribute>
			</xsl:if>
			<td align="left"><font color="YELLOW">حساب:</font></td>
			<td align="right" height="25px">
				<font color="YELLOW"><xsl:value-of select="./customer/id"/> - <xsl:value-of select="./customer/accountTitle"/></font>
			</td>
			<td align="left">
				<font color="YELLOW">
					<xsl:if test="./isOrder!='yes'">استعلام گيرنده:</xsl:if>
					<xsl:if test="./isOrder='yes'">سفارش گيرنده:</xsl:if>
				</font>
			</td>
			<td>
				<input type="Text" readonly="readonly" name="SalesPerson" style="font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px">
					<xsl:attribute name="value"><xsl:value-of select="./salesPerson"/></xsl:attribute>
				</input>
			</td>
			<td align="left">
				<font color="yellow">
					<xsl:if test="./isOrder!='yes'">نوع استعلام:</xsl:if>
					<xsl:if test="./isOrder='yes'">نوع سفارش:</xsl:if>
				</font>
			</td>
			<td>
				<font color="red">
					<b><xsl:value-of select="./orderTypeName"/></b>
				</font>
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
				<font color="YELLOW">
					<xsl:if test="./isOrder!='yes'">شماره استعلام:</xsl:if>
					<xsl:if test="./isOrder='yes'">شماره سفارش:</xsl:if>
				</font>
			</td>
			<td align="right">
				<input disabled="disabled" type="text" name="orderID" maxlength="6" size="8" dir="LTR">
					<xsl:attribute name="value"><xsl:value-of select="./orderID"/></xsl:attribute>
				</input>
			</td>
			<td align="left"><font color="YELLOW">تاريخ:</font></td>
			<td>
				<span>
					<input disabled="disabled" type="text" maxlength="10" size="10" value="1391/04/22">
						<xsl:attribute name="value"><xsl:value-of select="./today/date"/></xsl:attribute>
					</input>
				</span>
				<span>
					<font color="YELLOW"><xsl:value-of select="./today/wName"/></font>
				</span>
			</td>
			<td align="left"><font color="YELLOW">ساعت:</font></td>
			<td align="right">
				<input disabled="disabled" type="text" maxlength="5" size="3" dir="LTR">
					<xsl:attribute name="value"><xsl:value-of select="./today/time"/></xsl:attribute>
				</input>
			</td>
		</tr>
			<tr class="grayColor">
				<xsl:if test="./isOrder!='yes'">
					<td align="left">زمان توليد:</td>
					<td>
						<input type="text" name="productionDuration" size="5" class="forceErr" onblur="checkForceHead(this);">
							<xsl:attribute name="value"><xsl:value-of select="./productionDuration"/></xsl:attribute>
						</input>
						<span>روز(عددصحيح)</span>
					</td>
				</xsl:if>
				<xsl:if test="./isOrder='yes'">
					<td align="right" colspan="2">
						<span style="margin-right: 18px;">تاريخ تحويل، فعلا معلوم نيست!</span>
						<input name="returnDateNull" type="checkbox">
							<xsl:if test="./today/retIsNull='yes'">
								<xsl:attribute name="checked">true</xsl:attribute>
							</xsl:if>
						</input>
					</td>
				</xsl:if>
				<td align="left">
					<xsl:if test="./isOrder!='yes'">موعد اعتبار:</xsl:if>
					<xsl:if test="./isOrder='yes'">تاريخ تحويل:</xsl:if>
				</td>
				<td>
					<input type="text" name="ReturnDate" onblur="acceptDate(this);checkForceHead(this);" maxlength="10" size="10" dir="LTR">
						<xsl:attribute name="value"><xsl:value-of select="./today/retDate"/></xsl:attribute>
					</input>
				</td>
				<td align="left">
					<xsl:if test="./isOrder!='yes'">ساعت اعتبار:</xsl:if>
					<xsl:if test="./isOrder='yes'">ساعت تحويل:</xsl:if>
				</td>
				<td align="right">
					<input type="text" name="ReturnTime" maxlength="5" size="3" dir="LTR" onblur="acceptTime(this);">
					<xsl:if test="concat(./today/retIsNull,'')!='yes'">
						<xsl:if test="concat(./today/retTime,'')!=''">
							<xsl:attribute name="value"><xsl:value-of select="./today/retTime"/></xsl:attribute>
						</xsl:if>
						<xsl:if test="concat(./today/retTime,'')=''">
							<xsl:attribute name="value">16:30</xsl:attribute>
						</xsl:if>
					</xsl:if>
					</input>
				</td>
			</tr>
			
			<tr class="grayColor">
				<td align="left">سايز:</td>
				<td>
					<input type="text" name="paperSize" size="5" onblur="acceptPaper(this);checkForceHead(this);">
						<xsl:attribute name="value"><xsl:value-of select="./paperSize"/></xsl:attribute>
					</input>
					<span>مانند: 21X30</span>
				</td>
				<td align="left">عنوان كار داخل فايل:</td>
				<td align="right" colspan="3">
					<input type="text" name="OrderTitle" maxlength="255" size="50" style="width:100%" class="forceErr" onblur="checkForceHead(this);"> 
						<xsl:attribute name="value"><xsl:value-of select="./orderTitle"/></xsl:attribute>
					</input>
				</td>
			</tr>
			<tr class="grayColor">
				<td align="left">تيراژ:</td>
				<td>
					<input type="text" name="qtty" size="5" onblur="checkForceHead(this);">
						<xsl:attribute name="value"><xsl:value-of select="./qtty"/></xsl:attribute>
					</input>
					<span>يك عدد صحيح</span>
				</td>
				<td align="left">توضيحات بيشتر:</td>
				<td align="right" colspan="3" rowspan="3">
					<textarea name="Notes" row="6" class="orderNotes"><xsl:value-of select="./notes"/></textarea>
				</td>
			</tr>
			<tr class="grayColor">
				<td align="left">قيمت كل:</td>
				<td colspan="2">
					<input type="text" name="totalPrice" id="totalPrice" readonly="readonly">
						<xsl:attribute name="value"><xsl:value-of select="./totalPrice"/></xsl:attribute>
					</input>
				</td>
			</tr>
			<tr class="grayColor">
				<td align="left">قيمت كل با ماليات:</td>
				<td colspan="2">
					<input type="text" name="totalVatedPrice" id="totalVatedPrice" readonly="readonly"/>
					<span id="sumDis" class="dis"></span>
					<span id="sumOver" class="over"></span>
					<span id="sumReverse" class="reverse"></span>
				</td>
			</tr>
			<tr class="grayColor">
				<td align="left">پيش پرداخت:</td>
				<td align="right">
					<input name="deposit" size="10" onblur="checkForceHead(this);">
						<xsl:attribute name="value">
							<xsl:value-of select="./deposit"/>
						</xsl:attribute>
					</input>
					<span> ريال</span>
				</td>
				<td align="left">سررسيد:</td>
				<td align="right">
					<input name="dueDate" size="2" onblur="checkForceHead(this);">
						<xsl:attribute name="value">
							<xsl:value-of select="./dueDate"/>
						</xsl:attribute>
						<xsl:if test="concat(./dueDate/@readOnly,'')='yes'">
							<xsl:attribute name="readonly">readonly</xsl:attribute>	
						</xsl:if>
					</input>
					<span> روز</span>
				</td>
				<td align="left">مهلت دريافت چك:</td>
				<td align="right">
					<input name="chequeDueDate" size="2" onblur="checkForceHead(this);">
						<xsl:attribute name="value">
							<xsl:value-of select="./chequeDueDate"/>
						</xsl:attribute>
						<xsl:if test="concat(./chequeDueDate/@readOnly,'')='yes'">
							<xsl:attribute name="readonly">readonly</xsl:attribute>	
						</xsl:if>
					</input>
					<span> روز</span>
				</td>
			</tr>
	</table>
	<div id="errMsg"></div>
	<input type="hidden" id="outXML" name="myXML"/>
</xsl:template>
</xsl:stylesheet>