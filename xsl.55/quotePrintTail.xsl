<?xml version="1.0" encoding="windows-1256"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/head">
	<p>شايان ذكر است مي‌باشد مبلغ فوق بدون احتساب ماليات بر ارزش افزوده مي‌باشد</p>
	<p>اعتبار استعلام فوق تا تاريخ <xsl:value-of select="./today/date"/> مي‌باشد</p>
	<p>زمان تحويل سفارش پس از تاييد نمونه و پرداخت <xsl:value-of select="./deposit"/> ريال به عنوان پيش پرداخت <xsl:value-of select="./productionDuration"/> روز كاري مي‌باشد.</p>
	<xsl:if test="./dueDate=0">
		<p>زمان سررسيد تسويه حداكثر در هنگام تحويل خواهد بود</p>
	</xsl:if>
	<xsl:if test="./dueDate>0">
		<p>زمان سررسيد تسويه حداكثر <xsl:value-of select="./dueDate"/> روز پس از تحويل خواهد بود</p>
	</xsl:if>
	<xsl:if test="./chequeDueDate=0">
		<p>زمان سررسيد چك حداكثر در روز تحويل خواهد بود</p>
	</xsl:if>
	<xsl:if test="./chequeDueDate>0">
		<p>مهلت سر رسيد چك <xsl:value-of select="./chequeDueDate"/> روز پس از تحويل خواهد بود</p>
	</xsl:if>	
	<p>مسئول پيگيري: <xsl:value-of select="./salesPerson"/> داخلي <xsl:value-of select="./extention"/></p>
	<div class="sign">
		<p>با تشكر<br/><xsl:value-of select="./approvedByName"/><br/>خانه چاپ و طرح</p>
	</div>
	<div class="tail">
		<p>تهران | خيابان آزادي | شماره 545 | تلفن 66042700 | فكس 66042704 | نشاني وب www.pdhco.com</p>
	</div>
</xsl:template>
</xsl:stylesheet>