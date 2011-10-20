<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Reports
PageTitle= "าวัิๅวํ ว฿ำแํ"
SubmenuItem=1
if not Auth("E" , 6) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<BR><BR>
<BLOCKQUOTE>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="inv_cost.asp">าวัิ ัํวแํ วไศวั</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="inv_details.asp">าวัิ ฮัๆฬ วา วไศวั (ศว ฬาํํวส)</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="orders_latency.asp">าวัิ สวฮํั ำÝวัิ ๅว</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="sales_items_by_customer.asp">าวัิ ยํสใ ๅวํ Ýัๆิ ศั อำศ ใิสัํวไ</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="Salse_by_customer_category.asp">าวัิ วไๆวฺ Ýัๆิ ศั อำศ ใิสัํวไ</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="Salse_by_customer_category.old.iqy">าวัิ วไๆวฺ Ýัๆิ ศั อำศ ใิสัํวไ 2</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="invoices_isA_by_customer.asp">าวัิ ฬใฺ Ýัๆิ ๆ ศัิส วแÝ ศั อำศ ใิสัํ</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="sales_by_customer.iqy">าวัิ ฬใฺ Ýัๆิ ๆ ใำฦๆแ ํํัํ ๆ ยฯัำๅวํ  ใิสัํวไ</A><BR><BR>
</BLOCKQUOTE>

<!--#include file="tah.asp" -->
