<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Reports
PageTitle= "����ԝ��� �����"
SubmenuItem=1
if not Auth("E" , 6) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<BR><BR>
<BLOCKQUOTE>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="inv_cost.asp">����� ����� �����</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="inv_details.asp">����� ���� �� ����� (�� ������)</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="orders_latency.asp">����� ����� ����� ��</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="sales_items_by_customer.asp">����� ���� ��� ���� �� ��� �������</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="Salse_by_customer_category.asp">����� ����� ���� �� ��� �������</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="Salse_by_customer_category.old.iqy">����� ����� ���� �� ��� ������� 2</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="invoices_isA_by_customer.asp">����� ��� ���� � �ѐ�� ��� �� ��� �����</A><BR><BR>
<IMG SRC="../images/excel_icon.gif" BORDER="0" ALT="Excel Report"> <A HREF="sales_by_customer.iqy">����� ��� ���� � ����� ����� � ���ӝ���  �������</A><BR><BR>
</BLOCKQUOTE>

<!--#include file="tah.asp" -->
