<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
Response.Buffer=True
%>
<!--#include file="../config.asp" -->
<%
sys=request.queryString("sys")
Item=request.queryString("Item")
if sys<>"" and IsNumeric(Item) then
	Set rs=Conn.Execute("Select Link, Type FROM "& sys & "Items WHERE (ID='"& Item & "')")
	if rs.eof then
		response.redirect "../AO/AccountReport.asp?errmsg=" & Server.URLEncode("äíä ÂíÊãí æÌæÏ äÏÇÑÏ.")
	else
		ItemLink=rs("Link")
		ItemType=rs("Type")
	end if
	rs.close
	set rs= Nothing
	select case ItemType
	case 1
		Action="showInvoice&invoice="
	case 2
		Action="showReceipt&receipt="
	case 3
		Action="showMemo&sys="& sys & "&memo="
	case 4
		Action="showInvoice&invoice="
	case 5
		Action="showPayment&payment="
	case 6
		Action="showVoucher&voucher="
	case else
		response.redirect "../" & sys & "/AccountReport.asp?errmsg=" & Server.URLEncode("ãÔÎÕÇÊ æÇÑÏ ÔÏå ÛáØ ÇÓÊ.")
	end select
	response.redirect "../" & sys & "/AccountReport.asp?act=" & Action & ItemLink
else
	response.redirect "../" & sys & "/AccountReport.asp?errmsg=" & Server.URLEncode("ãÔÎÕÇÊ æÇÑÏ ÔÏå ÛáØ ÇÓÊ.")
end if
%>