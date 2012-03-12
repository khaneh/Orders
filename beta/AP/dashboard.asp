<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= "ÏÔÈæÑÏ ÎÑíÏ"
SubmenuItem=8
if not Auth(7 , "B") then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<link type="text/css" href="/css/jquery-ui-1.8.16.custom.css" rel="stylesheet" />

<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.8.16.custom.min.js"></script>
<script type="text/javascript">
	$(document).ready(function () {
		$("#unFactBT").click(function() {
			if ($("#unFact").is(":hidden")) {
				$("#unFact").slideDown("slow");
			} else {
				$("#unFact").hide();
			}
		});
		$("#unVerifBT").click(function() {
			if ($("#unVerif").is(":hidden")) {
				$("#unVerif").slideDown("slow");
			} else {
				$("#unVerif").hide();
			}
		});
		$("#unPaidBT").click(function() {
			if ($("#unPaid").is(":hidden")) {
				$("#unPaid").slideDown("slow");
			} else {
				$("#unPaid").hide();
			}
		});
	});
</script>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style >
	.indent {margin-right: 30px;display: none;}
	.BT {cursor: pointer;}
</style>
<BR>
<form method="post" action="">
	<input name="fromDate" type="text" size="10" onblur="acceptDate(this);" value="<%=shamsiDate(dateadd("d",-180,date()))%>">
	<input type="submit" value="ÊÇííÏ">
</form>
<%
if request("fromDate")="" then 
	fromDate=shamsiDate(dateadd("d",-180,date()))
else
	fromDate=request("fromDate")
end if
%>
<a href="../inventory/">ÏÑÎæÇÓÊåÇí ßÇáÇ ÇÒ ÇäÈÇÑ ßå åäæÒ ÍæÇáå ÈÑÇí ÂäåÇ ÕÇÏÑ äÔÏå: </a>
<%
mySQL="select count(id) as [count] from InventoryItemRequests where status='new' and reqDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write  rs("count")&"<BR>"
%>
<a href="../purchase/outServiceOrder.asp">ÏÑÎæÇÓÊåÇí ÓİÇÑÔ ÎÑíÏ ÌÏíÏ: </a>
<%
mySQL="select count(id) as [count] from PurchaseRequests where status='new' and reqDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write rs("count")&"<BR>"
%>
<a href="../purchase/outServiceTrace.asp?lstOrd=NEW">ÓİÇÑÔåÇí ÎÑíÏ ÌÏíÏ: </a>
<%
mySQL="select count(id) as [count] from PurchaseOrders where status='NEW' and ordDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write rs("count")&"<BR>"
%>
<a href="../purchase/outServiceTrace.asp?lstOrd=OUT">ÓİÇÑÔåÇí ÎÑíÏ ßå ÈíÑæä åÓÊäÏ: </a>
<%
mySQL="select count(id) as [count] from PurchaseOrders where status='OUT' and ordDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write rs("count")&"<BR>"
%>
<label id='unFactBT' title="ÈÑÇí ãÔÇåÏå ÌÒÆíÇÊ ßáíß ßäíÏ" class="BT">ßá ÎÑíÏåÇí İÇßÊæÑ äÔÏå: </label>
<%
mySQL="select count(id) as [count] from PurchaseOrders where HasVoucher=0 and Status<>'CANCEL' and ordDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write rs("count")&"<BR>"
%>
<div class="indent" id='unFact'>
<%
mySQL="select vendor_id,accounts.AccountTitle,count(PurchaseOrders.id) as [count] from PurchaseOrders inner join Accounts on PurchaseOrders.Vendor_ID=accounts.ID where HasVoucher=0 and PurchaseOrders.Status<>'CANCEL' and PurchaseOrders.ordDate>=N'" & fromDate & "' group by Vendor_ID,AccountTitle"
set rs=Conn.Execute(mySQL)
while not rs.eof 
	response.write "<a href='../AP/voucherInput.asp?act=show&selectedCustomer="&rs("vendor_id")&"'>ÎÑíÏåÇí "& rs("accountTitle")&": </a>"&rs("count")&"<BR>"
	rs.moveNext
wend
%>
</div>
<label id='unVerifBT' title="ÈÑÇí ãÔÇåÏå ÌÒÆíÇÊ ßáíß ßäíÏ" class="BT">İÇßÊæÑåÇí ÎÑíÏ ÊÇííÏ äÔÏå: </label>
<%
mySQL="select count(id) as [count] from Vouchers where Voided=0 and Verified=0 and effectiveDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write rs("count")&"<BR>"
%>
<div class="indent" id='unVerif'>
<%
mySQL="select Vouchers.id,Vouchers.Title,Vouchers.comment,Vouchers.TotalPrice,accounts.AccountTitle from Vouchers inner join Accounts on vouchers.VendorID=accounts.id where Voided=0 and Verified=0 and Vouchers.effectiveDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
while not rs.eof 
	response.write "<a title='" & rs("comment") & "' href='../AP/AccountReport.asp?act=showVoucher&voucher=" & rs("id") & "'>İÇßÊæÑ ÎÑíÏ " & rs("title") & " ãÑÈæØ Èå " & rs("accountTitle") & " Èå ãÈáÛ " & rs("totalPrice") & "</a><BR>"
	rs.moveNext
wend
%>
</div>
<label id='unPaidBT' title="ÈÑÇí ãÔÇåÏå ÌÒÆíÇÊ ßáíß ßäíÏ" class="BT">İÇßÊæÑåÇí ÎÑíÏ ÑÏÇÎÊ äÔÏå: </label>
<%
mySQL="select count(id) as [count] from Vouchers where Voided=0 and paid=0 and effectiveDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write rs("count")&"<BR>"
%>
<div class="indent" id='unPaid'>
<%
mySQL="select Vouchers.id,Vouchers.Title,Vouchers.comment,Vouchers.TotalPrice,accounts.AccountTitle,isnull(ap.amount,0) as amount from Vouchers inner join Accounts on vouchers.VendorID=accounts.id left outer join (select APItems.link,sum(ap2.AmountOriginal) as amount from APItems inner join APItemsRelations on apItems.id=APItemsRelations.CreditAPItem inner join ApItems as ap2 on APItemsRelations.DebitAPItem=ap2.id where APItems.voided=0 and apItems.type=6 group by APItems.Link) as ap on ap.link=vouchers.id where Vouchers.Voided=0 and Vouchers.paid=0 and isnull(ap.amount,0)<Vouchers.TotalPrice and Vouchers.effectiveDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
while not rs.eof 
	response.write "<a title='" & rs("comment") & "' href='../AP/AccountReport.asp?act=showVoucher&voucher=" & rs("id") & "'>İÇßÊæÑ ÎÑíÏ " & rs("title") & " ãÑÈæØ Èå " & rs("accountTitle") & " Èå ãÈáÛ " & Separate(rs("totalPrice")) & "</a>"
	if CDbl(rs("amount"))>0 then response.write " ãÈáÛ ÏæÎÊå ÔÏå: " & Separate(rs("amount")) & "¡ ãÇäÏå " & Separate(CDbl(rs("totalPrice"))-CDbl(rs("amount")))
	response.write "<BR>"
	rs.moveNext
wend
%>
</div>