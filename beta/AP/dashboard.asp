<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= "œ‘»Ê—œ Œ—Ìœ"
SubmenuItem=8
if not Auth(7 , "B") then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->

<script type="text/javascript">
	$(document).ready(function () {
		$.ajaxSetup({
			cache: false
		});
		
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
	function go(url){
		window.open(url,'ap-dashboard');
	}
	function show(myAct){
		var loadUrl="dashboard_ajax.asp";
		$.getJSON(loadUrl,
			{act:myAct},
			function(json){
				out="";
				$("#list").html("<div></div>");
				$.each(json,function(i,msg){
					out+="<div class='row "
					if (msg.link!='') 
						out+="pointer' onclick=\"go('"+ msg.link +"');\"";
					else
						out+="'"	
					if (msg.title!='')
						out+=" title=\"" + msg.title + "\"";
					out+=">"+msg.div+"</div>";
				});
				$("#list").children("div:first").after(out);
			}
		);
	}
</script>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style >
	.indent {margin-right: 30px;display: none;}
	.row{margin-left: 20px;padding-top: 5px;padding-bottom: 5px;}
	.rspan1{width: 70px;}
	.rspan3{width: 260px;}
	.pointer {cursor: pointer;}
	#list div:nth-child(even) {background-color: #FFF;}
	#list div:nth-child(odd) {background-color: #DDD;}
	div.step{float: right;width: 120px;text-align: center;cursor: pointer;}
	div.inStep{}
	div.folder{width: 35px;height: 35px;background-image: url('/beta/images/folder1.gif');margin: 0 42px 0 0;}
	div.folder-close{background-image: url('/beta/images/folder0.gif');}
	div.float-close{clear: both;float: none;}
	#list{margin: 20px 10px 0 0;}
</style>
<BR>
<form method="post" action="">
	<!--input name="fromDate" type="text" size="10" onblur="acceptDate(this);" value="<%=shamsiDate(dateadd("d",-180,date()))%>">
	<input type="submit" value=" «ÌÌœ"-->
</form>
<%
if request("fromDate")="" then 
	fromDate=shamsiDate(dateadd("d",-180,date()))
else
	fromDate=request("fromDate")
end if
function getCount(mySQL)
	out=0
	set rs=Conn.Execute(mySQL)
	if not rs.eof then out=rs("count")
	rs.close
	set rs=nothing
	getCount = out
end function
startDate="1389/01/01"
c1=getCount("select count(id) as [count] from InventoryItems inner join (select ItemID,sum(qtty) as sumQtty from InventoryPickuplistItems inner join InventoryPickuplists on InventoryPickuplistItems.pickupListID=InventoryPickuplists.id and InventoryPickuplists.Status='new' group by ItemID) as pick on InventoryItems.ID=pick.itemID where enabled=1 and owner=-1 and qtty-sumQtty<=minim and qtty>minim")
c2=getCount("select count(id) as [count] from InventoryItems where enabled=1 and owner=-1 and qtty<minim")
c3=getCount("select count(id) as [count] from InventoryItemRequests where status='new'")
c4=getCount("SELECT count(id) as [count] FROM purchaseRequests WHERE (Status = 'new' AND TypeID <> 0 AND IsService=1)")
c5=getCount("SELECT count(id) as [count] FROM purchaseRequests WHERE (Status = 'new' AND IsService=0)")
c60=getCount("select count(id) as [count] from PurchaseOrders where status='NEW'") 
c61=getCount("select count(id) as [count] from PurchaseOrders where status='OUT'") 
c62=getCount("select count(id) as [count] from PurchaseOrders where ordDate>'"&startDate&"' and IsService=0 and status<>'CANCEL' and id not in (select distinct relatedID from InventoryLog where RelatedID>0 and isInput=1 and owner=-1 and type=1)")
c7=getCount("select count(id) as [count] from PurchaseOrders where ordDate>'"&startDate&"' and IsService=0 and status<>'CANCEL' and id not in (select VoucherLines.RelatedPurchaseOrderID from Vouchers inner join VoucherLines on VoucherLines.Voucher_ID=Vouchers.id where Vouchers.Voided=0)")
c8=getCount("select count(id) as [count] from PurchaseOrders where ordDate>'"&startDate&"' and IsService=1 and status<>'CANCEL' and id not in (select VoucherLines.RelatedPurchaseOrderID from Vouchers inner join VoucherLines on VoucherLines.Voucher_ID=Vouchers.id where Vouchers.Voided=0)")
c9=getCount("select count(VoucherLines.RelatedPurchaseOrderID) as [count] from Vouchers inner join VoucherLines on VoucherLines.Voucher_ID=Vouchers.id where Vouchers.Voided=0 and Vouchers.paid=0 and Vouchers.Verified=0")
c10=getCount("select count(Vouchers.id) as [count] from Vouchers inner join APItems on Vouchers.id=apItems.Link and apItems.Type=6 where APItems.Voided=0 and Vouchers.Voided=0 and Vouchers.paid=0 and APItems.FullyApplied=0 and apItems.effectiveDate>'" & startDate & "'")
%>

<div title="·Ì”  ﬂ«·«Ì „Ê—œ ‰Ì«“ ﬂÂ œ—ŒÊ«”  ﬂ«·« ‰œ«—œ" class="step">
	<div onclick="show('ifPickLow');">	
		<div class="folder <%if c1=0 then response.write "folder-close"%>"></div>
		<div><%=c1%></div>
		<span>œ— ’Ê—  «‰Ã«„ ”›«—‘ “Ì— Õœ«ﬁ· ŒÊ«Âœ ‘œ</span>		
	</div>
	<div onclick="show('invLow');">
		<div class="folder <%if c2=0 then response.write "folder-close"%>"></div>
		<div><%=c2%></div>
		<span>“Ì— Õœ«ﬁ· „ÊÃÊœÌ</span>
	</div>
</div>
<div class="step">
	<div title="œ—ŒÊ«” ùÂ«ÌÌ ﬂÂ Â‰Ê“ ”›«—‘ ‰œ«—‰œ" onclick="go('../inventory/');">	
		<div class="folder <%if c3=0 then response.write "folder-close"%>"></div>
		<div><%=c3%></div>
		<span>œ—ŒÊ«”  ﬂ«·«Ì «‰»«—</span>
	</div>
	<div title="œ—ŒÊ«” ùÂ«ÌÌ ﬂÂ Â‰Ê“ ”›«—‘ ‰œ«—‰œ" onclick="go('../purchase/outServiceOrder.asp');">
		<div class="folder <%if c4=0 then response.write "folder-close"%>"></div>
		<div><%=c4%></div>	
		<span>œ— ŒÊ«”  Œ—Ìœ Œœ„« </span>
	</div>
	<div title="œ—ŒÊ«” ùÂ«ÌÌ ﬂÂ Â‰Ê“ ”›«—‘ ‰œ«—‰œ" onclick="go('../purchase/outServiceOrder.asp');">
		<div class="folder <%if c5=0 then response.write "folder-close"%>"></div>
		<div><%=c5%></div>
		<span>œ— ŒÊ«”  Œ—Ìœ ﬂ«·«Ì «‰»«—</span>
	</div>
</div>
<div class="step">
	<div title="”›«—‘ùÂ«ÌÌ ﬂÂ Ê÷⁄Ì  ¬‰Â« ÃœÌœ «” " onclick="go('../purchase/outServiceTrace.asp?lstOrd=NEW');">
		<div class="folder <%if c60=0 then response.write "folder-close"%>"></div>
		<div><%=c60%></div>
		<span>”›«—‘ Œ—Ìœ ÃœÌœ</span>
	</div>
	<div title="”›«—‘ùÂ«ÌÌ ﬂÂ œ— Œ«—Ã «“ ‘—ﬂ  Â” ‰œ" onclick="go('../purchase/outServiceTrace.asp?lstOrd=OUT');">
		<div class="folder <%if c61=0 then response.write "folder-close"%>"></div>
		<div><%=c61%></div>
		<span>”›«—‘ Œ—Ìœ Œ«—Ã «“ ‘—ﬂ </span>
	</div>
	<div title="”›«—‘ùÂ«ÌÌ ﬂÂ Ê—Êœ »Â «‰»«— ‰‘œÂù«‰œ(«“ «» œ«Ì ”«· 89)" onclick="show('noInvRelation');">
		<div class="folder <%if c62=0 then response.write "folder-close"%>"></div>
		<div><%=c62%></div>
		<span>”›«—‘ »Â «‰»«— Ê«—œ ‰‘œÂ</span>
	</div>
</div>
<div class="step">
	<div title="ﬂÂ Â‰Ê“ ›«ﬂ Ê—  «„Ì‰ ﬂ‰‰œÂ ‰Ì«„œÂ(«“ «» œ«Ì ”«· 89)" onclick="show('invNOvo');">
		<div class="folder <%if c7=0 then response.write "folder-close"%>"></div>
		<div><%=c7%></div>
		<span>Ê—Êœ »Â «‰»«—</span>
	</div>
	<div title="ﬂÂ Â‰Ê“ ›«ﬂ Ê—  «„Ì‰ ﬂ‰‰œÂ ‰Ì«„œÂ(«“ «» œ«Ì ”«· 89)" onclick="show('serviceNoVo');">
		<div class="folder <%if c8=0 then response.write "folder-close"%>"></div>
		<div><%=c8%></div>
		<span> «ÌÌœ «‰Ã«„ Œœ„« </span>
	</div>
</div>
<div title="ﬂÂ Â‰Ê“  «ÌÌœ ‰‘œÂ" class="step" onclick="go('../AP/verify.asp');">
	<div class="folder <%if c9=0 then response.write "folder-close"%>"></div>
	<div><%=c9%></div>
	<span>›«ﬂ Ê—  «„Ì‰ ﬂ‰‰œÂ</span>
</div>
<div title="›«ﬂ Ê—Â«ÌÌ ﬂÂ  «ÌÌœ ‘œÂ Ê Â‰Ê“  ”ÊÌÂ ‰‘œÂ(«“ «» œ«Ì ”«· 89)" class="step" onclick="show('unPaid');">
	<div class="folder <%if c10=0 then response.write "folder-close"%>"></div>
	<div><%=c10%></div>
	<span>Å—œ«Œ  çﬂ</span>
</div>
<div class="float-close"></div>
<div id="list"><div></div></div>


<%
response.end
mySQL="select count(id) as [count] from PurchaseOrders where status='NEW' and ordDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write rs("count")&"<BR>"
%>
<a href="../purchase/outServiceTrace.asp?lstOrd=OUT">”›«—‘ùÂ«Ì Œ—Ìœ ﬂÂ »Ì—Ê‰ Â” ‰œ: </a>
<%
mySQL="select count(id) as [count] from PurchaseOrders where status='OUT' and ordDate>=N'" & fromDate & "'"
set rs=Conn.Execute(mySQL)
if not rs.eof then response.write rs("count")&"<BR>"
%>
<label id='unFactBT' title="»—«Ì „‘«ÂœÂ Ã“∆Ì«  ﬂ·Ìﬂ ﬂ‰Ìœ" class="BT">ﬂ· Œ—ÌœÂ«Ì ›«ﬂ Ê— ‰‘œÂ: </label>

<div class="indent" id='unFact'>
<%
mySQL="select vendor_id,accounts.AccountTitle,count(PurchaseOrders.id) as [count] from PurchaseOrders inner join Accounts on PurchaseOrders.Vendor_ID=accounts.ID where HasVoucher=0 and PurchaseOrders.Status<>'CANCEL' and PurchaseOrders.ordDate>=N'" & fromDate & "' group by Vendor_ID,AccountTitle"
set rs=Conn.Execute(mySQL)
while not rs.eof 
	response.write "<a href='../AP/voucherInput.asp?act=show&selectedCustomer="&rs("vendor_id")&"'>Œ—ÌœÂ«Ì "& rs("accountTitle")&": </a>"&rs("count")&"<BR>"
	rs.moveNext
wend
%>
</div>
<label id='unVerifBT' title="»—«Ì „‘«ÂœÂ Ã“∆Ì«  ﬂ·Ìﬂ ﬂ‰Ìœ" class="BT">›«ﬂ Ê—Â«Ì Œ—Ìœ  «ÌÌœ ‰‘œÂ: </label>
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
	response.write "<a title='" & rs("comment") & "' href='../AP/AccountReport.asp?act=showVoucher&voucher=" & rs("id") & "'>›«ﬂ Ê— Œ—Ìœ " & rs("title") & " „—»Êÿ »Â " & rs("accountTitle") & " »Â „»·€ " & rs("totalPrice") & "</a><BR>"
	rs.moveNext
wend
%>
</div>
<label id='unPaidBT' title="»—«Ì „‘«ÂœÂ Ã“∆Ì«  ﬂ·Ìﬂ ﬂ‰Ìœ" class="BT">›«ﬂ Ê—Â«Ì Œ—Ìœ Å—œ«Œ  ‰‘œÂ: </label>
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
	response.write "<a title='" & rs("comment") & "' href='../AP/AccountReport.asp?act=showVoucher&voucher=" & rs("id") & "'>›«ﬂ Ê— Œ—Ìœ " & rs("title") & " „—»Êÿ »Â " & rs("accountTitle") & " »Â „»·€ " & Separate(rs("totalPrice")) & "</a>"
	if CDbl(rs("amount"))>0 then response.write " „»·€ œÊŒ Â ‘œÂ: " & Separate(rs("amount")) & "° „«‰œÂ " & Separate(CDbl(rs("totalPrice"))-CDbl(rs("amount")))
	response.write "<BR>"
	rs.moveNext
wend
%>
</div>