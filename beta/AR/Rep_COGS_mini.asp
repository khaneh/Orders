<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<%
response.buffer= true
Response.CharSet = "windows-1256" 

if not Auth("C" , 7) then NotAllowdToViewThisPage()

if request("id")<>"" then
	Invoice = cdbl(request("id"))
%>
<center>
<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#888855" style="border:dashed 2px #888888;font-family:tahoma;font-size:8pt;">
<TR bgcolor="#EEFFCC">
	<TD width="65">”›«—‘ Œ—Ìœ</TD>
	<TD width="60"> «—ÌŒ</TD>
	<TD width="155">‰Ê⁄ Œœ„« </TD>
	<TD width="250">‰«„ ‘—ﬂ </TD>
	<TD width="70">ﬁÌ„ </TD>
</TR>
<%	
mySQL="SELECT InvoiceOrderRelations.Invoice, PurchaseOrders.ID AS PurchaseOrderID, PurchaseOrders.TypeName, PurchaseOrders.OrdDate, PurchaseOrders.Vendor_ID, Accounts.AccountTitle, VoucherLines.price, Vouchers.Id AS VoucherID FROM InvoiceOrderRelations INNER JOIN PurchaseOrders INNER JOIN PurchaseRequestOrderRelations INNER JOIN PurchaseRequests ON PurchaseRequestOrderRelations.Req_ID = PurchaseRequests.ID ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID ON InvoiceOrderRelations.[Order] = PurchaseRequests.Order_ID LEFT OUTER JOIN Accounts ON PurchaseOrders.Vendor_ID = Accounts.ID LEFT OUTER JOIN Vouchers INNER JOIN VoucherLines ON Vouchers.id = VoucherLines.Voucher_ID ON Vouchers.Voided = 0 AND PurchaseOrders.ID = VoucherLines.RelatedPurchaseOrderID WHERE (PurchaseRequests.Status <> N'del') AND (PurchaseOrders.Status <> N'CANCEL') AND (InvoiceOrderRelations.Invoice = " & Invoice & ")"
set RS1=Conn.Execute (mySQL)
	if not RS1.eof then
		while not RS1.eof
			tmpCounter=0
			if isnull(RS1("VoucherId")) then
				priceText = "<FONT COLOR=""red"">›«ﬂ Ê— ‰‘œÂ</FONT>"
			else
				priceText = "<a target=""_blank"" title=""›«ﬂ Ê— Œ—Ìœ „—»ÊÿÂ"" href=""../AP/AccountReport.asp?act=showVoucher&voucher=" & RS1("VoucherID") & """>" & Separate(RS1("price")) & "</a>"
			end if
%>
		<TR bgcolor="#FFFFFF">
			<TD><a target="_blank" title="œÌœ‰ ”›«—‘ Œ—Ìœ „—»ÊÿÂ" href="../purchase/outServiceTrace.asp?od=<%=RS1("PurchaseOrderID")%>"><%=RS1("PurchaseOrderID")%></a></TD>
			<TD DIR="LTR"><%=RS1("OrdDate")%></TD>
			<TD><%=RS1("TypeName")%></TD>
			<TD><a target="_blank" title="„‘Œ’«  Õ”«»" href="../CRM/AccountInfo.asp?act=show&tab=3&selectedCustomer=<%=RS1("Vendor_ID")%>"><%=RS1("AccountTitle")%></a></TD>
			<TD><%=priceText%>&nbsp;</TD>
		</TR>

<%
		RS1.moveNext
		wend
	else
%>		<TR>
			<TD colspan="10" align="center" style="font-size:9pt;color:#888888">«ÿ·«⁄«  Œ—Ìœ »—«Ì «Ì‰ ›«ﬂ Ê— ÃÊœ ‰œ«—œ</TD>
		</TR>
<%
	end if 
	RS1.close
%>
<TR bgcolor="#CCEEFF">
	<TD>Œ—ÊÃ «“ «‰»«—</TD>
	<TD>&nbsp;</TD>
	<TD colspan="2">‰«„ ﬂ«·«</TD>
	<TD>&nbsp;</TD>
</TR>
<%	
mySQL="SELECT InvoiceOrderRelations.Invoice, InventoryPickuplists.id AS PickupListID, InventoryPickuplists.CreationDate, InventoryPickuplistItems.ItemID, InventoryPickuplistItems.ItemName, InventoryPickuplistItems.Qtty, InventoryPickuplistItems.unit, InventoryItemsUnitPrice.UnitPrice FROM InventoryPickuplistItems INNER JOIN InventoryPickuplists ON InventoryPickuplistItems.pickupListID = InventoryPickuplists.id INNER JOIN InvoiceOrderRelations ON InventoryPickuplistItems.Order_ID = InvoiceOrderRelations.[Order] LEFT OUTER JOIN InventoryItemsUnitPrice ON InventoryPickuplists.CreationDate >= InventoryItemsUnitPrice.StartDate AND InventoryPickuplists.CreationDate <= InventoryItemsUnitPrice.EndDate AND InventoryPickuplistItems.ItemID = InventoryItemsUnitPrice.InventoryItem WHERE (InventoryPickuplistItems.CustomerHaveInvItem = 0) AND (NOT (InventoryPickuplists.Status = N'del')) AND (InvoiceOrderRelations.Invoice = " & Invoice & ")"

set RS1=Conn.Execute (mySQL)
	if not RS1.eof then
		while not RS1.eof
			tmpCounter=0
			if isnull(RS1("UnitPrice")) then
				priceText = "<FONT COLOR=""red"">»œÊ‰ ﬁÌ„ </FONT>"
			else
				priceText = "<a target=""_blank"" title=""„‘Œ’«  ﬂ«·«Ì «‰»«—"" href=""../inventory/editItem.asp?itemDetail=" & RS1("ItemID") & """>" & Separate(cdbl(RS1("UnitPrice")) * cdbl(RS1("Qtty"))) & "</a>"
			end if
%>
		<TR bgcolor="#FFFFFF">
			<TD><a target="_blank" title="Œ—ÊÃ «“ «‰»«— „—»ÊÿÂ" href="../inventory/default.asp?show=<%=RS1("PickupListID")%>"><%=RS1("PickupListID")%></a></TD>
			<TD DIR="LTR"><%=RS1("CreationDate")%></TD>
			<TD colspan="2"><a target="_blank" title="„‘Œ’«  ﬂ«·«Ì «‰»«—" href="../inventory/editItem.asp?itemDetail=<%=RS1("ItemID")%>"><%=RS1("Qtty") & " " & RS1("Unit") & " " & RS1("ItemName")%></a></TD>
			<TD><%=priceText%>&nbsp;</TD>
		</TR>

<%
		RS1.moveNext
		wend
	else
%>		<TR>
			<TD colspan="10" align="center" style="font-size:9pt;color:#888888">«ÿ·«⁄«  «‰»«— »—«Ì «Ì‰ ›«ﬂ Ê— ÃÊœ ‰œ«—œ</TD>
		</TR>
<%
	end if 
	RS1.close
%>
</TABLE>
</center>
<%
end if

Conn.Close
%>