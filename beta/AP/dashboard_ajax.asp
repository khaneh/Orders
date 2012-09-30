<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<!--#include File="../JSON_2.0.4.asp"-->
<%
startDate="1389/01/01"
Set result = jsArray()
select case request("act")
	case "ifPickLow":
		set rs=Conn.Execute("select InventoryItems.OldItemID,sumQtty-qtty as afterQtty,InventoryItems.Name,InventoryItems.Unit from InventoryItems inner join (select ItemID,sum(qtty) as sumQtty from InventoryPickuplistItems inner join InventoryPickuplists on InventoryPickuplistItems.pickupListID=InventoryPickuplists.id and InventoryPickuplists.Status='new' group by ItemID) as pick on InventoryItems.ID=pick.itemID where enabled=1 and owner=-1 and qtty-sumQtty<=minim and qtty>minim")
		while not rs.eof 
			set result(null) = jsObject()
			result(null)("div")= "<span class='rspan0'>" & rs("oldItemID") & "</span>"_
				& "<span class='rspan3'>" & rs("Name") & "</span>"_
				& "<span class='rspan0'><b>" & rs("afterQtty") & "</b></span>"_
				& "<span class='rspam2'>" & rs("unit") & " ﬂ”— ŒÊ«Âœ ¬„œ</span>"
			result(null)("link")=""
			result(null)("title")=""
			rs.moveNext
		wend
		rs.close
		set rs=nothing
	case "invLow":
		set rs=Conn.Execute ("select OldItemID,Name,Unit,Qtty,minim from InventoryItems where enabled=1 and owner=-1 and qtty<minim")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("div")="<span class='rspan0'>"_ 
				& rs("oldItemID") & "</span><span class='rspan3'>"_
				& rs("Name") & "</span><span class='rspan1'> „ÊÃÊœÌ: <b>"_ 
				& rs("Qtty") & "</b></span> <span class='rspan1'>Õœ«ﬁ·: "_
				& rs("minim") & "</span><span class='rspan0'>"& rs("Unit") &"</span>"
			result(null)("link")=""
			result(null)("title")=""
			rs.moveNext
		wend
	case "invNOvo":
		set rs=Conn.Execute ("select id,typeName,qtty,ordDate,comment from PurchaseOrders where ordDate>'" & startDate & "' and IsService=0 and status<>'CANCEL' and id not in (select VoucherLines.RelatedPurchaseOrderID from Vouchers inner join VoucherLines on VoucherLines.Voucher_ID=Vouchers.id where Vouchers.Voided=0)")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("div")= "<span class='rspan3'>" & rs("typeName") & "</span>"_
				& "<span class='rspan0'><b>" & rs("Qtty") & "</b></span>"
			result(null)("link")="../purchase/outServiceTrace.asp?od=" & rs("id")
			result(null)("title")= rs("comment") & " œ—  «—ÌŒ " & rs("ordDate")
			rs.moveNext
		wend
	case "serviceNoVo":
		set rs=Conn.Execute ("select id,typeName,qtty,ordDate,comment from PurchaseOrders where ordDate>'" & startDate & "' and IsService=1 and status<>'CANCEL' and id not in (select VoucherLines.RelatedPurchaseOrderID from Vouchers inner join VoucherLines on VoucherLines.Voucher_ID=Vouchers.id where Vouchers.Voided=0)")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("div")= "<span class='rspan3'>" & rs("typeName") & "</span>"_
				& "<span class='rspan0'><b>" & rs("Qtty") & "</b></span>"
			result(null)("link")="../purchase/outServiceTrace.asp?od=" & rs("id")
			result(null)("title")= rs("comment") & " œ—  «—ÌŒ " & rs("ordDate")
			rs.moveNext
		wend
	case "noInvRelation":
		set rs=Conn.Execute ("select id,typeName,qtty,ordDate,comment  from PurchaseOrders where ordDate>'" & startDate & "' and IsService=0 and status<>'CANCEL' and id not in (select distinct relatedID from InventoryLog where RelatedID>0 and isInput=1 and owner=-1 and type=1)")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("div")= "<span class='rspan3'>" & rs("typeName") & "</span>"_
				& "<span class='rspan0'><b>" & rs("Qtty") & "</b></span>"
			result(null)("link")="../purchase/outServiceTrace.asp?od=" & rs("id")
			result(null)("title")= rs("comment") & " œ—  «—ÌŒ " & rs("ordDate")
			rs.moveNext
		wend
	case "unPaid":
		 set result(null) = jsObject()
 		result(null)("div")="<span class='rspan1' title=''> «—ÌŒ</span>"&_
 			"<span class='rspan3'>‰«„ Õ”«»</span>" &_ 
 			"<span class='rspan1'>„»·€</span>" &_
 			"<span class='rspan1'>„«‰œÂ</span>" &_
 			"<span class='rspan1'>„«‰œÂ Õ”«»</span>"
 		result(null)("link")=""
 		result(null)("title")=""
		set rs=Conn.Execute ("select Vouchers.id, accounts.AccountTitle, apItems.AmountOriginal, apItems.EffectiveDate, apItems.RemainedAmount, accounts.APBalance from Vouchers inner join APItems on Vouchers.id=apItems.Link and apItems.Type=6 inner join Accounts on apItems.Account=accounts.ID where APItems.Voided=0 and Vouchers.Voided=0 and Vouchers.paid=0 and APItems.FullyApplied=0 and APItems.effectiveDate>'" & startDate & "'")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("div")="<span class='rspan1'>" & rs("EffectiveDate") & "</span>"_
				 & "<span class='rspan3'>" & rs("AccountTitle") & "</span>"_ 
 				& "<span class='rspan1'>" & rs("AmountOriginal") & "</span>"_
 				& "<span class='rspan1'>" & rs("RemainedAmount") & "</span>"_
 				& "<span class='rspan1'>" & rs("APBalance") & "</span>"
			result(null)("link")="AccountReport.asp?act=showVoucher&voucher="&rs("id")
			result(null)("title")=""
			rs.moveNext
		wend
end select
Response.Write toJSON(result)
%>