<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= "  «ÌÌœ ›«ﬂ Ê—"
SubmenuItem=2
if not Auth(7 , 2) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------- Submit Payment
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")=" «ÌÌœ" then

	'Check for permission for APPROVING the Voucher
	if NOT Auth(7 , 2) then NotAllowdToViewThisPage()

	ON ERROR RESUME NEXT
		VouchID = clng(request("VouchID"))
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— Œ—Ìœ „⁄ »— ‰Ì” .")
		end if
	ON ERROR GOTO 0

	set RSV=Conn.Execute ("SELECT * FROM Vouchers WHERE (id="& VouchID & ") and Verified = 0" )
	if RSV.EOF then
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br>ç‰Ì‰ ›«ﬂ Ê—Ì „ÊÃÊœ ‰Ì”  Ì« ﬁ»·«  «ÌÌœ ‘œÂ «” .")
	end if
	TotalPrice =	RSV("TotalPrice")
	VendorID =		RSV("VendorID")
	totalVat =		RSV("totalVat")

	effectiveDate = RSV("EffectiveDate")

	'---- Checking wether EffectiveDate is valid in current open GL
	if (effectiveDate < session("OpenGLStartDate")) OR (effectiveDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ „ÊÀ— «Ì‰ ›«ﬂ Ê— „—»Êÿ »Â ”«· „«·Ì Ã«—Ì ‘„« ‰Ì” .<br><br> Ì«  «—ÌŒ „ÊÀ— ›«ﬂ Ê— —« ⁄Ê÷ ﬂ‰Ìœ Ì« ”«· „«·Ì  «‰ —«.")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	Conn.Execute ("UPDATE Vouchers SET Verified = 1 WHERE (id = "& VouchID & ")")
	Conn.Execute ("UPDATE Accounts SET APBalance=APBalance+"& TotalPrice & " WHERE (ID = "& VendorID & ")")	

	creationDate=	shamsiToday()
	creationTime=	CurrentTime10()

	GLAccount=		"89099"	'This must be changed... (Misc. Purchase)
	firstGLAccount=	"41001"	'This must be changed... (Business Creditors)

	'*** Type = 6 means APItem is a Voucher

	conn.Execute("INSERT INTO APItems (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, AmountOriginal, createdDate, creationTime, CreatedBy, RemainedAmount, Vat) VALUES (" &_
	GLAccount & ", '"& OpenGL & "', '"& firstGLAccount & "', '"& VendorID & "', N'"& EffectiveDate & "', 1, 6, '"& VouchID & "', '"& TotalPrice & "', N'"& creationDate & "', N'"& creationTime & "', '"& session("ID") & "', '"& TotalPrice & "', '" & totalVat & "')" )

	Conn.close
	response.redirect "AccountReport.asp?act=showVoucher&voucher="& VouchID

end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- Show Voucher Details
'-----------------------------------------------------------------------------------------------------
if request("VouchID") <> "" then
	response.write "<br><br>"
	call showAlert("«‘ »«ÂÌ ÅÌ‘ ¬„œÂ «” .<br><br>‘„« ﬁ«⁄œ « ‰»«Ìœ »Â «Ì‰ ﬁ”„  „Ì ¬„œÌœ .<br><br>·ÿ›« ’›ÕÂ «Ì —« ﬂÂ «“ ¬‰Ã« »Â «Ì‰ ’›ÕÂ „‰ ﬁ· ‘œÂ «Ìœ —« »Â «›—«“ Ì« Õ«ÃÌ “«œÂ »êÊÌÌœ. <br><br>⁄–— ŒÊ«ÂÌ „Ì ﬂ‰Ì„.<br><br>»«Ìœ »Â <A HREF='AccountReport.asp?act=showVoucher&voucher="& request("VouchID") & "'><B>«Ì‰Ã«</B></A> „Ì —› Ìœ." , CONST_MSG_ALERT)
	response.end
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------ Main
'-----------------------------------------------------------------------------------------------------

'mySQL="SELECT DRV.OrdDate AS OrdDate, DRV.RelatedPurchaseOrderID, DRV.LineTitle, DRV.qtty, DRV.price, DRV.Invoice, DRV.status, SUM(InvoiceLines.Price - InvoiceLines.Discount - InvoiceLines.Reverse) AS SalePrice, COUNT(InvoiceLines.Price) AS COUNT FROM (SELECT VoucherLines.*, PurchaseOrders.OrdDate, InvoiceOrderRelations.Invoice AS Invoice, InvoiceItems.ID AS InvoiceItem, PurchaseOrders.status FROM InvoiceItems RIGHT OUTER JOIN InvoiceOrderRelations INNER JOIN PurchaseRequests INNER JOIN PurchaseRequestOrderRelations ON PurchaseRequests.ID = PurchaseRequestOrderRelations.Req_ID ON InvoiceOrderRelations.[Order] = PurchaseRequests.Order_ID RIGHT OUTER JOIN PurchaseOrders ON PurchaseRequestOrderRelations.Ord_ID = PurchaseOrders.ID ON InvoiceItems.RelatedInventoryItemID = PurchaseOrders.TypeID FULL OUTER JOIN Vouchers INNER JOIN 					 VoucherLines ON Vouchers.id = VoucherLines.Voucher_ID ON PurchaseOrders.ID = VoucherLines.RelatedPurchaseOrderID WHERE (PurchaseOrders.Status <> 'CANCEL') AND (Vouchers.Verified=0)) DRV LEFT OUTER JOIN InvoiceLines ON DRV.Invoice = InvoiceLines.Invoice AND DRV.InvoiceItem = InvoiceLines.Item GROUP BY DRV.OrdDate, DRV.RelatedPurchaseOrderID, DRV.LineTitle, DRV.qtty, DRV.price, DRV.Invoice, DRV.status ORDER BY DRV.OrdDate"
' S A M
mySQL="SELECT DRV.ID, DRV.OrdDate AS OrdDate, DRV.RelatedPurchaseOrderID, DRV.LineTitle, DRV.qtty, DRV.price, DRV.Invoice, DRV.status, SUM(InvoiceLines.Price - InvoiceLines.Discount - InvoiceLines.Reverse + InvoiceLines.Vat) AS SalePrice, COUNT(InvoiceLines.Price) AS COUNT FROM (SELECT Vouchers.ID, Voucherlines.RelatedPurchaseOrderID, Voucherlines.LineTitle, Voucherlines.Qtty, Voucherlines.Price, PurchaseOrders.OrdDate, InvoiceOrderRelations.Invoice AS Invoice, InvoiceItems.ID AS InvoiceItem, PurchaseOrders.status FROM InvoiceItems RIGHT OUTER JOIN InvoiceOrderRelations INNER JOIN PurchaseRequests INNER JOIN PurchaseRequestOrderRelations ON PurchaseRequests.ID = PurchaseRequestOrderRelations.Req_ID ON InvoiceOrderRelations.[Order] = PurchaseRequests.Order_ID RIGHT OUTER JOIN PurchaseOrders ON PurchaseRequestOrderRelations.Ord_ID = PurchaseOrders.ID ON InvoiceItems.RelatedInventoryItemID = PurchaseOrders.TypeID FULL OUTER JOIN Vouchers INNER JOIN VoucherLines ON Vouchers.id = VoucherLines.Voucher_ID ON PurchaseOrders.ID = VoucherLines.RelatedPurchaseOrderID WHERE (PurchaseOrders.Status <> 'CANCEL') AND (Vouchers.Verified = 0)) DRV LEFT OUTER JOIN InvoiceLines ON DRV.Invoice = InvoiceLines.Invoice AND DRV.InvoiceItem = InvoiceLines.Item GROUP BY DRV.ID, DRV.OrdDate, DRV.RelatedPurchaseOrderID, DRV.LineTitle, DRV.qtty, DRV.price, DRV.Invoice, DRV.status ORDER BY DRV.OrdDate DESC, DRV.ID"

set RSS=Conn.Execute (mySQL)	'"SELECT * FROM Vouchers WHERE (Verified = 0)" )
%>
<BR><BR><BR>
<style>
	.myTbl td {direction: LTR; text-align:right;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<FORM METHOD=POST ACTION="AccountReport.asp?act=showVoucher">
<TABLE dir=rtl style="border:2 solid #333399;" align=center width="90%">
	<TR bgcolor="#eeeeee" height=25>
		<TD align=center colspan=7><B>›«ﬂ Ê—Â«Ì  «ÌÌœ ‰‘œÂ</B></TD>
	</TR>
	<TR bgcolor="#eeeeee" height=25>
		<TD width=10> # </TD>
		<TD><INPUT TYPE="radio" NAME="" disabled></TD>
		<TD width=120>›—Ê‘‰œÂ </TD>
		<TD>⁄‰Ê«‰ ›«ﬂ Ê— </TD>
		<TD>„»·€ </TD>
		<TD width=60> «—ÌŒ Ê—Êœ </TD>
		<TD>
			<input style='width:120pt;border:0pt;background:transpanrent;font-size:7pt' value='ﬂ«·« Ì« ⁄„·Ì« '>
			<input style='width:25pt;border:0pt;background:transpanrent;font-size:7pt' value=' ⁄œ«œ'>
			<input style='width:35pt;border:0pt;background:transpanrent;font-size:7pt' value='ﬁÌ„ '>
			<input style='width:35pt;border:0pt;background:transpanrent;font-size:7pt' value='ﬁÌ„  ›—Ê‘'>
		</TD>
	</TR>
	<TR bgcolor="#5555BB" height="2">
		<TD colspan=7></TD>
	</TR>
	<%
	tmpCounter=0

	Do while not RSS.eof
		tmpCounter = tmpCounter + 1

		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 

		DisableFlag = ""
		vouchLines = "<table class='myTbl' >"

		Do 
			Voucher=		RSS("ID")

			mySQL="SELECT Vouchers.Title, Vouchers.Number, Vouchers.TotalPrice, Vouchers.comment, Vouchers.CreationDate, Vouchers.VendorID, Accounts.AccountTitle FROM Vouchers INNER JOIN Accounts ON Vouchers.VendorID = Accounts.ID WHERE (Vouchers.ID = "& Voucher & ")"
			set RST=Conn.Execute (mySQL)

			Comment =		RST("Comment")
			VoucherTitle =	RST("title")
			VoucherPrice =	RST("totalPrice")
			CreationDate =	RST("CreationDate")
			Vendor =		RST("VendorID")
			AccountTitle =	RST("AccountTitle")
			number =		RST("Number")

			RST.close

			vouchLines = vouchLines & "<tr><td><input readonly onclick='window.open(""../purchase/outServiceTrace.asp?od=" & RSS("RelatedPurchaseOrderID") & """,""_blank"","""")' type='' style='width:120pt;border:0pt;background:transpanrent;background:transpanrent;font:7pt blue;cursor:hand' value='" & RSS("LineTitle") & "'></td></a><td style='width:25pt;font-size:7pt;'>" & RSS("qtty") & "</td><td><input readonly type='' style='width:35pt;border:0pt;background:transpanrent;font-size:7pt;cursor:hand' value='" & Separate(RSS("price")) & "'></td><td> "
			if trim(" " & RSS("Invoice"))="" then
					vouchLines = vouchLines & "<span disabled style='font-size:7pt'>›«ﬂ Ê— ‰‘œÂ</span>" 
			else
					vouchLines = vouchLines & "<span style='font-size:7pt'><A HREF='../AR/AccountReport.asp?act=showInvoice&invoice="& RSS("Invoice")& "' target='_blank'>" & Separate(RSS("SalePrice"))
					if RSS("count") > 1 then 
						vouchLines = vouchLines & " (" & RSS("count") & " ”ÿ—)"
					end if
				if trim(RSS("SalePrice"))="" or isnull(RSS("SalePrice")) then
					vouchLines = vouchLines & "›«ﬂ Ê— œ«—œ" 
				end if
				vouchLines = vouchLines & "</A></span>"
			end if
			vouchLines = vouchLines & "</tr>"
			if RSS("status") <> "OK" then DisableFlag=" disabled"
			RSS.moveNext
			if NOT RSS.eof then
				nextVoucher=RSS("ID")
			else
				nextVoucher=-1
			end if
		Loop While Voucher= nextVoucher AND NOT RSS.eof

		vouchLines = vouchLines & "</table>"
		%>

		<TR bgcolor="<%=tmpColor%>" <%=DisableFlag%> Title="<% 
					if Comment<>"-" then
						response.write " Ê÷ÌÕ: " & Comment
					else
						response.write " Ê÷ÌÕ ‰œ«—œ"
					end if
				%>">
			<TD><INPUT TYPE="hidden" name=color1 value="<%=tmpColor%>"><INPUT TYPE="hidden" name=color2 value="<%=tmpColor2%>"><A target="_blank" HREF="AccountReport.asp?act=showVoucher&voucher=<%=Voucher%>"><%=Voucher%></A></TD>
			<TD><INPUT TYPE="radio" NAME="voucher" VALUE="<%=Voucher%>" onfocus="setColorFocus(this)" onblur="setColorBlur(this)" id="<%=tmpCounter-1%>"></TD>
			<TD><%=AccountTitle%></TD>
			<TD><%=VoucherTitle%></TD>
			<TD dir=LTR align=right><%=Separate(VoucherPrice)%></TD>
			<TD align=left><span dir=ltr><%=CreationDate%></span></TD>
			<TD><%=vouchLines%></TD>
		</TR>
		<% 
	Loop
	%>
</TABLE>
<br>
<br>
<center><INPUT TYPE="submit" Name="Submit" Value="«‰ Œ«» »—«Ì  «ÌÌœ" class="inputBut" style="width:150px;" tabIndex="14">
</form>
</center>


<SCRIPT LANGUAGE="JavaScript">
<!--
function setColorFocus(obj){
	ii=parseInt(obj.id) 
	theTR = obj.parentNode.parentNode
	theTR.setAttribute("bgColor",document.getElementsByName("color2")[ii].value)
}

function setColorBlur(obj){
	ii=parseInt(obj.id) 
	theTR = obj.parentNode.parentNode
	theTR.setAttribute("bgColor",document.getElementsByName("color1")[ii].value)
}

/*function setColor(obj)
{
	for(i=0; i<document.all.voucher.length; i++)
		{
		theTR = document.all.voucher[i].parentNode.parentNode
		theTR.setAttribute("bgColor","<%=AppFgColor%>")
		}
	theTR = obj.parentNode.parentNode
	theTR.setAttribute("bgColor","#FFFFFF")
}*/

//-->
</SCRIPT>

<!--#include file="tah.asp" -->
