<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= " ÊÌ—«Ì‘ ›«ﬂ Ê— "
SubmenuItem=1
' ATTENTION: This is the permission for VERIFY not EDIT
if not Auth(7 , 2) then NotAllowdToViewThisPage() 
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
if request("act")="editVoucher" and isnumeric(request("VoucherID")) then

	VoucherID=clng(request("VoucherID"))

	mySQL = "SELECT * FROM Vouchers WHERE (ID="& VoucherID & ")" 
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ÅÌœ« ‰‘œ.")
	else
		VendorID=		RS1("VendorID")
		Title=			RS1("Title")
		CreationDate=	RS1("CreationDate")
		CreationTime=	RS1("CreationTime")
		SavedFileName=	RS1("SavedFileName")
		Comment=		RS1("Comment")
		TotalPrice=		RS1("TotalPrice")
		Verified=		RS1("Verified")
		Paid=			RS1("Paid")
		EffectiveDate=	RS1("EffectiveDate")
		number =		RS1("Number")
		totalVat =		RS1("TotalVat")
	end if

	RS1.close

	if Verified then
		conn.close
		response.redirect "AccountReport.asp?act=showVoucher&voucher=" & VoucherID & "&errMsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— Œ—Ìœ ﬁ»·«  «ÌÌœ ‘œÂ «” .<br>·–« «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ.")
	end if

%>

<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>
<SCRIPT LANGUAGE="JavaScript">
<!--
function selectVendor(button){
	document.all.tmpDlgArg.value="#"
	document.all.tmpDlgTxt.value="‰«„ Õ”«»Ì —« ﬂÂ „Ì ŒÊ«ÂÌœ Ã” ÃÊ ﬂ‰Ìœ Ê«—œ ﬂ‰Ìœ:"
	window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
	if (document.all.tmpDlgTxt.value !="") {
		window.showModalDialog('../AP/dialog_SelectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value), document.all.tmpDlgArg, 'dialogWidth:780px; dialogHeight:500px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
		if (document.all.tmpDlgArg.value!="#"){
			Arguments=document.all.tmpDlgArg.value.split("#")
			button.parentNode.getElementsByTagName("input")[0].value=Arguments[0];
			button.parentNode.getElementsByTagName("input")[1].innerText=Arguments[1]+"["+Arguments[0]+"]";
		}
	}
}

//-->
</SCRIPT>
<BR><BR>
<FORM METHOD=POST ACTION="?act=submitEdit">
<TABLE align=center>
	<TR height=30>
		<TD width=100 align=left>›—Ê‘‰œÂ</td>
		<TD width=320>
			<% set RSV=Conn.Execute ("SELECT * FROM Accounts where ID = " & VendorID) 
			if RSV.eof then
				response.redirect "voucherInput.asp?act=search"
			end if
			%>
			<INPUT TYPE="hidden"  NAME="VendorID"  value="<%=VendorID%>">
			<INPUT TYPE="text" readonly size=50 NAME="vendorName" value="<%=RSV("AccountTitle")%>" style="height:20;border: 1 solid black;">
			<INPUT TYPE="Button" value=" €ÌÌ—"style="width:40;height:20;border: 1 solid black;" onClick="selectVendor(this);">
			<%
			RSV.close
			%>
		</TD>
		<TD width=150>
			&nbsp;  «—ÌŒ „ÊÀ— <INPUT TYPE="text" size=10 NAME="EffectiveDate" style="direction:LTR; text-align:right" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10" value="<%=EffectiveDate%>">
		</TD>
	</TR>
	<TR height=40>
		<TD align=left>⁄‰Ê«‰ ›«ﬂ Ê— </TD>
		<TD>
			<INPUT TYPE="text" size=60 NAME="title" style="direction:LTR; text-align:right" Value="<%=Title%>">
		</TD>
		<TD>
			&nbsp;‘„«—Â :
			<%=VoucherID%>
			<INPUT TYPE="hidden"  NAME="VoucherID"  value="<%=VoucherID%>">
		</TD>
	</TR>
	<TR>
		<TD align='left'>‘„«—Â ›«ò Ê— «’·Ì</TD>
		<TD><input type='text' size='14' name='Number' value='<%=Number%>'></TD>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP>”›«—‘ Â«Ì „—»Êÿ </TD>
		<TD colspan=2 VALIGN=TOP>
		<%
		if VendorID="-1" then 
			response.write "<hr><br><br>"
		else
			set RSS=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (Status <> 'CANCEL' and  HasVoucher = 0 and Vendor_ID="& VendorID& ") order by OrdDate" )
			%>			
			<TABLE dir=rtl style="border:2 solid #333399;">
			<TR bgcolor="eeeeee" >
				<TD></TD>
				<TD><SMALL>‘„«—Â</SMALL></TD>
				<TD><SMALL>ﬂ«·« Ì« ⁄„·Ì« </SMALL></TD>
				<TD><SMALL>ﬁÌ„  (—Ì«·)</SMALL></TD>
				<TD><SMALL> ⁄œ«œ</SMALL></TD>
				<TD> «—ÌŒ ”›«—‘</SMALL></TD>
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

				%>
				<TR bgcolor="<%=tmpColor%>" >
					<TD><INPUT TYPE="hidden" name=color1 value="<%=tmpColor%>"><INPUT TYPE="hidden" name=color2 value="<%=tmpColor2%>"><INPUT TYPE="checkbox" onclick="setPrice(this)"  name="pcheck" id="<%=tmpCounter-1%>" value="<%=RSS("id")%>" ><!checked></TD>
					<TD><INPUT TYPE="hidden" name="purchaseOrderID" value="0"><A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("id")%>"><%=RSS("id")%></A></TD>
					<TD><INPUT TYPE="text" NAME="TypeName" style="border:none; background: transparent" size=40 value="<%=RSS("TypeName")%>"></TD>
					<TD><INPUT TYPE="text" NAME="price"  style="border:none; background: transparent "  onblur="setPrice(this)" onKeyPress="return maskNumber(this);" size=8 value="<%=Separate(RSS("price"))%>"></TD>
					<TD><INPUT TYPE="text" NAME="qtty"   style="border:none; background: transparent "  onKeyPress="return maskNumber(this);" size=3 value="<%=RSS("qtty")%>"></TD>
					<TD  align=left><span dir=ltr><%=RSS("OrdDate")%></span></TD>
				</TR>
					  
				<% 
				RSS.moveNext
			Loop

			set RS1=Conn.Execute ("SELECT VoucherLines.*, PurchaseOrders.OrdDate FROM VoucherLines INNER JOIN PurchaseOrders ON VoucherLines.RelatedPurchaseOrderID = PurchaseOrders.ID WHERE (Voucher_ID = "& VoucherID & ")" )

		'---	Already Selected Orders
			totalprice=0
			Do while not RS1.eof
				totalprice = totalprice + cdbl(RS1("price"))
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 

				RelatedPurchaseOrderID = RS1("RelatedPurchaseOrderID")
				%>
				<TR bgcolor="<%=tmpColor2%>" >
					<TD><INPUT TYPE="hidden" name=color1 value="<%=tmpColor%>"><INPUT TYPE="hidden" name=color2 value="<%=tmpColor2%>"><INPUT TYPE="checkbox" onclick="setPrice(this)"  name="pcheck" id="<%=tmpCounter-1%>" value="<%=RelatedPurchaseOrderID%>" checked></TD>
					<TD><INPUT TYPE="hidden" name="purchaseOrderID" value="<%=RelatedPurchaseOrderID%>"><A HREF="../purchase/outServiceTrace.asp?od=<%=RelatedPurchaseOrderID%>"><%=RelatedPurchaseOrderID%></A></TD>
					<TD><INPUT TYPE="text" NAME="TypeName" style="border:none; background: transparent" size=40 value="<%=RS1("LineTitle")%>"></TD>
					<TD><INPUT TYPE="text" NAME="price"  style="border:none; background: transparent "  onblur="setPrice(this)" onKeyPress="return maskNumber(this);" size=8 value="<%=Separate(RS1("price"))%>"></TD>
					<TD><INPUT TYPE="text" NAME="qtty"   style="border:none; background: transparent "  onKeyPress="return maskNumber(this);" size=3 value="<%=RS1("qtty")%>"></TD>
					<TD  align=left><span dir=ltr><%=RS1("OrdDate")%></span></TD>
				</TR>
<% 
				RS1.moveNext
			Loop
		'---
%>
			<TR bgcolor="#5555BB" height="2">
				<TD colspan=7></TD>
			</TR>
			<TR bgcolor="#eeeeee" height="2">
				<TD colspan='3' align='left'>„«·Ì«  »— «—“‘ «›“ÊœÂ: </TD>
				<TD colspan=4><INPUT style="border:none; background:transparent;" TYPE="text" NAME="totalVat"  size=16 value="<%=totalVat%>" onchange='setPrice(this)'>  —Ì«·</td>
			</TR>
			<TR bgcolor="#eeeeee" height="2">
				<TD colspan=3 align=left>Ã„⁄ ﬂ·</TD>
				<TD colspan=3><INPUT  style="border:none; background:transparent;"  TYPE="text" NAME="totalPrice" readonly size=16 value="<%=Separate(cdbl(totalPrice) + cdbl(totalVat))%>">  —Ì«·</td>
			</TR>
			</TABLE><br>
<%
		end if
%>
		</TD>
	</TR>
	<TR disabled>
		<TD  align=left VALIGN=TOP>›«Ì·  ’ÊÌ— ›«ﬂ Ê— </TD>
		<TD colspan=2 VALIGN=TOP>
			<INPUT TYPE="Text" size=62 NAME="ImageFile" Value="<%=SavedFileName%>"><BR><BR>
		</TD>
	</TR>
	<TR>
		<TD  align=left VALIGN=TOP> Ê÷ÌÕ« </TD>
		<TD colspan=2 VALIGN=TOP>
			<TEXTAREA NAME="comment" ROWS="7" COLS="48"><%=Comment%></TEXTAREA>
		</TD>
	</TR>
	<TR>
		<TD colspan=3 align=center VALIGN=TOP>
			<br>
			<INPUT class=inputBut TYPE="submit" Name="Submit" Value=" –ŒÌ—Â  €ÌÌ—«  " style="width:125px;" <%if VendorID="-1" then response.write " disabled " %>>
		</TD>
	</TR>
</TABLE>

</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--

function setPrice(obj)
{
	a= obj.type
	if (a=="text")
		obj.value=val2txt(txt2val(obj.value))
	else
		{
		ii=parseInt(obj.id) 
		if(obj.checked)
			{
			document.getElementsByName("purchaseOrderID")[ii].value = obj.value
			theTR = obj.parentNode.parentNode
			theTR.setAttribute("bgColor",document.getElementsByName("color2")[ii].value)
			}
		else
			{
			document.getElementsByName("purchaseOrderID")[ii].value = 0
			theTR = obj.parentNode.parentNode
			theTR.setAttribute("bgColor",document.getElementsByName("color1")[ii].value)
			}
		}
	totalPrice = 0;
	checkBoxList = document.getElementsByName("pcheck")
	for(i=0;i<document.getElementsByName("price").length;i++) {
			if(checkBoxList[i].checked)
			totalPrice = totalPrice + txt2val(document.getElementsByName("price")[i].value);
		}
	totalPrice = totalPrice + txt2val(document.getElementsByName("totalVat")[0].value);
	document.all.totalPrice.value = val2txt(totalPrice) ;

}
//-->
</SCRIPT>
<% 
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------ Submit an Inventory Item request For Buy
'-----------------------------------------------------------------------------------------------------
elseif request("act") = "submitEdit" then
	VoucherID		=	clng(request.Form("VoucherID"))
	VendorID		=	clng(request.Form("VendorID"))
	Title			=	sqlSafe(request.Form("title"))
	Totalprice		=	text2value(request.Form("Totalprice"))
	Comment			=	sqlSafe(request.Form("Comment"))
	EffectiveDate	=	sqlSafe(request.Form("EffectiveDate"))
	number			=	sqlSafe(request.form("number"))
	totalVat		=	text2value(request.form("totalVat"))

	EditDate		=	shamsiToday()
	EditBy			=	session("id")

	if title = "" then title = "-"
	if comment = "" then comment = "-"

	'---- Checking wether EffectiveDate is valid in current open GL
	if (EffectiveDate < session("OpenGLStartDate")) OR (EffectiveDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?act=editVoucher&VoucherID="& VoucherID & "&errMsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ œ— „ÕœÊœÂ ”«· „«·Ì Ã«—Ì ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	mySQL = "SELECT * FROM Accounts WHERE (ID="& VendorID & ")" 
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("›—Ê‘‰œÂ „Ê—œ‰Ÿ— ÊÃÊœ ‰œ«—œ.")
	end if
	RS1.close

	mySQL = "SELECT * FROM Vouchers WHERE (ID="& VoucherID & ")" 
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ÅÌœ« ‰‘œ.")
	else
		OldVendorID=	RS1("VendorID")
		Verified=		RS1("Verified")
	end if
	RS1.close

	if Verified then
		conn.close
		response.redirect "AccountReport.asp?act=showVoucher&voucher=" & VoucherID & "&errMsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— Œ—Ìœ ﬁ»·«  «ÌÌœ ‘œÂ «” .<br>·–« «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ.")
	end if


	mysql = "UPDATE Vouchers SET VendorID='" & VendorID & "', Title=N'"& Title & "', TotalPrice='"& Totalprice & "', comment=N'"& Comment & "', LastEditDate=N'"& EditDate & "', LastEditBy='"& EditBy & "', EffectiveDate=N'"& EffectiveDate & "', Number=N'" & number &"', TotalVat=N'" & totalVat & "' WHERE (ID='"& VoucherID & "')"
	Conn.Execute (mysql)

	'===================================================
	'	Removing old voucherLines AND releasing Purchase Orders 
	'===================================================
		mysql = "UPDATE PurchaseOrders SET HasVoucher=0 WHERE (ID IN (SELECT RelatedPurchaseOrderID FROM VoucherLines WHERE (Voucher_ID = "& VoucherID & ")))"
		Conn.Execute (mysql)

		mySQL = "DELETE FROM VoucherLines WHERE (Voucher_ID = "& VoucherID & ")"
		Conn.Execute (mysql)

	'===================================================
	'	Adding new voucherLines 
	'===================================================
	for i=1 to request.Form("purchaseOrderID").count
		POrderID	=	request.Form("purchaseOrderID")(i)

		OTypeName	=	sqlSafe(request.Form("TypeName")(i))
		price		=	text2value(request.Form("price")(i))
		qtty		=	request.Form("qtty")(i)

		if POrderID <>"0" then
			mysql = "INSERT INTO VoucherLines(Voucher_ID, RelatedPurchaseOrderID, LineTitle, price, qtty) VALUES ("& VoucherID & ", "& POrderID & ",N'"& OTypeName & "',"& price & ","& qtty & ")"
			Conn.Execute (mysql)

			mysql = "UPDATE PurchaseOrders SET HasVoucher = 1 WHERE (ID = "& POrderID & ")"
			Conn.Execute (mysql)
		end if
	next

	'===================================================
	response.redirect "AccountReport.asp?act=showVoucher&voucher=" & VoucherID & "&msg=" &Server.URLEncode("›«ﬂ Ê— À»  ‘œ")

elseif request("act") = "delVoucher" and isnumeric(request("VoucherID")) then

	VoucherID=clng(request("VoucherID"))

	mySQL = "SELECT * FROM Vouchers WHERE (ID="& VoucherID & ")" 
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ÅÌœ« ‰‘œ.")
	else
		VendorID=		RS1("VendorID")
		Title=			RS1("Title")

		CreationDate=	RS1("CreationDate")
		CreationTime=	RS1("CreationTime")
		SavedFileName=	RS1("SavedFileName")
		Comment=		RS1("Comment")
		TotalPrice=		RS1("TotalPrice")
		Verified=		RS1("Verified")
		Paid=			RS1("Paid")
	end if

	RS1.close

	if Verified then
		conn.close
		response.redirect "AccountReport.asp?act=showVoucher&voucher=" & VoucherID & "&errMsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— Œ—Ìœ ﬁ»·«  «ÌÌœ ‘œÂ «” .<br>·–« «„ﬂ«‰ Õ–› ¬‰ ÊÃÊœ ‰œ«—œ.")
	end if

	'===================================================
	'	Removing voucherLines AND releasing Purchase Orders 
	'===================================================
		mysql = "UPDATE PurchaseOrders SET HasVoucher=0 WHERE (ID IN (SELECT RelatedPurchaseOrderID FROM VoucherLines WHERE (Voucher_ID = "& VoucherID & ")))"
		Conn.Execute (mysql)

		mySQL = "DELETE FROM VoucherLines WHERE (Voucher_ID = "& VoucherID & ")"
		Conn.Execute (mysql)

	'===================================================
	'	Deleting Voucher
	'===================================================
		'mySQL = "DELETE FROM Vouchers WHERE (ID = "& VoucherID & ")"
		'Changed By Kid 830929 
		mySQL = "UPDATE Vouchers SET Voided=1 WHERE (ID = "& VoucherID & ")"
		Conn.Execute (mysql)

	'===================================================

	conn.close
	response.redirect "top.asp?msg=" & Server.URLEncode("›«ﬂ Ê— Œ—Ìœ „Ê—œ ‰Ÿ— Õ–› ‘œ.")
elseif request("act") = "voidVoucher" then
	if NOT Auth(7 , 9) then 
		'Doesn't have the Priviledge to APPROVE the Voucher
		response.write "<br>" 
		call showAlert ("‘„« „Ã«“ »Â «»ÿ«· ›«ﬂ Ê— Œ—Ìœ ‰Ì” Ìœ",CONST_MSG_ERROR) 
		response.end
	end if

	ON ERROR RESUME NEXT
		VoucherID = clng(request("VoucherID"))
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— Œ—Ìœ „⁄ »— ‰Ì” .")
		end if
	ON ERROR GOTO 0

	mySQL = "SELECT * FROM Vouchers WHERE (ID="& VoucherID & ")" 
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ÅÌœ« ‰‘œ.")
	else
		VendorID=		RS1("VendorID")
		Title=			RS1("Title")
		CreationDate=	RS1("CreationDate")
		CreationTime=	RS1("CreationTime")
		SavedFileName=	RS1("SavedFileName")
		Comment=		RS1("Comment")
		TotalPrice=		RS1("TotalPrice")
		Verified=		RS1("Verified")
		Paid=			RS1("Paid")
		if RS1("Voided") then
			response.redirect "AccountReport.asp?act=showVoucher&voucher=" & VoucherID & "&errMsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·« œ—  «—ÌŒ <span dir='LTR'>"& RS1("VoidedDate") & "</span> »«ÿ· ‘œÂ «” .")
		end if
	end if

	RS1.close

	if NOT Verified then
		conn.close
		response.redirect "AccountReport.asp?act=showVoucher&voucher=" & VoucherID & "&errMsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— Œ—Ìœ ﬁ»·«  «ÌÌœ ‰‘œÂ «” .<br>·–« «„ﬂ«‰ «»ÿ«· ¬‰ ÊÃÊœ ‰œ«—œ.")
	end if

	'===================================================
	'	Voiding Voucher
	'===================================================
	mySQL="UPDATE Vouchers SET Voided=1, VoidedDate=N'"& shamsiToday() & "', VoidedBy='"& session("ID") & "' WHERE (ID = "& VoucherID & ")"
	conn.Execute(mySQL)

	'===================================================
	'	releasing Purchase Orders 
	'===================================================
	mysql = "UPDATE PurchaseOrders SET HasVoucher=0 WHERE (ID IN (SELECT RelatedPurchaseOrderID FROM VoucherLines WHERE (Voucher_ID = "& VoucherID & ")))"
	Conn.Execute (mysql)

  '**********************************	 Voiding APItem of Voucher  **************************
  '*** Type = 6 means APItem is a Voucher
  '***
	'*********  Finding the APItem of Voucher
	mySQL="SELECT ID FROM APItems WHERE (Type = 6) AND (Link='"& VoucherID & "')"
	Set RS1=conn.Execute(mySQL)
	voidedAPItem=RS1("ID")

	'*********  Finding other APItems related to this Item
	mySQL="SELECT ID AS RelationID, DebitAPItem, Amount FROM APItemsRelations WHERE (CreditAPItem = '"& voidedAPItem & "')"
	Set RS1=conn.Execute(mySQL)
	Do While not (RS1.eof)
		'*********  Adding back the amount in the relation, to the debit APItem ...
		conn.Execute("UPDATE APItems SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("DebitAPItem") & "')")

		'*********  Deleting the relation
		conn.Execute("DELETE FROM APItemsRelations WHERE ID='"& RS1("RelationID") & "'")
		
		RS1.movenext
	Loop
	RS1.close

	'*********  Voiding APItem 
	mySQL="UPDATE APItems SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& voidedAPItem & "')"
	conn.Execute(mySQL)

	'*********  Affecting Account's AP Balance  
	mySQL="UPDATE Accounts SET APBalance = APBalance - "& TotalPrice & "  WHERE (ID = "& VendorID & ")"
	conn.Execute(mySQL)
	
  '***
  '***----------------------------- End of Voiding APItem of Voucher  ------------------------

	response.redirect "AccountReport.asp?act=showVoucher&voucher=" & VoucherID
else
	response.write "<br><br>"
	call showAlert (" Œÿ« ! ",CONST_MSG_ERROR) 
end if
%>
<!--#include file="tah.asp" -->
