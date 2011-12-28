<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= " Ê—Êœ ›«ﬂ Ê—"
SubmenuItem=1
if not Auth(7 , 1) then NotAllowdToViewThisPage()

DONT_LOG_THIS = true
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
if request("act")="search" or request("act")="" then
	%>
	<br>
	<FORM METHOD=POST ACTION="?act=select" onsubmit="if (document.all.search.value=='') return false;">
	<div dir='rtl'><B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«»</B>
		<INPUT TYPE="text" NAME="search">&nbsp;
		<INPUT TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.search.focus();
	//-->
	</SCRIPT>
	<%
elseif request("act")="select" then
	if request("search") <> "" then
		SA_TitleOrName=request("search")
		SA_Action="return true;"
		'SA_Action="return selectOperations();"
		SA_SearchAgainURL="voucherInput.asp"
		SA_StepText="" '"ê«„ œÊ„ : «‰ Œ«» Õ”«»"
		SA_ActName			= "select"	
		SA_SearchBox	="search"		
		SA_IsVendor = 1
%>
		<FORM METHOD=POST ACTION="?act=show">
		<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	end if


'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------ Submit an Inventory Item request For Buy
'-----------------------------------------------------------------------------------------------------
elseif request("act") = "submitForm" then

	'	ASP Upload Style:
	'
	'set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
	'mySmartUpload.Codepage = 1256
	'mySmartUpload.Upload
	'vendor			=	mySmartUpload.Form("vendor") 
	'title			=	mySmartUpload.Form("title") 
	'totalPrice		=	text2value(mySmartUpload.Form("totalPrice"))
	'ImageFile		=	mySmartUpload.Form("ImageFile") 
	'comment			=	mySmartUpload.Form("comment") 
	'EffectiveDate	=	mySmartUpload.Form("EffectiveDate") 

	vendor			=	clng(request.Form("vendor"))
	title			=	request.Form("title") 
	totalPrice		=	text2value(request.Form("totalPrice"))
	comment			=	request.Form("comment") 
	EffectiveDate	=	request.Form("EffectiveDate") 
	totalVat		=	text2value(request.form("totalVat"))
	number			=	request.form("Number")

	CreationDate	=	shamsiToday()
	CreationTime	=	currentTime10()
	CreatedBy		=	session("id")

	'---- Checking wether EffectiveDate is valid in current open GL
	if (EffectiveDate < session("OpenGLStartDate")) OR (EffectiveDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?act=show&selectedCustomer="& vendor & "&errMsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ œ— „ÕœÊœÂ ”«· „«·Ì Ã«—Ì ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	'	ASP Upload Style:
	'  ( Uploading Files )
	'
	'For each file In mySmartUpload.Files
	'
	'	If not file.IsMissing Then
	'		if file.FileExt <> "" then
	'			 ext = "." & file.FileExt 
	'		else
	'			ext = ""
	'		end if
	'
	'		SavedFileName = replace(shamsiToday(),"/","-") & "_" & replace(currentTime10(),":", "-") &  ext
	'       file.SaveAs("./vouchers/" & SavedFileName)
	'
	'		ImageFile   = file.FileName 
	'		'Response.Write("FilePathName = " & file.FilePathName & "<BR>")
	'	End If
	'Next
	

	response.write "<BR><BR><TABLE align=center style='border: solid 2pt black'><TR><TD>"

	'===================================================
	'create a voucher 
	'===================================================
	if title = "" then title = "-"
	if number = "" then number = "-"
	if ImageFile = "" then ImageFile = "-"
	if comment = "" then comment = "-"

	mysql = "INSERT INTO Vouchers (ImageFileName, SavedFileName, VendorID, Title, TotalPrice, comment, paid, CreationTime, CreationDate, CreatedBy, EffectiveDate, totalVat, Number) VALUES (N'"& ImageFile & "', N'"& SavedFileName & "',  "& vendor & ", N'"& title & "', "& totalPrice & ", N'"& comment & "', 0, N'"& CreationTime & "', N'"& CreationDate & "', "& CreatedBy & ", N'"& EffectiveDate & "', '" & totalVat & "', '" & number & "')"

	Conn.Execute (mysql)
	'response.write mysql
	response.write "<li>  «—ÌŒ : " &  CreationDate & "&nbsp;-&nbsp; " & CreationTime
	response.write "<hr>"
	

	'===================================================
	'create many voucherLines for an voucher item
	'===================================================
	set RSX=Conn.Execute ("SELECT * FROM Vouchers WHERE CreationTime = N'"& CreationTime & "' and CreationDate=N'" & CreationDate & "' and VendorID=" & vendor )	
	VoucherID = RSX("id")
	RSX.close

	for i=1 to request.Form("purchaseOrderID").count
		POrderID	=	request.Form("purchaseOrderID")(i)
		OTypeName	=	request.Form("TypeName")(i)
		price		=	text2value(request.Form("price")(i))
		qtty		=	request.Form("qtty")(i)

		if POrderID <>"0" then
			mysql = "INSERT INTO VoucherLines(Voucher_ID, RelatedPurchaseOrderID, LineTitle, price, qtty) VALUES ("& VoucherID & ", "& POrderID & ",N'"& OTypeName & "',"& price & ","& qtty & ")"
			Conn.Execute (mysql)
			'response.write "<span dir=ltr>"
			'response.write mysql
			'response.write "</span dir=ltr>"
			'response.end

			response.write "<li> " & OTypeName & " ( ⁄œ«œ: " & Qtty & " - ﬁÌ„ : " &  price & ")"
		end if
	next


	'===================================================
	'change status of related purchaseOrder
	'===================================================
	for i=1 to request.Form("purchaseOrderID").count
		POrderID	=	request.Form("purchaseOrderID")(i)

		if POrderID <>"0" then
			'===================================================
			' check if STATUS of related Purchase order is OK
			'===================================================
			isTotalyOKforPay = "yes"
			set RSOD=Conn.Execute ("SELECT * FROM purchaseOrders WHERE id = "& POrderID )	
			if NOT RSOD("Status") = "OK" then
				isTotalyOKforPay = "no"
			end if 

			mysql = "UPDATE PurchaseOrders SET HasVoucher = 1 WHERE (ID = "& POrderID & ")"
			Conn.Execute (mysql)
		end if
	next

	'---------------------------------------------------------------------------------------------------
	'------ This 3 lines has been Commented by Alix - 82-02-18 
	'------ Routine changed: Account will be updated when Voucher verified. (only in AP/Verify.asp page)
	'---------------------------------------------------------------------------------------------------

	'if isTotalyOKforPay="yes" then
	'	Conn.Execute ("UPDATE Accounts SET APBalance=APBalance+"& totalPrice & "  WHERE (ID = "& vendor & ")")	
	'end if
	'===================================================



	response.write "</TD></TR></TABLE><center>"
	response.write "<BR><BR>›«ﬂ Ê— ›Êﬁ À»  ‘œ</center>"
'	response.end

	response.redirect "AccountReport.asp?act=showVoucher&voucher="& VoucherID & "&msg=" &Server.URLEncode("›«ﬂ Ê— À»  ‘œ")

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------ Main
'-----------------------------------------------------------------------------------------------------
elseif request("act")="show" AND request("selectedCustomer") <> "" then
	venID=request("selectedCustomer")

%>
<BR><BR>
<FORM METHOD=POST ACTION="?act=submitForm" >
<TABLE align=center width="90%">
	<TR>
		<TD align=left VALIGN=TOP>›—Ê‘‰œÂ</td>
		<TD align=right VALIGN=TOP>
			<% set RSV=Conn.Execute ("SELECT * FROM Accounts where id = " & venID) 
			if RSV.eof then
				response.redirect "?act=search"
			end if
			%>
			<INPUT TYPE="text" size=62 NAME="vendorName" readonly value="<%=RSV("AccountTitle")%>">
			<INPUT TYPE="hidden"  NAME="vendor"  value="<%=RSV("id")%>">&nbsp;
			 «—ÌŒ „ÊÀ— <INPUT TYPE="text" size=10 NAME="EffectiveDate" style="direction:LTR;" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10" value="<%=shamsiToday()%>">
			<%
			RSV.close
			%>
			<br><br>
		</TD>
	</TR>
	<TR>
		<TD  align=left VALIGN=TOP>⁄‰Ê«‰ ›«ﬂ Ê— </TD>
		<TD align=right VALIGN=TOP>
			<INPUT TYPE="text" size=88 NAME="title" style="direction:RTL;text-align:right">
		</TD>
	</TR>
	<TR>
		<TD align='left'>‘„«—Â ›«ﬂ Ê— «’·Ì </TD>
		<TD>
			<input type='text' size='14' name='Number'>
		</TD>
	</TR>
	<TR>
		<TD  align=left VALIGN=TOP>”›«—‘ Â«Ì „—»Êÿ </TD><br><br>
		<TD align=right VALIGN=TOP>
		<%
		if venID="-1" then 
			response.write "<hr><br><br>"
		else
			totalprice=0
			'Changed By kid 82/12/26 - Show Related Invoice Items If Exist AND Show the Invoice ID anyway.
			'S A M
			mySQL="SELECT DRV.*, InvoiceLines.Price - InvoiceLines.Discount - InvoiceLines.Reverse + InvoiceLines.Vat AS SalePrice FROM (SELECT PurchaseOrders.*, InvoiceOrderRelations.Invoice AS Invoice, InvoiceItems.ID AS InvoiceItem FROM InvoiceItems RIGHT OUTER JOIN InvoiceOrderRelations INNER JOIN PurchaseRequests INNER JOIN PurchaseRequestOrderRelations ON PurchaseRequests.ID = PurchaseRequestOrderRelations.Req_ID ON  InvoiceOrderRelations.[Order] = PurchaseRequests.Order_ID RIGHT OUTER JOIN PurchaseOrders ON PurchaseRequestOrderRelations.Ord_ID = PurchaseOrders.ID ON  InvoiceItems.RelatedInventoryItemID = PurchaseOrders.TypeID WHERE (PurchaseOrders.Status <> 'CANCEL') AND (PurchaseOrders.HasVoucher = 0) AND (PurchaseOrders.Vendor_ID = "& VenID& "))  DRV LEFT OUTER JOIN InvoiceLines ON DRV.Invoice = InvoiceLines.Invoice AND DRV.InvoiceItem = InvoiceLines.Item left outer JOIN INVOICES on InvoiceLines.Invoice = Invoices.id where invoices.voided is null or invoices.voided<>1 ORDER BY DRV.OrdDate,DRV.ID"

'			response.write "<br>" & mySQL
'			response.end
			set RSS=Conn.Execute (mySQL)
			%>			
			<TABLE dir=rtl style="border:2 solid #333399;">
			<TR bgcolor="#eeeeee" height=25>
				<TD><INPUT TYPE="checkbox" disabled checked></TD>
				<TD>‘„«—Â</TD>
				<TD>ﬂ«·« Ì« ⁄„·Ì« </TD>
				<TD>ﬁÌ„  (—Ì«·)</TD>
				<TD> ⁄œ«œ</TD>
				<TD> «—ÌŒ ”›«—‘</TD>
				<TD>ﬁÌ„  ›—Ê‘</TD>
			</TR>
			<TR bgcolor="#5555BB" height="2">
				<TD colspan=7></TD>
			</TR>
			<%
			tmpCounter=0
			lastPurchaseOrder = 0
			Do while not RSS.eof
				tmpCounter = tmpCounter + 1

				purchaseOrder =		RSS("id")
				totalprice =		totalprice + cdbl(RSS("price"))
				Invoice =			RSS("Invoice")
				SalePrice =			RSS("SalePrice")



				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" >
					<TD><INPUT TYPE="hidden" name=color1 value="<%=tmpColor%>">
						<INPUT TYPE="hidden" name=color2 value="<%=tmpColor2%>">
						<INPUT TYPE="checkbox" onclick="setPrice(this)"  name="pcheck" id="<%=tmpCounter-1%>" value="<%=RSS("id")%>">
					</TD>
					<TD dir=LTR align=right>
						<INPUT TYPE="hidden" name="purchaseOrderID" value="0">
						<A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("id")%>" target="_blank"><%=RSS("id")%></A>
					</TD>
					<TD>
						<INPUT TYPE="text" NAME="TypeName" style="border:none; background: transparent" size=40 value="<%=RSS("TypeName")%>">
					</TD>
					<TD>
						<INPUT TYPE="text" NAME="price"  style="border:none; background: transparent; direction:LTR;text-align:right;"  onblur="setPrice(this)" onKeyPress="return maskNumber(this);" size=8 value="<%=Separate(RSS("price"))%>">
					</TD>
					<TD>
						<INPUT TYPE="text" NAME="qtty"   style="border:none; background: transparent; direction:LTR;text-align:right;"  onKeyPress="return maskNumber(this);" size=3 value="<%=RSS("qtty")%>">
					</TD>
					<TD align=left><span dir=ltr><%=RSS("OrdDate")%></span></TD>
					<TD dir=LTR align=right>
<%				
					lastPurchaseOrder = purchaseOrder
					
					Do
						if trim("" & Invoice)="" then
							response.write "<span disabled>›«ﬂ Ê— ‰‘œÂ</span>" 
						else
%>							<A HREF="../AR/AccountReport.asp?act=showInvoice&invoice=<%=Invoice%>" title="›«ﬂ Ê— ‘„«—Â <%=Invoice%>" target="_blank"><%=Separate(SalePrice)%>
<%							if trim("" & SalePrice)="" then
								response.write "›«ﬂ Ê— œ«—œ" 
							end if
%>							</A>
<%						end if
						
						RSS.moveNext
						if NOT RSS.eof then
							purchaseOrder =		RSS("id")
							Invoice =			RSS("Invoice")
							SalePrice =			RSS("SalePrice")
						else
							purchaseOrder =		-1
						end if
						response.write "<br>"
					Loop While lastPurchaseOrder = purchaseOrder AND NOT RSS.eof 
%>
					</TD>
				</TR>
<%				
			Loop
%>
			<TR bgcolor="#5555BB" height="2">
				<TD colspan=7></TD>
			</TR>
			<TR bgcolor="#eeeeee" height="2">
				<TD colspan='3' align='left'>„«·Ì«  »— «—“‘ «›“ÊœÂ: </TD>
				<TD colspan=4><INPUT TYPE="text" NAME="totalVat"  size=16 value="0" onchange='setPrice(this)'> —Ì«·</td>
			</TR>
			<TR bgcolor="#eeeeee" height="2">
				<TD colspan=3 align=left>Ã„⁄ ﬂ·:</TD>
				<TD colspan=4><INPUT  style="border:none; background:transparent; direction:LTR;text-align:right;" TYPE="text" NAME="totalPrice" readonly size=16 value="<%=totalprice%>">  —Ì«·</td>
			</TR>
			</TABLE><br>
			<%
		end if
		%>
		</TD>
	</TR>
	<TR disabled=true>
		<TD  align=left VALIGN=TOP>›«Ì·  ’ÊÌ— ›«ﬂ Ê— </TD>
		<TD align=right VALIGN=TOP>
			<INPUT TYPE="FILE" size=49 NAME="ImageFile">
		</TD>
	</TR>
	<TR>
		<TD  align=left VALIGN=TOP> Ê÷ÌÕ« </TD>
		<TD align=right VALIGN=TOP>
			<TEXTAREA NAME="comment" ROWS="7" COLS="48"></TEXTAREA>
		</TD>
	</TR>
	<TR>
		<TD  align=left VALIGN=TOP></TD>
		<TD align=center VALIGN=TOP>
			<br>
			<INPUT class=inputBut TYPE="submit" Name="Submit" Value="À»  ›«ﬂ Ê—" style="width:125px;" tabIndex="14"<%
			if venID="-1" then
				response.write " disabled "
			end if
			%>>
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
end if
%>
<!--#include file="tah.asp" -->
