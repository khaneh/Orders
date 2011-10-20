<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= "ê“«—‘ Õ”«»"
SubmenuItem=5
if not Auth(7 , 5) then NotAllowdToViewThisPage()

fixSys = "AP"
SA_IsVendor = 1
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {padding:5;border:1pt solid gray;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTableTitle {background-color: #CCCCFF; text-align: center; font-weight:bold; height:50;}
	.RepTableHeader {background-color: #BBBBBB; text-align: center; font-weight:bold;}
	.RepTableFooter {background-color: #BBBBBB; direction: LTR; }
	.RepTR1 {background-color: #DDDDDD;}
	.RepTR2 {background-color: #FFFFFF;}
	.RepDescSpan {overflow:auto;border:none; width:250px; height:23px; font-size:7pt;}
	.RepGenInput { font-family:tahoma; font-size: 8pt; border: 1 solid black; direction:LTR; width:70px; height:19px;}
</STYLE>
<style>
	.InvTable { font-size: 9pt;}
	.InvRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.InvHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.InvRowInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0FFF0; text-align:right;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right; direction: RTL;}
	.InvGenInput { font-family:tahoma; font-size: 9pt; border: none; }
</style>
<style>
	.RcpTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; }
	.RcpMainTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; background-color: #558855; text-align:right; direction: RTL;}
	.RcpMainTableTH { background-color: #C3C300;}
	.RcpMainTableTR { background-color: #CCCC88; border: 0; }
	.RcpRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.RcpRowInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.RcpHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.RcpHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.RcpHeadInput3 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #D0E0FF; text-align:right; direction: RTL;}
	.RcpGenInput { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction: LTR;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<style>
	.MmoTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; direction: LTR;}
	.MmoMainTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; background-color: #558855; text-align:right; direction: RTL;}
	.MmoMainTableTH { background-color: #C3C300;}
	.MmoMainTableTR { background-color: #CCCC88; border: 0; }
	.MmoRowInput { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.MmoRowInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.MmoHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.MmoHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.MmoHeadInput3 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #D0E0FF; text-align:right; direction: RTL;}
	.MmoGenInput { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction: LTR;}
</style>


<%
	CustomerID=request("selectedCustomer")
	'sys = request("sys")
	'if sys = "" then sys = "AR"
	reason = request("Reason")
	if reason = "" then reason = 1

	mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
	else
		Sys=			RS1("Acron")
		firstGLAccount=	RS1("GLAccount")
	end if
	RS1.close

	if not fixSys = "-" then sys = fixSys

	mySQL="SELECT AccountTitle FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)
	customerName=RS1("AccountTitle")
	RS1.close
	
	EndDate=sqlSafe(request("EndDate"))
	StartDate=sqlSafe(request("StartDate"))
	if StartDate="" then
		StartDate=session("OpenGLStartDate") 'yearBeginnig=left(shamsiToday(),4)&"/01/01"
	end if
	if EndDate="" then
		if session("OpenGLEndDate") > shamsiToday() then
			EndDate=shamsiToday()
		else
			EndDate=session("OpenGLEndDate")
		end if
	end if

	nextYear=cint(left(StartDate,4)) + 1 
	prevYear=cint(left(StartDate,4)) - 1 

	nextStartDate=	nextYear & "/01/01" 'right(StartDate,6)
	nextEndDate=	nextYear & "/12/30" 'right(EndDate,6)
	 
	prevStartDate=	prevYear & "/01/01" 'right(StartDate,6)
	prevEndDate=	prevYear & "/12/30" 'right(EndDate,6)

	nextYear=right(nextYear,2)
	prevYear=right(prevYear,2)
%>
	<br>

	<TABLE class="RepTable" width='90%' align='center'>
	<TR>
		<TD colspan=8 dir='rtl' align='center'>
		</TD>
	</TR>
	<TR>
		<TD class="RepTableTitle" colspan=8 dir='rtl' align='center'>
			<br>
			<FORM METHOD=POST ACTION="?act=show&sys=<%=sys%>&selectedCustomer=<%=CustomerID%>" ID="dateForm">
			ê“«—‘ Œ—Ìœ «·› <A target="_blank" HREF="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=CustomerID%>"><%=customerName%> [<%=CustomerID%>]</A>
			<br>
			<br>
			«“  «—ÌŒ <INPUT class="RepGenInput" TYPE="text" NAME="StartDate" Value="<%=StartDate%>" OnBlur="return acceptDate(this);">  « <INPUT class="RepGenInput" TYPE="text" NAME="EndDate" Value="<%=EndDate%>"OnBlur="return acceptDate(this);"> <INPUT Class="GenButton" TYPE="button" Value=" ‰„«Ì‘ "onclick="if(acceptDate(document.getElementsByName('StartDate')[0]) && acceptDate(document.getElementsByName('EndDate')[0])) document.getElementById('dateForm').submit()">
			<BR><BR>
			<INPUT Class="GenButton" TYPE="button" Value=" <- <%=prevYear%> " onclick="window.location='?act=show&reason=<%=reason%>&sys=<%=sys%>&selectedCustomer=<%=CustomerID%>&startDate=<%=prevStartDate%>&endDate=<%=prevEndDate%>';">
			<INPUT Class="GenButton" TYPE="button" Value=" <%=nextYear%> -> " onclick="window.location='?act=show&reason=<%=reason%>&sys=<%=sys%>&selectedCustomer=<%=CustomerID%>&startDate=<%=nextStartDate%>&endDate=<%=nextEndDate%>';">
			</FORM>

		</TD>
	</TR>
	<TR class="RepTableHeader">
		<TD>#</TD>
		<TD> «—ÌŒ</TD>
		<TD>›«ﬂ Ê— Œ—Ìœ</TD>
		<TD>›«ﬂ Ê— ›—Ê‘</TD>
		<TD> Ê÷ÌÕ« </TD>
		<TD>„»·€</TD>
	</TR>
<%

	'mySQL="SELECT Vouchers.*, Invoices.Number AS InvoiceNumber, Invoices.ID AS InvoiceID FROM VoucherLines INNER JOIN Vouchers ON VoucherLines.Voucher_ID = Vouchers.id INNER JOIN PurchaseOrders ON VoucherLines.RelatedPurchaseOrderID = PurchaseOrders.ID INNER JOIN PurchaseRequestOrderRelations ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID INNER JOIN PurchaseRequests ON PurchaseRequestOrderRelations.Req_ID = PurchaseRequests.ID INNER JOIN Orders ON PurchaseRequests.Order_ID = Orders.ID INNER JOIN InvoiceOrderRelations ON Orders.ID = InvoiceOrderRelations.[Order] INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE Invoices.isA = 1 and Invoices.Voided = 0 and Vouchers.Verified = 1 and Vouchers.vendorID = " & CustomerID
	'Changed By Kid 830625

	mySQL="SELECT DISTINCT * FROM Vouchers INNER JOIN (SELECT Invoices.Number AS InvoiceNumber, Invoices.ID AS InvoiceID, VoucherLines.Voucher_ID FROM VoucherLines INNER JOIN PurchaseOrders ON VoucherLines.RelatedPurchaseOrderID = PurchaseOrders.ID INNER JOIN PurchaseRequestOrderRelations ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID INNER JOIN PurchaseRequests ON PurchaseRequestOrderRelations.Req_ID = PurchaseRequests.ID INNER JOIN Orders ON PurchaseRequests.Order_ID = Orders.ID INNER JOIN InvoiceOrderRelations ON Orders.ID = InvoiceOrderRelations.[Order] INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (Invoices.IsA = 1) AND (Invoices.Voided = 0) AND (Invoices.Issued = 1)) VoucherLine_AInvoice ON  VoucherLine_AInvoice.Voucher_ID = Vouchers.id WHERE (Vouchers.Verified = 1) AND (Vouchers.VendorID = " & CustomerID &") AND (Vouchers.EffectiveDate >= '" & StartDate &"') AND (Vouchers.EffectiveDate <= '" & EndDate &"') ORDER BY Vouchers.EffectiveDate, Vouchers.ID"

	Set RS1 = conn.execute(mySQL)
	
	Remained = 0
	Totalcredit = 0
	TotalDebit = 0
	tempCounter = 0
	if Not (RS1.EOF) then
		While Not (RS1.EOF)
			tempCounter=tempCounter+1
			
			sourceLink="<a href='AccountReport.asp?act=showVoucher&voucher="& RS1("id") & "' target='_blank'>" & replace(RS1("Title"),"/",".") & "</a>"
			sourceLink2="<a href='../AR/AccountReport.asp?act=showInvoice&invoice="& RS1("InvoiceID") & "' target='_blank'>" & RS1("InvoiceID") & "</a>"

			if not isnull(RS1("comment")) then 
				Description = replace(RS1("comment"),chr(13),"<br>") + " - ›«ﬂ Ê— " & RS1("InvoiceNumber")
				Description = replace(Description,"/",".")
			else
				Description = ""
			end if

			if RS1("Voided") then 

%>			<TR bgcolor=#FFEEEE style='color:#999999'>
				<td> # <%tempCounter = tempCounter -1 %></td>
				<td width=60 dir='LTR' align='right'><%=RS1("EffectiveDate")%></td>
				<td width=140><%=sourceLink%></td>
				<td width=70><%=sourceLink2%></td>
				<td width=200><span ><%=Description%></span></td>
				<td dir='LTR'>
					<div style="position:absolute;width:650;"><hr style="color:red;"></div>
					<div align=right><%=Separate(RS1("TotalPrice"))%></div>
				</td>
			</TR>
<%
			else
				TotalDebit = TotalDebit + cdbl(Debit)
				Totalcredit = Totalcredit + cdbl(Credit)
				Remained = Remained + Cdbl(Credit) - Cdbl(Debit)
%>			<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
				<td><%=tempCounter %></td>
				<td width=60 dir='LTR' align='right'><%=RS1("EffectiveDate")%></td>
				<td width=140><%=sourceLink%></td>
				<td width=70><%=sourceLink2%></td>
				<td width=200><span ><%=Description%></span></td>
				<td dir='LTR' align='right'><%=Separate(RS1("TotalPrice"))%></td>
			</TR>
<%
			end if
			total = total + Cdbl(RS1("TotalPrice"))
			RS1.MoveNext
		Wend
		if Remained>=0 then
			remainedColor="green"
		else
			remainedColor="red"
		end if

	else
		total= " - "
%>
		<TR class="RepTR1">
			<TD colspan=6 align=center>ÂÌç</TD>
		</TR>
<%	
	end if
%>
		<TR>
			<TD class="RepTableFooter" colspan='5'>&nbsp;&nbsp; : Ã„⁄</span></td>
			<TD class="RepTableFooter" align='right'><%=Separate(total)%></td>
		</TR>
	</TABLE>
	<br>

<!--#include file="tah.asp" -->