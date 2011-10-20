<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ ›—Ê‘"
SubmenuItem=11
if not Auth("C" , 5) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; height:100%; background-color:#C3DBEB;}
</STYLE>
<%
if request("act")="show" then

	ON ERROR RESUME NEXT

		isA =			clng(request("isA"))

		FromDate =		sqlSafe(request("FromDate"))
		ToDate =		sqlSafe(request("ToDate"))

		ResultsInPage =	cint(request("ResultsInPage"))
		
		if FromDate="" AND ToDate="" then
			pageTitle="»Â ’Ê—  ﬂ«„·"
		elseif FromDate="" then
			pageTitle="«“ «» œ«  «  «—ÌŒ " & replace (ToDate,"/",".")
		elseif ToDate="" then
			pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  « «‰ Â« "
		else
			pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  «  «—ÌŒ " & replace (ToDate,"/",".")
		end if
	
		if ToDate = "" then ToDate = "9999/99/99"
		
		if request("ShowRemained")="on" OR request("ShowRemained")="True" then
			ShowRemained=true
		else
			ShowRemained=false
		end if
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "OtherReports.asp?errMsg=" & Server.URLEncode("Œÿ« œ— Ê—ÊœÌ.")
		end if
	ON ERROR GOTO 0
	Ord=request("Ord")

	select case Ord
	case "1":
		order="Receipts.ID"
	case "-1":
		order="Receipts.ID DESC"
	case "2":
		order="AccountTitle"
	case "-2":
		order="AccountTitle DESC"
	case "3":
		order="Customer"
	case "-3":
		order="Customer DESC"
	case "4":
		order="InvoiceId"
	case "-4":
		order="InvoiceId DESC"
	case "5"
		order="Number DESC"
	case "-5":
		order="Number DESC"
	case "6"
		order="ReceiptAmount DESC"
	case "-6":
		order="ReceiptAmount DESC"
	case "7"
		order="Amount DESC"
	case "-7":
		order="Amount DESC"
	case "8"
		order="EffectiveDate DESC"
	case "-8":
		order="EffectiveDate DESC"
	case else:
		order="Receipts.ID"
		Ord=1
	end select
	if ord<0 then
		style="background-color: #33CC99;"
		arrow="<br><span style='font-family:webdings'>6</span>"
	else
		style="background-color: #33CC99;"
		arrow="<br><span style='font-family:webdings'>5</span>"
	end if

	otherCriteria=""
	if isA = 1 then
		isAName = "»«  Ìﬂ «·›"
		otherCriteria = " AND (Invoices.IsA = 1) "
	elseif isA = 2 then
		isAName = "»œÊ‰  Ìﬂ «·›"
		otherCriteria = " AND (Invoices.IsA = 0) "
	end if

		
	'mySQL="SELECT * From (SELECT ARItems_1.GL,Receipts.ID AS ReceiptId, Accounts.AccountTitle, Invoices.Customer, Invoices.ID AS InvoiceId, Invoices.Number, ARItems.AmountOriginal AS ReceiptAmount, ARItemsRelations.Amount, Receipts.EffectiveDate FROM ARItems INNER JOIN Receipts ON ARItems.Link = Receipts.ID INNER JOIN ARItemsRelations ON ARItems.ID = ARItemsRelations.CreditARItem INNER JOIN ARItems ARItems_1 ON ARItemsRelations.DebitARItem = ARItems_1.ID INNER JOIN Invoices ON ARItems_1.Link = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Receipts.Voided=0) AND (Invoices.IsA = 1) AND (Receipts.EffectiveDate >= N'"& FromDate & "') AND (Receipts.EffectiveDate<= N'"& ToDate & "') AND (ARItems.Type = 2))drv ORDER BY " & order
	' Changed By Kid 851026 Adding the option of with/without isA tick
	mySQL="SELECT ARItems_1.GL,Receipts.ID AS ReceiptId, Accounts.AccountTitle, Invoices.Customer, Invoices.ID AS InvoiceId, Invoices.Number, ARItems.AmountOriginal AS ReceiptAmount, ARItemsRelations.Amount, Receipts.EffectiveDate FROM ARItems INNER JOIN Receipts ON ARItems.Link = Receipts.ID INNER JOIN ARItemsRelations ON ARItems.ID = ARItemsRelations.CreditARItem INNER JOIN ARItems ARItems_1 ON ARItemsRelations.DebitARItem = ARItems_1.ID INNER JOIN Invoices ON ARItems_1.Link = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Receipts.Voided=0) AND (Receipts.EffectiveDate >= N'"& FromDate & "') AND (Receipts.EffectiveDate<= N'"& ToDate & "') AND (ARItems.Type = 2) AND (ARItems_1.Type = 1) " & otherCriteria & " ORDER BY " & order

	'mySQLsum ="SELECT Count(ReceiptId) As ReceiptId,Sum(ReceiptAmount) As ReceiptAmount,Sum(Amount) As Amount From (SELECT ARItems_1.GL,Receipts.ID AS ReceiptId, Accounts.AccountTitle, Invoices.Customer, Invoices.ID AS InvoiceId, Invoices.Number, ARItems.AmountOriginal AS ReceiptAmount, ARItemsRelations.Amount, Receipts.EffectiveDate FROM ARItems INNER JOIN Receipts ON ARItems.Link = Receipts.ID INNER JOIN ARItemsRelations ON ARItems.ID = ARItemsRelations.CreditARItem INNER JOIN ARItems ARItems_1 ON ARItemsRelations.DebitARItem = ARItems_1.ID INNER JOIN Invoices ON ARItems_1.Link = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Receipts.Voided=0) AND (Invoices.IsA = 1) AND (Receipts.EffectiveDate >= N'"& FromDate & "') AND (Receipts.EffectiveDate<= N'"& ToDate & "') AND (ARItems.Type = 2))drv "
	' Changed By Kid 851026 Adding the option of with/without isA tick
	'mySQLsum ="SELECT Count(ReceiptId) As ReceiptId,Sum(ReceiptAmount) As ReceiptAmount,Sum(Amount) As Amount From (SELECT ARItems_1.GL,Receipts.ID AS ReceiptId, Accounts.AccountTitle, Invoices.Customer, Invoices.ID AS InvoiceId, Invoices.Number, ARItems.AmountOriginal AS ReceiptAmount, ARItemsRelations.Amount, Receipts.EffectiveDate FROM ARItems INNER JOIN Receipts ON ARItems.Link = Receipts.ID INNER JOIN ARItemsRelations ON ARItems.ID = ARItemsRelations.CreditARItem INNER JOIN ARItems ARItems_1 ON ARItemsRelations.DebitARItem = ARItems_1.ID INNER JOIN Invoices ON ARItems_1.Link = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Receipts.Voided=0) AND (Receipts.EffectiveDate >= N'"& FromDate & "') AND (Receipts.EffectiveDate<= N'"& ToDate & "') AND (ARItems.Type = 2) AND (ARItems_1.Type = 1) " & otherCriteria & ")drv "
	mySQLsum="SELECT COUNT(*) AS ReceiptCNT, SUM(ReceiptAmount) AS SumReceiptAmount, SUM(RelatedAmount) AS SumRelatedAmount FROM (SELECT Receipts.ID, Receipts.TotalAmount AS ReceiptAmount, SUM(ARItemsRelations.Amount) AS RelatedAmount FROM ARItems INNER JOIN Receipts ON ARItems.Link = Receipts.ID INNER JOIN ARItemsRelations ON ARItems.ID = ARItemsRelations.CreditARItem INNER JOIN ARItems ARItems_1 ON ARItemsRelations.DebitARItem = ARItems_1.ID INNER JOIN Invoices ON ARItems_1.Link = Invoices.ID WHERE (Receipts.Voided = 0) AND (Receipts.EffectiveDate >= N'"& FromDate & "') AND (Receipts.EffectiveDate <= N'"& ToDate & "') AND (ARItems.Type = 2) AND (ARItems_1.Type = 1) " & otherCriteria & " GROUP BY Receipts.ID, Receipts.TotalAmount) DRV"


	Set Rs2 = Conn.Execute(mySQLsum)
	if (Not Rs2.EOF) Then
	ReceiptCNT = Rs2("ReceiptCNT")
	SumReceitpAmount = Rs2("SumReceiptAmount")
	SumRelatedAmount = Rs2("SumRelatedAmount")

%>
	<br><br>
	<table class="CustTable" cellspacing='1' style='width:90%;' align='center'>
		<tr>
			<td colspan="8" class="CusTableHeader" style="text-align:right;height:35;">ÕÃ„ —”Ìœ Â«Ì «·› <%=isAName %> &nbsp; <%=GroupName%>  (<%=pageTitle%>)</td>
		</tr>
		<tr>
			<td class="CusTD3" width=100> ⁄œ«œ —”Ìœ Â«</td>
			<td class="CusTD3">Ã„⁄ „»·€</td>
			<td class="CusTD3">Ã„⁄ Å«” ‘œÂ</td>
		</tr>
		<tr bgcolor="white">
			<TD dir=LTR align=center><%=Separate(ReceiptCNT)%></TD>
			<TD dir=LTR align=center><%=Separate(SumReceitpAmount)%></TD>
			<TD dir=LTR align=center><%=Separate(SumRelatedAmount)%></TD>
		</tr>
	</table>
	<br><br>
<%
	end if 
	Rs2.Close
%>

	<table class="CustTable" cellspacing='1' style='width:90%;' align='center'>
		<tr>
			<td colspan="10" class="CusTableHeader" style="text-align:right;height:35;">—”Ìœ Â«Ì «·› (<%=pageTitle%>)</td>
		</tr>
<%
		Set RS1 = Server.CreateObject("ADODB.Recordset")

		PageSize = ResultsInPage
		RS1.PageSize = PageSize 

		RS1.CursorLocation=3 'in ADOVBS_INC adUseClient=3
		RS1.Open mySQL ,Conn,3
		TotalPages = RS1.PageCount

		CurrentPage=1

		if isnumeric(Request.QueryString("p")) then
			pp=clng(Request.QueryString("p"))
			if pp <= TotalPages AND pp > 0 then
				CurrentPage = pp
			end if
		end if

		if not RS1.eof then
			RS1.AbsolutePage=CurrentPage
		end if


		if RS1.eof then
%>			<tr>
				<td colspan="10" class="CusTD3" style='direction:RTL;'>ÂÌç ›«ﬂ Ê—Ì ÅÌœ« ‰‘œ.</td>
			</tr>
<%		else
%>
			<tr>
				<td class="CusTD3" >#</td>
				<td class="CusTD3" onclick='go2Page(1,1);' style="cursor:hand; <%if abs(ord)=1 then response.write style%>" style='direction:RTL;'># —”Ìœ <%if abs(ord)=1 then response.write arrow%></td>
				<td class="CusTD3" onclick='go2Page(1,2);' style="cursor:hand; <%if abs(ord)=2 then response.write style%>">⁄‰Ê«‰ Õ”«» <%if abs(ord)=2 then response.write arrow%></td>
				<td class="CusTD3" onclick='go2Page(1,3);' style="cursor:hand; <%if abs(ord)=3 then response.write style%>">‘„«—Â Õ”«» <%if abs(ord)=3 then response.write arrow%></td>
				<td class="CusTD3" onclick='go2Page(1,4);' style="cursor:hand; <%if abs(ord)=4 then response.write style%>" style='direction:RTL;'># ›«ﬂ Ê— <%if abs(ord)=4 then response.write arrow%></td>
				<td class="CusTD3" onclick='go2Page(1,5);' style="cursor:hand; <%if abs(ord)=5 then response.write style%>" >‘„«—Â «·› <%if abs(ord)=5 then response.write arrow%></td>
				<td class="CusTD3" >GL</td>
				<td class="CusTD3" onclick='go2Page(1,6);' style="cursor:hand; <%if abs(ord)=6 then response.write style%>" >„»·€ <%if abs(ord)=6 then response.write arrow%></td>
				<td class="CusTD3" onclick='go2Page(1,7);' style="cursor:hand; <%if abs(ord)=7 then response.write style%>" >Å«” ‘œÂ <%if abs(ord)=7 then response.write arrow%></td>
				<td class="CusTD3" onclick='go2Page(1,8);' style="cursor:hand; <%if abs(ord)=8 then response.write style%>" > «—ÌŒ ’œÊ— <%if abs(ord)=8 then response.write arrow%></td>
			</tr>
<%			tmpCounter=(CurrentPage - 1) * PageSize
			SumAmount = 0
			SumReceiptAmount = 0
			LastReceipt = 0

			Do while NOT RS1.eof AND (RS1.AbsolutePage = CurrentPage)

				Number=	RS1("Number")

				ReceiptID = RS1("ReceiptID")

				Amount = cdbl(RS1("Amount"))

				if ReceiptID = LastReceipt then
					ReceiptAmount = 0
					ReceiptAmountText = ""

				else
					ReceiptAmount = cdbl(RS1("ReceiptAmount"))
					ReceiptAmountText = Separate(ReceiptAmount)

					tmpCounter = tmpCounter + 1

				end if

				SumReceiptAmount = SumReceiptAmount + ReceiptAmount
				SumAmount = SumAmount + Amount

				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 

%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showReceipt&receipt=<%=RS1("ReceiptID")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD style="height:30px;"><%=ReceiptID%></TD>
					<TD><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD><%=RS1("Customer")%>&nbsp;</TD>
					<TD style="height:30px;"><%=RS1("InvoiceId")%></TD>
					<TD dir="LTR" align='right' <%=tmpColor3%> ><%=Number%>&nbsp;</TD>
					<TD dir="LTR" align='right' <%=tmpColor3%> ><%=RS1("GL")%></TD>
					<TD><%=ReceiptAmountText%>&nbsp;</TD>
					<TD><%=Separate(Amount)%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("EffectiveDate")%>&nbsp;</TD>
				</TR>
<%
				LastReceipt = ReceiptID 
				RS1.moveNext
			Loop
%>
			<tr dir=rtl  bgcolor=#CCCCCC >
				<td  colspan=7 align=left >„Ã„Ê⁄ &nbsp;</td>
				<td ><B><%=Separate(SumReceiptAmount)%></B></td>
				<td ><B><%=Separate(SumAmount)%></B></td>
				<td ></td>
			</tr>

<%
			if ToDate="9999/99/99" then ToDate="" 

			if TotalPages > 1 then
				pageCols=20
%>			
				<TR class="RepTableTitle">
					<TD bgcolor='#33AACC' height="30" colspan="10">
					<table width=100% cellspacing=0 style="cursor:hand;color:#444444">
					<tr>
						<td style="height:25;border-bottom:1 solid black;" colspan=<%=pagecols%>>
							<b>’›ÕÂ <%=CurrentPage%> «“ <%=TotalPages%></b>
							&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>,0);">’›ÕÂ »⁄œ &gt;</a>
						</td>
					</tr>
					<tr>
<%					for i=1 to TotalPages 
						if i = CurrentPage then 
%>							<td style="color:black;"><b>[<%=i%>]</b></td>
<%						else
%>							<td onclick="go2Page(<%=i%>,0);"><%=i%></td>
<%						end if 
						if i mod pageCols = 0 then response.write "</tr><tr>" 
					next 

%>					</tr>
					</table>

					<SCRIPT LANGUAGE="JavaScript">
					<!--
					function go2Page(p,ord) {
						if(ord==0){
						ord=<%=Ord%>;
						}
						else if(ord==<%=Ord%>){
						ord= 0-ord;
						}
						window.location="?act=show&ResultsInPage=<%=ResultsInPage%>&p="+p+"&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>&isA=<%=isA%>&ShowRemained=<%=ShowRemained%>&Ord="+escape(ord);
					}
					//-->
					</SCRIPT>

					</TD>
				</TR>
<%			end if
		end if
		RS1.close
		Set RS1 = Nothing
%>
	</table>
	<br>
<%
end if%>
<!--#include file="tah.asp" -->