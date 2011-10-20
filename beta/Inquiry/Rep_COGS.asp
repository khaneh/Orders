<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle="ê“«—‘ ﬁÌ„   „«„ ‘œÂ"
SubmenuItem=11
if not Auth("C" , 7) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {border:1pt solid white;vertical-align:top;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTable2 th {font-size:9pt; background-color:#666699;height:25px;}
	.RepTable2 input {font-family:tahoma; font-size:9pt; border:1 solid black;}
</STYLE>
<BR>
<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Search Form
'-----------------------------------------------------------------------------------------------------
if request("act")="MoeenRep" OR request("act")="show" then

	ON ERROR RESUME NEXT

		ResultsInPage =	cint(request("ResultsInPage"))

		FromDate =		sqlSafe(request("FromDate"))
		ToDate =		sqlSafe(request("ToDate"))
		
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
		
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "OtherReports.asp?errMsg=" & Server.URLEncode("Œÿ« œ— Ê—ÊœÌ.")
		end if
	ON ERROR GOTO 0

	Ord=request("Ord")

	select case Ord
	case "1":
		order="Invoices.ID"
	case "-1":
		order="Invoices.ID DESC"
	case "2":
		order="AccountTitle"
	case "-2":
		order="AccountTitle DESC"
	case "3":
		order="InventoryCost"
	case "-3":
		order="InventoryCost DESC"
	case "4":
		order="OutServiceCost"
	case "-4":
		order="OutServiceCost DESC"
	case "5":
		order="TotalReceivable"
	case "-5":
		order="TotalReceivable DESC"
	case "6":
		order="Margin"
	case "-6":
		order="Margin DESC"
	case else:
		order="Invoices.ID"
		Ord=1
	end select

%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	if (window.XMLHttpRequest) {
		var objHTTP=new XMLHttpRequest();
	} else if (window.ActiveXObject) {
		var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
	}
	var detailPosition=0;
	function showDetails(index){
		if (detailPosition==0){
			index=index+1;
		}
		else{
			COGStbl.removeChild(COGStbl.getElementsByTagName("tr")[detailPosition])
		}
		if (detailPosition < index)
			index=index-1;
		

		detailPosition=index+1;

		COGStbl=document.getElementById('COGS').getElementsByTagName("TBODY")[0];
		inv = COGStbl.getElementsByTagName("tr")[index].getElementsByTagName("td")[0].innerText;

		objHTTP.Open('GET','Rep_COGS_mini.asp?id='+inv,false);
		objHTTP.Send()

		newRow=document.createElement("tr");
		newRow.setAttribute("bgColor", '#f0f0f0');

		tempTD=document.createElement("td");
		tempTD.colSpan=6;
		tempTD.setAttribute("align", 'center');
		tempTD.innerHTML=objHTTP.responseText;
		newRow.appendChild(tempTD);

		COGStbl.insertBefore(newRow,COGStbl.getElementsByTagName("tr")[detailPosition]);

		return;
	}
	//-->
	</SCRIPT>
	<TABLE id="COGS" dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
<%
	if ord<0 then
		style="background-color: #33CC99;"
		arrow="<br><span style='font-family:webdings'>6 6 6</span>"
	else
		style="background-color: #33CC99;"
		arrow="<br><span style='font-family:webdings'>5 5 5</span>"
	end if
%>
	<TR bgcolor="eeeeee" style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
		<TD width=50 onclick='go2Page(1,1);' style="<%if abs(ord)=1 then response.write style%>">›«ﬂ Ê—			<%if abs(ord)=1 then response.write arrow%></TD>
		<TD width='*' onclick='go2Page(1,2);' style="<%if abs(ord)=2 then response.write style%>">⁄‰Ê«‰ Õ”«»		<%if abs(ord)=2 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-3);' style="<%if abs(ord)=3 then response.write style%>">Â“Ì‰Â «‰»«—		<%if abs(ord)=3 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-4);' style="<%if abs(ord)=4 then response.write style%>">Â“Ì‰Â »Ì—Ê‰Ì	<%if abs(ord)=4 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-5);' style="<%if abs(ord)=5 then response.write style%>">ﬁ«»· Å—œ«Œ 	<%if abs(ord)=5 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-6);' style="<%if abs(ord)=6 then response.write style%>">”Êœ	<%if abs(ord)=6 then response.write arrow%></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=6 height=2 bgcolor=0></TD>
	</TR>
<%		
	SumCredit=0
	SumDebit=0
	SumCreditRemained=0
	SumDebitRemained=0
	tmpCounter=0

	mySQL="SELECT Invoices.ID, InventoryCost.SumPrice AS InventoryCost, OutServiceCost.SumPrice AS OutServiceCost, Invoices.TotalReceivable, Invoices.TotalReceivable - ISNULL(NULLIF (OutServiceCost.SumPrice, - 1), 0) - ISNULL(NULLIF (InventoryCost.SumPrice, - 1), 0) AS Margin, Accounts.AccountTitle FROM Invoices INNER JOIN Accounts ON Invoices.Customer = Accounts.ID LEFT OUTER JOIN (SELECT InvoiceOrderRelations.Invoice, CASE WHEN COUNT(*) = COUNT(VoucherLines.price) THEN SUM(VoucherLines.price) ELSE - 1 END AS SumPrice FROM Vouchers INNER JOIN VoucherLines ON Vouchers.id = VoucherLines.Voucher_ID RIGHT OUTER JOIN InvoiceOrderRelations INNER JOIN PurchaseOrders INNER JOIN PurchaseRequestOrderRelations INNER JOIN PurchaseRequests ON PurchaseRequestOrderRelations.Req_ID = PurchaseRequests.ID ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID ON InvoiceOrderRelations.[Order] = PurchaseRequests.Order_ID ON Vouchers.Voided = 0 AND VoucherLines.RelatedPurchaseOrderID = PurchaseOrders.ID WHERE (PurchaseRequests.Status <> N'del') AND (PurchaseOrders.Status <> N'CANCEL') GROUP BY InvoiceOrderRelations.Invoice) OutServiceCost ON Invoices.ID = OutServiceCost.Invoice LEFT OUTER JOIN (SELECT InvoiceOrderRelations.Invoice, CASE WHEN COUNT(*) = COUNT(InventoryItemsUnitPrice.UnitPrice) THEN SUM(InventoryPickuplistItems.Qtty * InventoryItemsUnitPrice.UnitPrice) ELSE - 1 END AS SumPrice FROM InventoryPickuplistItems INNER JOIN InventoryPickuplists ON InventoryPickuplistItems.pickupListID = InventoryPickuplists.id INNER JOIN InvoiceOrderRelations ON InventoryPickuplistItems.Order_ID = InvoiceOrderRelations.[Order] LEFT OUTER JOIN InventoryItemsUnitPrice ON InventoryPickuplists.CreationDate >= InventoryItemsUnitPrice.StartDate AND InventoryPickuplists.CreationDate <= InventoryItemsUnitPrice.EndDate AND InventoryPickuplistItems.ItemID = InventoryItemsUnitPrice.InventoryItem WHERE (InventoryPickuplistItems.CustomerHaveInvItem = 0) AND (NOT (InventoryPickuplists.Status = N'del')) GROUP BY InvoiceOrderRelations.Invoice) InventoryCost ON Invoices.ID = InventoryCost.Invoice WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "' AND Invoices.IssuedDate <= N'"& ToDate & "') ORDER BY "& order
'response.write "<div dir=LTR>" & mySQL & "</div><br>"

	Set rs=Server.CreateObject("ADODB.Recordset")'Conn.Execute(mySQL)

	PageSize = 50
	rs.PageSize = PageSize

	rs.CursorLocation=3 'in ADOVBS_INC adUseClient=3
	rs.Open mySQL ,Conn,3
	TotalPages = rs.PageCount

	CurrentPage=1

	if isnumeric(Request.QueryString("p")) then
		pp=clng(Request.QueryString("p"))
		if pp <= TotalPages AND pp > 0 then
			CurrentPage = pp
		end if
	end if

	if not rs.eof then
		rs.AbsolutePage=CurrentPage
	end if

	if rs.eof then
%>		<tr>
			<td bgcolor="#BBBBBB" height="30" colspan="7" align=center><b>ÂÌç .</b></td>
		</tr>
<%	else
		Do While NOT rs.eof AND (rs.AbsolutePage = CurrentPage)
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

			marginInvalid=false

			if isnull(rs("InventoryCost"))  then 
				InventoryCost	=	""
			elseif rs("InventoryCost")="-1" then
				InventoryCost	=	"<font color='red'> ‰«ﬁ’ </font>"
				marginInvalid=true
			else
				InventoryCost	=	cdbl(rs("InventoryCost"))
			end if

			if isnull(rs("OutServiceCost"))  then 
				OutServiceCost	=	""
			elseif rs("OutServiceCost")="-1" then
				OutServiceCost	=	"<font color='red'> ‰«ﬁ’ </font>"
				marginInvalid=true
			else
				OutServiceCost	=	cdbl(rs("OutServiceCost"))
			end if
			TotalReceivable	=	cdbl(rs("TotalReceivable"))
			Margin			=	cdbl(rs("Margin"))

			if marginInvalid then
				MarginColor = "red"
			else
				MarginColor = ""
			end if

			

	%>
			<TR bgcolor="<%=tmpColor%>" onclick="showDetails(this.rowIndex);" style="cursor:pointer;">
				<TD dir=ltr align=right><A target="_blank" onclick="event.cancelBubble=true;" HREF="AccountReport.asp?act=showInvoice&invoice=<%=rs("ID")%>"><%=rs("ID")%></A></TD>
				<TD><%=rs("AccountTitle")%></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(InventoryCost)%></span></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(OutServiceCost)%></span></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(TotalReceivable)%></span></TD>
				<TD dir=ltr align=right><span dir=ltr style="color:<%=MarginColor%>"><%=Separate(Margin)%></span></TD>
			</TR>
			 
	<% 
		rs.moveNext
		Loop

		if TotalPages > 1 then
			pageCols=20
%>			
			<TR bgcolor="eeeeee" >
				<TD colspan=6 height=2 bgcolor=0></TD>
			</TR>
			<TR class="RepTableTitle">
				<TD bgcolor="#CCCCEE" height="30" colspan="6">
				<table width=100% cellspacing=0 style="cursor:hand;color:gray;">
				<tr>
					<td style="height:25;border-bottom:1 solid black;" colspan=<%=pagecols%>>
						<b>’›ÕÂ <%=CurrentPage%> «“ <%=TotalPages%></b>
						&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>,0);">’›ÕÂ »⁄œ &gt;</a>
					</td>
				</tr>
				<tr>
<%				for i=1 to TotalPages 
					if i = CurrentPage then 
%>						<td style="color:black;"><b>[<%=i%>]</b></td>
<%					else
%>						<td onclick="go2Page(<%=i%>,0);"><%=i%></td>
<%					end if 
					if i mod pageCols = 0 then response.write "</tr><tr>" 
				next 

%>				</tr>
				</table>
				</TD>
			</TR>
<%		end if
%>
		</TABLE><br>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		function go2Page(p,ord) {
			if(ord==0){
				ord=<%=Ord%>;
			}
			else if(ord==<%=Ord%>){
				ord= 0-ord;
			}
			str='?act=MoeenRep&GLAccount='+escape('<%=GLAccount%>')+'&FromDate='+escape('<%=FromDate%>')+'&ToDate='+escape('<%=ToDate%>')+'&FromTafsil='+escape('<%=FromTafsil%>')+'&ToTafsil='+escape('<%=ToTafsil%>')+'&Ord='+escape(ord)+'&p='+escape(p) //+'& ='+escape(' ')+'& ='+escape(' ')+'& ='+escape(' ')
			window.location=str;
		}
		//-->
		</SCRIPT>
<%
	end if
end if


%>

<!--#include file="tah.asp" -->
