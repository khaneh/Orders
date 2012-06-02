<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'	AR (6)
PageTitle="«’·«Õ ›«ﬂ Ê— "
SubmenuItem=1
'if not Auth(6 , 1) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
function ShowErrorMessage(msg)
	response.write "<table align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> Œÿ« ! <br>"& msg & "<br></td></tr></table><br>"
end function

function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/TraceOrder.asp?act=show&order="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

function Link2TraceQuote(QuoteNo)
	Link2TraceQuote = "<A HREF='../order/Inquiry.asp?act=show&quote="& QuoteNo & "' target='_balnk'>"& QuoteNo & "</A>"
end function

%>
<style>
	Table { font-size: 9pt;}
	.InvRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.InvHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.InvRowInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0FFF0; text-align:right;}
	.InvRowInput4 { font-family:tahoma; font-size: 9pt; border: none; background-color: #FFD3A8; direction:LTR; text-align:right;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.InvHeadInput4 { font-family:tahoma; font-size: 9pt; border: none; background-color: #FF9900; text-align:center;}
	.InvGenInput  { font-family:tahoma; font-size: 9pt; border: none; }
	.InvGenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;
var currentRow=0;
//-->
</SCRIPT>
<%
if request("act")="search" then
	if isnumeric(request("order")) then
		OrderID = clng(request("order"))
		mySQL="SELECT InvoiceOrderRelations.Invoice FROM InvoiceOrderRelations INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (InvoiceOrderRelations.[Order] = '"& OrderID & "') AND (Invoices.IsReverse = 0) AND (Invoices.Voided = 0)"
		Set RS1=Conn.Execute(mySQL)
		if RS1.eof then
			Conn.close
			response.redirect "?errmsg=" & Server.URLEncode("«Ì‰ ”›«—‘ ﬁ»·« ›«ﬂ Ê— ‰‘œÂ «” .")
		else
			theInvoice=RS1("Invoice")
			Conn.close
			response.redirect "?act=editInvoice&invoice=" & theInvoice
		end if
	elseif isnumeric(request("invoice")) then
		response.redirect "?act=editInvoice&invoice=" & request("invoice")
	else
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ”›«—‘ ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
	end if
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------- Edit Invoice
'-----------------------------------------------------------------------------------------------------
elseif request("act")="editInvoice" then
	if isnumeric(request("invoice")) then
		InvoiceID=clng(request("invoice"))
		mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("›«ﬂ Ê— ÅÌœ« ‰‘œ.")
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
	end if

	customerID=		RS1("Customer")
	creationDate=	RS1("CreatedDate")
	IssuedDate=		RS1("IssuedDate")
	totalPrice=		cdbl(RS1("totalPrice"))
	totalDiscount=	cdbl(RS1("totalDiscount"))
	totalReverse=	cdbl(RS1("totalReverse"))
	totalVat =		cdbl(RS1("totalVat"))
	Voided=			RS1("Voided")
	Issued=			RS1("Issued")
	Approved=		RS1("Approved")
	isReverse=		RS1("IsReverse")
	IsA=			RS1("IsA")
	InvoiceNo=		RS1("Number")

	mySQL="SELECT ID,AccountTitle FROM Accounts WHERE (ID='"& customerID & "')"
	Set RS1 = conn.Execute(mySQL)
	AccountNo=RS1("ID")
	customerName=RS1("AccountTitle")

	RS1.close

	if isReverse then
		'Check for permission for EDITTING Rev. Invoice
		if not Auth(6 , 5) then NotAllowdToViewThisPage()
		itemTypeName="›«ﬂ Ê— »—ê‘ "
		HeaderColor="#FF9900"
	else
		'Check for permission for EDITTING Invoice
		if not Auth(6 , 3) then NotAllowdToViewThisPage()
		itemTypeName="›«ﬂ Ê—"
		HeaderColor="#C3C300"
	end if

	if Voided then
		Conn.close
		response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— »«ÿ· ‘œÂ «” .")
	elseif Issued then
		if Auth(6 , "A") then
			' Has the Priviledge to change the Invoice
			response.write "<BR>"
			if Auth(6,"N") then 
				call showAlert ("«Ì‰ ›«ﬂ Ê— ’«œ— ‘œÂ «” .<br>Â—ç‰œ ﬂÂ ‘„« «Ã«“Â œ«—Ìœ «Ì‰ ›«ﬂ Ê— —« œç«—  €ÌÌ—«  ﬂ·Ì ﬂ‰Ìœ<br>„”Ê·Ì  „‘ﬂ·«  «Õ „«·Ì —« „Ì Å–Ì—Ìœø",CONST_MSG_INFORM)
			else 
				call showAlert ("«Ì‰ ›«ﬂ Ê— ’«œ— ‘œÂ «” .<br>Â—ç‰œ ﬂÂ ‘„« «Ã«“Â œ«—Ìœ ﬂÂ «Ì‰ ›«ﬂ Ê— —«  €ÌÌ— »œÂÌœ<br>„”Ê·Ì  „‘ﬂ·«  «Õ „«·Ì —« „Ì Å–Ì—Ìœø",CONST_MSG_INFORM) 
			end if
		else
			Conn.close
			response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ’«œ— ‘œÂ «” .")
		end if
	end if 
%>
<!-- Ê—Êœ «ÿ·«⁄«  ›«ﬂ Ê— -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<input type="hidden" name="VatRate" id="VatRate" value="<%=session("VatRate")%>">
	<FORM METHOD=POST ACTION="?act=submitEdit">
		<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<tr bgcolor='<%=HeaderColor%>'>
			<td colspan='2'>
				<TABLE width='100%'>
				<TR>
					<TD align="left" >‘„«—Â <%=itemTypeName%>:</TD>
					<TD align="right" width='15%'>&nbsp;<INPUT readonly class="InvGenInput" NAME="InvoiceID" value="<%=InvoiceID%>" style="direction:ltr" TYPE="text" maxlength="10" size="10"></TD>
				</TR>
				</TABLE></td>
			</tr>
			<tr bgcolor='<%=HeaderColor%>'>
			<td colspan="2">
				<TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL">
				<TR>
					<TD><table>
						<tr>
							<td align="left">Õ”«»:</td>
							<td align="right">
								<span id="customer"><%' after any changes in this span "./Customers.asp" must be revised%>
									<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><span><%=CustomerName%></span>.
								</span></td>
							<td><INPUT class="InvGenButton" TYPE="button" value=" €ÌÌ—" onClick="selectCustomer();"></td>
						</tr>
						</table></TD>
					<TD align="left"><table>
						<tr>
							<td align="left"> «—ÌŒ ’œÊ—:</td>
							<td dir="LTR">
								<INPUT class="InvGenInput" NAME="issueDate" TYPE="text" maxlength="10" size="10" value="<%=IssuedDate%>"  onblur="acceptDate(this)"></td>
							<td dir="RTL"><%="ø‘‰»Â"%></td>
						</tr>
						</table></TD>
				</TR></TABLE>
			</td>
			</tr>
			<tr bgcolor='<%=HeaderColor%>'>
				<TD align="right" width="50%">
					„—»Êÿ »Â ”›«—‘(Â«Ì):
					<span id="orders">
<%
					tempWriteAnd=""

					mySQL="SELECT * FROM InvoiceOrderRelations WHERE (Invoice='"& InvoiceID & "')"
					Set RS1 = conn.Execute(mySQL)
					while not(RS1.eof) 
						response.write "<input type='hidden' name='selectedOrders' value='"& RS1("Order") & "'>"
						response.write tempWriteAnd & Link2Trace(RS1("Order"))
						tempWriteAnd=" Ê "
						RS1.moveNext
					wend
%>					</span>&nbsp;
					<INPUT class="InvGenButton" TYPE="button" value=" €ÌÌ—" onClick="selectOrder();">
				</TD>
				<TD align="left"><table>
					<tr>
						<td align="left">‘„«—Â:</td>
						<td dir="LTR">
							<INPUT class="InvGenInput" NAME="InvoiceNo" value="<%=InvoiceNo%>" style="border:1px solid black;" TYPE="text" maxlength="10" size="10"></td>
						<td dir="RTL">
							<INPUT TYPE="checkbox" NAME="IsA" onClick='checkIsA();' <% if IsA then response.write " checked " %>> «·› &nbsp;</td>
					</tr>
					</table></TD>
			</tr>
<%
			tempWriteAnd=""
			relatedQuotesText=""
			hasRelatedQuotes = false
			mySQL="SELECT InvoiceQuoteRelations.[Quote], Quotes.[order_kind], Quotes.[order_title] FROM InvoiceQuoteRelations INNER JOIN Quotes ON InvoiceQuoteRelations.[Quote] = Quotes.[ID] WHERE (InvoiceQuoteRelations.[Invoice] = '"& InvoiceID & "')"

			Set RS1 = conn.Execute(mySQL)
			While Not(RS1.eof) 
				relatedQuotesText = relatedQuotesText & " <input type='hidden' name='selectedQuotes' value='"& RS1("Quote") & "'>"
				relatedQuotesText = relatedQuotesText & tempWriteAnd & Link2TraceQuote(RS1("Quote"))
				'relatedQuotesText = relatedQuotesText & " [ " & RS1("Order_Title") & " (" & RS1("Order_Kind") & ") ]"
				tempWriteAnd=" Ê "
				hasRelatedQuotes = true
				RS1.moveNext
			Wend
			RS1.close
			Set RS1 = Nothing 

			If hasRelatedQuotes Then 
%>
				<tr bgcolor='#AAAAEE' height="30">
					<td width="50%">
						„—»Êÿ »Â «” ⁄·«„ :
						<span id="quotes"><%=relatedQuotesText%></span>&nbsp;
						<INPUT class="InvGenButton" TYPE="button" value=" €ÌÌ—" onClick="selectQuote();">
						</TD>
					<td width="50%">
						&nbsp; 
						<SCRIPT LANGUAGE="JavaScript">
						<!--
						function js_Link2Trace(num){
							return "<A HREF='../order/Inquiry.asp?act=show&quote="+ num + "' target='_balnk'>"+ num + "</A>"
						}
						function selectQuote(){
							theSpan=document.getElementById("quotes");
							document.all.tmpDlgArg.value="";
							window.showModalDialog('Quotes.asp?act=select&customer='+document.all.customerID.value,document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
							if (document.all.tmpDlgArg.value!="[Esc]"){
								theSpan.innerHTML="";
								Arguments=document.all.tmpDlgArg.value.split("#")
								tempWriteAnd=""
								for (i=1;i<=Arguments[0];i++){
									theSpan.innerHTML += "<input type='hidden' name='selectedQuotes' value='"+Arguments[i]+"'>" + tempWriteAnd + js_Link2Trace(Arguments[i])
									tempWriteAnd=" Ê "
								}
							}
						}
							
						//-->
						</SCRIPT>
					</td>
				</tr>
<%			End If %>

			<tr bgcolor='#CCCC88'>
			<TD colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<tr bgcolor='#CCCC88'>
				<td align='center' width="25px"> # </td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" value="¬Ì „" size="3" ></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" value=" Ê÷ÌÕ« " size="30"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="ÿÊ·" size="2"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="⁄—÷" size="2"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value=" ⁄œ«œ" size="3"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="›—„" size="2"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value=" ⁄œ«œ „ÊÀ—" size="6"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="›Ì" size="7"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="ﬁÌ„ " size="9"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value=" Œ›Ì›"size="7"></td><!-- S A M -->
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="»—ê‘ " size="5"></td><!-- S A M -->
				<td><INPUT class="InvHeadInput4" readonly TYPE="text" Value="„«·Ì« " size="6"></td><!-- S A M -->
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="ﬁ«»· Å—œ«Œ " size="9"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
			<tr bgcolor='#CCCC88'>
			<TD colspan="10"><div style="overflow:auto; height:190px; width:*;">
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<Tbody id="InvoiceLines">

<%		
	i=0
	mySQL="SELECT *,InvoiceLines.ID as lineID,isnull(invoiceItems.fee,0) as ItemFee FROM InvoiceLines LEFT OUTER JOIN invoiceItems ON InvoiceLines.item = invoiceItems.id WHERE (Invoice='"& InvoiceID & "') "
	Set RS1 = conn.Execute(mySQL)
	while not(RS1.eof) 
	if RS1("Item") <> 39999 then
		i=i+1
%>
			<tr bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
				<td align='center' width="25px"><%=i%></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Items" value="<%=RS1("Item")%>" size="3" Maxlength="6" onKeyPress="return mask(this,event);" onfocus="setCurrentRow(this.parentNode.parentNode.rowIndex);" onChange='return check(this);'>
				<INPUT TYPE="hidden" name="type" value="<%=RS1("type")%>">
				<INPUT TYPE="hidden" name="fee" value="<%=RS1("ItemFee")%>">
				<input type='hidden' name='hasVat' value='<%=text2value(RS1("hasVat"))%>'>
				<input type='hidden' name='lineID' value='<%=RS1("lineID")%>'>
				</td>
				<td dir="RTL"><INPUT class="InvRowInput2" TYPE="text" NAME="Descriptions" value="<%=RS1("Description")%>" size="30"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Lengths" value="<%=RS1("Length")%>" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Widths" value="<%=RS1("Width")%>" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Qttys" value="<%=RS1("Qtty")%>" size="3" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Sets" value="<%=RS1("Sets")%>" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="AppQttys" value="<%=Separate(RS1("AppQtty"))%>" size="6" onBlur="setPrice(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Fees" <% if (cdbl(rs1("ItemFee"))>0 and cint(RS1("type"))<>0) then response.write " readonly='readonly' "%> value="<%if RS1("AppQtty") <> 0 then response.write Separate(RS1("Price")/RS1("AppQtty")) else response.write "0"%>" size="7" onBlur="setPrice(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Prices" value="<%=Separate(RS1("Price"))%>" size="9" readonly tabIndex="9999"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Discounts" value="<%=Separate(RS1("Discount"))%>" size="7" onBlur="setPrice(this);"></td><!-- S A M -->
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Reverses" value="<%=Separate(RS1("Reverse"))%>" size="5" onBlur="setPrice(this);" onfocus="setCurrentRow(this.parentNode.parentNode.rowIndex);"></td><!-- S A M -->
				<td dir="LTR"><INPUT class="InvRowInput4" TYPE="text" Name="Vat" value="<%=Separate(RS1("Vat"))%>" size="6" readonly></td><!-- S A M -->
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="AppPrices" value="<%=Separate(RS1("Price") - RS1("Discount") - RS1("Reverse") + RS1("Vat"))%>" size="9" readonly tabIndex="9999"></td><!-- S A M -->
			</tr>
<%
	end if
		RS1.moveNext
	wend
	RS1.close
%>
			<tr bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
				<td colspan="15">
					<INPUT class="InvGenButton" TYPE="button" value="«÷«›Â" onkeyDown="if(event.keyCode==9) {setCurrentRow(this.parentNode.parentNode.rowIndex); return false;};" onClick="addRow();">
				</td>
			</tr>
			</Tbody></TABLE></div>
			</TD>
			</tr>
			<tr bgcolor='#CCCC88'>
			<TD colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#CCCC88">
			<tr bgcolor='#CCCC88'>
				<td colspan='9' width='500px'>*** Œ›Ì› —‰œ ›«ﬂ Ê— »⁄œ «“ À»  œ—Ã ŒÊ«Âœ ‘œ***</td>
				<!--td align='center' width="25px"> &nbsp; </td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="3" ></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="30"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="3"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="6"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="7"></td-->
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalPrice" value="<%=Separate(totalPrice)%>" TYPE="text" size="9"></td>
				<!-- S A M -->
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalDiscount" value="<%=Separate(totalDiscount)%>" TYPE="text" size="7"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalReverse" value="<%=Separate(totalReverse)%>" TYPE="text" size="5"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalVat" value="<%=Separate(totalVat)%>" TYPE="text" size="6"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="Payable" value="<%=Separate(totalPrice - totalDiscount - totalReverse + totalVat)%>" TYPE="text" size="9"></td>
			</tr>
			<tr bgcolor='#CCCC88'>
				<td colspan="9"> &nbsp; </td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput" TYPE="text" size="9"></td>
				<!-- S A M -->
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" Name="TPDiscount" value="<%=Pourcent(totalDiscount,totalPrice) & "% Œ›Ì›"%>" size="7"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" Name="TPReverse" value="<%=Pourcent(totalReverse,totalPrice) & "%»—ê‘ "%>" size="5"></td>
				<td dir="LTR"><INPUT readonly calss="InvHeadINput" TYPE="text" size="6" value="<%=session("VatRate")%>%„«·Ì« "></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput" TYPE="text" size="9" value="—‰œ ‘œÂ"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
		</table>
		<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
		<tr>
			<td align='center'>&nbsp;<!-- <INPUT class="InvGenButton" TYPE="button" value=" «ÌÌœ ›«ﬂ Ê—" onclick="ApproveInvoice();"> --></td>
			<td width="40">&nbsp;</td>
			<td align='center'>&nbsp;<!-- <INPUT class="InvGenButton" TYPE="button" value="’œÊ— ›«ﬂ Ê—" onclick="IssueInvoice();"> --></td>
			<td width="40">&nbsp;</td>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="–ŒÌ—Â " onclick="submitOperations();"></td>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="«‰’—«›" onclick="window.location='AccountReport.asp?act=showInvoice&invoice=<%=InvoiceID%>';"></td>
		</tr>
		</TABLE>
		</FORM>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			//document.getElementsByName("Items")[0].focus();
		//-->
		</SCRIPT>

<%
elseif request("act")="submitEdit" then

	'******************** Checking and Preparing Input ****************
	errorFound=false
	ON ERROR RESUME NEXT

		InvoiceID=		clng(request.form("InvoiceID"))
		CustomerID=		clng(request.form("CustomerID"))

		issueDate=	request.form("issueDate")

		if request.form("IsA") = "on" then 
			IsA=1 
			InvoiceNo=request.form("InvoiceNo")
			if InvoiceNo <> "" then InvoiceNo = clng(InvoiceNo)
		else 
			IsA=0
			InvoiceNo=""
		end if
			
		for i=1 to request.form("selectedOrders").count
			theOrder=		clng(request.form("selectedOrders")(i))
			mySQL="SELECT ID FROM Orders WHERE ID=" & theOrder 
			Set rs=conn.Execute(mySQL)
			if rs.eof then
				errorFound=True
				exit for
			end if
			rs.close
		next

		for i=1 to request.form("selectedQuotes").count
			theQuote=		clng(request.form("selectedQuotes")(i))
			mySQL="SELECT ID FROM Quotes WHERE ID=" & theQuote 
			Set rs=conn.Execute(mySQL)
			if rs.eof then
				errorFound=True
				exit for
			end if
			rs.close
		next

		Set rs= Nothing

		mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
		Set rs= conn.Execute(mySQL)
		if NOT rs.eof then
			voided=		rs("Voided")
			issued=		rs("Issued")
			approved=	rs("Approved")
			isReverse=	rs("IsReverse")
			ApprovedBy=	rs("ApprovedBy")
		else
			errorFound=True
		end if

		if Err.Number<>0 then
			Err.clear
			errorFound=True
		end if

		if NOT errorFound then
			TotalPrice	=		0
			TotalDiscount =		0
			TotalReverse =		0
			TotalReceivable =	0
			TotalVat =			0
			RFD =				0

			for i=1 to request.form("Items").count 

				theItem =			clng(text2value(request.form("Items")(i)))
				mySQL="SELECT ID FROM InvoiceItems WHERE ID=" & theItem 
				Set rs=conn.Execute(mySQL)
				if rs.eof then
					errorFound=True
					exit for
				end if
				rs.close

				theDescription =	left(sqlSafe(request.form("Descriptions")(i)),100)

				theAppQtty =		cdbl(text2value(request.form("AppQttys")(i)))
				thePrice =			clng(text2value(request.form("Prices")(i)))

				theDiscount =		text2value(request.form("Discounts")(i))
				theReverse =		text2value(request.form("Reverses")(i))

				theLength =			text2value(request.form("Lengths")(i))
				theWidth =			text2value(request.form("Widths")(i))
				theQtty =			text2value(request.form("Qttys")(i))
				theSets =			text2value(request.form("Sets")(i))
				theVat =			text2value(request.form("Vat")(i))
				'theHasVat =			text2value(request.form("hasVat")(i))

				if theDiscount <>"" then theDiscount= clng(theDiscount)
				if theReverse <> "" then theReverse = clng(theReverse)

				if theLength <>	"" then  theLength	= cdbl(theLength)
				if theWidth <> ""  then	 theWidth	= cdbl(theWidth) 
				if theQtty <> ""   then	 theQtty	= clng(theQtty)
				if theSets <> ""   then	 theSets	= clng(theSets)

				TotalPrice	=		TotalPrice + thePrice
				TotalDiscount =		TotalDiscount + theDiscount
				TotalReverse =		TotalReverse + theReverse
				TotalReceivable =	TotalReceivable + (thePrice - theDiscount - theReverse + theVat)
				TotalVat =			TotalVat + theVat

			next
		RFD = TotalReceivable - fix(TotalReceivable / 1000) * 1000
		'RFD = RFD / 1.03
		TotalReceivable = TotalReceivable - RFD
		TotalDiscount = TotalDiscount + RFD
		end if

		if Err.Number<>0 then
			Err.clear
			errorFound=True
		end if

	ON ERROR GOTO 0

	if errorFound then
		response.write "<br>" 
		call showAlert ("Œÿ« œ— Ê—ÊœÌ",CONST_MSG_ERROR) 
		response.end
	end if
	'^^^^---------------- Checking Input ------------^^^^

	if isReverse then
		'Check for permission for EDITTING Rev. Invoice
		if not Auth(6 , 5) then NotAllowdToViewThisPage()
		itemType=4 
	else
		'Check for permission for EDITTING Invoice
		if not Auth(6 , 3) then NotAllowdToViewThisPage()
		itemType=1
	end if

	if voided then
		Conn.close
		response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·« »«ÿ· ‘œÂ «” .")
	elseif issued then
		if Auth(6 , "A") then
			' Has the Priviledge to change the Invoice / Reverse Invoice

			'mySQL="SELECT ID FROM ARItems WHERE (Type='"& itemType & "') AND (GL_Update=1) AND (Link='"& InvoiceID & "')"
			'Changed by Kid ! 831124
			mySQL="SELECT ARItems.ID, ARItems.GL_Update, EffGLRows.GL, EffGLRows.GLDocID FROM ARItems LEFT OUTER JOIN (SELECT Link, GL, GLDocID FROM EffectiveGLRows WHERE SYS = 'AR') EffGLRows ON ARItems.ID = EffGLRows.Link WHERE (ARItems.Type = '"& itemType & "') AND (ARItems.Link = '"& InvoiceID & "')"

			Set RS2 = conn.Execute(mySQL)

			if RS2.eof then
				Conn.close
				response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("Œÿ« !! <br><br> ÅÌœ« ‰‘œ.")
			else
				if RS2("GL_Update") = False  then
					tmpGL=RS2("GL")
					tmpGLDoc=RS2("GLDocID")
					Conn.close
					response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("ﬁ»·« »—«Ì «Ì‰ ›«ﬂ Ê— ”‰œ Õ”«»œ«—Ì ’«œ— ‘œÂ «” .<br><br>œ› — ﬂ·:"& tmpGL & " ”‰œ ‘„«—Â: "& tmpGLDoc & " .")
				else
					ARItemID=RS2("ID")
					IssuedButEdit=true
				end if
			end if
			RS2.close
		else
			Conn.close
			response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·« ’«œ— ‘œÂ «” .")
		end if
	elseif approved then 
		call UnApproveInvoice ( InvoiceID , ApprovedBy )
	end if
	

	'******************* Editing  *******************
	' ****
	if IssuedButEdit then
		' Only Updating  IssuedDate, Number & no IsA
		'				 and related Orders

		'---- Checking wether issueDate is valid in current open GL
		If Not CheckDateFormat(issueDate) Then
			Conn.close
			response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode(" «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” .")
		end if

		if (issueDate < session("OpenGLStartDate")) OR (issueDate > session("OpenGLEndDate")) then
			Conn.close
			response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” . <br>(œ— ”«· „«·Ì Ã«—Ì ‰Ì” )")
		end if 
		'----
	'----- Check GL is closed

	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----

		mySQL="UPDATE Invoices SET IssuedDate=N'" & issueDate & "', Number='"& InvoiceNo & "', IsA='"&IsA&"',issuedDate_en=dbo.udf_date_solarToDate(cast(substring('" & issueDate & "',1,4) as int),cast(substring('" & issueDate & "',6,2) as int),cast(substring('" & issueDate & "',9,2) as int)) WHERE (ID='"& InvoiceID & "')"
		'response.write mySQL
		'response.end
		conn.Execute(mySQL)
'---------------------------------------------------------------------------------------------------
		for i=1 to request.form("Items").count 
			if request.form("lineID") <> "" then
			'response.write text2value(request.form("lineID")(i))
				theID = 			clng(text2value(request.form("lineID")(i)))
				theItem =			clng(text2value(request.form("Items")(i)))
				theDescription =	left(sqlSafe(request.form("Descriptions")(i)),100)
				'theAppQtty =		cdbl(text2value(request.form("AppQttys")(i)))
				theLength =			text2value(request.form("Lengths")(i))
				theWidth =			text2value(request.form("Widths")(i))
				theQtty =			text2value(request.form("Qttys")(i))
				theSets =			text2value(request.form("Sets")(i))
				if theLength <>	"" then	theLength	= cdbl(theLength)
				if theWidth <> ""  then	 theWidth	= cdbl(theWidth) 
				if theQtty <> ""	  then	 theQtty	= clng(theQtty)
				if theSets <> ""	  then	 theSets	= clng(theSets)
	
				mySQL="UPDATE InvoiceLines SET Item='" & theItem & "', Description=N'" & theDescription & "', Length='" & theLength & "', Width='" & theWidth & "', Qtty='" & theQtty & "', Sets='" & theSets & "' WHERE ID = " & theID
				conn.Execute(mySQL)
'response.write "<br>" & mySQL
'response.end				
			end if
		next
		 

'---------------------------------------------------------------------------------------------------
		if IsA then
			GLAccount=	"91001"	'This must be changed... (Sales A)
		else
			GLAccount=	"91002"	'This must be changed... (Sales B)
		end if
		
		' Changed By Kid 860118 , seasing to use Sales B

		'GLAccount=	"91001"	'This must be changed... (Sales A)

		conn.Execute("UPDATE ARItems SET GL='"& OpenGL & "', EffectiveDate='" & issueDate & "', GLAccount='"& GLAccount & "' WHERE (ID='" & ARItemID & "')")

		'**************** Updating Invoice-Order Relations ****************
		'mySQL="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice= '" & InvoiceID & "'))"
		'Changed By Kid ! 840509 
		'set orders which are ONLY related to this invoice, "Open"
		'that means, orders which are related to this invoice and are NOT related to any OTHER issued invoices.
		mySQL ="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Order] NOT IN (SELECT InvoiceOrderRelations.[ORDER] FROM Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "'))))"
		conn.Execute(mySQL)

		mySQL ="DELETE FROM InvoiceOrderRelations WHERE (Invoice='" & InvoiceID & "')"
		'mySQL ="DELETE FROM InvoiceOrderRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Order] NOT IN (SELECT InvoiceOrderRelations.[ORDER] FROM Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "')))"
		conn.Execute(mySQL)
		
		for i=1 to request.form("selectedOrders").count
			theOrder=	clng(request.form("selectedOrders")(i))
			mySQL="INSERT INTO InvoiceOrderRelations (Invoice,[Order]) VALUES ('" & InvoiceID & "', '" & theOrder & "')"
			conn.Execute(mySQL)
		next

		conn.Execute("UPDATE Orders SET Closed=1 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice='" & InvoiceID & "'))")
		'^^^^------------ Updating Invoice-Order Relations ------------^^^^


		'**************** Updating Invoice-Quote Relations ****************
		mySQL ="UPDATE Quotes SET Closed=0 WHERE ID IN (SELECT [Quote] FROM InvoiceQuoteRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Quote] NOT IN (SELECT InvoiceQuoteRelations.[Quote] FROM Invoices INNER JOIN InvoiceQuoteRelations ON Invoices.ID = InvoiceQuoteRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "'))))"
		conn.Execute(mySQL)

		mySQL ="DELETE FROM InvoiceQuoteRelations WHERE (Invoice='" & InvoiceID & "')"
		conn.Execute(mySQL)

		for i=1 to request.form("selectedQuotes").count
			theQuote=	clng(request.form("selectedQuotes")(i))
			mySQL="INSERT INTO InvoiceQuoteRelations (Invoice,[Quote]) VALUES ('" & InvoiceID & "', '" & theQuote & "')"
			conn.Execute(mySQL)
		next

		conn.Execute("UPDATE Quotes SET Closed=1 WHERE ID IN (SELECT [Quote] FROM InvoiceQuoteRelations WHERE (Invoice='" & InvoiceID & "'))")
		'^^^^------------ Updating Invoice-Quote Relations ------------^^^^
		mySQL = "SELECT * FROM ARItems WHERE Type=1 AND Link=" & InvoiceID
		set RSSS=conn.Execute(mySQL)
		if RSSS.eof or not Auth(6 , "N") then 
			response.redirect "AccountReport.asp?act=showInvoice&invoice=" & InvoiceID & "&msg=" &Server.URLEncode("«Ì‰ ›«ﬂ‰Ê— ’«œ— ‘œÂ Ê ‘„« „Ã«“ »Â  €ÌÌ—«  ﬂ·Ì œ— ¬‰ ‰Ì” Ìœ.")
		else
			if RSSS("GL_Update") then 
			' In this secssion we update all issued invoice has not GLs
			'---------------------------------------------------------------------------------------------------
			'---------------------------------------------------------------------------------------------------
			'---------------------------------------------------------------------------------------------------
			'---------------------------------------------------------------------------------------------------
			'---------------------------------------------------------------------------------------------------	
				set RSinvoice=conn.Execute("SELECT * FROM Invoices WHERE ID="&InvoiceID)
				'response.write TotalReceivable & "<br>"
				'response.write totalVat & "<br>"
				'response.write TotalPrice & "<br>"
				'response.write TotalDiscount & "<br>"
				'response.end
				if (not RSinvoice.eof) and (clng(RSinvoice("TotalReceivable"))<>TotalReceivable or Cdbl(RSinvoice("TotalVat"))<>TotalVat) then
					oldTotalReceivable = CLng(RSinvoice("TotalReceivable"))
					
					
					mySQL="UPDATE Invoices SET Customer='"& CustomerID & "', Number='"& InvoiceNo & "', TotalPrice='"& TotalPrice & "', TotalDiscount='"& TotalDiscount & "', TotalReverse='"& TotalReverse & "', TotalReceivable='"& TotalReceivable & "' , IsA='"& IsA & "', TotalVat='" & totalVat & "' WHERE (ID='"& InvoiceID & "')"
					conn.Execute(mySQL)
			
					mySQL="DELETE FROM InvoiceLines WHERE (Invoice='"& InvoiceID & "')"
					conn.Execute(mySQL)
			
					'**************************** Inserting Invoice Lines ****************
					for i=1 to request.form("Items").count 
						theItem =			clng(text2value(request.form("Items")(i)))
						theDescription =	left(sqlSafe(request.form("Descriptions")(i)),100)
			
						theAppQtty =		cdbl(text2value(request.form("AppQttys")(i)))
						thePrice =			clng(text2value(request.form("Prices")(i)))
			
						theDiscount =		text2value(request.form("Discounts")(i))
						theReverse =		text2value(request.form("Reverses")(i))
			
						theLength =			text2value(request.form("Lengths")(i))
						theWidth =			text2value(request.form("Widths")(i))
						theQtty =			text2value(request.form("Qttys")(i))
						theSets =			text2value(request.form("Sets")(i))
						theVat =			clng(text2value(request.form("Vat")(i)))
						theHasVat =			text2value(request.form("hasVat")(i))
			
						if theDiscount <>"" then theDiscount= clng(theDiscount)
						if theReverse <> "" then theReverse = clng(theReverse)
			
						if theLength <>	"" then	theLength	= cdbl(theLength)
						if theWidth <> ""  then	 theWidth	= cdbl(theWidth) 
						if theQtty <> ""	  then	 theQtty	= clng(theQtty)
						if theSets <> ""	  then	 theSets	= clng(theSets)
			
						mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat) VALUES ('"& InvoiceID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "', "& theHasVat &")"
						conn.Execute(mySQL)
					next 
					if RFD > 0 then
						theItem =			39999
						theDescription =	" Œ›Ì› —‰œ ›«ﬂ Ê—"
			
						theAppQtty =		0
						thePrice =			0
			
						theDiscount =		RFD
						theReverse =		0
			
						theLength =			0
						theWidth =			0
						theQtty =			0
						theSets =			0
						theVat =			0
						mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat) VALUES ('"& InvoiceID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "')"
						conn.Execute(mySQL)
					end if
					mySQL="UPDATE ARItems SET AmountOriginal='"& TotalReceivable &"', RemainedAmount='"& TotalReceivable &"' ,Vat='"& totalVat &"' WHERE ID="&RSSS("ID")
					conn.Execute(mySQL)
					mySQL="UPDATE Accounts SET ARBalance=ARBalance - "&TotalReceivable - oldTotalReceivable&" WHERE id=" & CustomerID
					conn.Execute(mySQL)
				end if
			'end if
			'****
			'^^^^--------------- Editing  ---------------^^^^
		'--------------------------------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------------
			response.redirect "AccountReport.asp?act=showInvoice&invoice=" & InvoiceID & "&msg=" &Server.URLEncode("›«ﬂ Ê— ‘„« œç«—  €ÌÌ—«  ﬂ·Ì ‘œ")
		'---------------------------------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------------
			else
				response.redirect "AccountReport.asp?act=showInvoice&invoice=" & InvoiceID & "&msg=" &Server.URLEncode("»—«Ì «Ì‰ ›«ﬂ Ê— ﬁ»·« ”‰œ Õ”«»œ«—Ì ’«œ— ‘œÂ° Å” ¬‰—« ‰„Ìù Ê«‰ ÊÌ—«Ì‘ ﬂ—œ.")
			end if
		end if
		conn.close
		response.redirect "AccountReport.asp?act=showInvoice&invoice=" & InvoiceID & "&msg=" &Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·« ’«œ— ‘œÂ »Êœ.<br>Â—ç‰œ ﬂÂ ‘„« «Ã«“Â œ«—Ìœ ﬂÂ «Ì‰ ›«ﬂ Ê— —«  €ÌÌ— »œÂÌœ<br>»Â — «”  ﬂÂ «Ì‰ ﬂ«— —«  ﬂ—«— ‰ﬂ‰Ìœ.")
	else
' S A M
'response.write(totalDiscount)
'response.end
		mySQL="UPDATE Invoices SET Customer='"& CustomerID & "', Number='"& InvoiceNo & "', TotalPrice='"& TotalPrice & "', TotalDiscount='"& TotalDiscount & "', TotalReverse='"& TotalReverse & "', TotalReceivable='"& TotalReceivable & "' , IsA='"& IsA & "', TotalVat='" & totalVat & "' WHERE (ID='"& InvoiceID & "')"
		conn.Execute(mySQL)

		mySQL="DELETE FROM InvoiceLines WHERE (Invoice='"& InvoiceID & "')"
		conn.Execute(mySQL)

		'**************************** Inserting Invoice Lines ****************
		for i=1 to request.form("Items").count 
			theItem =			clng(text2value(request.form("Items")(i)))
			theDescription =	left(sqlSafe(request.form("Descriptions")(i)),100)

			theAppQtty =		cdbl(text2value(request.form("AppQttys")(i)))
			thePrice =			clng(text2value(request.form("Prices")(i)))

			theDiscount =		text2value(request.form("Discounts")(i))
			theReverse =		text2value(request.form("Reverses")(i))

			theLength =			text2value(request.form("Lengths")(i))
			theWidth =			text2value(request.form("Widths")(i))
			theQtty =			text2value(request.form("Qttys")(i))
			theSets =			text2value(request.form("Sets")(i))
			theVat =			clng(text2value(request.form("Vat")(i)))
			theHasVat =			text2value(request.form("hasVat")(i))

			if theDiscount <>"" then theDiscount= clng(theDiscount)
			if theReverse <> "" then theReverse = clng(theReverse)

			if theLength <>	"" then  theLength	= cdbl(theLength)
			if theWidth <> ""  then	 theWidth	= cdbl(theWidth) 
			if theQtty <> ""   then	 theQtty	= clng(theQtty)
			if theSets <> ""   then	 theSets	= clng(theSets)

			mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat) VALUES ('"& InvoiceID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "', "& theHasVat &")"
			conn.Execute(mySQL)
		next 
		if RFD > 0 then
			theItem =			39999
			theDescription =	" Œ›Ì› —‰œ ›«ﬂ Ê—"

			theAppQtty =		0
			thePrice =			0

			theDiscount =		RFD
			theReverse =		0

			theLength =			0
			theWidth =			0
			theQtty =			0
			theSets =			0
			theVat =			0
			mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat) VALUES ('"& InvoiceID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "')"
			conn.Execute(mySQL)
		end if

		'**************** Updating Invoice-Order Relations ****************
		mySQL="DELETE FROM InvoiceOrderRelations WHERE (Invoice='" & InvoiceID & "')"
		conn.Execute(mySQL)
'		response.write "aaaa:"
'		response.write request.form("selectedOrders").count
		for i=1 to request.form("selectedOrders").count
			theOrder=	clng(request.form("selectedOrders")(i))
			mySQL="INSERT INTO InvoiceOrderRelations (Invoice,[Order]) VALUES ('" & InvoiceID & "', '" & theOrder & "')"
			'--------------------------SAM------------------------------
'			response.write mySQL
			'-----------------------------------------------------------
			conn.Execute(mySQL)
		next
		'^^^^------------ Updating Invoice-Order Relations ------------^^^^

		'**************** Updating Invoice-Quote Relations ****************
		mySQL="DELETE FROM InvoiceQuoteRelations WHERE (Invoice='" & InvoiceID & "')"
		conn.Execute(mySQL)
		for i=1 to request.form("selectedQuotes").count
			theQuote=	clng(request.form("selectedQuotes")(i))
			mySQL="INSERT INTO InvoiceQuoteRelations (Invoice,[Quote]) VALUES ('" & InvoiceID & "', '" & theQuote & "')"
			conn.Execute(mySQL)
		next
		'^^^^------------ Updating Invoice-Quote Relations ------------^^^^

	end if
	'****
	'^^^^--------------- Editing  ---------------^^^^

	response.redirect "AccountReport.asp?act=showInvoice&invoice=" & InvoiceID
'--------------------------------------------------------------------------------------------------------------------
elseif request("act")="approveInvoice" then

	if not Auth(6 , "C") then		
		'Doesn't have the Priviledge to APPROVE the Invoice 
		response.write "<br>" 
		call showAlert ("‘„« „Ã«“ »Â  «ÌÌœ ›«ﬂ Ê— ‰Ì” Ìœ",CONST_MSG_ERROR) 
		response.end
	end if
	
	'-------------------------------------------------------------------------
	'---- CHECK pickup list
	'-------------------------------------------------------------------------
	mySQL="select distinct sales.item, sales.description, sales.appQtty,pik.itemID,pik.itemName,pik.qtty from (select invoiceLines.item,max(invoiceLines.description) as description,sum(invoiceLines.appQtty) as appQtty, InventoryInvoiceRelations.inventoryItem from InvoiceLines inner join Invoices on invoices.ID=invoiceLines.Invoice inner join InvoiceOrderRelations on InvoiceOrderRelations.invoice=invoices.id inner join InventoryInvoiceRelations on InventoryInvoiceRelations.invoiceItem=invoiceLines.Item WHERE Invoices.id=" & InvoiceID & " group by invoiceLines.item, InventoryInvoiceRelations.inventoryItem) as sales full outer join (SELECT InventoryPickuplistItems.itemID, max(InventoryPickuplistItems.ItemName) AS ItemName, sum(InventoryPickuplistItems.Qtty) AS Qtty, InventoryInvoiceRelations.invoiceItem FROM InventoryPickuplistItems INNER JOIN InventoryPickuplists ON InventoryPickuplistItems.pickupListID = InventoryPickuplists.id INNER JOIN InvoiceOrderRelations ON InventoryPickuplistItems.Order_ID = InvoiceOrderRelations.[Order] inner join invoices on InvoiceOrderRelations.invoice=invoices.id inner join InventoryInvoiceRelations on InventoryInvoiceRelations.inventoryItem=InventoryPickuplistItems.itemID WHERE NOT InventoryPickuplists.Status = N'del' and InventoryPickuplistItems.CustomerHaveInvItem=0 and Invoices.id=" & InvoiceID & " group by InventoryPickuplistItems.itemID, InventoryInvoiceRelations.invoiceItem  ) as pik on pik.itemID=sales.inventoryItem or pik.invoiceItem=sales.item"
	set rs=Conn.Execute (mySQL)
	msg=""
	while not rs.eof
		if IsNull(rs("appQtty")) then 
			msg = msg & "»—«Ì ÕÊ«·ÂùÌ "  & rs("itemName") & " ‘„« ÂÌç ¬Ì „Ì œ— ›«ﬂ Ê— À»  ‰ﬂ—œÌœ<br>"
		else
			if IsNull(rs("qtty")) then 
				msg = msg & "»—«Ì " & rs("description") & " ‘„« ÂÌç ÕÊ«·Âù«Ì ’«œ— ‰ﬂ—œÂù«Ìœ<br>"
			else
				if CDbl(rs("qtty")) <> CDbl(rs("appQtty")) then msg = msg & " ⁄œ«œ ÕÊ«·Â <b>" & rs("itemName") & "</b> »«  ⁄œ«œ <b>" & rs("description") & "</b> „€«Ì— «” .<br>" 
			end if
		end if
		
		rs.moveNext
	wend
	set rs=nothing
	
	'-------------------------------------------------------------------------
	'---- CHECK out Service
	'-------------------------------------------------------------------------
	mySQL="select * from (SELECT InvoiceItems.id,max(PurchaseRequests.TypeName) as TypeName, sum(PurchaseOrders.Qtty) as Qtty FROM PurchaseOrders FULL OUTER JOIN PurchaseRequestOrderRelations RIGHT OUTER JOIN PurchaseRequests INNER JOIN InvoiceOrderRelations ON PurchaseRequests.Order_ID = InvoiceOrderRelations.[Order] ON PurchaseRequestOrderRelations.Req_ID = PurchaseRequests.ID ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID inner join InvoiceItems on PurchaseOrders.TypeID=InvoiceItems.RelatedInventoryItemID and InvoiceItems.Type=5 WHERE (InvoiceOrderRelations.Invoice =216989) and PurchaseRequests.Status<> 'del' group by InvoiceItems.id) as outService full outer join (select item,max(description) as description,sum(appQtty) as appQtty from InvoiceLines inner join InvoiceItems on InvoiceItems.ID=invoiceLines.Item and InvoiceItems.Type=5 where Invoice=" & InvoiceID & " group by item) as sales on sales.item=outService.id"
	set rs=Conn.Execute (mySQL)
	while not rs.eof
		if IsNull(rs("appQtty")) then 
			msg = msg & "»—«Ì Œœ„«  "  & rs("typeName") & " ‘„« ÂÌç ¬Ì „Ì œ— ›«ﬂ Ê— À»  ‰ﬂ—œÌœ<br>"
		else
			if IsNull(rs("qtty")) then 
				msg = msg & "»—«Ì " & rs("description") & " ‘„« ÂÌç œ—ŒÊ«”  Œœ„« Ì À»  ‰ﬂ—œÂù«Ìœ<br>"
			else
				if CDbl(rs("qtty")) <> CDbl(rs("appQtty")) then msg = msg & " ⁄œ«œ  <b>" & rs("typeName") & "</b> »«  ⁄œ«œ <b>" & rs("description") & "</b> „€«Ì— «” .<br>" 
			end if
		end if
		
		rs.moveNext
	wend
	set rs=nothing
	if (msg<>"") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode(msg)
	end if 
	'-------------------------------------------------------------------------
	'---------------------------------------------
	'-------------------------------------------------------------------------
	if request("invoice")<>"" then

		InvoiceID=request("invoice")
		if not(isnumeric(request("invoice"))) then
			ShowErrorMessage("Œÿ«")
			response.end
		end if
		mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			ShowErrorMessage("ÅÌœ« ‰‘œ ")
			response.end
		else
			if RS1("Voided") = True then
				Conn.close
				response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·« »«ÿ· ‘œÂ «” .")
			elseif RS1("Issued") = True then
				Conn.close
				response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·« ’«œ— ‘œÂ «” .")
			elseif RS1("Approved") = True then
				Conn.close
				response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·«  «ÌÌœ ‘œÂ «” .")
			end if
		end if
	else
		ShowErrorMessage("Œÿ«")
		response.end
	end if
	'--------------------SAM, Iqnore multi invoice in one order
	'mySQL="SELECT COUNT(*) AS OrderCount FROM InvoiceOrderRelations WHERE [Order] IN (SELECT [Order] FROM InvoiceOrderRelations WHERE Invoice=" & InvoiceID &")"-----------SAM coorect this on 9 mar 2011
	mySQL = "SELECT COUNT(Invoice) AS OrderCount FROM (SELECT DISTINCT InvoiceOrderRelations.Invoice FROM InvoiceOrderRelations inner join Invoices on InvoiceOrderRelations.Invoice = Invoices.ID WHERE InvoiceOrderRelations.[Order] IN (SELECT [Order] FROM InvoiceOrderRelations WHERE Invoice=204133) and Invoices.Voided=0) tbl"
	set RS1= conn.Execute(mySQL)
'	response.write rs1("OrderCount")
'	response.end
	if RS1("OrderCount")>1 then 
		conn.close
		response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("»Â «Ì‰ ”›«—‘ œÊ ›«ﬂ Ê— „—»Êÿ ‘œÂ! ·ÿ›« »Â Â— ”›«—‘ ›ﬁÿ Ìﬂ ›«ﬂ Ê— „ ’· ﬂ‰Ìœ.")
	end if
	'------------------------------------------------
	mySQL="UPDATE Invoices SET Approved=1, ApprovedDate=N'"& shamsiToday() & "', ApprovedBy='"& session("ID") & "' WHERE (ID='"& InvoiceID & "')"
	conn.Execute(mySQL)

	response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID 
'-------------------------------------------------------------------------------------------------------
elseif request("act")="IssueInvoice" then

	if not Auth(6 , "D") then		
		'Doesn't have the Priviledge to ISSUE the Invoice 
		response.write "<br>" 
		call showAlert ("‘„« „Ã«“ »Â ’œÊ— ›«ﬂ Ê— ‰Ì” Ìœ",CONST_MSG_ERROR) 
		response.end
	end if

	ON ERROR RESUME NEXT
		InvoiceID=		clng(request("Invoice"))

		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
		end if
	ON ERROR GOTO 0

	creationDate=	shamsiToday() 
	issueDate=		SqlSafe(request("issueDate"))
	if issueDate="" then issueDate=creationDate

	if issueDate<>creationDate then
		if Auth(6 , "I") then
			' can ISSUE the Invoice / Rev. Invoice on another Date

			'---- Checking wether issueDate is valid in current open GL
			If Not CheckDateFormat(issueDate) Then
				Conn.close
				response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode(" «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” .")
			end if

			if (issueDate < session("OpenGLStartDate")) OR (issueDate > session("OpenGLEndDate")) then
				Conn.close
				response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” . <br>(œ— ”«· „«·Ì Ã«—Ì ‰Ì” )")
			end if 
			'----
		else
			'Doesn't have the Priviledge to ISSUE the Invoice on another Date
			response.write "<br>" 
			call showAlert ("‘„« „Ã«“ »Â ’œÊ— ›«ﬂ Ê— œ— «Ì‰  «—ÌŒ ‰Ì” Ìœ",CONST_MSG_ERROR) 
			response.end
		end if
	end if
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----

	'---- Checking wether issueDate is valid in current open GL
	If Not CheckDateFormat(issueDate) Then
		Conn.close
		response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ ’œÊ— „⁄ »— ‰Ì” .")
	end if

	if (issueDate < session("OpenGLStartDate")) OR (issueDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "AccountReport.asp?act=showInvoice&invoice=" & InvoiceID & "&errmsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ ’œÊ— „⁄ »— ‰Ì” . <br>(œ— ”«· „«·Ì Ã«—Ì ‰Ì” )")
	end if 
	'----
	'---- CHECK pickup list
	mySQL="select distinct sales.item, sales.description, sales.appQtty,pik.itemID,pik.itemName,pik.qtty from (select invoiceLines.item,max(invoiceLines.description) as description,sum(invoiceLines.appQtty) as appQtty, InventoryInvoiceRelations.inventoryItem from InvoiceLines inner join Invoices on invoices.ID=invoiceLines.Invoice inner join InvoiceOrderRelations on InvoiceOrderRelations.invoice=invoices.id inner join InventoryInvoiceRelations on InventoryInvoiceRelations.invoiceItem=invoiceLines.Item WHERE Invoices.id=" & InvoiceID & " group by invoiceLines.item, InventoryInvoiceRelations.inventoryItem) as sales full outer join (SELECT InventoryPickuplistItems.itemID, max(InventoryPickuplistItems.ItemName) AS ItemName, sum(InventoryPickuplistItems.Qtty) AS Qtty, InventoryInvoiceRelations.invoiceItem FROM InventoryPickuplistItems INNER JOIN InventoryPickuplists ON InventoryPickuplistItems.pickupListID = InventoryPickuplists.id INNER JOIN InvoiceOrderRelations ON InventoryPickuplistItems.Order_ID = InvoiceOrderRelations.[Order] inner join invoices on InvoiceOrderRelations.invoice=invoices.id inner join InventoryInvoiceRelations on InventoryInvoiceRelations.inventoryItem=InventoryPickuplistItems.itemID WHERE NOT InventoryPickuplists.Status = N'del' and InventoryPickuplistItems.CustomerHaveInvItem=0 and Invoices.id=" & InvoiceID & " group by InventoryPickuplistItems.itemID, InventoryInvoiceRelations.invoiceItem  ) as pik on pik.itemID=sales.inventoryItem or pik.invoiceItem=sales.item"
	set rs=Conn.Execute (mySQL)
	msg=""
	while not rs.eof
		if IsNull(rs("appQtty")) then 
			msg = msg & "»—«Ì ÕÊ«·ÂùÌ "  & rs("itemName") & " ‘„« ÂÌç ¬Ì „Ì œ— ›«ﬂ Ê— À»  ‰ﬂ—œÌœ<br>"
		else
			if IsNull(rs("qtty")) then 
				msg = msg & "»—«Ì " & rs("description") & " ‘„« ÂÌç ÕÊ«·Âù«Ì ’«œ— ‰ﬂ—œÂù«Ìœ<br>"
			else
				if CDbl(rs("qtty")) <> CDbl(rs("appQtty")) then msg = msg & " ⁄œ«œ ÕÊ«·Â <b>" & rs("itemName") & "</b> »«  ⁄œ«œ <b>" & rs("description") & "</b> „€«Ì— «” .<br>" 
			end if
		end if
		
		rs.moveNext
	wend
	set rs=nothing
	
	'----
	'---- CHECK out Service
	mySQL="select * from (SELECT InvoiceItems.id,max(PurchaseRequests.TypeName) as TypeName, sum(PurchaseOrders.Qtty) as Qtty FROM PurchaseOrders FULL OUTER JOIN PurchaseRequestOrderRelations RIGHT OUTER JOIN PurchaseRequests INNER JOIN InvoiceOrderRelations ON PurchaseRequests.Order_ID = InvoiceOrderRelations.[Order] ON PurchaseRequestOrderRelations.Req_ID = PurchaseRequests.ID ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID inner join InvoiceItems on PurchaseOrders.TypeID=InvoiceItems.RelatedInventoryItemID and InvoiceItems.Type=5 WHERE (InvoiceOrderRelations.Invoice =216989) and PurchaseRequests.Status<> 'del' group by InvoiceItems.id) as outService full outer join (select item,max(description) as description,sum(appQtty) as appQtty from InvoiceLines inner join InvoiceItems on InvoiceItems.ID=invoiceLines.Item and InvoiceItems.Type=5 where Invoice=" & InvoiceID & " group by item) as sales on sales.item=outService.id"
	set rs=Conn.Execute (mySQL)
	while not rs.eof
		if IsNull(rs("appQtty")) then 
			msg = msg & "»—«Ì Œœ„«  "  & rs("typeName") & " ‘„« ÂÌç ¬Ì „Ì œ— ›«ﬂ Ê— À»  ‰ﬂ—œÌœ<br>"
		else
			if IsNull(rs("qtty")) then 
				msg = msg & "»—«Ì " & rs("description") & " ‘„« ÂÌç œ—ŒÊ«”  Œœ„« Ì À»  ‰ﬂ—œÂù«Ìœ<br>"
			else
				if CDbl(rs("qtty")) <> CDbl(rs("appQtty")) then msg = msg & " ⁄œ«œ  <b>" & rs("typeName") & "</b> »«  ⁄œ«œ <b>" & rs("description") & "</b> „€«Ì— «” .<br>" 
			end if
		end if
		
		rs.moveNext
	wend
	set rs=nothing
	if (msg<>"") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode(msg)
	end if 
	'---------------------------------------------
	
	mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«! ›«ﬂ Ê— »« ‘„«—Â" & InvoiceID & " ÅÌœ« ‰‘œ.")
	else
		voided=			RS1("Voided")
		issued=			RS1("Issued")
		approved=		RS1("Approved")
		isReverse=		RS1("IsReverse")
		customerID=		RS1("Customer")
		invoiceFee=		RS1("TotalReceivable")
		IsA =			RS1("IsA")
		Vat =			RS1("TotalVat")' sam

		if voided then
			Conn.close
			response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·« »«ÿ· ‘œÂ «” .")
		elseif issued then
			Conn.close
			response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ﬁ»·« ’«œ— ‘œÂ «” .")
		elseif not approved then
			Conn.close
			response.redirect "?act=editInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê—  «ÌÌœ ‰‘œÂ «” .")
		end if
	end if

	mySQL="UPDATE Invoices SET Issued=1, IssuedDate=N'"& issueDate & "', IssuedBy='"& session("ID") & "' WHERE (ID='"& InvoiceID & "')"
	conn.Execute(mySQL)

	if isReverse then
		isCredit=1
		itemType=4 
	else
		isCredit=0
		itemType=1
		'----------------------- Declaring the related orders as closed --------------
		conn.Execute("UPDATE Orders SET Closed=1 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice='" & InvoiceID & "'))")

		'----------------------- Declaring the related Quotes as closed --------------
		conn.Execute("UPDATE Quotes SET Closed=1 WHERE ID IN (SELECT [Quote] FROM InvoiceQuoteRelations WHERE (Invoice='" & InvoiceID & "'))")
	end if

	'**************************** Creating ARItem for Invoice / Reverse Invoice  ****************
	'*** Type = 1 means ARItem is an Invoice
	'*** Type = 4 means ARItem is a Reverse Invoice

	firstGLAccount=	"13003"	'This must be changed... (Business Debitors)
	if IsA then
		GLAccount=	"91001"	'This must be changed... (Sales A)
	else
		GLAccount=	"91002"	'This must be changed... (Sales B)
	end if
	'
	' Changed By Kid 860118 , seasing to use Sales B

	'GLAccount=	"91001"	'This must be changed... (Sales A)
	
	mySQL="INSERT INTO ARItems (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount, Vat) VALUES ('" &_
	GLAccount & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& CustomerID & "', N'"& issueDate & "', '"& isCredit & "', '"& itemType & "', '"& InvoiceID & "', '"& invoiceFee & "', N'"& creationDate & "', '"& session("ID") & "', '"& invoiceFee & "', '" & Vat & "')"
	conn.Execute(mySQL)

	if isReverse then
		'*** ATTENTION: Increasing AR Balance ....
		mySQL="UPDATE Accounts SET ARBalance = ARBalance + '"& invoiceFee & "' WHERE (ID='"& CustomerID & "')"
	else
		'*** ATTENTION: Decreasing AR Balance ....
		mySQL="UPDATE Accounts SET ARBalance = ARBalance - '"& invoiceFee & "' WHERE (ID='"& CustomerID & "')"
	end if
	conn.Execute(mySQL)
	conn.close
	response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID 
'---------------------------------------------------------------------------------------------------------
elseif request("act")="voidInvoice" then
	if not Auth(6 , "F") then		
		'Doesn't have the Priviledge to VOID the Invoice 
		response.write "<br>" 
		call showAlert ("‘„« „Ã«“ »Â «»ÿ«· ›«ﬂ Ê— ‰Ì” Ìœ",CONST_MSG_ERROR) 
		response.end
	end if

	comment=sqlSafe(request("comment"))

	InvoiceID=request("invoice")
	if InvoiceID="" or not(isnumeric(InvoiceID)) then
		response.write "<br>" 
		call showAlert ("Œÿ« œ— ‘„«—Â ›«ﬂ Ê—",CONST_MSG_ERROR) 
		response.end
	end if

	InvoiceID=clng(InvoiceID)
	
	mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		ShowErrorMessage("ÅÌœ« ‰‘œ ")
		response.end
	else
		voided=			RS1("Voided")
		issued=			RS1("Issued")
		issuedBy=		RS1("IssuedBy")
		isReverse=		RS1("IsReverse")
		customerID=		RS1("Customer")
		invoiceFee=		RS1("TotalReceivable")
		IsA =			RS1("IsA")
		if voided then
			ShowErrorMessage("«Ì‰ ›«ﬂ Ê— ﬁ»·« œ—  «—ÌŒ <span dir='LTR'>"& RS1("VoidedDate") & "</span> »«ÿ· ‘œÂ «” .")
			response.end
		end if
	end if
	
	mySQL="UPDATE Invoices SET Voided=1, VoidedDate=N'"& shamsiToday() & "', VoidedBy='"& session("ID") & "' WHERE (ID='"& InvoiceID & "')"
	conn.Execute(mySQL)
	
	if isReverse then
		isCredit=1
		itemType=4 
		itemTypeName="›«ﬂ Ê— »—ê‘  «“ ›—Ê‘"
	else
		isCredit=0
		itemType=1
		itemTypeName="›«ﬂ Ê—"
		'---------- Declaring the related orders as Open  -------------------
		'mySQL="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice= '" & InvoiceID & "'))"
		'Changed By Kid ! 840509 
		'set orders which are ONLY related to this invoice, "Open"
		'that means, orders which are related to this invoice and are NOT related to any OTHER issued invoices.
		mySQL ="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Order] NOT IN (SELECT InvoiceOrderRelations.[ORDER] FROM Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "'))))"
		conn.Execute(mySQL)

		'---------- Declaring the related Quotes as Open  -------------------
		'set Quotes which are ONLY related to this invoice, "Open"
		'that means, Quotes which are related to this invoice and are NOT related to any OTHER issued invoices.
		mySQL ="UPDATE Quotes SET Closed=0 WHERE ID IN (SELECT [Quote] FROM InvoiceQuoteRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Quote] NOT IN (SELECT InvoiceQuoteRelations.[Quote] FROM Invoices INNER JOIN InvoiceQuoteRelations ON Invoices.ID = InvoiceQuoteRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "'))))"
		conn.Execute(mySQL)

	end if

  '**************************** Voiding ARItem of Invoice / Reverse Invoice ****************
  '*** Type = 1 means ARItem is an Invoice
  '*** Type = 4 means ARItem is a Reverse Invoice
  '***
	'*********  Finding the ARItem of Invoice / Reverse Invoice
	mySQL="SELECT ID FROM ARItems WHERE (Type = '"& itemType & "') AND (Link='"& InvoiceID & "')"
	Set RS1=conn.Execute(mySQL)
	voidedARItem=RS1("ID")
	'*********  Finding other ARItems related to this Item
	if isReverse then
		mySQL="SELECT ID AS RelationID, DebitARItem, Amount FROM ARItemsRelations WHERE (CreditARItem = '"& voidedARItem & "')"
		Set RS1=conn.Execute(mySQL)
		Do While not (RS1.eof)
			'*********  Adding back the amount in the relation, to the credit ARItem ...
			conn.Execute("UPDATE ARItems SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("DebitARItem") & "')")

			'*********  Deleting the relation
			conn.Execute("DELETE FROM ARItemsRelations WHERE ID='"& RS1("RelationID") & "'")
			
			RS1.movenext
		Loop
	else
		mySQL="SELECT ID AS RelationID, CreditARItem, Amount FROM ARItemsRelations WHERE (DebitARItem = '"& voidedARItem & "')"
		Set RS1=conn.Execute(mySQL)
		Do While not (RS1.eof)
			'*********  Adding back the amount in the relation, to the credit ARItem ...
			conn.Execute("UPDATE ARItems SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("CreditARItem") & "')")

			'*********  Deleting the relation
			conn.Execute("DELETE FROM ARItemsRelations WHERE ID='"& RS1("RelationID") & "'")
			
			RS1.movenext
		Loop
	end if

	'*********  Voiding ARItem 
	conn.Execute("UPDATE ARItems SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& voidedARItem & "')")

	'**************************************************************
	'*				Affecting Account's AR Balance  
	'**************************************************************
	if isReverse then
		mySQL="UPDATE Accounts SET ARBalance = ARBalance - '"& invoiceFee & "' WHERE (ID='"& CustomerID & "')"
	else
		mySQL="UPDATE Accounts SET ARBalance = ARBalance + '"& invoiceFee & "' WHERE (ID='"& CustomerID & "')"
	end if

	conn.Execute(mySQL)
	
	'***
	'***---------------- End of  Voiding ARItem of Invoice / Reverse Invoice ----------------

	' Sending a Message to Issuer ...
	if trim(comment)<>"" then comment = chr(13) & chr(10) & "[" & comment & "]"
	MsgTo			=	issuedBy
	msgTitle		=	"Invoice Voided"
	msgBody			=	"›«ﬂ Ê— ›Êﬁ  Ê”ÿ "& session("CSRName") & " »«ÿ· ‘œ." & comment
	RelatedTable	=	"invoices"
	relatedID		=	invoiceID
	replyTo			=	0
	IsReply			=	0
	urgent			=	1
	MsgFrom			=	session("ID")
	MsgDate			=	shamsiToday()
	MsgTime			=	currentTime10()
	Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")


	' Copying the PreInvoice Data...
	response.redirect "InvoiceInput.asp?act=copyInvoice&invoice="& InvoiceID & "&msg=" & Server.URLEncode(itemTypeName & " ‘„«—Â "& InvoiceID & " »«ÿ· ‘œ.")
'-----------------------------------------------------------------------------------------------
'--------------------------------------------S A M----------------------------------------------
'-----------------------------------------------------------------------------------------------
elseif request("act")="voidInvoiceOnly" then
	if not Auth(6 , "M") then		
		'Doesn't have the Priviledge to VOID the Invoice 
		response.write "<br>" 
		call showAlert ("‘„« „Ã«“ »Â «»ÿ«· ›«ﬂ Ê— ‰Ì” Ìœ",CONST_MSG_ERROR) 

		response.end
	end if

	comment=sqlSafe(request("comment"))

	InvoiceID=request("invoice")
	if InvoiceID="" or not(isnumeric(InvoiceID)) then
		response.write "<br>" 
		call showAlert ("Œÿ« œ— ‘„«—Â ›«ﬂ Ê—",CONST_MSG_ERROR) 
		response.end
	end if

	InvoiceID=clng(InvoiceID)
	
	mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		ShowErrorMessage("ÅÌœ« ‰‘œ ")
		response.end
	else
		voided=			RS1("Voided")
		issued=			RS1("Issued")
		issuedBy=		RS1("IssuedBy")
		isReverse=		RS1("IsReverse")
		customerID=		RS1("Customer")
		invoiceFee=		RS1("TotalReceivable")
		IsA =			RS1("IsA")
		if voided then
			ShowErrorMessage("«Ì‰ ›«ﬂ Ê— ﬁ»·« œ—  «—ÌŒ <span dir='LTR'>"& RS1("VoidedDate") & "</span> »«ÿ· ‘œÂ «” .")
			response.end
		end if
		if isReverse then 
			itemType=4 
		else
			itemType=1
		end if
		set myRS=conn.Execute("SELECT * FROM ARItems WHERE (Type = '"& itemType & "') AND (Link='"& InvoiceID & "') AND GL_Update=1")
		if myRS.eof then 
			ShowErrorMessage("»—«Ì «Ì‰ ›«ﬂ Ê— ﬁ»·« ”‰œ Õ”«»œ«—Ì ’«œ— ‘œÂ «”  ›·–« ‘„« Õﬁ «»ÿ«· ¬‰—« ‰œ«—Ìœ")
			response.end
		end if
	end if
	
	mySQL="UPDATE Invoices SET Voided=1, VoidedDate=N'"& shamsiToday() & "', VoidedBy='"& session("ID") & "' WHERE (ID='"& InvoiceID & "')"
	conn.Execute(mySQL)
	
	if isReverse then
		isCredit=1
		itemType=4 
		itemTypeName="›«ﬂ Ê— »—ê‘  «“ ›—Ê‘"
	else
		isCredit=0
		itemType=1
		itemTypeName="›«ﬂ Ê—"
		'---------- Declaring the related orders as Open  -------------------
		'mySQL="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice= '" & InvoiceID & "'))"
		'Changed By Kid ! 840509 
		'set orders which are ONLY related to this invoice, "Open"
		'that means, orders which are related to this invoice and are NOT related to any OTHER issued invoices.
		mySQL ="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Order] NOT IN (SELECT InvoiceOrderRelations.[ORDER] FROM Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "'))))"
		conn.Execute(mySQL)

		'---------- Declaring the related Quotes as Open  -------------------
		'set Quotes which are ONLY related to this invoice, "Open"
		'that means, Quotes which are related to this invoice and are NOT related to any OTHER issued invoices.
		mySQL ="UPDATE Quotes SET Closed=0 WHERE ID IN (SELECT [Quote] FROM InvoiceQuoteRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Quote] NOT IN (SELECT InvoiceQuoteRelations.[Quote] FROM Invoices INNER JOIN InvoiceQuoteRelations ON Invoices.ID = InvoiceQuoteRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "'))))"
		conn.Execute(mySQL)

	end if

  '**************************** Voiding ARItem of Invoice / Reverse Invoice ****************
  '*** Type = 1 means ARItem is an Invoice
  '*** Type = 4 means ARItem is a Reverse Invoice
  '***
	'*********  Finding the ARItem of Invoice / Reverse Invoice
	mySQL="SELECT ID FROM ARItems WHERE (Type = '"& itemType & "') AND (Link='"& InvoiceID & "')"
	Set RS1=conn.Execute(mySQL)
	voidedARItem=RS1("ID")
	'*********  Finding other ARItems related to this Item
	if isReverse then
		mySQL="SELECT ID AS RelationID, DebitARItem, Amount FROM ARItemsRelations WHERE (CreditARItem = '"& voidedARItem & "')"
		Set RS1=conn.Execute(mySQL)
		Do While not (RS1.eof)
			'*********  Adding back the amount in the relation, to the credit ARItem ...
			conn.Execute("UPDATE ARItems SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("DebitARItem") & "')")

			'*********  Deleting the relation
			conn.Execute("DELETE FROM ARItemsRelations WHERE ID='"& RS1("RelationID") & "'")
			
			RS1.movenext
		Loop
	else
		mySQL="SELECT ID AS RelationID, CreditARItem, Amount FROM ARItemsRelations WHERE (DebitARItem = '"& voidedARItem & "')"
		Set RS1=conn.Execute(mySQL)
		Do While not (RS1.eof)
			'*********  Adding back the amount in the relation, to the credit ARItem ...
			conn.Execute("UPDATE ARItems SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("CreditARItem") & "')")

			'*********  Deleting the relation
			conn.Execute("DELETE FROM ARItemsRelations WHERE ID='"& RS1("RelationID") & "'")
			
			RS1.movenext
		Loop
	end if

	'*********  Voiding ARItem 
	conn.Execute("UPDATE ARItems SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& voidedARItem & "')")

	'**************************************************************
	'*				Affecting Account's AR Balance  
	'**************************************************************
	if isReverse then
		mySQL="UPDATE Accounts SET ARBalance = ARBalance - '"& invoiceFee & "' WHERE (ID='"& CustomerID & "')"
	else
		mySQL="UPDATE Accounts SET ARBalance = ARBalance + '"& invoiceFee & "' WHERE (ID='"& CustomerID & "')"
	end if

	conn.Execute(mySQL)
	
	'***
	'***---------------- End of  Voiding ARItem of Invoice / Reverse Invoice ----------------

	' Sending a Message to Issuer ...
	if trim(comment)<>"" then comment = chr(13) & chr(10) & "[" & comment & "]"
	MsgTo			=	issuedBy
	msgTitle		=	"Invoice Voided"
	msgBody			=	"›«ﬂ Ê— ›Êﬁ  Ê”ÿ "& session("CSRName") & " »«ÿ· ‘œ." & comment
	RelatedTable	=	"invoices"
	relatedID		=	invoiceID
	replyTo			=	0
	IsReply			=	0
	urgent			=	1
	MsgFrom			=	session("ID")
	MsgDate			=	shamsiToday()
	MsgTime			=	currentTime10()
	Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")


	' Copying the PreInvoice Data...
	response.redirect "InvoiceInput.asp?act=copyInvoice&invoice="& InvoiceID & "&msg=" & Server.URLEncode(itemTypeName & " ‘„«—Â "& InvoiceID & " »«ÿ· ‘œ.")
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------

elseif request("act")="removePreInvoice" then
	response.write "<br>" 

	if not Auth(6 , "G") then		
		'Doesn't have the Priviledge to REMOVE the Pre-Invoice 
		call showAlert ("‘„« „Ã«“ »Â Õ–› ÅÌ‘ ‰ÊÌ” ‰Ì” Ìœ",CONST_MSG_ERROR) 
		response.end
	end if

	InvoiceID=request("invoice")
	if InvoiceID="" or not(isnumeric(InvoiceID)) then
		call showAlert ("Œÿ« œ— ‘„«—Â ÅÌ‘ ‰ÊÌ”",CONST_MSG_ERROR) 
		response.end
	end if
	InvoiceID=clng(InvoiceID)

	mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		call showAlert ("ÅÌ‘ ‰ÊÌ” ÅÌœ« ‰‘œ.",CONST_MSG_ERROR) 
	else
		voided=			RS1("Voided")
		issued=			RS1("Issued")
		isReverse=		RS1("IsReverse")
		customerID=		RS1("Customer")
		IsA =			RS1("IsA")
		if issued then
			call showAlert ("«Ì‰ ›«ﬂ Ê— ’«œ— ‘œÂ «” .",CONST_MSG_ERROR) 
		elseif voided then
			call showAlert ("«Ì‰ ›«ﬂ Ê— »«ÿ· ‘œÂ «” .",CONST_MSG_ERROR) 
		else
			Conn.execute("DELETE FROM InvoiceOrderRelations Where (Invoice='" & InvoiceID & "')")

			Conn.execute("DELETE FROM InvoiceQuoteRelations Where (Invoice='" & InvoiceID & "')")

			'Conn.execute("DELETE FROM InvoiceLines Where (Invoice='" & InvoiceID & "')")
			'Conn.execute("DELETE FROM Invoices Where (ID='" & InvoiceID & "')")
			'Changed By Kid 830929
			'Conn.execute("UPDATE Invoices SET Voided=1 Where (ID='" & InvoiceID & "')")
			'Changed By Kid 8400502 also adding VoidedBy & VoidedDate
			Conn.execute("UPDATE Invoices SET Voided=1, VoidedDate=N'"& shamsiToday() & "', VoidedBy='"& session("ID") & "' WHERE (ID='"& InvoiceID & "')")

			call showAlert ("ÅÌ‘ ‰ÊÌ” Õ–› ‘œ.",CONST_MSG_INFORM) 
		end if
	end if
	response.end
end if
conn.Close
%>
<!--#include file="include_JS_for_Invoices.asp" -->
<SCRIPT LANGUAGE="JavaScript">
<!--
function ApproveInvoice(){
	if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« ' «ÌÌœ' ﬂ‰Ìœø\n\n( ÊÃÂ:  €ÌÌ—«  –ŒÌ—Â ‰„Ì ‘Ê‰œ)\n"))
		window.location="?act=approveInvoice&invoice=<%=InvoiceID%>";
}
function IssueInvoice(){
	if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '’«œ—' ﬂ‰Ìœø\n\n( ÊÃÂ:  €ÌÌ—«  –ŒÌ—Â ‰„Ì ‘Ê‰œ)\n"))
		window.location="?act=IssueInvoice&invoice=<%=InvoiceID%>";
}
function VoidInvoice(){
	if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '»«ÿ·' ﬂ‰Ìœø\n"))
		window.location="?act=voidInvoice&invoice=<%=InvoiceID%>";
}
//-->
</SCRIPT>
<!--#include file="tah.asp" -->
