<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="���� ������"
SubmenuItem=1
if not (Auth(6 , 1) OR Auth(6 , 4)) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
function ShowErrorMessage(msg)
	response.write "<table align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> ��� ! <br>"& msg & "<br></td></tr></table><br>"
end function

function Link2Trace(OrderNo)
	Link2Trace = "<A HREF='../order/TraceOrder.asp?act=show&order="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
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

if request("act")="submitsearch" then
	if trim(request("query")) <> "" then
		if isnumeric(request("query")) then
			SO_Order=clng(request("query"))
			SO_Action="return true;"
			SO_StepText="<br>��� ��� : ������ ����� ��� ������"
%>
			<FORM METHOD=POST ACTION="?act=getInvoice">
			<!--#include File="include_SelectOrdersByOrder.asp"-->
			</FORM>
<%if Auth(6 , "K") then ' Has the priviledge to create a ReverseInvoice without an ORDER %>
			<p align='center' style='font-size:9pt;' dir='RTL'><a href='?act=getInvoice&selectedCustomer=<%=SO_Customer%>'>����� �� ����� ���� ����...</a></p>
<%end if%>
			<br>
<%		else
			SA_TitleOrName=request("query")
			SA_Action="return true;"
			SA_SearchAgainURL="InvoiceInput.asp"
			SA_StepText="��� ��� : ������ ����"
%>
			<FORM METHOD=POST ACTION="?act=selectOrder">
			<!--#include File="include_SelectAccount.asp"-->
			</FORM>
<%
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("���� �э �������!")
	end if
elseif request("act")="selectOrder" then
	if request("selectedCustomer") <> "" then
		SO_Customer=request("selectedCustomer")
		SO_Action="return true;"
		SO_StepText="<br>��� ��� : ������ ����� ��� ������"
%>
		<FORM METHOD=POST ACTION="?act=getInvoice">
		<!--#include File="include_SelectOrder.asp"-->
		</FORM>
<%if Auth(6 , "K") then ' Has the priviledge to create a ReverseInvoice without an ORDER %>
			<p align='center' style='font-size:9pt;' dir='RTL'><a href='?act=getInvoice&selectedCustomer=<%=SO_Customer%>'>����� �� ����� ���� ����...</a></p>
<%end if%>
		<br>
<%
	end if
elseif request("act")="getInvoice" then
%>
<!--#include file="include_JS_for_Invoices.asp" -->
<%
	'******************** Checking Input ****************
	errorFound=false
	ON ERROR RESUME NEXT

		customerID=	clng(request("selectedCustomer"))

		mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
		Set rs = conn.Execute(mySQL)
		if rs.eof then
			errorFound=True
		else
			customerName=rs("AccountTitle")
			isAdefault=rs("IsADefault")
		end if
		rs.close
		
		if Err.Number<>0 then
			Err.clear
			errorFound=True
		end if
	ON ERROR GOTO 0

	if errorFound then
		response.write "<br>" 
		call showAlert ("��� �� �����",CONST_MSG_ERROR) 
		response.end
	end if
	'^^^^---------------- Checking Input ------------^^^^

	
	creationDate=shamsiToday()
'	creationTime=Hour(creationTime)&":"&Minute(creationTime)
'	if instr(creationTime,":")<3 then creationTime="0" & creationTime
'	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)

	InvoiceLinesNo=1
%>
<!-- ���� ������� ������ -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<input type="hidden" name="VatRate" id="VatRate" value="<%=session("VatRate")%>">
		<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
		<FORM METHOD=POST ACTION="?act=submitInvoice">
			<tr bgcolor='#C3C300'>
			<td colspan="2"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL"><TR>
				<TD><table>
					<tr>
						<td align="left">����:</td>
						<td align="right">
							<span id="customer"><%' after any changes in this span "./Customers.asp" must be revised%>
								<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><span><%=CustomerName%></span>.
							</span></td>
						<td><INPUT class="InvGenButton" TYPE="button" value="�����" onClick="selectCustomer();"></td>
					</tr>
					</table></TD>
				<TD align="left"><table>
					<tr>
						<td align="left">�����:</td>
						<td dir="LTR">
							<INPUT class="InvGenInput" NAME="InvoiceDate" TYPE="text" maxlength="10" size="10" value="<%=CreationDate%>"></td>
						<td dir="RTL"><%=weekdayname(weekday(date))%></td>
					</tr>
					</table></TD>
				</TR></TABLE>
			</td>
			</tr>
			<tr bgcolor='#C3C300'>
				<td align="right" width="50%">
					&nbsp; ����� �� �����(���):
					<span id="orders">
<%
					tempWriteAnd=""
					for i=1 to request("selectedOrders").count
						response.write "<input type='hidden' name='selectedOrders' value='"& request("selectedOrders")(i)& "'>"
						response.write tempWriteAnd & Link2Trace(request("selectedOrders")(i))
						tempWriteAnd=" � "
					next
%>
					</span>&nbsp;
					<input class="InvGenButton" TYPE="button" value="�����" onClick="selectOrder();">
				</td>
				<td align="left"><TABLE>
					<TR>
						<TD align="left">�����:</TD>
						<TD dir="LTR">
							<INPUT class="InvGenInput" NAME="InvoiceNo" style="border:1px solid black;" TYPE="text" maxlength="10" size="10"></TD>
						<TD dir="RTL" <% if not Auth(6,"N") then response.write "title='��� ���� �� ����� ��� ���� ������!'"%>>
							<INPUT TYPE="checkbox" 
							<% if not Auth(6,"N") then 
									response.write " onclick='this.checked="
									if IsADefault then 
										response.write "true;'"
									else
										response.write "false;'"
									end if
								else 
									response.write " onClick='checkIsA();'"
								end if
				    	%> <% if IsADefault then response.write " checked" %> NAME="IsA"> ��� &nbsp;
						</TD>
					</TR>
					</TABLE></td>
			</tr>
<%		If request("selectedQuotes").count > 0 Then %>
			<tr bgcolor='#AAAAEE' height="30">
				<td align="right" width="50%" colspan="2">
					&nbsp; ����� �� ������� :
					<span id="quotes">
<%
					tempWriteAnd=""
					for i=1 to request("selectedQuotes").count
						response.write "<input type='hidden' name='selectedQuotes' value='"& request("selectedQuotes")(i)& "'>"
						response.write tempWriteAnd & Link2TraceQuote(request("selectedQuotes")(i))
						tempWriteAnd=" � "
					next
%>
					</span>&nbsp;
					<!--input class="InvGenButton" TYPE="button" value="�����" onClick="selectQuote();"-->
				</td>
			</tr>
<%		End If %>
			<tr bgcolor='#CCCC88'>
			<TD colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<tr bgcolor='#CCCC88'>
				<td align='center' width="32px"> # </td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" value="����" size="3" ></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" value="�������" size="30"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="���" size="2"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="���" size="2"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="�����" size="3"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="���" size="2"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="����� ����" size="6"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="��" size="7"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="����" size="9"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="�����"size="7"></td><!-- S A M -->
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="�ѐ��" size="5"></td><!-- S A M -->
				<td><INPUT class="InvHeadInput4" readonly TYPE="text" Value="������" size="6"></td><!-- S A M -->
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="���� ������" size="9"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
			<tr bgcolor='#CCCC88'>
			<TD colspan="10"><div style="overflow:auto; height:250px; width:*;">
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<Tbody id="InvoiceLines">

<%		
		for i=1 to 1
%>
			<tr bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
				<td align='center' width="25px"><%=i%></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Items" size="3" Maxlength="6" onKeyPress="return mask(this,event);" onfocus="setCurrentRow(this.parentNode.parentNode.rowIndex);" onChange='return check(this);'>
				<INPUT TYPE="hidden" name="type" value=0>
				<INPUT TYPE="hidden" name="fee" value=0>
				<INPUT type='hidden' name='hasVat' value=0>
				</td>
				<td dir="RTL"><INPUT class="InvRowInput2" TYPE="text" NAME="Descriptions" size="30"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Lengths" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Widths" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Qttys" size="3" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Sets" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="AppQttys" size="6" onBlur="setPrice(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Fees" size="7" onBlur="setPrice(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Prices" size="9" readonly tabIndex="9999"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Discounts" size="7" onBlur="setPrice(this);"></td><!-- S A M -->
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Reverses" size="5" onBlur="setPrice(this);" onfocus="setCurrentRow(this.parentNode.parentNode.rowIndex);"></td><!-- S A M -->
				<td dir="LTR"><INPUT class="InvRowInput4" TYPE="text" NAME="Vat" size="6" readonly></td><!-- S A M -->
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="AppPrices" size="9" readonly tabIndex="9999"></td>
			</tr>
<%
		next
%>
			<tr bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
				<td colspan="15">
					<INPUT class="InvGenButton" TYPE="button" value="�����" onkeyDown="if(event.keyCode==9) {setCurrentRow(this.parentNode.parentNode.rowIndex); return false;};" onClick="addRow();">
				</td>
			</tr>
			</Tbody></TABLE></div>
			</TD>
			</tr>
			<tr bgcolor='#CCCC88'>
			<TD colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#CCCC88">
			<tr bgcolor='#CCCC88'>
				
				
				<td colspan='9' width='500px'>***����� ��� ������ ��� �� ��� ��� ����� ��***</td>
				<!--td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="3"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="6"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="7"></td-->
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalPrice" TYPE="text" size="9"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalDiscount" TYPE="text" size="7"></td><!-- S A M -->
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalReverse" TYPE="text" size="5"></td><!-- S A M -->
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalVat" TYPE="test" size="6"></td><!-- S A M -->
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="Payable" TYPE="text" size="9"></td>
			</tr>
			<tr bgcolor='#CCCC88'>
				<td colspan="9"> &nbsp; </td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput" TYPE="text" size="9"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" Name="TPDiscount" size="7"></td><!-- S A M -->
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" Name="TPReverse" size="5"></td><!-- S A M -->
				<td dir="LTR"><INPUT readonly class="invHeadInput3" TYPE="text" size="6" value="<%=session("VatRate")%>������"></td><!-- S A M -->
				<td dir="LTR"><INPUT readonly class="InvHeadInput" TYPE="text" size="9" value="��� ���"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
		</table><br>
		<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
		<tr>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="��� ��� ���� ������..." onclick="submitOperations();"></td>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="������" onclick="window.close();"></td>
		</tr>
		</TABLE>
		</FORM>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			document.getElementsByName("Items")[0].focus();
		//-->
		</SCRIPT>
<%elseif request("act")="submitInvoice" then
	'******************** Checking Input ****************
	errorFound=false
	ON ERROR RESUME NEXT

		InvoiceDate=	request.form("InvoiceDate")

		If Not CheckDateFormat(InvoiceDate) Then
			errorFound=True
		end if

		CustomerID=		clng(request.form("CustomerID"))

		if request.form("IsA") = "on" then 
			IsA=1 
			InvoiceNo=request.form("InvoiceNo")
			if InvoiceNo <> "" then InvoiceNo = clng(InvoiceNo)
		else 
			IsA=0
			InvoiceNo=""
		end if
			
		'OLD Code:
		'TotalPrice=		cdbl(text2value(request.form("TotalPrice")))
		'TotalDiscount=	cdbl(text2value(request.form("TotalDiscount")))
		'TotalReverse=	cdbl(text2value(request.form("TotalReverse")))
		'TotalReceivable=cdbl(text2value(request.form("TotalAppPrice")))
		'
		'Changed By Kid 831021: calculating totals ServerSide
		TotalPrice	=		0
		TotalDiscount =		0
		TotalReverse =		0
		TotalReceivable =	0
		TotalVat =			0
		RFD	=				0

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
			theVat =			clng(text2value(request.form("Vat")(i)))

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
			TotalVat = TotalVat + theVat

		next 
		RFD = TotalReceivable - fix(TotalReceivable / 1000) * 1000
		'RFD = RFD / 1.03
		TotalReceivable = TotalReceivable - RFD
		TotalDiscount = TotalDiscount + RFD

		for i = 1 to request("selectedOrders").count
			theOrder=		clng(request("selectedOrders")(i))
			mySQL = "SELECT ID FROM Orders WHERE ID=" & theOrder 
			Set rs = conn.Execute(mySQL)
			if rs.eof then
				errorFound=True
				exit for
			end if
			rs.close
		next

		for i = 1 to request("selectedQuotes").count
			theOrder= clng(request("selectedQuotes")(i))
			mySQL = "SELECT ID FROM Quotes WHERE ID=" & theOrder 
			Set rs = conn.Execute(mySQL)
			if rs.eof then
				errorFound=True
				exit for
			end if
			rs.close
		next

		Set rs = Nothing

		if Err.Number<>0 then
			Err.clear
			errorFound=True
		end if
	ON ERROR GOTO 0

	if errorFound then
		response.write "<br>" 
		call showAlert ("��� �� �����",CONST_MSG_ERROR) 
		response.end
	end if
	'^^^^---------------- Checking Input ------------^^^^

	'**************************** Inserting new Invoice ****************
	mySQL="INSERT INTO Invoices (CreatedDate, CreatedBy, Customer, Number, TotalPrice, TotalDiscount, TotalReverse, TotalReceivable, IsA, TotalVat) VALUES (N'"& InvoiceDate & "', '"& session("ID") & "', '"& CustomerID & "', '"& InvoiceNo & "', '"& TotalPrice & "', '"& TotalDiscount & "', '"& TotalReverse & "', '"& TotalReceivable & "', '"& IsA & "', '" & TotalVat & "');SELECT @@Identity AS NewInvoice"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	InvoiceID = RS1 ("NewInvoice")
	RS1.close	

	'**************************** Inserting Invoice Lines ****************
	for i=1 to request.form("Items").count 
		theItem =			clng(text2value(request.form("Items")(i)))
		theDescription =	left(sqlSafe(request.form("Descriptions")(i)),100)

		theAppQtty =		cdbl(text2value(request.form("AppQttys")(i)))
		thePrice =			clng(text2value(request.form("Prices")(i)))

		theDiscount =		text2value(request.form("Discounts")(i))
		theReverse =		text2value(request.form("Reverses")(i))
		theVat =			text2value(request.form("Vat")(i))
		theHasVat =			text2value(request.form("hasVat")(i))

		theLength =			text2value(request.form("Lengths")(i))
		theWidth =			text2value(request.form("Widths")(i))
		theQtty =			text2value(request.form("Qttys")(i))
		theSets =			text2value(request.form("Sets")(i))

		if theDiscount <>"" then theDiscount= clng(theDiscount)
		if theReverse <> "" then theReverse = clng(theReverse)

		if theLength <>	"" then  theLength	= cdbl(theLength)
		if theWidth <> ""  then	 theWidth	= cdbl(theWidth) 
		if theQtty <> ""   then	 theQtty	= clng(theQtty)
		if theSets <> ""   then	 theSets	= clng(theSets)

		mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat) VALUES ('"& InvoiceID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "', "&theHasVat&")"
		conn.Execute(mySQL)
	next 
	if RFD > 0 then
		theItem =			39999
		theDescription =	"����� ��� ������"

		theAppQtty =		0
		thePrice =			0

		theDiscount =		RFD
		theReverse =		0
		theVat =			0

		theLength =			0
		theWidth =			0
		theQtty =			0
		theSets =			0
		mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat) VALUES ('"& InvoiceID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "')"
		conn.Execute(mySQL)
	end if

	'**************************** Making Invoice-Order Relations ****************
	for i=1 to request.form("selectedOrders").count
		theOrder = clng(request.form("selectedOrders")(i))
		mySQL="INSERT INTO InvoiceOrderRelations (Invoice,[Order]) VALUES ('" & InvoiceID & "', '" & theOrder & "')"
		conn.Execute(mySQL)
	next

	'**************************** Making Invoice-Quote Relations ****************
	for i=1 to request.form("selectedQuotes").count
		theQuote = clng(request.form("selectedQuotes")(i))
		mySQL="INSERT INTO InvoiceQuoteRelations (Invoice,Quote) VALUES ('" & InvoiceID & "', '" & theQuote & "')"
		conn.Execute(mySQL)
	next

	conn.close
	response.redirect "AccountReport.asp?act=showInvoice&invoice="	 & InvoiceID

elseif request("act")="copyInvoice" then
	InvoiceID=request.queryString("invoice")
	if InvoiceID="" OR not isnumeric(InvoiceID) then
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("����� ������ ���� ���� ��� ����.")
	end if

	InvoiceID=clng(InvoiceID)
	mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("������ ���� ���.")
	end if

	customerID=		RS1("Customer")
	creationDate=	shamsiToday()
	totalPrice=		cdbl(RS1("totalPrice"))
	totalDiscount=	cdbl(RS1("totalDiscount"))
	totalReverse=	cdbl(RS1("totalReverse"))
	totalVat =		cdbl(RS1("totalVat"))
	TotalReceivable = totalPrice - totalDiscount - totalReverse + totalVat
	Issued =		RS1("Issued") 
	Voided =		RS1("Voided") 

	if RS1("IsReverse") then
		isReverse=	1
	else
		isReverse=	0
	end if
	if RS1("IsA") then
		IsA=		1
	else
		IsA=		0
	end if
	InvoiceNo=		RS1("Number")

	RS1.close

	if isReverse AND not Auth(6 , 4) then
		'Doesn't have the permission for ADDING Rev. Invoice
		response.write "<br>" 
		call showAlert ("��� ���� �� ���� ������ �ѐ�� ������",CONST_MSG_ERROR) 
		response.end
	elseif not Auth(6 , 1) then 
		'Doesn't have the permission for ADDING Invoice
		response.write "<br>" 
		call showAlert ("��� ���� �� ���� ������ ������",CONST_MSG_ERROR) 
		response.end
	end if

	'Copying Invoice ...
	mySQL="INSERT INTO Invoices (CreatedDate, CreatedBy, Customer, IsA, Number, TotalPrice, TotalDiscount, TotalReverse, TotalReceivable, IsReverse, TotalVat) VALUES (N'"& creationDate & "', '"& session("ID") & "', '"& CustomerID & "', '"& IsA & "', '"& InvoiceNo & "', '"& TotalPrice & "', '"& TotalDiscount & "', '"& TotalReverse & "', '"& TotalReceivable & "', '"& IsReverse & "', '" & TotalVat & "'); SELECT @@Identity AS NewInvoice"
	Set RS1 = conn.Execute(mySQL).NextRecordSet
	NewInvoice = RS1("NewInvoice")

	'Copying InvoiceLines ...
	mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat) SELECT '"& NewInvoice & "' AS Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat FROM InvoiceLines WHERE (Invoice = '"& InvoiceID & "')"
	conn.Execute(mySQL)

	'Copying InvoiceOrderRelations ...
	mySQL="INSERT INTO InvoiceOrderRelations (Invoice, [Order]) SELECT '"& NewInvoice & "' AS Invoice, [Order] FROM InvoiceOrderRelations WHERE (Invoice = '"& InvoiceID & "')"
	conn.Execute(mySQL)

	'Copying InvoiceQuoteRelations ...
	mySQL="INSERT INTO InvoiceQuoteRelations (Invoice, [Quote]) SELECT '"& NewInvoice & "' AS Invoice, [Quote] FROM InvoiceQuoteRelations WHERE (Invoice = '"& InvoiceID & "')"
	conn.Execute(mySQL)

	if Issued AND voided then
	' if it's copied from a voided invoice, place a comment for the voided invoice 
		MsgTo			=	0
		msgTitle		=	"Invoice Copied"
		msgBody			=	"������ ����� "& NewInvoice & " �� ��� ������ ��� ߁� ��." 
		RelatedTable	=	"invoices"
		relatedID		=	InvoiceID
		replyTo			=	0
		IsReply			=	0
		urgent			=	0
		MsgFrom			=	session("ID")
		MsgDate			=	shamsiToday()
		MsgTime			=	currentTime10()
		Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")
		
	end if	

	' Place a comment for the new invoice 
	MsgTo			=	0
	msgTitle		=	"Invoice Copied"
	msgBody			=	"��� ���� �� ������ ����� "& InvoiceID & " ߁� ��." 
	RelatedTable	=	"invoices"
	relatedID		=	NewInvoice 
	replyTo			=	0
	IsReply			=	0
	urgent			=	0
	MsgFrom			=	session("ID")
	MsgDate			=	shamsiToday()
	MsgTime			=	currentTime10()
	Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")

	response.redirect "AccountReport.asp?act=showInvoice&invoice=" & NewInvoice & "&msg=" & Server.URLEncode(request.queryString("msg") & "<br>��� ���� �� ��� ������ �� ������� ����� "& InvoiceID & " ߁� ��.")

end if
conn.Close
%>
<!--#include file="tah.asp" -->
