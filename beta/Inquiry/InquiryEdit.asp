<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="����� ������ "
SubmenuItem=1
'if not Auth(6 , 1) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
function ShowErrorMessage(msg)
	response.write "<table align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> ��� ! <br>"& msg & "<br></td></tr></table><br>"
end function

function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/TraceOrder.asp?act=show&order="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
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
		mySQL="SELECT InquiryOrderRelations.Inquiry FROM InquiryOrderRelations INNER JOIN Inquiries ON InquiryOrderRelations.Inquiry = Inquiries.ID WHERE (InquiryOrderRelations.[Order] = '"& OrderID & "') AND (Inquiries.IsReverse = 0) AND (Inquiries.Voided = 0)"
		Set RS1=Conn.Execute(mySQL)
		if RS1.eof then
			Conn.close
			response.redirect "?errmsg=" & Server.URLEncode("��� ����� ���� ������ ���� ���.")
		else
			theInquiry=RS1("Inquiry")
			Conn.close
			response.redirect "?act=editInquiry&inquiry=" & theInquiry
		end if
	elseif isnumeric(request("inquiry")) then
		response.redirect "?act=editInquiry&inquiry=" & request("inquiry")
	else
		response.redirect "?errmsg=" & Server.URLEncode("����� ����� ���� ���� ��� ����.")
	end if
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------- Edit Inquiry
'-----------------------------------------------------------------------------------------------------
elseif request("act")="editInquiry" then
	if isnumeric(request("inquiry")) then
		InquiryID=clng(request("inquiry"))
		mySQL="SELECT * FROM Inquiries WHERE (ID='"& InquiryID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("������ ���� ���.")
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("����� ������ ���� ���� ��� ����.")
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
	InquiryNo=		RS1("Number")

	mySQL="SELECT ID,AccountTitle FROM Accounts WHERE (ID='"& customerID & "')"
	Set RS1 = conn.Execute(mySQL)
	AccountNo=RS1("ID")
	customerName=RS1("AccountTitle")

	RS1.close

	if isReverse then
		'Check for permission for EDITTING Rev. Inquiry
		if not Auth(6 , 5) then NotAllowdToViewThisPage()
		itemTypeName="�ǘ��� �ѐ��"
		HeaderColor="#FF9900"
	else
		'Check for permission for EDITTING Inquiry
		if not Auth(6 , 3) then NotAllowdToViewThisPage()
		itemTypeName="�ǘ���"
		HeaderColor="#C3C300"
	end if

	if Voided then
		Conn.close
		response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ��� ���.")
	elseif Issued then
		if Auth(6 , "A") then
			' Has the Priviledge to change the Inquiry
			response.write "<BR>"
			call showAlert ("��� ������ ���� ��� ���.<br>�э�� �� ��� ����� ����� �� ��� �ǘ��� �� ����� �����<br>������ �Ԙ��� ������� �� �� �����Ͽ",CONST_MSG_INFORM) 
		else
			Conn.close
			response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ��� ���.")
		end if
	end if 
%>
<!-- ���� ������� ������ -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
		<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
		<FORM METHOD=POST ACTION="?act=submitEdit">
			<tr bgcolor='<%=HeaderColor%>'>
			<td colspan='2'>
				<TABLE width='100%'>
				<TR>
					<TD align="left" >����� <%=itemTypeName%>:</TD>
					<TD align="right" width='15%'>&nbsp;<INPUT readonly class="InvGenInput" NAME="InquiryID" value="<%=InquiryID%>" style="direction:ltr" TYPE="text" maxlength="10" size="10"></TD>
				</TR>
				</TABLE></td>
			</tr>
			<tr bgcolor='<%=HeaderColor%>'>
			<td colspan="2">
				<TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL">
				<TR>
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
							<td align="left">����� ����:</td>
							<td dir="LTR">
								<INPUT class="InvGenInput" NAME="issueDate" TYPE="text" maxlength="10" size="10" value="<%=IssuedDate%>"  onblur="acceptDate(this)"></td>
							<td dir="RTL"><%="�����"%></td>
						</tr>
						</table></TD>
				</TR></TABLE>
			</td>
			</tr>
			<tr bgcolor='<%=HeaderColor%>'>
				<TD align="right" width="50%">
					����� �� �����(���):
					<span id="orders">
<%
					tempWriteAnd=""

					mySQL="SELECT * FROM InquiryOrderRelations WHERE (Inquiry='"& InquiryID & "')"
					Set RS1 = conn.Execute(mySQL)
					while not(RS1.eof) 
						response.write "<input type='hidden' name='selectedOrders' value='"& RS1("Order") & "'>"
						response.write tempWriteAnd & Link2Trace(RS1("Order"))
						tempWriteAnd=" � "
						RS1.moveNext
					wend
%>					</span>&nbsp;
					<INPUT class="InvGenButton" TYPE="button" value="�����" onClick="selectOrder();">
				</TD>
				<TD align="left"><table>
					<tr>
						<td align="left">�����:</td>
						<td dir="LTR">
							<INPUT class="InvGenInput" NAME="InquiryNo" value="<%=InquiryNo%>" style="border:1px solid black;" TYPE="text" maxlength="10" size="10"></td>
						<td dir="RTL"><INPUT TYPE="checkbox" onClick="checkIsA();" NAME="IsA" <%if IsA then response.write "checked"%>> ��� &nbsp;</td>
					</tr>
					</table></TD>
			</tr>
			<tr bgcolor='#CCCC88'>
			<TD colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<tr bgcolor='#CCCC88'>
				<td align='center' width="25px"> # </td>
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
			<TD colspan="10"><div style="overflow:auto; height:190px; width:*;">
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<Tbody id="InquiryLines">

<%		
	i=0
	mySQL="SELECT * FROM InquiryLines LEFT OUTER JOIN invoiceItems ON InquiryLines.item = invoiceItems.id WHERE (Inquiry='"& InquiryID & "') "
	Set RS1 = conn.Execute(mySQL)
	while not(RS1.eof) 
	if RS1("Item") <> 39999 then
		i=i+1
%>
			<tr bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
				<td align='center' width="25px"><%=i%></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Items" value="<%=RS1("Item")%>" size="3" Maxlength="6" onKeyPress="return mask(this);" onfocus="setCurrentRow(this.parentNode.parentNode.rowIndex);" onChange='return check(this);'>
				<INPUT TYPE="hidden" name="type" value="<%=RS1("type")%>">
				<INPUT TYPE="hidden" name="fee" value="<%=RS1("fee")%>">
				<input type='hidden' name='hasVat' value='<%=text2value(RS1("hasVat"))%>'>
				</td>
				<td dir="RTL"><INPUT class="InvRowInput2" TYPE="text" NAME="Descriptions" value="<%=RS1("Description")%>" size="30"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Lengths" value="<%=RS1("Length")%>" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Widths" value="<%=RS1("Width")%>" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Qttys" value="<%=RS1("Qtty")%>" size="3" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Sets" value="<%=RS1("Sets")%>" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="AppQttys" value="<%=Separate(RS1("AppQtty"))%>" size="6" onBlur="setPrice(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Fees" value="<%if RS1("AppQtty") <> 0 then response.write Separate(RS1("Price")/RS1("AppQtty")) else response.write "0"%>" size="7" onBlur="setPrice(this);"></td>
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
				<td colspan='9' width='500px'>***����� ��� �ǘ��� ��� �� ��� ��� ����� ��***</td>
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
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" Name="TPDiscount" value="<%=Pourcent(totalDiscount,totalPrice) & "%�����"%>" size="7"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" Name="TPReverse" value="<%=Pourcent(totalReverse,totalPrice) & "%�ѐ��"%>" size="5"></td>
				<td dir="LTR"><INPUT readonly calss="InvHeadINput" TYPE="text" size="6" value="3%������"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput" TYPE="text" size="9" value="��� ���"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
		</table>
		<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
		<tr>
			<td align='center'>&nbsp;<!-- <INPUT class="InvGenButton" TYPE="button" value="����� ������" onclick="ApproveInvoice();"> --></td>
			<td width="40">&nbsp;</td>
			<td align='center'>&nbsp;<!-- <INPUT class="InvGenButton" TYPE="button" value="���� ������" onclick="IssueInvoice();"> --></td>
			<td width="40">&nbsp;</td>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="����� " onclick="submitOperations();"></td>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="������" onclick="window.location='AccountReport.asp?act=showInquiry&inquiry=<%=InquiryID%>';"></td>
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

		InquiryID=		clng(request.form("InquiryID"))
		CustomerID=		clng(request.form("CustomerID"))

		issueDate=	request.form("issueDate")

		if request.form("IsA") = "on" then 
			IsA=1 
			InquiryNo=request.form("InquiryNo")
			if InquiryNo <> "" then InquiryNo = clng(InquiryNo)
		else 
			IsA=0
			InquiryNo=""
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

		Set rs= Nothing

		mySQL="SELECT * FROM Inquiries WHERE (ID='"& InquiryID & "')"
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

		if NOT errorFound AND NOT issued then
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
		call showAlert ("��� �� �����",CONST_MSG_ERROR) 
		response.end
	end if
	'^^^^---------------- Checking Input ------------^^^^

	if isReverse then
		'Check for permission for EDITTING Rev. Invoice
		if not Auth(6 , 5) then NotAllowdToViewThisPage()
		itemType=4 
	else
		'Check for permission for EDITTING Inquiry
		if not Auth(6 , 3) then NotAllowdToViewThisPage()
		itemType=1
	end if

	if voided then
		Conn.close
		response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ���� ��� ���.")
	elseif issued then
		if Auth(6 , "A") then
			' Has the Priviledge to change the Inquiry / Reverse Inquiry

			'mySQL="SELECT ID FROM ARItems WHERE (Type='"& itemType & "') AND (GL_Update=1) AND (Link='"& InvoiceID & "')"
			'Changed by Kid ! 831124
			mySQL="SELECT ARItems.ID, ARItems.GL_Update, EffGLRows.GL, EffGLRows.GLDocID FROM ARItems LEFT OUTER JOIN (SELECT Link, GL, GLDocID FROM EffectiveGLRows WHERE SYS = 'AR') EffGLRows ON ARItems.ID = EffGLRows.Link WHERE (ARItems.Type = '"& itemType & "') AND (ARItems.Link = '"& InquiryID & "')"

			Set RS2 = conn.Execute(mySQL)

			if RS2.eof then
				Conn.close
				response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� !! <br><br> ���� ���.")
			else
				if RS2("GL_Update") = False  then
					tmpGL=RS2("GL")
					tmpGLDoc=RS2("GLDocID")
					Conn.close
					response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("���� ���� ��� ������ ��� �������� ���� ��� ���.<br><br>���� ��:"& tmpGL & " ��� �����: "& tmpGLDoc & " .")
				else
					ARItemID=RS2("ID")
					IssuedButEdit=true
				end if
			end if
			RS2.close
		else
			Conn.close
			response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ���� ��� ���.")
		end if
	elseif approved then 
		call UnApproveInquiry ( InquiryID , ApprovedBy )
	end if
	

	'******************* Editing  *******************
	' ****
	if IssuedButEdit then
		' Only Updating  IssuedDate, Number & IsA
		'				 and related Orders

		'---- Checking wether issueDate is valid in current open GL
		If Not CheckDateFormat(issueDate) Then
			Conn.close
			response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("����� ���� ��� ����� ����.")
		end if

		if (issueDate < session("OpenGLStartDate")) OR (issueDate > session("OpenGLEndDate")) then
			Conn.close
			response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("���!<br>����� ���� ��� ����� ����. <br>(�� ��� ���� ���� ����)")
		end if 
		'----

		mySQL="UPDATE Inquiries SET IssuedDate=N'" & issueDate & "', Number='"& InquiryNo & "', IsA='"& IsA & "' WHERE (ID='"& InquiryID & "')"
		conn.Execute(mySQL)

		'if IsA then
		'	GLAccount=	"91001"	'This must be changed... (Sales A)
		'else
		'	GLAccount=	"91002"	'This must be changed... (Sales B)
		'end if
		'
		' Changed By Kid 860118 , seasing to use Sales B

		GLAccount=	"91001"	'This must be changed... (Sales A)

		conn.Execute("UPDATE ARItems SET GL='"& OpenGL & "', EffectiveDate='" & issueDate & "', GLAccount='"& GLAccount & "' WHERE (ID='" & ARItemID & "')")

		'**************** Updating Inquiry-Order Relations ****************
		'mySQL="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice= '" & InvoiceID & "'))"
		'Changed By Kid ! 840509 
		'set orders which are ONLY related to this invoice, "Open"
		'that means, orders which are related to this invoice and are NOT related to any OTHER issued invoices.
		mySQL ="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InquiryOrderRelations WHERE (Inquiry = '" & InquiryID & "') AND ([Order] NOT IN (SELECT InquiryOrderRelations.[ORDER] FROM Inquiries INNER JOIN InquiryOrderRelations ON Inquiries.ID = InquirieOrderRelations.Inquiry WHERE (Inquiries.Issued = 1) AND (Inquiries.Voided = 0) AND (Inquiries.isReverse = 0) AND (Inquiries.ID <> '" & InquiryID & "'))))"
		conn.Execute(mySQL)

		mySQL ="DELETE FROM InquiryOrderRelations WHERE (Inquiry='" & InquiryID & "')"
		'mySQL ="DELETE FROM InvoiceOrderRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Order] NOT IN (SELECT InvoiceOrderRelations.[ORDER] FROM Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "')))"
		conn.Execute(mySQL)
		

		for i=1 to request.form("selectedOrders").count
			theOrder=	clng(request.form("selectedOrders")(i))
			mySQL="INSERT INTO InquiryOrderRelations (Inquiry,[Order]) VALUES ('" & InquiryID & "', '" & theOrder & "')"
			conn.Execute(mySQL)
		next

		conn.Execute("UPDATE Orders SET Closed=1 WHERE ID IN (SELECT [Order] FROM InquiryOrderRelations WHERE (Inquiry='" & InquiryID & "'))")
		'^^^^------------ Updating Inquiry-Order Relations ------------^^^^

		conn.close
		response.redirect "AccountReport.asp?act=showInquiry&inquiry=" & InquiryID & "&msg=" &Server.URLEncode("��� ������ ���� ���� ��� ���.<br>�э�� �� ��� ����� ����� �� ��� �ǘ��� �� ����� �����<br>���� ��� �� ��� ��� �� ʘ��� ����.")
	else
' S A M
'response.write(totalDiscount)
'response.end
		mySQL="UPDATE Inquiries SET Customer='"& CustomerID & "', Number='"& InquiryNo & "', TotalPrice='"& TotalPrice & "', TotalDiscount='"& TotalDiscount & "', TotalReverse='"& TotalReverse & "', TotalReceivable='"& TotalReceivable & "' , IsA='"& IsA & "', TotalVat='" & totalVat & "' WHERE (ID='"& InquiryID & "')"
		conn.Execute(mySQL)

		mySQL="DELETE FROM InquiryLines WHERE (Inquiry='"& InquiryID & "')"
		conn.Execute(mySQL)

		'**************************** Inserting Inquiry Lines ****************
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

			mySQL="INSERT INTO InquiryLines (Inquiry, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat) VALUES ('"& InquiryID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "', "& theHasVat &")"
			conn.Execute(mySQL)
		next 
		if RFD > 0 then
			theItem =			39999
			theDescription =	"����� ��� �ǘ���"

			theAppQtty =		0
			thePrice =			0

			theDiscount =		RFD
			theReverse =		0

			theLength =			0
			theWidth =			0
			theQtty =			0
			theSets =			0
			theVat =			0
			mySQL="INSERT INTO InquiryLines (Inquiry, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat) VALUES ('"& InquiryID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "')"
			conn.Execute(mySQL)
		end if

		'**************** Updating Inquiry-Order Relations ****************
		mySQL="DELETE FROM InquiryOrderRelations WHERE (Inquiry='" & InquiryID & "')"
		conn.Execute(mySQL)
		for i=1 to request.form("selectedOrders").count
			theOrder=	clng(request.form("selectedOrders")(i))
			mySQL="INSERT INTO InquiryOrderRelations (Inquiry,[Order]) VALUES ('" & InquiryID & "', '" & theOrder & "')"
			conn.Execute(mySQL)
		next
		'^^^^------------ Updating Inquiry-Order Relations ------------^^^^

	end if
	'****
	'^^^^--------------- Editing  ---------------^^^^

	response.redirect "AccountReport.asp?act=showInquiry&inquiry=" & InquiryID

elseif request("act")="approveInquiry" then

	if not Auth(6 , "C") then		
		'Doesn't have the Priviledge to APPROVE the Inquiry 
		response.write "<br>" 
		call showAlert ("��� ���� �� ����� ������ ������",CONST_MSG_ERROR) 
		response.end
	end if

	if request("inquiry")<>"" then

		InquiryID=request("inquiry")
		if not(isnumeric(request("inquiry"))) then
			ShowErrorMessage("���")
			response.end
		end if
		mySQL="SELECT * FROM Inquiries WHERE (ID='"& InquiryID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			ShowErrorMessage("���� ��� ")
			response.end
		else
			if RS1("Voided") = True then
				Conn.close
				response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ���� ��� ���.")
			elseif RS1("Issued") = True then
				Conn.close
				response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ���� ��� ���.")
			elseif RS1("Approved") = True then
				Conn.close
				response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ����� ��� ���.")
			end if
		end if
	else
		ShowErrorMessage("���")
		response.end
	end if
	
	mySQL="UPDATE Inquiries SET Approved=1, ApprovedDate=N'"& shamsiToday() & "', ApprovedBy='"& session("ID") & "' WHERE (ID='"& InquiryID & "')"
	conn.Execute(mySQL)

	response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID 

elseif request("act")="IssueInquiry" then

	if not Auth(6 , "D") then		
		'Doesn't have the Priviledge to ISSUE the Inquiry
		response.write "<br>" 
		call showAlert ("��� ���� �� ���� ������ ������",CONST_MSG_ERROR) 
		response.end
	end if

	ON ERROR RESUME NEXT
		InquiryID=		clng(request("Inquiry"))

		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("���!")
		end if
	ON ERROR GOTO 0

	creationDate=	shamsiToday() 
	issueDate=		SqlSafe(request("issueDate"))
	if issueDate="" then issueDate=creationDate

	if issueDate<>creationDate then
		if Auth(6 , "I") then
			' can ISSUE the Inquiry / Rev. Inquiry on another Date

			'---- Checking wether issueDate is valid in current open GL
			If Not CheckDateFormat(issueDate) Then
				Conn.close
				response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("����� ���� ��� ����� ����.")
			end if

			if (issueDate < session("OpenGLStartDate")) OR (issueDate > session("OpenGLEndDate")) then
				Conn.close
				response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("���!<br>����� ���� ��� ����� ����. <br>(�� ��� ���� ���� ����)")
			end if 
			'----
		else
			'Doesn't have the Priviledge to ISSUE the Inquiry on another Date
			response.write "<br>" 
			call showAlert ("��� ���� �� ���� ������ �� ��� ����� ������",CONST_MSG_ERROR) 
			response.end
		end if
	end if

	'---- Checking wether issueDate is valid in current open GL
	If Not CheckDateFormat(issueDate) Then
		Conn.close
		response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("���!<br>����� ���� ����� ����.")
	end if

	if (issueDate < session("OpenGLStartDate")) OR (issueDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("���!<br>����� ���� ����� ����. <br>(�� ��� ���� ���� ����)")
	end if 
	'----

	mySQL="SELECT * FROM Inquiries WHERE (ID='"& InquiryID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "top.asp?errMsg=" & Server.URLEncode("���! ������ �� �����" & InquiryID & " ���� ���.")
	else
		voided=			RS1("Voided")
		issued=			RS1("Issued")
		approved=		RS1("Approved")
		isReverse=		RS1("IsReverse")
		customerID=		RS1("Customer")
		inquiryFee=		RS1("TotalReceivable")
		IsA =			RS1("IsA")
		Vat =			RS1("TotalVat")' sam

		if voided then
			Conn.close
			response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ���� ��� ���.")
		elseif issued then
			Conn.close
			response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ���� ���� ��� ���.")
		elseif not approved then
			Conn.close
			response.redirect "?act=editInquiry&inquiry="& InquiryID & "&errmsg=" & Server.URLEncode("��� ������ ����� ���� ���.")
		end if
	end if

	mySQL="UPDATE Inquiries SET Issued=1, IssuedDate=N'"& issueDate & "', IssuedBy='"& session("ID") & "' WHERE (ID='"& InquiryID & "')"
	conn.Execute(mySQL)

	if isReverse then
		isCredit=1
		itemType=4 
	else
		isCredit=0
		itemType=1
		'----------------------- Declaring the related orders as closed --------------
		conn.Execute("UPDATE Orders SET Closed=1 WHERE ID IN (SELECT [Order] FROM InquiryOrderRelations WHERE (Inquiry='" & InquiryID & "'))")
	end if

	'**************************** Creating ARItem for Inquiry / Reverse Inquiry  ****************
	'*** Type = 1 means ARItem is an Inquiry
	'*** Type = 4 means ARItem is a Reverse Inquiry

	firstGLAccount=	"13003"	'This must be changed... (Business Debitors)
	'if IsA then
	'	GLAccount=	"91001"	'This must be changed... (Sales A)
	'else
	'	GLAccount=	"91002"	'This must be changed... (Sales B)
	'end if
	'
	' Changed By Kid 860118 , seasing to use Sales B

	GLAccount=	"91001"	'This must be changed... (Sales A)
	
	mySQL="INSERT INTO ARItems (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount, Vat) VALUES ('" &_
	GLAccount & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& CustomerID & "', N'"& issueDate & "', '"& isCredit & "', '"& itemType & "', '"& InquiryID & "', '"& inquiryFee & "', N'"& creationDate & "', '"& session("ID") & "', '"& inquiryFee & "', '" & Vat & "')"
	conn.Execute(mySQL)

	if isReverse then
		'*** ATTENTION: Increasing AR Balance ....
		mySQL="UPDATE Accounts SET ARBalance = ARBalance + '"& inquiryFee & "' WHERE (ID='"& CustomerID & "')"
	else
		'*** ATTENTION: Decreasing AR Balance ....
		mySQL="UPDATE Accounts SET ARBalance = ARBalance - '"& inquiryFee & "' WHERE (ID='"& CustomerID & "')"
	end if
	conn.Execute(mySQL)
	conn.close
	response.redirect "AccountReport.asp?act=showInquiry&inquiry="& InquiryID 

elseif request("act")="voidInquiry" then
	if not Auth(6 , "F") then		
		'Doesn't have the Priviledge to VOID the Inquiry 
		response.write "<br>" 
		call showAlert ("��� ���� �� ����� ������ ������",CONST_MSG_ERROR) 
		response.end
	end if

	comment=sqlSafe(request("comment"))

	InquiryID=request("inquiry")
	if InquiryID="" or not(isnumeric(InquiryID)) then
		response.write "<br>" 
		call showAlert ("��� �� ����� ������",CONST_MSG_ERROR) 
		response.end
	end if

	InquiryID=clng(InquiryID)
	
	mySQL="SELECT * FROM Inquiries WHERE (ID='"& InquiryID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		ShowErrorMessage("���� ��� ")
		response.end
	else
		voided=			RS1("Voided")
		issued=			RS1("Issued")
		issuedBy=		RS1("IssuedBy")
		isReverse=		RS1("IsReverse")
		customerID=		RS1("Customer")
		inquiryFee=		RS1("TotalReceivable")
		IsA =			RS1("IsA")
		if voided then
			ShowErrorMessage("��� ������ ���� �� ����� <span dir='LTR'>"& RS1("VoidedDate") & "</span> ���� ��� ���.")
			response.end
		end if
	end if
	
	mySQL="UPDATE Inquiries SET Voided=1, VoidedDate=N'"& shamsiToday() & "', VoidedBy='"& session("ID") & "' WHERE (ID='"& InquiryID & "')"
	conn.Execute(mySQL)
	
	if isReverse then
		isCredit=1
		itemType=4 
		itemTypeName="������ �ѐ�� �� ����"
	else
		isCredit=0
		itemType=1
		itemTypeName="������"
		'---------- Declaring the related orders as Open  -------------------
		'mySQL="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice= '" & InvoiceID & "'))"
		'Changed By Kid ! 840509 
		'set orders which are ONLY related to this invoice, "Open"
		'that means, orders which are related to this invoice and are NOT related to any OTHER issued invoices.
		mySQL ="UPDATE Orders SET Closed=0 WHERE ID IN (SELECT [Order] FROM InquiryOrderRelations WHERE (Inquiry = '" & InquiryID & "') AND ([Order] NOT IN (SELECT InquiryOrderRelations.[ORDER] FROM Inquiries INNER JOIN InquiryOrderRelations ON Inquiries.ID = InquiryOrderRelations.Inquiry WHERE (Inquiries.Issued = 1) AND (Inquiries.Voided = 0) AND (Inquiries.isReverse = 0) AND (Inquiries.ID <> '" & InquiryID & "'))))"
		conn.Execute(mySQL)

	end if

  '**************************** Voiding ARItem of Inquiry / Reverse Inquiry ****************
  '*** Type = 1 means ARItem is an Inquiry
  '*** Type = 4 means ARItem is a Reverse Inquiry
  '***
	'*********  Finding the ARItem of Inquiry / Reverse Inquiry
	mySQL="SELECT ID FROM ARItems WHERE (Type = '"& itemType & "') AND (Link='"& InquiryID & "')"
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
		mySQL="UPDATE Accounts SET ARBalance = ARBalance - '"& inquiryFee & "' WHERE (ID='"& CustomerID & "')"
	else
		mySQL="UPDATE Accounts SET ARBalance = ARBalance + '"& inquiryFee & "' WHERE (ID='"& CustomerID & "')"
	end if

	conn.Execute(mySQL)
	
	'***
	'***---------------- End of  Voiding ARItem of Inquiry / Reverse Inquiry ----------------

	' Sending a Message to Issuer ...
	if trim(comment)<>"" then comment = chr(13) & chr(10) & "[" & comment & "]"
	MsgTo			=	issuedBy
	msgTitle		=	"Inquiry Voided"

	msgBody			=	"�ǘ��� ��� ���� "& session("CSRName") & " ���� ��." & comment
	RelatedTable	=	"inquiries"
	relatedID		=	inquiryID
	replyTo			=	0
	IsReply			=	0
	urgent			=	1
	MsgFrom			=	session("ID")
	MsgDate			=	shamsiToday()
	MsgTime			=	currentTime10()
	Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")


	' Copying the PreInquiry Data...
	response.redirect "InquiryInput.asp?act=copyInquiry&inquiry="& InquiryID & "&msg=" & Server.URLEncode(itemTypeName & " ����� "& InquiryID & " ���� ��.")

elseif request("act")="removePreInquiry" then
	response.write "<br>" 

	if not Auth(6 , "G") then		
		'Doesn't have the Priviledge to REMOVE the Pre-Inquiry 
		call showAlert ("��� ���� �� ��� ��� ���� ������",CONST_MSG_ERROR) 
		response.end
	end if

	InquiryID=request("inquiry")
	if InquiryID="" or not(isnumeric(InquiryID)) then
		call showAlert ("��� �� ����� ��� ����",CONST_MSG_ERROR) 
		response.end
	end if
	InquiryID=clng(InquiryID)

	mySQL="SELECT * FROM Inquiries WHERE (ID='"& InquiryID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		call showAlert ("��� ���� ���� ���.",CONST_MSG_ERROR) 
	else
		voided=			RS1("Voided")
		issued=			RS1("Issued")
		isReverse=		RS1("IsReverse")
		customerID=		RS1("Customer")
		IsA =			RS1("IsA")
		if issued then
			call showAlert ("��� ������ ���� ��� ���.",CONST_MSG_ERROR) 
		elseif voided then
			call showAlert ("��� ������ ���� ��� ���.",CONST_MSG_ERROR) 
		else
			Conn.execute("DELETE FROM InquiryOrderRelations Where (Inquiry='" & InquiryID & "')")

			'Conn.execute("DELETE FROM InvoiceLines Where (Invoice='" & InvoiceID & "')")
			'Conn.execute("DELETE FROM Invoices Where (ID='" & InvoiceID & "')")
			'Changed By Kid 830929
			'Conn.execute("UPDATE Invoices SET Voided=1 Where (ID='" & InvoiceID & "')")
			'Changed By Kid 8400502 also adding VoidedBy & VoidedDate
			Conn.execute("UPDATE Inquiries SET Voided=1, VoidedDate=N'"& shamsiToday() & "', VoidedBy='"& session("ID") & "' WHERE (ID='"& InquiryID & "')")

			call showAlert ("��� ���� ��� ��.",CONST_MSG_INFORM) 
		end if
	end if
	response.end
end if
conn.Close
%>
<!--#include file="include_JS_for_Inquiries.asp" -->
<SCRIPT LANGUAGE="JavaScript">
<!--
function ApproveInquiry(){
	if (confirm("��� ����� ����� �� �� ������ ��� ������ �� '�����' ���Ͽ\n\n(����: ������� ����� ��� ����)\n"))
		window.location="?act=approveInquiry&inquiry=<%=InquiryID%>";
}
function IssueInquiry(){
	if (confirm("��� ����� ����� �� �� ������ ��� ������ �� '����' ���Ͽ\n\n(����: ������� ����� ��� ����)\n"))
		window.location="?act=IssueInquiry&inquiry=<%=InquiryID%>";
}
function VoidInquiry(){
	if (confirm("��� ����� ����� �� �� ������ ��� ������ �� '����' ���Ͽ\n"))
		window.location="?act=voidInquiry&inquiry=<%=InquiryID%>";
}
//-->
</SCRIPT>
<!--#include file="tah.asp" -->
