<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="Ê—Êœ «” ⁄·«„"
SubmenuItem=1
if not (Auth(6 , 1) OR Auth(6 , 4)) then NotAllowdToViewThisPage()
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
			SO_StepText="<br>ê«„ ”Ê„ : «‰ Œ«» ”›«—‘ Â«Ì „—»ÊÿÂ"
%>
			<FORM METHOD=POST ACTION="?act=getInquiry">
			<!--#include File="include_SelectOrdersByOrder.asp"-->
			</FORM>
<%if Auth(6 , "K") then ' Has the priviledge to create a ReverseInquiry without an ORDER %>
			<p align='center' style='font-size:9pt;' dir='RTL'><a href='?act=getInquiry&selectedCustomer=<%=SO_Customer%>'>„—»Êÿ »Â ”›«—‘ Œ«’Ì ‰Ì” ...</a></p>
<%end if%>
			<br>
<%		else
			SA_TitleOrName=request("query")
			SA_Action="return true;"
			SA_SearchAgainURL="InquiryInput.asp"
			SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
			<FORM METHOD=POST ACTION="?act=selectOrder">
			<!--#include File="include_SelectAccount.asp"-->
			</FORM>
<%
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("Œ«·Ì ”—ç ‰„Ìﬂ‰Ì„!")
	end if
elseif request("act")="selectOrder" then
	if request("selectedCustomer") <> "" then
		SO_Customer=request("selectedCustomer")
		SO_Action="return true;"
		SO_StepText="<br>ê«„ ”Ê„ : «‰ Œ«» ”›«—‘ Â«Ì „—»ÊÿÂ"
%>
		<FORM METHOD=POST ACTION="?act=getInquiry">
		<!--#include File="include_SelectOrder.asp"-->
		</FORM>
<%if Auth(6 , "K") then ' Has the priviledge to create a ReverseInquiry without an ORDER %>
			<p align='center' style='font-size:9pt;' dir='RTL'><a href='?act=getInquiry&selectedCustomer=<%=SO_Customer%>'>„—»Êÿ »Â ”›«—‘ Œ«’Ì ‰Ì” ...</a></p>
<%end if%>
		<br>
<%
	end if
elseif request("act")="getInquiry" then
%>
<!--#include file="include_JS_for_Inquiries.asp" -->
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
		end if
		rs.close
		
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

	
	creationDate=shamsiToday()
'	creationTime=Hour(creationTime)&":"&Minute(creationTime)
'	if instr(creationTime,":")<3 then creationTime="0" & creationTime
'	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)

	InquiryLinesNo=1
%>
<!-- Ê—Êœ «ÿ·«⁄«  ›«ﬂ Ê— -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
		<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
		<FORM METHOD=POST ACTION="?act=submitInquiry">
			<tr bgcolor='#C3C300'>
			<td colspan="2"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL"><TR>
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
						<td align="left"> «—ÌŒ:</td>
						<td dir="LTR">
							<INPUT class="InvGenInput" NAME="InquiryDate" TYPE="text" maxlength="10" size="10" value="<%=CreationDate%>"></td>
						<td dir="RTL"><%=weekdayname(weekday(date))%></td>
					</tr>
					</table></TD>
				</TR></TABLE>
			</td>
			</tr>
			<tr bgcolor='#C3C300'>
				<TD align="right" width="50%">
					„—»Êÿ »Â ”›«—‘(Â«Ì):
					<span id="orders">
<%
					tempWriteAnd=""
					for i=1 to request.form("selectedOrders").count
						response.write "<input type='hidden' name='selectedOrders' value='"& request.form("selectedOrders")(i)& "'>"
						response.write tempWriteAnd & Link2Trace(request.form("selectedOrders")(i))
						tempWriteAnd=" Ê "
					next
%>
					</span>&nbsp;
					<INPUT class="InvGenButton" TYPE="button" value=" €ÌÌ—" onClick="selectOrder();">
				</TD>
				<TD align="left"><table>
					<tr>
						<td align="left">‘„«—Â:</td>
						<td dir="LTR">
							<INPUT class="InvGenInput" NAME="InquiryNo" style="border:1px solid black;" TYPE="text" maxlength="10" size="10"></td>
						<td dir="RTL"><INPUT TYPE="checkbox" onClick="checkIsA();" NAME="IsA"> «·› &nbsp;</td>
					</tr>
					</table></TD>
			</tr>
			<tr bgcolor='#CCCC88'>
			<TD colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<tr bgcolor='#CCCC88'>
				<td align='center' width="32px"> # </td>
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
			<TD colspan="10"><div style="overflow:auto; height:250px; width:*;">
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<Tbody id="InquiryLines">

<%		
		for i=1 to 1
%>
			<tr bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
				<td align='center' width="25px"><%=i%></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Items" size="3" Maxlength="6" onKeyPress="return mask(this);" onfocus="setCurrentRow(this.parentNode.parentNode.rowIndex);" onChange='return check(this);'>
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
				
				
				<td colspan='9' width='500px'>*** Œ›Ì› —‰œ ›«ò Ê— »⁄œ «“ À»  œ—Ã ŒÊ«Âœ ‘œ***</td>
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
				<td dir="LTR"><INPUT readonly class="invHeadInput3" TYPE="text" size="6" value="3%„«·Ì« "></td><!-- S A M -->
				<td dir="LTR"><INPUT readonly class="InvHeadInput" TYPE="text" size="9" value="—‰œ ‘œÂ"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
		</table><br>
		<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
		<tr>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="À»  ÅÌ‘ ‰ÊÌ” «” ⁄·«„..." onclick="submitOperations();"></td>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="«‰’—«›" onclick="window.close();"></td>
		</tr>
		</TABLE>
		</FORM>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			document.getElementsByName("Items")[0].focus();
		//-->
		</SCRIPT>
<%elseif request("act")="submitInquiry" then
	'******************** Checking Input ****************
	errorFound=false
	ON ERROR RESUME NEXT

		InquiryDate=	request.form("InquiryDate")

		If Not CheckDateFormat(InquiryDate) Then
			errorFound=True
		end if

		CustomerID=		clng(request.form("CustomerID"))

		if request.form("IsA") = "on" then 
			IsA=1 
			InquiryNo=request.form("InquiryNo")
			if InquiryNo <> "" then InquiryNo = clng(InquiryNo)
		else 
			IsA=0
			InquiryNo=""
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

		for i = 1 to request.form("selectedOrders").count
			theOrder=		clng(request.form("selectedOrders")(i))
			mySQL = "SELECT ID FROM Orders WHERE ID=" & theOrder 
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
		call showAlert ("Œÿ« œ— Ê—ÊœÌ",CONST_MSG_ERROR) 
		response.end
	end if
	'^^^^---------------- Checking Input ------------^^^^

	'**************************** Inserting new Inquiry ****************
	mySQL="INSERT INTO Inquiries (CreatedDate, CreatedBy, Customer, Number, TotalPrice, TotalDiscount, TotalReverse, TotalReceivable, IsA, TotalVat) VALUES (N'"& InquiryDate & "', '"& session("ID") & "', '"& CustomerID & "', '"& InquiryNo & "', '"& TotalPrice & "', '"& TotalDiscount & "', '"& TotalReverse & "', '"& TotalReceivable & "', '"& IsA & "', '" & TotalVat & "');SELECT @@Identity AS NewInquiry"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	InquiryID = RS1 ("NewInquiry")
	RS1.close	

	'**************************** Inserting Inquiry Lines ****************
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

		mySQL="INSERT INTO InquiryLines (Inquiry, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat) VALUES ('"& InquiryID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "', "&theHasVat&")"
		conn.Execute(mySQL)
	next 
	if RFD > 0 then
		theItem =			39999
		theDescription =	" Œ›Ì› —‰œ ›«ò Ê—"

		theAppQtty =		0
		thePrice =			0

		theDiscount =		RFD
		theReverse =		0
		theVat =			0

		theLength =			0
		theWidth =			0
		theQtty =			0
		theSets =			0
		mySQL="INSERT INTO InquiryLines (Inquiry, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat) VALUES ('"& InquiryID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "')"
		conn.Execute(mySQL)
	end if

	'**************************** Making Inquiry-Order Relations ****************
	for i=1 to request.form("selectedOrders").count
		theOrder=		clng(request.form("selectedOrders")(i))
		mySQL="INSERT INTO InquiryOrderRelations (Inquiry,[Order]) VALUES ('" & InquiryID & "', '" & theOrder & "')"
		conn.Execute(mySQL)
	next

	response.redirect "AccountReport.asp?act=showInquiry&inquiry="	 & InquiryID

elseif request("act")="copyInquiry" then
	InquiryID=request.queryString("inquiry")
	if InquiryID="" OR not isnumeric(InquiryID) then
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â «” ⁄·«„ ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
	end if

	InquiryID=clng(InquiryID)
	mySQL="SELECT * FROM Inquiries WHERE (ID='"& InquiriID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("›«ﬂ Ê— ÅÌœ« ‰‘œ.")
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
	InquiryNo=		RS1("Number")

	RS1.close

	if isReverse AND not Auth(6 , 4) then
		'Doesn't have the permission for ADDING Rev. Inquiry
		response.write "<br>" 
		call showAlert ("‘„« „Ã«“ »Â Ê—Êœ «” ⁄·«„ »—ê‘  ‰Ì” Ìœ",CONST_MSG_ERROR) 
		response.end
	elseif not Auth(6 , 1) then 
		'Doesn't have the permission for ADDING Inquiry
		response.write "<br>" 
		call showAlert ("‘„« „Ã«“ »Â Ê—Êœ «” ⁄·«„ ‰Ì” Ìœ",CONST_MSG_ERROR) 
		response.end
	end if

	'Copying Inquiry ...
	mySQL="INSERT INTO Inquiries (CreatedDate, CreatedBy, Customer, IsA, Number, TotalPrice, TotalDiscount, TotalReverse, TotalReceivable, IsReverse, TotalVat) VALUES (N'"& creationDate & "', '"& session("ID") & "', '"& CustomerID & "', '"& IsA & "', '"& InquiryNo & "', '"& TotalPrice & "', '"& TotalDiscount & "', '"& TotalReverse & "', '"& TotalReceivable & "', '"& IsReverse & "', '" & TotalVat & "'); SELECT @@Identity AS NewInquiry"
	Set RS1 = conn.Execute(mySQL).NextRecordSet
	NewInquiry = RS1("NewInquiry")

	'Copying InquiryLines ...
	mySQL="INSERT INTO InquiryLines (Inquiry, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat) SELECT '"& NewInquiry & "' AS Inquiry, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat FROM InquiryLines WHERE (Inquiry = '"& InquiryID & "')"
	conn.Execute(mySQL)

	'Copying InquiryOrderRelations ...
	mySQL="INSERT INTO InquiryOrderRelations (Inquiry, [Order]) SELECT '"& NewInquiry & "' AS Inquiry, [Order] FROM InquiryOrderRelations WHERE (Inquiry = '"& InquiryID & "')"
	conn.Execute(mySQL)

	if Issued AND voided then
	' if it's copied from a voided Inquiry, place a comment for the voided Inquiry 
		MsgTo			=	0
		msgTitle		=	"Inquiry Copied"
		msgBody			=	"«” ⁄·«„ ‘„«—Â "& NewInquiry & " «“ —ÊÌ «” ⁄·«„ ›Êﬁ ﬂÅÌ ‘œ." 
		RelatedTable	=	"Inquiries"
		relatedID		=	InquiryID
		replyTo			=	0
		IsReply			=	0
		urgent			=	0
		MsgFrom			=	session("ID")
		MsgDate			=	shamsiToday()
		MsgTime			=	currentTime10()
		Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")
		
	end if	

	' Place a comment for the new Inquiry 
	MsgTo			=	0
	msgTitle		=	"Inquiry Copied"
	msgBody			=	"ÅÌ‘ ‰ÊÌ” «“ «” ⁄·«„ ‘„«—Â "& InquiryID & " ﬂÅÌ ‘œ." 
	RelatedTable	=	"Inquiries"
	relatedID		=	NewInquiry 
	replyTo			=	0
	IsReply			=	0
	urgent			=	0
	MsgFrom			=	session("ID")
	MsgDate			=	shamsiToday()
	MsgTime			=	currentTime10()
	Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")

	response.redirect "AccountReport.asp?act=showInquiry&inquiry=" & NewInquiry & "&msg=" & Server.URLEncode(request.queryString("msg") & "<br>ÅÌ‘ ‰ÊÌ” «“ —ÊÌ «” ⁄·«„ Ì« ÅÌ‘‰ÊÌ” ‘„«—Â "& InquiryID & " ﬂÅÌ ‘œ.")

end if
conn.Close
%>
<!--#include file="tah.asp" -->
