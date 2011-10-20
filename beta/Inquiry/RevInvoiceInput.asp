<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="Ê—Êœ ›«ﬂ Ê— »—ê‘  «“ ›—Ê‘"
SubmenuItem=4
if not Auth(6 , 4) then NotAllowdToViewThisPage()

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
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
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
			<FORM METHOD=POST ACTION="?act=getInvoice">
			<!--#include File="include_SelectClosedOrdersByOrder.asp"-->
			</FORM>
<%if Auth(6 , "K") then ' Has the priviledge to create a ReverseInvoice without an ORDER %>
			<p align='center' style='font-size:9pt;' dir='RTL'><a href='?act=getInvoice&selectedCustomer=<%=SO_Customer%>'>„—»Êÿ »Â ”›«—‘ Œ«’Ì ‰Ì” ...</a></p>
<%end if%>
			<br>
<%		else
			SA_TitleOrName=request("query")
			SA_Action="return true;"
			SA_SearchAgainURL="RevInvoiceInput.asp"
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
		<FORM METHOD=POST ACTION="?act=getInvoice">
		<!--#include File="include_SelectClosedOrder.asp"-->
		</FORM>
<%if Auth(6 , "K") then ' Has the priviledge to create a ReverseInvoice without an ORDER %>
			<p align='center' style='font-size:9pt;' dir='RTL'><a href='?act=getInvoice&selectedCustomer=<%=SO_Customer%>'>„—»Êÿ »Â ”›«—‘ Œ«’Ì ‰Ì” ...</a></p>
<%end if%>
		<br>
<%
	end if
elseif request("act")="getInvoice" then
%>
<!--#include file="include_JS_for_Invoices.asp" -->
<%
	customerID=request("selectedCustomer")

	mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)
	customerName=RS1("AccountTitle")
	
	creationDate=shamsiToday()
'	creationTime=Hour(creationTime)&":"&Minute(creationTime)
'	if instr(creationTime,":")<3 then creationTime="0" & creationTime
'	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)

	InvoiceLinesNo=1
%>
<!-- Ê—Êœ «ÿ·«⁄«  ›«ﬂ Ê— »—ê‘  «“ ›—Ê‘ -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
		<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
		<FORM METHOD=POST ACTION="?act=submitInvoice">
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
							<INPUT class="InvGenInput" NAME="InvoiceDate" TYPE="text" maxlength="10" size="10" value="<%=CreationDate%>"></td>
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
							<INPUT class="InvGenInput" NAME="InvoiceNo" style="border:1px solid black;" TYPE="text" maxlength="10" size="10"></td>
						<td dir="RTL"><INPUT TYPE="checkbox" NAME="IsA"> «·› &nbsp;</td>
					</tr>
					</table></TD>
			</tr>
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
				<!--S A M-->
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value=" Œ›Ì›"size="7"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="»—ê‘ " size="5"></td>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" Value="„«·Ì« " size="6"></td>
				<td><INPUT class="InvHeadInput2" readonly TYPE="text" Value="ﬁ«»· Å—œ«Œ " size="9"></td>
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
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Items" size="3" onKeyPress="return mask(this);" onfocus="setCurrentRow(this.parentNode.parentNode.rowIndex);" onBlur='check(this);'>
				<INPUT TYPE="hidden" name="type" value=0>
				<INPUT TYPE="hidden" name="fee" value=0>
				</td>
				<td dir="RTL"><INPUT class="InvRowInput2" TYPE="text" NAME="Descriptions" size="30"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Lengths" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Widths" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Qttys" size="3" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput2" TYPE="text" NAME="Sets" size="2" onBlur="setFeeQtty(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="AppQttys" size="6" onBlur="setPrice(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Fees" size="7" onBlur="setPrice(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Prices" size="9" readonly tabIndex="9999"></td>
				<!--S A M-->
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Discounts" size="7" onBlur="setPrice(this);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Reverses" size="5" onBlur="setPrice(this);" onfocus="setCurrentRow(this.parentNode.parentNode.rowIndex);"></td>
				<td dir="LTR"><INPUT class="InvRowInput" TYPE="text" NAME="Vat" size="6" readonly></td>
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
				<td align='center' width="25px"> &nbsp; </td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="3" ></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="30"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="3"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="2"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="6"></td>
				<td><INPUT readonly class="InvHeadInput" TYPE="text" size="7"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalPrice" TYPE="text" size="9"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalDiscount" TYPE="text" size="7"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalReverse" TYPE="text" size="5"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalVat" TYPE="text" size="6"
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" Name="TotalAppPrice" TYPE="text" size="9"></td>
			</tr>
			<tr bgcolor='#CCCC88'>
				<td colspan="9"> &nbsp; </td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput" TYPE="text" size="9"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" Name="TPDiscount" size="7"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" Name="TPReverse" size="5"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput3" TYPE="text" value="3%„«·Ì« " size="6"></td>
				<td dir="LTR"><INPUT readonly class="InvHeadInput" TYPE="text" size="9"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
		</table><br>
		<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
		<tr>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="À»  ÅÌ‘ ›«ﬂ Ê— »—ê‘  «“ ›—Ê‘..." onclick="submitOperations();"></td>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="«‰’—«›" onclick="window.close();"></td>
		</tr>
		</TABLE>
		</FORM>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			document.getElementsByName("Items")[0].focus();
		//-->
		</SCRIPT>
<%elseif request("act")="submitInvoice" then

	InvoiceDate=request.form("InvoiceDate")
	CustomerID=request.form("CustomerID") 
	InvoiceNo=request.form("InvoiceNo") 
	if request.form("IsA") = "on" then IsA=1 else IsA=0
		
	TotalPrice=text2value(request.form("TotalPrice"))
	TotalDiscount=text2value(request.form("TotalDiscount"))
	TotalReverse=text2value(request.form("TotalReverse"))
	TotalReceivable=text2value(request.form("TotalAppPrice"))
	TotalVat = text2value(request.form("TotalVat"))' S A M CHANGE THIS

	mySQL="INSERT INTO Invoices (IsReverse, CreatedDate, CreatedBy, Customer, Number, TotalPrice, TotalDiscount, TotalReverse, TotalReceivable, IsA, TotalVat) VALUES (1, N'"& InvoiceDate & "', '"& session("ID") & "', '"& CustomerID & "', '"& InvoiceNo & "', '"& TotalPrice & "', '"& TotalDiscount & "', '"& TotalReverse & "', '"& TotalReceivable & "', '"& IsA & "', '" & TotalVat & "')"	
	conn.Execute(mySQL)

	mySQL="SELECT MAX(ID) AS lastID FROM Invoices WHERE (IsReverse=1 AND CreatedBy='"& session("ID") & "' AND Customer='"& CustomerID & "' AND TotalPrice='"& TotalPrice & "')"

	Set RS1 = conn.Execute(mySQL)
	InvoiceID=RS1("lastID")
	RS1.close
%>
	<hr>
	<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" align="center">
	<tr bgcolor='#CCCC88'>
		<td align='center' width="25px"> # </td>
		<td><INPUT readonly class="InvHeadInput" TYPE="text" value="¬Ì „" size="4" ></td>
		<td><INPUT readonly class="InvHeadInput2" TYPE="text" value=" Ê÷ÌÕ« " size="30"></td>
		<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="ÿÊ·" size="2"></td>
		<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="⁄—÷" size="2"></td>
		<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value=" ⁄œ«œ" size="3"></td>
		<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="›—„" size="2"></td>
		<td><INPUT readonly class="InvHeadInput" TYPE="text" Value=" ⁄œ«œ „ÊÀ—" size="6"></td>
		<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="›Ì" size="7"></td>
		<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="ﬁÌ„ " size="9"></td>
		<!--S A M-->
		<td><INPUT readonly class="InvHeadInput" TYPE="text" Value=" Œ›Ì›"size="7"></td>
		<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="»—ê‘ " size="5"></td>
		<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="„«·Ì« " size="6"></td>
		<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="ﬁ«»· Å—œ«Œ " size="9"></td>
	</tr>
<%	for i=1 to request.form("Items").count 
		theItem = text2value(request.form("Items")(i))
		theDescription = request.form("Descriptions")(i)
		theLength = text2value(request.form("Lengths")(i))
		theWidth = text2value(request.form("Widths")(i))
		theQtty = text2value(request.form("Qttys")(i))
		theSets = text2value(request.form("Sets")(i))
		theAppQtty = text2value(request.form("AppQttys")(i))
		theFee = text2value(request.form("Fees")(i))
		thePrice = text2value(request.form("Prices")(i))
		theDiscount = text2value(request.form("Discounts")(i))
		theReverse = text2value(request.form("Reverses")(i))
		theAppPrice	= text2value(request.form("AppPrices")(i))
		theVat = text2value(reguest.form("Vat")(i))


		mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat) VALUES ('"& InvoiceID & "', '" & theItem & "', N'" & theDescription & "', '" & theLength & "', '" & theWidth & "', '" & theQtty & "', '" & theSets & "', '" & theAppQtty & "', '" & thePrice & "', '" & theDiscount & "', '" & theReverse & "', '" & theVat & "')"
		conn.Execute(mySQL)
%>
	<tr bgcolor='#F0F0F0' height="20px">
		<td align='center' width="25px"><%=i%></td>
		<td align='right' dir="LTR"><%=Separate(theItem)%></td>
		<td dir="RTL" width="170px"><%=theDescription%></td>
		<td align='right' dir="LTR"><%=Separate(theLength)%></td>
		<td align='right' dir="LTR"><%=Separate(theWidth)%></td>
		<td align='right' dir="LTR"><%=Separate(theQtty)%></td>
		<td align='right' dir="LTR"><%=Separate(theSets)%></td>
		<td align='right' dir="LTR"><%=Separate(theAppQtty)%></td>
		<td align='right' dir="LTR"><%=Separate(theFee)%></td>
		<td align='right' dir="LTR"><%=Separate(thePrice)%></td>
		<td align='right' dir="LTR"><%=Separate(theDiscount)%></td>
		<td align='right' dir="LTR"><%=Separate(theReverse)%></td>
		<td align='right' dir="LTR"><%=Separate(theVat)%></td>
		<td align='right' dir="LTR"><%=Separate(theAppPrice)%></td>
	</tr>
<%	next %>
	<tr>
		<td colspan="13"></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="9"> &nbsp; </td>
		<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalPrice)%></td>
		<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalDiscount)%></td>
		<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalReverse)%></td>
		<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalVat)%></td>
		<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalReceivable)%></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="10"> &nbsp; </td>
		<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Pourcent(TotalDiscount,TotalPrice) & "% Œ›Ì›"%></td>
		<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Pourcent(TotalReverse,TotalPrice) & "%»—ê‘ "%></td>
		<td align='right' dir="LTR" bgcolor="#F0F0F0">„«·Ì« </td>
		<td>&nbsp;</td>
	</tr>
	</TABLE>
	<hr>
<%
	for i=1 to request.form("selectedOrders").count
		theOrder=request.form("selectedOrders")(i)
		mySQL="INSERT INTO InvoiceOrderRelations (Invoice,[Order]) VALUES ('" & InvoiceID & "', '" & theOrder & "')"
		conn.Execute(mySQL)
	next
%>

<!-- «ÿ·«⁄«  ›«ﬂ Ê— »—ê‘  «“ ›—Ê‘ À»  ‘œ... -->
	<div dir='rtl'><B>«ÿ·«⁄«  ›«ﬂ Ê— »—ê‘  «“ ›—Ê‘ À»  ‘œ...</B>
	<br>„‘ —Ì:'<%=CustomerID%>'
	<br>›«ﬂ Ê— »—ê‘  «“ ›—Ê‘:'<%=InvoiceID%>'<br>
	<input class="InvGenInput" type="button" value="Å«Ì«‰" onclick="window.close();">
	</div>
<%
	response.redirect "AccountReport.asp?act=showInvoice&invoice="	 & InvoiceID

else%>
<!-- Ã” ÃÊ »—«Ì ‰«„ Õ”«» »« ‘„«—Â ”›«—‘-->
	<%if Auth(6 , 4) then %>
	<br>
	<br>
	<FORM METHOD=POST ACTION="?act=submitsearch" onsubmit="if (document.all.query.value=='') return false;">
	<div dir='rtl'><B><FONT SIZE="" COLOR="red"> &nbsp;›«ﬂ Ê— ÃœÌœ: </FONT><BR>Ã” ÃÊ »—«Ì ‰«„ Õ”«» Ì« ‘„«—Â ”›«—‘</B>
		<INPUT TYPE="text" NAME="query">&nbsp;
		<INPUT TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>
	<%end if %>
<!-- Ã” ÃÊ »—«Ì ›«ﬂ Ê— -->
	<%if Auth(6 , 5) then %>
	<br>
	<hr>
	<br>
	<br>
	<FORM METHOD=POST ACTION="InvoiceEdit.asp?act=search" onsubmit="if (document.all.order.value=='' && document.all.invoice.value=='') return false;">
	<div dir='rtl'>&nbsp;<B><FONT SIZE="" COLOR="red">«’·«Õ ›«ﬂ Ê—:</FONT><BR> ‘„«—Â ”›«—‘ —« Ê«—œ ﬂ‰Ìœ</B>
		<INPUT style="font-family:Tahoma;" TYPE="text" NAME="order">&nbsp;
		<INPUT style="font-family:Tahoma;" TYPE="submit" value="«œ«„Â..."><br><br>
		<Blockquote>
			‘„«—Â ›«ﬂ Ê— —« „Ìœ«‰„:<br><INPUT style="font-family:Tahoma;width:100px;" TYPE="text" NAME="invoice"><br>
			<INPUT style="font-family:Tahoma;" TYPE="submit" value="«œ«„Â..."><br><br>
		</Blockquote>
	</div>
	</FORM>
	<%end if %>

	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.query.focus();
	//-->
	</SCRIPT>
<%
end if
conn.Close
%>
<!--#include file="tah.asp" -->