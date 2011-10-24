<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CashRegister (9)
PageTitle="œ—Ì«› "
SubmenuItem=1
if not Auth(9 , 1) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/orderEdit.asp?e=n&radif="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

%>
<style>
	.RcpTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; }
	.RcpMainTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; background-color: #558855; text-align:right; direction: RTL;}
	.RcpMainTableTH { background-color: #C3C300;}
	.RcpMainTableTR { background-color: #CCCC88; border: 0; }
	.RcpRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.RcpRowInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.RcpHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.RcpHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.RcpHeadInput3 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #D0E0FF; text-align:right; direction: right-to-left;}
	.RcpGenInput { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction: LTR;}
	.GenTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; }
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
	.GrayOutLine {border: 1px solid gray; }
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;
var currentRow=null;
//-->
</SCRIPT>
<!--#include File="../include_UtilFunctions.asp"-->
<%
if request("act")="ShowReceipt" then
	Receipt=clng(request("id"))
%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	var status = -1;
	var timeoutID= 0;
	function openCashDrawer()
	{
		drawerPresent = true;
		try
		{
			document.getElementById('Drawer').OpenUPort(1);
		}
		catch(e)
		{
			drawerPresent = false;
		}

		if ( drawerPresent )
		{ 
			//document.getElementById('Drawer').OpenUPort(1);
			status = document.getElementById('Drawer').UCashDrawerStatus(1);
			if (status>-1)
			{
				document.getElementById('Drawer').OpenUCashDrawer(1);
			}
			//document.getElementById('Drawer').CloseUPort(1);
		}
	}
	function closeCashDrawer()
	{
		drawerPresent = true;
		try
		{
			document.getElementById('Drawer').OpenUPort(1);
		}
		catch(e)
		{
			drawerPresent = false;
		}

		if ( drawerPresent )
		{ 
			//document.getElementById('Drawer').OpenUPort(1);
			document.getElementById('Drawer').OpenUPort(1);
			// Status: 1:Open, 0:Close
			while ( document.getElementById('Drawer').UCashDrawerStatus(1)==1 )
			{
				alert('·ÿ›« ﬂ‘ÊÌ ’‰œÊﬁ —« »»‰œÌœ');
			}
			//document.getElementById('Drawer').CloseUPort(1);
		}
	}
	//-->
	</SCRIPT>

	<OBJECT ID="Drawer" WIDTH="1" HEIGHT="1" CLASSID="CLSID:A46E44C7-AC96-4EFE-B8D7-EE7D67990B6F">
	</OBJECT>

	<BR><BR><BR>	
	<CENTER>
		<% 	ReportLogRow = PrepareReport ("Receipt.rpt", "Recept_ID", Receipt, "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="openCashDrawer();closeCashDrawer();printThisReport(this,<%=ReportLogRow%>);">
	</CENTER>
	<BR><iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:700; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no ></iframe>
	<BR>
	<BR>
<%
	response.end
end if

mySQL="SELECT * FROM CashRegisters WHERE (IsOpen=1) AND (Cashier='"& session("ID") & "')"
Set RS1= conn.Execute(mySQL)
if RS1.eof then 
%><br><br>
	<TABLE width='70%' align='center'>
	<TR>
		<TD align=center bgcolor=#FFBBBB style='border: solid 1pt black'><BR><b>Ã‰«» <%=CSRName%> ‘„« ’‰œÊﬁ »«“ ‰œ«—Ìœ Ö <br><br>«“ œ”  „‰ ﬂ«—Ì »—«Ì ‘„« »— ‰„Ì ¬Ìœ <br>„ «”›„."</b><BR><BR></TD>
	</TR>
	</TABLE>
<%	conn.close
	response.end
else
	CashRegID=RS1("ID")
	theBanker=RS1("Banker")
	CashRegName=RS1("NameDate")
	Set RS1=nothing
end if

if request("act")="submitsearch" then
	if isnumeric(request("SearchBox")) then
		orderID=clng(request("SearchBox"))
		mySQL="SELECT Invoices.Issued, Invoices.Approved, Invoices.ID, Invoices.TotalReceivable, Invoices.Customer FROM InvoiceOrderRelations INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE ([Order] = '"& orderID & "')"
		Set RS1=Conn.Execute(mySQL)
		if RS1.eof then
			Conn.close
			response.redirect "?errmsg=" & Server.URLEncode("«Ì‰ ”›«—‘ ﬁ»·« ›«ﬂ Ê— ‰‘œÂ «” .")
		else
			theInvoice=RS1("ID")
			theCustomer=RS1("Customer")
			if RS1("Issued") then
				Conn.close
				response.redirect "?act=getReceipt&selectedCustomer=" & theCustomer & "&selectedInvoice=" & theInvoice
			else
				response.write "<br><br>" 

				if RS1("Approved") then
					extraDesc="<br>(Â—ç‰œ ﬂÂ  «ÌÌœ ‘œÂ «” )"
				end if

				call showAlert ("<b>›«ﬂ Ê— „—»Êÿ »Â «Ì‰ ”›«—‘ ’«œ— ‰‘œÂ «” ."& extraDesc & " <br>(„»·€: "& Separate(RS1("TotalReceivable")) & ")",CONST_MSG_ALERT) 
				response.write "<Blockquote>" 
				response.write "<br><br>çﬂ«— „Ì ﬂ‰Ìœø<br></b>" 
				response.write "<br><br><li><A target='_blank' HREF='../AR/AccountReport.asp?act=showInvoice&invoice="& theInvoice & "'>‰‘«‰ »œÂ »»Ì‰„</A>" 
				response.write "<br><br><li><A HREF='?act=getReceipt&selectedCustomer=" & theCustomer & "'>«œ«„Â  „Ì œÂ„.</A>" 
				if RS1("Approved") then
					response.write "<br><br><li> «Ê· ›«ﬂ Ê— —« <A target='_blank' HREF='../AR/InvoiceEdit.asp?act=search&order="& request("SearchBox") & "'>’«œ— „Ì ﬂ‰„</A> " 
					response.write "Ê »⁄œ «“ ¬‰  <A HREF='?act=submitsearch&SearchBox="& request("SearchBox") & "'>«œ«„Â „Ì œÂ„.</A>" 
				end if
				response.write "</Blockquote>" 
			end if 
		end if
	elseif request("SearchBox") <> "" then
		SA_TitleOrName=request("SearchBox")
		SA_Action="return true;"
		SA_SearchAgainURL="ReceiptInput.asp"
		SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<FORM METHOD=POST ACTION="?act=getReceipt">
		<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	else
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ”›«—‘ ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
	end if
elseif request("act")="getReceipt" then
	if not isnumeric(request("selectedCustomer")) then
		response.write "EORROR"
		response.end
	end if
	if request("Reason")="" then
		Reason=1
	else
		Reason=cint(request("Reason"))
	end if
	if Reason=1 then
		sys="AR"
	elseif Reason=2 then
		sys="AP"
	else
		sys="AO"
	end if

	customerID=clng(request("selectedCustomer"))
	selectedInvoice=request.queryString("selectedInvoice")
	mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)
	AccountNo=cdbl(RS1("ID"))
	customerName=RS1("AccountTitle")
	
	creationDate=shamsiToday()

%>
	<br><div dir='rtl'><B>ê«„ ”Ê„ : Ê—Êœ «ÿ·«⁄«  œ—Ì«›  ÅÊ·</B>
	</div>
<!-- Ê—Êœ «ÿ·«⁄«  œ—Ì«›  ÅÊ· -->
	<hr>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<TABLE Cellspacing="0" Cellpadding="10" align="center">
	<TR><TD valign='top'>
		<table class="RcpMainTable" Cellspacing="1" Cellpadding="0" Width="500" align="center">
		<FORM METHOD=POST ACTION="?act=submitReceipt" onsubmit="return submitCeck();" name="f1">
			<tr class="RcpMainTableTH">
			<td colspan="10" align='center' height='25px'>’‰œÊﬁ <span dir='LTR'><%=CashRegName%></span> - <%=CSRName%> 
				<INPUT TYPE="hidden" Name="CashRegID" Value="<%=CashRegID%>">
			</td>
			</tr>
			<tr class="RcpMainTableTH">
			<td colspan="10"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL"><TR>
				<TD><table>
					<tr>
						<td align="left">Õ”«»:</td>
						<td align="right">
							<span id="customer">
								<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><span><%=CustomerName%></span>.
							</span></td>
					</tr>
					</table></TD>
				<td><table>
					<tr>
						<td>«·›</td>
						<td><input name="isA" value="1" type="radio"></td>
						<td>»</td>
						<td><input name="isA" value="0" type="radio"></td>
						<td>ÿ»ﬁ ›«ﬂ Ê—</td>
						<td><input id="asInvoice" name="isA" value="-1" type="radio" checked="checked"></td>
					</tr>
				</table></td>
				<TD align="left"><table>
					<tr>
						<td align="left"> «—ÌŒ:</td>
						<td dir="LTR">
							<INPUT class="RcpGenInput" style="text-align:Left;" NAME="ReceiptDate" TYPE="text" maxlength="10" size="10" value="<%=CreationDate%>" onblur="acceptDate(this)"></td>
						<td dir="RTL"><%=weekdayname(weekday(date))%></td>
					</tr>
					</table></TD>
				</TR></TABLE>
			</td>
			</tr>
			<tr class="RcpMainTableTR">
			<TD colspan="10"><div>
			<TABLE class="RcpTable" Cellspacing="1" Cellpadding="0" Dir="RTL">
			<tr height="40">
				<td colspan="2" align="left"> „»·€ ‰ﬁœ: </td>
				<td colspan="15"><INPUT dir="LTR" class="RcpRowInput2" TYPE="text" Name="CashAmount" size="15" onKeyPress="return maskNumber(this);" onBlur="setPrice(this)"> ‘—Õ: <INPUT dir="RTL" class="RcpRowInput2" TYPE="text" Name="CashDescription" size="50" ></td>
			</tr>
			<tr height="30">
				<td colspan="15"> çﬂ: <br></td>
			</tr>
			</TABLE></div></TD>
			</TR>
			<tr class="RcpMainTableTR">
			<TD colspan="10"><div>
			<TABLE class="RcpTable" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<tr>
				<td class="RcpHeadInput" align='center' width="25px"> # </td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value="‘„«—Â çﬂ" size="12" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value=" «—ÌŒ" size="10" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="»«‰ﬂ" size="10" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="‘—Õ" size="20" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="„»·€" size="15" tabindex="9999"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
			<tr class="RcpMainTableTR">
			<TD colspan="10" dir="RTL"><div style="overflow:auto; height:130px;width:500px;">
			<TABLE class="RcpTable" Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<Tbody id="ChequeLines">

<%		
		for i=1 to 1
%>
			<tr bgcolor='#F0F0F0' onclick="currentRow=this.rowIndex;" >
				<td align='center' width="25px"><%=i%></td>
				<td dir="LTR"><INPUT class="RcpRowInput" TYPE="text" NAME="ChequeNos" size="12" onKeyPress="return maskNumber(this);"></td>
				<td dir="LTR"><INPUT class="RcpRowInput" style="text-align:left;" TYPE="text" NAME="ChequeDates" maxlength="10" size="10" onblur="acceptDate(this)"></td>
				<td dir="RTL"><INPUT class="RcpRowInput" TYPE="text" NAME="Banks" size="10"></td>
				<td dir="RTL"><INPUT class="RcpRowInput" TYPE="text" NAME="Descriptions" size="20"></td>
				<td dir="LTR"><INPUT class="RcpRowInput" TYPE="text" NAME="Amounts" size="15" onKeyPress="return maskNumber(this);" onBlur="setPrice(this);"></td>
			</tr>
<%
		next
%>
			<tr bgcolor='#F0F0F0' onclick="currentRow=this.rowIndex;" >
				<td colspan="15">
					<INPUT class="RcpGenInput" TYPE="button" value="«÷«›Â" onkeyDown="if(event.keyCode==9) return false;" onClick="addRow(this.parentNode.parentNode.rowIndex);">
				</td>
			</tr>
			</Tbody></TABLE></div>
			</TD>
			</tr>
			<tr class="RcpMainTableTR">
			<TD colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL">
			<tr>
				<td class="RcpHeadInput" align='center' width="25px"> &nbsp; </td>
				<td class="RcpHeadInput" colspan=2>»œÂÌùÂ«Ì «‰ Œ«» ‘œÂ:</td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="10" tabindex="9999" name="SelectedTotalPrice" onclick="document.all.CashAmount.value=document.all.SelectedTotalPrice.value;document.all.CashAmount.focus();"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="Ã„⁄:" size="20" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput3" readonly dir="LTR" TYPE="text" Name="TotalAmount" Value="" size="15" tabindex="9999"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
		</table>
		<TABLE class="RcpTable" Border="0" Cellspacing="5" Cellpadding="1" Dir="RTL">
		<tr>
			<td align='center' bgcolor="#000000"><INPUT class="RcpGenInput" style="text-align:center" TYPE="button" value="–ŒÌ—Â" onclick="submitCeck()"></td>
			<td align='center' bgcolor="#000000"><INPUT class="RcpGenInput" style="text-align:center" TYPE="button" value="«‰’—«›" onclick="window.location='';"></td>
		</tr>
		</TABLE>
	</TD>
	<TD valign='top'>
		<span style='width:100%;font-size:12pt;text-align:center;background-color:#FFDDDD;border:1px red solid;height:25px;'>»œÂÌ Â«</span><br>
		<span style='width:100%;font-size:9pt;text-align:center;background-color:#FFDDDD;border:1px red solid;height:25px;'>
		„—»Êÿ »Â <SELECT NAME="Reason" style='font-family:Tahoma;font-size:8pt;height:25px;width:100px;' onchange="changeSystem();">
<%			mySQL="SELECT * FROM AXItemReasons WHERE Display=1 ORDER BY ID"
			Set RS1=Conn.Execute(mySQL)
			Do while not RS1.eof
				if Reason = RS1("ID") then
					ifSelected="selected"
				else
					ifSelected=""
				end if
				response.write "<OPTION value='"& RS1("ID") & "' "& ifSelected & ">"& RS1("Name") & "</option>"
				RS1.MoveNext
			Loop
%>		
			</SELECT></span>
		<TABLE class="GenTable" align="center" cellspacing="0" cellpadding="2" dir="RTL" border="3">
		<Tbody ID="DebitsTAble">
<% '-----------------  The  Debit Items
	if Sys="AR" then
		mySQL="SELECT * From "& Sys & "Items WHERE (Account='"& AccountNo & "' AND IsCredit='0' AND FullyApplied='0') ORDER BY EffectiveDate , Link "
		Set RS1 = conn.Execute(mySQL)

		if (RS1.eof) then 
			response.write "<tr><td bgcolor=white width=150><br><table align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> ÅÌœ« ‰‘œ <br></td></tr></table><br></td></tr>"
		else
			TotalDebit=0
			while Not (RS1.EOF)
				if RS1("Type")=1 then 'invoice
					set rs2 = conn.Execute("select * from Invoices where id=" & RS1("Link"))
					if rs2("isA") then 
						isA="(«·›)"
					else
						isA="(»)"
					end if
					sourceLink="<a style='text-decoration:none; color:red;' href='../"& Sys & "/AccountReport.asp?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— ‘„«—Â " & RS1("Link") & isA & "</a>"
					rs2.close
				elseif RS1("Type")=2 then 'receipt
					sourceLink="<a style='text-decoration:none; color:red;' href='../"& Sys & "/AccountReport.asp?act=showReceipt&receipt="& RS1("Link") & "' target='_blank'>" & "œ—Ì«›  ‘„«—Â " & RS1("Link") & "</a>"
				elseif RS1("Type")=3 then 'memo
					sourceLink="<a style='text-decoration:none; color:red;' href='../"& Sys & "/AccountReport.asp?act=showMemo&Sys="& Sys & "&memo="& RS1("Link") & "' target='_blank'>" & "”‰œ ‘„«—Â " & RS1("Link") & "</a>"
				else	 ' i dunno
					sourceLink="<a style='text-decoration:none; color:red;' href='javascript:void(0);'>" & "‰« ‘‰«” [T:" & RS1("Type")& "] [L:"& RS1("Link") & "]</a>"
				end if

%>			<tr bgcolor="white" valign='top'><td  class="GrayOutLine">
				<table class='GenTable' cellspacing='4' cellpadding='1' width='100%'>
				<%if RS1("Type")=1 and clng(RS1("Link"))=clng(selectedInvoice) then%>
					<tr bgcolor='#FFFFBB'>
						<td class="GrayOutLine">
							<INPUT TYPE="checkbox" checked NAME="DebitItems" value='<%=RS1("ID")%>' onclick="setPrice2(this)"><%=sourceLink%></td>
					</tr>
				<%else%>
					<tr bgcolor='#FFDDDD'>
						<td class="GrayOutLine">
							<INPUT TYPE="checkbox" NAME="DebitItems" value='<%=RS1("ID")%>' onclick="setPrice2(this)"><%=sourceLink%></td>
					</tr>
				<%end if%>
					<tr>
						<td class="GrayOutLine">„«‰œÂ: <%=Separate(RS1("RemainedAmount"))%><INPUT TYPE="hidden" name="price2" value="<%=RS1("RemainedAmount")%>"></td>
					</tr>
				</table></td></tr>
<%
				TotalDebit = TotalDebit + cdbl(RS1("RemainedAmount"))
				RS1.movenext
			wend
		end if
		 '-----------------  End of The  Debit Items
	end if
%>		
		</Tbody>
		</TABLE></TD>
	</TR>
		</FORM>
	</TABLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
document.all.CashAmount.focus();

tmpColor="#FFDDDD"
tmpColor2="#FFFFBB"

function changeSystem(){	
	window.location='?act=getReceipt&selectedCustomer=<%=customerID%>&Reason='+document.all.Reason.value;
}

function setPrice2(obj){
	document.getElementById('asInvoice').checked=true;
	a= obj.type
	ii=parseInt(obj.id) 
	if(obj.checked){
		theTR = obj.parentNode.parentNode
		theTR.setAttribute("bgColor",tmpColor2)
	}
	else{
		theTR = obj.parentNode.parentNode
		theTR.setAttribute("bgColor",tmpColor)
	}
	addAllPrice2();
}
function addAllPrice2(){
	totalPrice = 0 
	va=""
	description="»«»  "
	checkBoxList = document.getElementsByName("DebitItems")
	for(i=0;i<document.getElementsByName("price2").length;i++) {
		if(checkBoxList[i].checked){
			totalPrice =  txt2val(totalPrice) + txt2val(document.getElementsByName("price2")[i].value) ;
			description = description + va + document.getElementsByName("price2")[i].parentNode.parentNode.parentNode.getElementsByTagName("td")[0].innerText;
			va = " Ê "
		}
	}
	if (description == "»«»  ") description=""
	document.all.SelectedTotalPrice.value = val2txt(totalPrice)
	document.all.CashDescription.value  = description;
}
function submitCeck(){	
	//alert(submitCeck2())
	if (submitCeck2()) 
	document.all.f1.submit();
}

function submitCeck2(){
	if (document.getElementsByName('TotalAmount')[0].value=='0') return false;
	if (document.all.SelectedTotalPrice.value==0) return true;
	//if (document.all.AccountTitle.value=='') return false;
	if (document.all.SelectedTotalPrice.value!=document.all.TotalAmount.value)
		return confirm("„»·€ œ—Ì«› Ì »« »œÂÌ Â«Ì «‰ Œ«» ‘œÂ »—«»— ‰Ì” . «œ«„Â „Ì œÂÌœø")
	return true;
}
addAllPrice2();
//-->
</SCRIPT>

<%
elseif request("act")="submitReceipt" then

	ON ERROR RESUME NEXT
		Reason=			cint(request.form("Reason"))
		CustomerID=		clng(request.form("CustomerID"))
		TotalAmount=	cdbl(text2value(request.form("TotalAmount")))
		CashAmount=		cdbl(text2value(request.form("CashAmount")))
		DepositAmount=	cdbl(text2value(request.form("DepositAmount")))
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
		end if
	ON ERROR GOTO 0

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

	CashDescription=	SqlSafe(request.form("CashDescription"))
	CashDescription=	left(CashDescription,500) ' The Description field in ReceivedCash table in DB is 500 Bytes
	
	creationDate=		shamsiToday()
	creationTime=		CurrentTime10()
	' ######################################################### 
	'		Find Receipts Related A or B Invoice or set it as Default in Accounts Ö
	' #########################################################
	' ------ All this added by SAM
	
	if request("DebitItems").count>0 then 
		if sys="AR" then
			isAs=""
			for i=1 to request("DebitItems").count
				mySQL="select isnull(invoices.IsA,2) as isA from ARItems inner join Invoices on arItems.link=invoices.ID where arItems.type=1 and arItems.id=" & request("DebitItems")(i)
				set RSSS=conn.Execute(mySQL)
				if RSSS("isA") = "True" then 
					isAlast = 1
				elseif RSSS("isA") = "False" then 
					isAlast = 0
				else 
					isAlast = 2
				end if
'				response.write mySQL & "<br>"
'				response.write isAlast & "<br>"
'				response.write RSSS("isA") & "<br>"
'				response.end
				if isAlast =2 then 
					if request("isA")<>"-1" then 
						isA=CInt(request("isA"))
					else
						response.redirect "?errMsg=" & server.URLEncode("Œÿ«! ›«ﬂ Ê— ÅÌœ« ‰„Ì‘Â° »Â„ »êÊ ﬂÂ «·›Â Ì« »")
					end if
				else
					isAs=isAs & CStr(isAlast)
				end if
				RSSS.close
			next
'			response.write isAs
'			response.end
			if isAlast  = 1 then 
				if InStr(isAs,"0")>0 then 
					response.redirect "?errMsg=" & server.URLEncode("Œÿ«! ›«ﬂ Ê— «·› Ê » —Ê œ—Â„  Ìﬂ ‰“‰")
				else 
					isA=1
				end if
			else
				if InStr(isAs,"1")>0 then 
					response.redirect "?errMsg=" & server.URLEncode("Œÿ«! ›«ﬂ Ê— «·› Ê » —Ê œÊÂ„  Ìﬂ ‰“‰")
				else
					isA=0
				end if
			end if
		else
			if request("isA")<>"-1" then 
				isA=CInt(request("isA"))
			else
				response.redirect "?errMsg=" & server.URLEncode("Œÿ«! “Ì— ”Ì „ ›—Ê‘ ‰Ì” ° ‰„Ìù Ê‰„ »›Â„„ ﬂÂ «Ì‰ œ—Ì«›  «·›Â Ì« ». ·ÿ›« »Â„ »êÊ")
			end if
		end if 
	else
		if request("isA")<>"-1" then 
			isA=CInt(request("isA"))
		else
			response.redirect "?errMsg=" & server.URLEncode("Œÿ«! œ— ’Ê— Ì ﬂÂ ‰„ÌùœÊ“Ìœ° »«Ìœ «·› Ì« » »Êœ‰ Å—œ«Œ  —« „⁄Ì‰ ﬂ‰Ìœ")
		end if
	end if	
'	response.write isAs & ", " & isA
'	response.end
	'###########################################################
	'###########################################################
	'###########################################################
	if isA then 
		GLAccount=			"11007"		'This must be changedÖ (Cashier A)
	else
		GLAccount=			"11005"		'This must be changedÖ (Cashier B)
	end if
	effectiveDate=		sqlSafe(request.form("ReceiptDate"))

	'---- Checking wether EffectiveDate is valid in current open GL
	if (effectiveDate < session("OpenGLStartDate")) OR (effectiveDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?act=getReceipt&selectedCustomer="& CustomerID & "&Reason="& Reason & "&errMsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	a = TotalAmount
	TotalAmountLetter = ConvertIT(a)

	' ######################################################### 
	'					INSERTING RECEIPT  Ö
	' #########################################################

	'	NUMBER must be Receipt Number
	NUMBER =0
	'	CHQAMOUNT Ö
	CHQAMOUNT =0
	'	CHQQTTY must be Receipt Qtty of Cheques
	CHQQTTY =0

'	Changed By Kid 821016
	mySQL="INSERT INTO Receipts (SYS, CreatedDate, CreatedBy, EffectiveDate, Number, Customer, CashAmount, DepositAmount, ChequeAmount, ChequeQtty, TotalAmount, TotalAmountLetter) VALUES ('"& Sys & "', N'" &_
	creationDate & "', '"& session("ID") & "', N'" & effectiveDate & "', '"& NUMBER & "', '"& CustomerID & "', '"& CashAmount & "', '"& DepositAmount & "', '"& CHQAMOUNT & "', '"& CHQQTTY & "', '"& TotalAmount & "', N'"& TotalAmountLetter & "');SELECT @@Identity AS NewReceipt"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	ReceiptID=RS1("NewReceipt")
	RS1.close

	' ######################################################### 
	'					INSERTING CASH  Ö
	' #########################################################
	mySQL="INSERT INTO ReceivedCash (Receipt, Amount, Banker, Description) VALUES ('"&_
			ReceiptID & "', '" & CashAmount & "', '" & theBanker & "', N'"& CashDescription & "')"
	conn.Execute(mySQL)

	chequeCount=0
	totalChequeAmount=0

	firstStatus=1 	'----------------- FIRST STATUS FOR CHEQUES
	for i=1 to request.form("ChequeNos").count 
		theChequeNo = text2value(request.form("ChequeNos")(i))
		theChequeDate = request.form("ChequeDates")(i)
		theOriginBank = request.form("Banks")(i)
		theOriginBank = left(theOriginBank,50) ' The BankOfOrigin field in ReceivedCheques table in DB is 50 Bytes long
		theDescription = request.form("Descriptions")(i)
		theAmount = text2value(request.form("Amounts")(i))
		if theAmount <> 0 then

			' ######################################################### 
			'					INSERTING CHEQUES Ö
			'				(Note: There is a TRIGGER)
			' #########################################################

			mySQL="INSERT INTO ReceivedCheques (Receipt, ChequeNo, ChequeDate, Description, BankOfOrigin, Amount, LastStatus, LastBanker, LastUpdatedDate, LastUpdatedBy) VALUES ('"&_
					ReceiptID & "', N'" & theChequeNo & "', N'" & theChequeDate & "', N'" & theDescription & "', N'" & theOriginBank & "', '" & theAmount & "', '" & firstStatus & "', '" & theBanker & "', N'" & creationDate & "', '" & session("ID") & "')"
			conn.Execute(mySQL)
			chequeCount = chequeCount + 1
			totalChequeAmount = totalChequeAmount + theAmount
		end if
	next
	' ######################################################### 
	'		CREATING AN ITEM for RECEIPT Ö
	' #########################################################
	'*** Type = 2 means Item is a Receipt
	
	mySQL="INSERT INTO "& Sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, Reason, IsCredit, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount) VALUES ('" &_
	GLAccount & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& CustomerID & "', '"& EffectiveDate & "', '"& Reason & "', 1, 2, '"& ReceiptID & "', '"& TotalAmount - DepositAmount & "', N'"& creationDate & "', '"& session("ID") & "', '"& TotalAmount - DepositAmount & "'); SELECT @@Identity AS NewItem;"
	Set RS1=conn.Execute(mySQL).NextRecordSet
	theReceiptItem=RS1("NewItem")
	RS1.close
	' ######################################################### 
	'		CREATING ITEMS RELATIONS  Ö
	' #########################################################
	response.write "<div dir=LTR>selected debit items count :<br>"
	OnHandAmount = TotalAmount
	for i = 1 to request.form("DebitItems").count
		mySQL="SELECT RemainedAmount FROM "& Sys & "Items WHERE (ID='"& request.form("DebitItems")(i) & "')"
		Set RS1=Conn.execute(mySQL)
		theAmount=cdbl(RS1("RemainedAmount"))
		if theAmount > cdbl(OnHandAmount) then theAmount=OnHandAmount
		OnHandAmount = OnHandAmount - theAmount
		mySQL="INSERT INTO "& Sys & "ItemsRelations (CreatedDate, CreatedBy, Credit"& Sys & "Item, Debit"& Sys & "Item, Amount) VALUES (N'"& creationDate & "', '"& session("ID") & "', '"& theReceiptItem &"', '"& request.form("DebitItems")(i) &"', '"& theAmount &"')"
		conn.Execute(mySQL)
		mySQL="UPDATE "& Sys & "Items SET RemainedAmount=RemainedAmount-'"& theAmount & "' WHERE (ID='"& request.form("DebitItems")(i) & "')"
		conn.Execute(mySQL)
		if OnHandAmount=0 then exit for
	next 
	mySQL="UPDATE "& Sys & "Items SET RemainedAmount='"& OnHandAmount & "' WHERE (ID='"& theReceiptItem & "')"
	conn.Execute(mySQL)
	
	mySQL="UPDATE "& Sys & "Items SET FullyApplied=1 WHERE (RemainedAmount=0) AND (voided=0) AND (FullyApplied=0)"
	conn.Execute(mySQL)
	' ######################################################### 
	'		INSERTING RECEIPT INTO CASH REGISTER LINES Ö
	'		Note:	Inserted CashRegisterLineType is '1'
	'				That means "Receipt"
	' #########################################################
	mySQL="INSERT INTO CashRegisterLines (CashReg, [Date], [Time], Type, Link,isA) VALUES ('"& CashRegID & "', N'"& creationDate & "', N'"& creationTime & "', '1', '"& ReceiptID & "', '"& isA &"')"
'	response.write mySQL
	conn.Execute(mySQL)
	
	' ######################################################### 
	'		UPDATING CASH REGISTERS Ö
	' #########################################################
	if isA then 
		mySQL="UPDATE CashRegisters SET CashAmountA=CashAmountA+'"& CashAmount & "', ChequeAmount=ChequeAmount+'"& totalChequeAmount & "', ChequeQtty=ChequeQtty+'"& chequeCount & "' WHERE ID='"& CashRegID & "'"
	else
		mySQL="UPDATE CashRegisters SET CashAmountB=CashAmountB+'"& CashAmount & "', ChequeAmount=ChequeAmount+'"& totalChequeAmount & "', ChequeQtty=ChequeQtty+'"& chequeCount & "' WHERE ID='"& CashRegID & "'"
	end if
	conn.Execute(mySQL)

	'*****************************************************************
	'*********** Creating Deposit if needed should go here ***********
	'*****************************************************************

	' ######################################################### 
	'		UPDATING Account Balance Ö
	' #########################################################
	mySQL="UPDATE Accounts SET "& Sys & "Balance = "& Sys & "Balance + '"& TotalAmount - DepositAmount & "' WHERE (ID='"& CustomerID & "')"
	conn.Execute(mySQL)
	Conn.close
	response.redirect "ReceiptInput.asp?act=ShowReceipt&id=" & ReceiptID
'response.end
else%>
<!-- Ã” ÃÊ »—«Ì ‰«„ Õ”«» --><BR><BR>
	<FORM METHOD=POST ACTION="?act=submitsearch" onsubmit="if (document.all.SearchBox.value=='') return false;">
	<div dir='rtl'>&nbsp;<B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«» Ì« ‘„«—Â ”›«—‘ </B>
		<INPUT TYPE="text" NAME="SearchBox">&nbsp;
		<INPUT class="GenButton" TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.SearchBox.focus();
	//-->
	</SCRIPT>
<%
end if
conn.Close
%>
<% if request("act")="getReceipt" then %>

<script language="JavaScript">
<!--
function delRow(rowNo){
	chqTable=document.getElementById("ChequeLines");
	theRow=chqTable.getElementsByTagName("tr")[rowNo];
	chqTable.removeChild(theRow);

	for (rowNo=0; rowNo < document.getElementsByName("ChequeNos").length; rowNo++){
		chqTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0].innerText= rowNo+1;
	}
}
function addRow(rowNo){
	chqTable=document.getElementById("ChequeLines");
	theRow=chqTable.getElementsByTagName("tr")[rowNo];
	newRow=document.createElement("tr");
	newRow.setAttribute("bgColor", '#f0f0f0');

	tempTD=document.createElement("td");
	tempTD.innerHTML=rowNo+1
	tempTD.setAttribute("align", 'center');
	tempTD.setAttribute("width", '25');
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="ChequeNos" size="12" onKeyPress="return maskNumber(this);">'
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML='<INPUT class="RcpRowInput" style="text-align:left;" TYPE="text" NAME="ChequeDates" maxlength="10" size="10" onblur="acceptDate(this)">'
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="Banks" size="10">'
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="Descriptions" size="20">'
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="Amounts" size="15" onKeyPress="return maskNumber(this);" onBlur="setPrice(this);">'
	newRow.appendChild(tempTD);

	chqTable.insertBefore(newRow,theRow);

	chqTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
}

function setPrice(src){
	myRow=src.parentNode.parentNode.rowIndex

	if (src.name=="Amounts" && document.getElementsByName("ChequeNos")[myRow].value==''){
		src.value=0;
	}
	else{
		src.value=val2txt(txt2val(src.value));
	}
	cashAmount=parseInt(txt2val(document.getElementsByName("CashAmount")[0].value));
	totalAmount = cashAmount;

	for (rowNo=0; rowNo < document.getElementsByName("Amounts").length; rowNo++){
		totalAmount += parseInt(txt2val(document.getElementsByName("Amounts")[rowNo].value));
	}
	document.all.TotalAmount.value = val2txt(totalAmount);
}

//-->
</script>
<%end if%>
<!--#include file="tah.asp" -->
