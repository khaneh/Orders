<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CashRegister (9)
PageTitle="Å—œ«Œ  ‰ﬁœ"
SubmenuItem=2
if not Auth(9 , 2) then NotAllowdToViewThisPage()

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
	.CPayTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; }
	.CPayMainTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; background-color: #558855; text-align:right; direction: RTL;}
	.CPayMainTable TH { background-color: #C3C300; font-size: 9pt; font-weight:normal;}
	.CPayMainTableTR { background-color: #CCCC88; border: 0; }
	.CPayInput1 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; direction:RTL;}
	.CPayInput2 { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction:LTR;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;
var currentRow=null;
//-->
</SCRIPT>
<font face="tahoma">
<%
mySQL="SELECT * FROM CashRegisters WHERE (IsOpen=1) AND (Cashier='"& session("ID") & "')"
Set RS1= conn.Execute(mySQL)
if RS1.eof then 
%><br><br>
	<TABLE width='70%' align='center'>
	<TR>
		<TD align=center bgcolor=#FFBBBB style='border: solid 1pt black'><BR><b>Ã‰«» <%=CSRName%> ‘„« ’‰œÊﬁ »«“ ‰œ«—Ìœ ... <br><br>«“ œ”  „‰ ﬂ«—Ì »—«Ì ‘„« »— ‰„Ì ¬Ìœ <br>„ «”›„."</b><BR><BR></TD>
	</TR>
	</TABLE>
<%	conn.close
	response.end
else
	CashRegID=RS1("ID")
	theBanker=RS1("Banker")
	CashRegName=RS1("NameDate")
	cashAmountA=cdbl(RS1("cashAmountA"))
	cashAmountB=cdbl(RS1("cashAmountB"))
	if cashAmountA=0 and cashAmountB=0 then 
		%><br><br>
			<TABLE width='70%' align='center'>
			<TR>
				<TD align=center bgcolor=#FFBBBB style='border: solid 1pt black'><BR><b>‘„« ﬂÂ Â‰Ê“ ÅÊ·Ì ‰ê—› Ì çÂ ÿÊ—Ì „ÌùŒÊ«Ì Å—œ«Œ  œ«‘ Â »«‘Ì!</b><BR><BR></TD>
			</TR>
			</TABLE>
		<%	
		conn.close
		response.end
	end if
	Set RS1=nothing
end if

if request("act")="submitsearch" then
	if request("CustomerNameSearchBox") <> "" then
		SA_TitleOrName=request("CustomerNameSearchBox")
		SA_Action="return true;"
		SA_SearchAgainURL=""
		SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<FORM METHOD=POST ACTION="?act=getPayment">
		<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	end if
elseif request("act")="selectOrder" then
	if request("selectedCustomer") <> "" then
		SO_Customer=request("selectedCustomer")
		SO_Action="return true;"
		SO_StepText="ê«„ ”Ê„ :"
%>
		<FORM METHOD=POST ACTION="?act=getPayment">
		<!--#include File="../AR/include_SelectOrder.asp"-->
		</FORM>
<%
	end if
elseif request("act")="getPayment" then
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

	AccountID=clng(request("selectedCustomer"))

	mySQL="SELECT * FROM Accounts WHERE (ID='"& AccountID & "')"
	Set RS1 = conn.Execute(mySQL)
	AccountNo=RS1("ID")
	customerName=RS1("AccountTitle")
	
%>
	<br><div dir='rtl'><B>ê«„ ”Ê„ : Ê—Êœ «ÿ·«⁄«  œ—Ì«›  ÅÊ·</B>
	</div>
<!-- Ê—Êœ «ÿ·«⁄«  œ—Ì«›  ÅÊ· -->
	<hr>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
		<table class="CPayMainTable" Cellspacing="1" Cellpadding="0" Width="500" align="center">
		<FORM METHOD=POST ACTION="?act=submitCashPayment" onsubmit="if (document.all.AccountTitle.value=='') return false;">
			<tr>
			<th colspan="10" align='center' height='25px'>’‰œÊﬁ <span dir='LTR'><%=CashRegName%></span> - <%=CSRName%> <INPUT TYPE="hidden" Name="CashRegID" Value="<%=CashRegID%>">
			</th>
			</tr>
			<tr>
			<th colspan="10"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL"><TR>
				<TD><table>
					<tr>
						<td align="left">Õ”«»:</td>
						<td align="right">
							<span id="customer">
								<INPUT TYPE="hidden" NAME="AccountID" value="<%=AccountID%>"><span><%=CustomerName%></span>.
							</span></td>
					</tr>
					</table></TD>
				<TD><table>
					<tr>
						<td align="left">„—»Êÿ »Â :</td>
						<td align="right">
							<SELECT NAME="Reason" style='font-family:Tahoma;font-size:8pt;height:25px;width:100px;' onchange="changeSystem();">
<%							mySQL="SELECT * FROM AXItemReasons WHERE Display=1 ORDER BY ID"
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
%>							</SELECT>
						</td>
					</tr>
					</table></TD>
				<td>«·›
				<input name="isA" type="checkbox" 
					<% 
					if cashAmountA=0 or cashAmountB=0 then 
						if cashAmountB=0 then 
							response.write " checked='checked' "
							response.write " onclick='this.checked=true;' "
							response.write " title='›ﬁÿ ’‰œÊﬁ «·› „ÊÃÊœÌ œ«—œ' "
						else
							response.write " onclick='this.checked=false;' "
							response.write " title='›ﬁÿ ’‰œÊﬁ » „ÊÃÊœÌ œ«—œ' "
						end if
					else 
						response.write " checked='checked' "	
					end if
					%>
				>
				</td>
				<TD align="left"><table>
					<tr>
						<td align="left"> «—ÌŒ:</td>
						<td dir="LTR">

							<INPUT class="CPayInput2" style="text-align:left;direction:LTR;" NAME="PaymentDate" TYPE="text" maxlength="10" size="10" value="<%=shamsiToday()%>" onblur="acceptDate(this)"></td>
						<td dir="RTL"><%=weekdayname(weekday(date))%></td>
					</tr>
					</table></TD>
				</TR></TABLE>
				<input type="hidden" Name='InvoiceID' value='<%=InvoiceID%>'>
			</th>
			</tr>
			<tr class="CPayMainTableTR">
			<TD colspan="10" ><div>
				<TABLE class="CPayTable" Cellspacing="1" Cellpadding="0" Dir="RTL">
				<tr height="40">
					<td colspan="1" align="left"> „ﬁœ«— ‰ﬁœ: </td>
					<td colspan="1"><INPUT dir="LTR" class="CPayInput1" TYPE="text" Name="CashAmount" size="15" onKeyPress="return maskNumber(this);" onBlur="setPrice(this)"></td>
				</tr>
				<tr height="40">
					<td colspan="1" align="left"> ‘—Õ: </td>
					<td colspan="1"><INPUT dir="LTR" class="CPayInput1" TYPE="text" Name="Description" size="60"></td>
				</tr>
				</TABLE></div></TD>
			</tr>
			<tr>
			<th colspan="10" align='center' height='25px'>
				<INPUT class="GenButton" style="text-align:center" TYPE="button" value="–ŒÌ—Â" onclick="if (document.getElementsByName('CashAmount')[0].value!='0') submit();">
				<INPUT class="GenButton" style="text-align:center" TYPE="button" value="«‰’—«›" onclick="window.location='';">
			</th>
		</table>
		</FORM>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			document.all.CashAmount.focus();

		function changeSystem(){	
			window.location='?act=getPayment&selectedCustomer=<%=AccountID%>&Reason='+document.all.Reason.value;
		}

		//-->
		</SCRIPT>


<%elseif request("act")="submitCashPayment" then

	ON ERROR RESUME NEXT
		Reason=			cint(request.form("Reason"))
		AccountID=		clng(request.form("AccountID"))
		CashAmount=		cdbl(text2value(request.form("CashAmount")))
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
		end if
	ON ERROR GOTO 0
	if request("isA")="on" then 
		isA=1
	else 
		isA=0
	end if
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

	Description=sqlSafe(request.form("Description"))

	creationDate=	shamsiToday()
	creationTime=	CurrentTime10()
	if isA=0 then 
		GLAccount=		"11005"		'This must be changed... (Cashier B)
	else
		GLAccount=		"11007"		'This must be changed... (Cashier A)
	end if

	effectiveDate=	sqlSafe(request.form("PaymentDate"))

	'---- Checking wether EffectiveDate is valid in current open GL
	if (effectiveDate < session("OpenGLStartDate")) OR (effectiveDate > session("OpenGLEndDate")) then
		Conn.close

		response.redirect "?act=getPayment&selectedCustomer="& AccountID & "&Reason="& Reason & "&errMsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'------- Check total cash A or B amount larger than this payment :)
	if isA=1 then
		if cashAmountA < CashAmount then  
			Conn.close
			response.redirect "?errMsg=" & Server.URLEncode("œﬁ ! „ÊÃÊœÌ ’‰œÊﬁ «·› ‘„« «“ „Ì“«‰ Å—œ«Œ Ì ﬂ„ — «” ")
		end if
	else
		if cashAmountB < CashAmount then  
			Conn.close
			response.redirect "?errMsg=" & Server.URLEncode("œﬁ ! „ÊÃÊœÌ ’‰œÊﬁ » ‘„« «“ „Ì“«‰ Å—œ«Œ Ì ﬂ„ — «” ")
		end if
	end if
	'----
	' ######################################################### 
	'					INSERTING PAYMENT  ...
	' #########################################################

	'	NUMBER must be Payment Number
	NUMBER =0
	mySQL="INSERT INTO Payments (SYS, Account, CashAmount, ChequeAmount, CreatedBy, CreationDate, CreationTime, EffectiveDate) VALUES ('"& Sys & "', '" &_
	AccountID & "', '"& CashAmount & "', 0, '"& session("ID") & "', N'"& creationDate & "', N'"& creationTime & "', N'"& effectiveDate & "');SELECT @@Identity AS NewPayment"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	paymentID = RS1 ("NewPayment")
	RS1.close

	' ######################################################### 
	'					INSERTING CASH  ...
	' #########################################################
	mySQL="INSERT INTO PaidCash (Payment, Description, Amount, Banker) VALUES ('"&_
			PaymentID & "', N'" & Description & "', '" & CashAmount & "', '" & theBanker & "')"
	conn.Execute(mySQL)

	' ######################################################### 
	'		INSERTING PAYMENT INTO CASH REGISTER LINES ...
	'		Note:	Inserted CashRegisterLineType is '2'
	'				That means "Payment"
	' #########################################################
	mySQL="INSERT INTO CashRegisterLines (CashReg, [Date], [Time], Type, Link,isA) VALUES ('"& CashRegID & "', N'"& creationDate & "', N'"& creationTime & "', '2', '"& PaymentID & "',"& isA &")"
	conn.Execute(mySQL)

	' ######################################################### 
	'		UPDATING CASH REGISTERS ...
	' #########################################################
	if isA then 
		mySQL="UPDATE CashRegisters SET CashAmountA=CashAmountA-'"& CashAmount & "' WHERE ID='"& CashRegID & "'"
	else
		mySQL="UPDATE CashRegisters SET CashAmountB=CashAmountB-'"& CashAmount & "' WHERE ID='"& CashRegID & "'"
	end if
	conn.Execute(mySQL)

	'**************************** Creating AN Item for Payment  ****************
	'*** Type = 5 means Item is a Payment

	mySQL="INSERT INTO "& Sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, Reason, IsCredit, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount) VALUES ('" &_
	GLAccount & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& AccountID & "', '"& effectiveDate & "', '"& Reason & "', 0, 5, '"& PaymentID & "', '"& CashAmount & "', N'"& creationDate & "', '"& session("ID") & "', '"& CashAmount & "')"	
	conn.Execute(mySQL)
	'***-------------------End of Creating Item for Payment  ----------------

	' ######################################################### 
	'		UPDATING Account Balance ...
	' #########################################################
	mySQL="UPDATE Accounts SET "& Sys & "Balance = "& Sys & "Balance - '"& CashAmount & "' WHERE (ID='"& AccountID & "')"
	conn.Execute(mySQL)

	Conn.close
	response.redirect "../"& Sys & "/AccountReport.asp?Sys="& Sys & "&act=showPayment&payment="& PaymentID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  Å—œ«Œ  ÅÊ· À»  ‘œ.")

else%>
<!-- Ã” ÃÊ »—«Ì ‰«„ Õ”«» --><BR><BR>
	<FORM METHOD=POST ACTION="?act=submitsearch" onsubmit="if (document.all.CustomerNameSearchBox.value=='') return false;">
	<div dir='rtl'>&nbsp;<B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«»</B>
		<INPUT TYPE="text" NAME="CustomerNameSearchBox">&nbsp;
		<INPUT class="GenButton" TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.CustomerNameSearchBox.focus();
	//-->
	</SCRIPT>
<%
end if
conn.Close
%>
</font>
</BODY>
</HTML>
<% if request("act")="getPayment" then %>

<script language="JavaScript">
<!--
function setPrice(src){
/* 	src.value=val2txt(txt2val(src.value)); */
	$(src).val(echoNum(getNum($(src).val())));
}

//-->
</script>
<%end if%>
<!--#include file="tah.asp" -->
