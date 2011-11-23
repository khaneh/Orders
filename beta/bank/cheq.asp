<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "’œÊ— çﬂ"
SubmenuItem=1
if not Auth("A" , 1) then NotAllowdToViewThisPage()

sendTo = session("id")
editFlag = 0

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
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
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
	.CustGenTable  { font-family:tahoma; font-size: 9pt;}
	.CustGenInput { font-family:tahoma; font-size: 9pt;}
</STYLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;
var currentRow=null;

var tempKeyBuffer;
function myKeyDownHandler(){
	tempKeyBuffer=window.event.keyCode;
}
function myKeyPressHandler(){
//	alert (tempKeyBuffer)
	if (tempKeyBuffer>=65 && tempKeyBuffer<=90){
		window.event.keyCode=tempKeyBuffer+32;
	}
	else if(tempKeyBuffer==186){
		window.event.keyCode=59;
	}
	else if(tempKeyBuffer==188){
		window.event.keyCode=44;
	}
	else if(tempKeyBuffer==190){
		window.event.keyCode=46;
	}
	else if(tempKeyBuffer==191){
		window.event.keyCode=47;
	}
	else if(tempKeyBuffer==192){
		window.event.keyCode=96;
	}
	else if(tempKeyBuffer>=219 && tempKeyBuffer<=221){
		window.event.keyCode=tempKeyBuffer-128;
	}
	else if(tempKeyBuffer==222){
		window.event.keyCode=39;
	}
}
//-->
</SCRIPT>
<font face="tahoma">
<!-- <div dir='rtl'><B>Ê—Êœ ”›«—‘ </B>
</div> -->
<%
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Submit Search
'-----------------------------------------------------------------------------------------------------
if request("act")="submit search" then
	if request("CustomerNameSearchBox") <> "" then
		SA_TitleOrName=sqlSafe(request("CustomerNameSearchBox"))
		mySQL="SELECT * FROM Accounts WHERE (REPLACE(AccountTitle, ' ', '') LIKE REPLACE(N'%"& SA_TitleOrName & "%', ' ', '') ) ORDER BY AccountTitle"
		Set RS1 = conn.Execute(mySQL)

		if (RS1.eof) then 
			response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='../CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")

		' Not Found	%>
		<div dir='rtl'><B>ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ</B> &nbsp; <A HREF="cheq.asp" style='font-size:7pt;'>Ã” ÃÊÌ „Ãœœ</A>
		</div><br>
	<%	else 
			SA_TitleOrName=request("CustomerNameSearchBox")
			SA_Action="return true;"
			SA_SearchAgainURL="cheq.asp"
			SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<FORM METHOD=POST ACTION="cheq.asp?act=enterCheque">
			<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
		end if
	end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------- Submit Payment
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submitPayment" or editFlag = 1 then

	ON ERROR RESUME NEXT

		errorFound=false
		CustomerID=			clng(request.form("CustomerID"))
		accountID=			clng(request.form("accountID"))
		TotalAmount=		cdbl(text2value(request.form("TotalAmount")))
		VouchIDS=			request.form("VouchID")
		VoucherTotalPrices= request.form("VoucherTotalPrice")
		SumVoucherTotalPrice= cdbl(request.form("SumVoucherTotalPrice"))
		Reason=				cint(request.form("Reason"))

		if Err.Number<>0 then
			'Err.clear
			errorFound=	true
			'errorMsg=	Server.URLEncode("Œÿ«!")
			errorMsg=	Err.Description
			response.write errorMsg
			response.end

		end if

		if NOT errorFound then
			effectiveDate=	sqlSafe(request.form("PaymentDate"))
			'---- Checking wether EffectiveDate is valid in current open GL
			if (effectiveDate < session("OpenGLStartDate")) OR (effectiveDate > session("OpenGLEndDate")) then
				errorFound=	true
				errorMsg=	Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” .")
			end if 
			'----
			'----- Check GL is closed
			if (session("IsClosed")="True") then
				'Conn.close
				errorFound=true
				errorMsg= Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
			end if 
			'----			
			if TotalAmount=0 then 
				errorFound=	true
				errorMsg=	Server.URLEncode("<B>Œÿ«!  </B><BR>„Ì“«‰ Å—œ«Œ  ’›— «” <BR>")
			end if 
			CheqsCount=request.form("ChequeNos").count 
			for i = 1 to CheqsCount
				ChequeNo	= clng(request.form("ChequeNos")(i))
				ChequeDate	= sqlSafe(request.form("ChequeDates")(i))
				Banker		= cint(request.form("Banks")(i))
				Description = sqlSafe(request.form("Description")(i))
				Amount		= cdbl(text2value(request.form("Amounts")(i)))

				if ChequeNo=0 or Banker=-1 or Amount=0 then  ' -- or ChequeDate="" 
					errorFound=	true
					errorMsg=	Server.URLEncode("<B>Œÿ«!  </B><BR>«ÿ·«⁄«  çﬂ ‰«ﬁ’ «” <BR>")
				end if 
				'---------------------SAM---------------------------------------------------------------
				 
				set che = conn.execute("Select COUNT(*) AS cheque From PaidCheques inner join Payments on PaidCheques.payment = Payments.ID Where PaidCheques.ChequeNo = " & ChequeNo & " AND PaidCheques.ChequeDate = N'" & ChequeDate & "' AND Payments.Voided = 0")
				if cdbl(che("cheque")) > 0 then
					errorFound = true
					errorMsg = Server.URLEncode("<B>Œÿ«! </B><br> «ÿ·«⁄«  «Ì‰ çﬂ ﬁ»·« Ê«—œ ‘œÂ «” <br>")
					call showAlert ("<B>Œÿ«! </B><br> «ÿ·«⁄«  «Ì‰ çﬂ ﬁ»·« Ê«—œ ‘œÂ «” <b>",CONST_MSG_ERROR)
				end if
				che.close
			next
		end if
		if Err.Number<>0 then
			'Err.clear
			errorFound=	true
			'errorMsg=	Server.URLEncode("Œÿ«!")
			errorMsg=	Err.Description
		end if
	ON ERROR GOTO 0

	if errorFound then
		conn.close
		response.redirect "?act=getPayment&selectedCustomer="& CustomerID & "&Reason="& Reason & "&errMsg=" & errorMsg
	end if

	mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
	else
		Sys=			RS1("Acron")
		firstGLAccount = accountID
	end if
	RS1.close

	RemainedTotalAmount = TotalAmount

	creationDate=	shamsiToday()
	creationTime=	currentTime10()
	GLAccount=		"NULL"

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	dim APRemainedAmount(10)
	dim APCreditID(10)
	dim relationAmount(10)
	dim n

	if SumVoucherTotalPrice <> 0 and VouchIDS<>"" then		'Balance this payment with a previous APItem.

		VouchID = split(VouchIDS, ",")
		VoucherTotalPrice = split(VoucherTotalPrices, ",")
		n = ubound(VouchID)

		for vch = 0 to n  
			set RSB=Conn.Execute ("SELECT * FROM  APItems WHERE (Type=6 and Link="& VouchID(vch) & ")" )
			if RSB.EOF then
				response.write "<br><BR>"
				call showAlert ("<B>Œÿ«!  </B><BR>«‘ﬂ«·Ì œ— Å—œ«Œ  «Ì‰ ›«ﬂ Ê— ÊÃÊœ œ«—œ. <br>ÂÌç Å—œ«Œ  Ì« çﬂÌ À»  ‰ê—œÌœ.<BR>",CONST_MSG_ERROR)
				response.end
			end if
			if not trim(VoucherTotalPrice(vch)) = trim(RSB("AmountOriginal")) then
				response.write "<br><BR>"
				call showAlert ("<B>Œÿ«!  </B><BR>«‘ﬂ«·Ì œ— Å—œ«Œ  «Ì‰ ›«ﬂ Ê— ÊÃÊœ œ«—œ. <br>ÂÌç Å—œ«Œ  Ì« çﬂÌ À»  ‰ê—œÌœ.<BR>",CONST_MSG_ERROR)
				response.end
			end if
			APRemainedAmount(vch) = cdbl(RSB("RemainedAmount"))
			APCreditID(vch) = RSB("ID")
		next

		for vch = 0 to n  
			if RemainedTotalAmount = 0 then
				relationAmount(vch) = 0
			else
				if APRemainedAmount(vch) > RemainedTotalAmount then
					mySQL="UPDATE "& sys & "Items SET RemainedAmount='"& APRemainedAmount(vch) - RemainedTotalAmount & "' WHERE (ID='"& APCreditID(vch) & "')"
					conn.Execute(mySQL)	
					relationAmount(vch) = RemainedTotalAmount
					RemainedTotalAmount = 0 
				else
					mySQL="UPDATE "& sys & "Items SET RemainedAmount='"& 0 & "' WHERE (ID='"& APCreditID(vch) & "')"
					conn.Execute(mySQL)	
					RemainedTotalAmount   = RemainedTotalAmount - APRemainedAmount(vch)
					Conn.Execute ("UPDATE Vouchers SET paid = 1 WHERE (id = "& VouchID(vch) & ")")
					relationAmount(vch) = APRemainedAmount(vch)
				end if
			end if 
		next
	end if

	'************* Inserting the Payment
	mySQL="INSERT INTO Payments (SYS, Account, CashAmount, ChequeAmount, CreatedBy, CreationDate, CreationTime, EffectiveDate) VALUES ('"& sys & "', "& customerID & ", 0, " & TotalAmount & ","& session("id") & ", N'"& creationDate & "', N'"& creationTime & "', N'"& effectiveDate & "');SELECT @@Identity AS NewPayment"
	set RSE = Conn.execute(mySQL).NextRecordSet
	paymentID = RSE ("NewPayment")
	RSE.close
	set RSE = Nothing

	'************* Inserting Cheques
	for i = 1 to CheqsCount
		ChequeNo	= clng(request.form("ChequeNos")(i))
		ChequeDate	= sqlSafe(request.form("ChequeDates")(i))
		Banker		= cint(request.form("Banks")(i))
		Description = sqlSafe(request.form("Description")(i))
		Amount		= cdbl(text2value(request.form("Amounts")(i)))

		conn.Execute("INSERT INTO PaidCheques (Payment, CreatedDate, CreatedBy, ChequeNo, ChequeDate, Amount, Banker, Status, StatusSetBy, StatusSetDate, Description) VALUES  ("& paymentID & ",N'"& creationDate & "','"& session("ID") & "',"& ChequeNo& ",N'"& ChequeDate& "',"& Amount& ","& Banker& ", 0, "& session("ID")& ", N'"& creationDate & "', N'"& Description & "')")

	next

	'*************	Creating an Item for the Payment Ö
	'*** Type = 5 means Item is a Payment
	'*** GLAccount = NULL because it can't be identified right now, it depends on where the cheques areÖ
	mySQL="INSERT INTO "& sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, Reason, IsCredit, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount) VALUES ("&_
		GLAccount & ", '"& OpenGL & "', '"& firstGLAccount & "', '"& customerID & "', N'"& effectiveDate & "', '"& reason & "', 0, 5, '"& paymentID & "', '"& TotalAmount & "', N'"& creationDate & "', '"& session("ID") & "', '"& RemainedTotalAmount & "');SELECT @@Identity AS NewItem"
	set RSF = Conn.execute(mySQL).NextRecordSet
	APDebitID = RSF ("NewItem")
	RSF.close 
	Set RSF = Nothing

	mySQL="UPDATE Accounts SET "& sys & "Balance="& sys & "Balance-"& TotalAmount & "  WHERE (ID = "& customerID & ")"
	conn.Execute(mySQL)

	'*************	Create the relations between 
	'				the Payment and Selected Vouvhers 
	if SumVoucherTotalPrice<> 0 and VouchIDS<>"" then
		for vch = 0 to n  
			if relationAmount(vch) = 0 then exit for
			mySQL="INSERT INTO "& sys & "ItemsRelations (CreatedDate, CreatedBy, Credit"& sys & "Item, Debit"& sys & "Item, Amount) VALUES (N'"& creationDate & "', '"& session("ID") & "', '"& APCreditID(vch) &"', '"& APDebitID &"', '"& relationAmount(vch) &"')"
			conn.Execute(mySQL)
		next
	end if
	''''------------------------------------------------

	Conn.close
	response.redirect "../"& sys & "/AccountReport.asp?act=showPayment&payment="& paymentID & "&msg=" & Server.URLEncode("Å—œ«Œ  «‰Ã«„ ‘œ.")

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Input cheques
'-----------------------------------------------------------------------------------------------------
elseif request("act")="enterCheque" then 
	customerID=request("selectedCustomer")
	VouchIDS = request("VouchID")
	lastCustomerID = -1
	SumAPAmountOriginal=0
	SumAPRemainedAmount=0
	SumVoucherTotalPrice=0
	VoucherTotalPrice = -1
	Reason = 0 ' Old Code: Reason = 2 ' Sys: AP
	if request("Reason")<>"" then Reason=request("Reason")

	if VouchIDS<>"" then
		VouchIDarray = split(VouchIDS,",")
		for i=0 to ubound(VouchIDarray) 
			VouchID = VouchIDarray(i)

			set RSV=Conn.Execute ("SELECT * FROM  Vouchers WHERE (id="& VouchID & ")" )
			if not RSV.eof then
				set RSB=Conn.Execute ("SELECT * FROM  APItems WHERE (Type=6 and Link="& VouchID & ")" )
				if RSB.EOF then
					response.write "<br><BR>"
					call showAlert ("<B>Œÿ«!  </B><BR>«„ﬂ«‰ Å—œ«Œ  «Ì‰ ›«ﬂ Ê— ÊÃÊœ ‰œ«—œ",CONST_MSG_ERROR)
					response.end
				end if
				customerID = RSV("VendorID")
				if lastCustomerID <> -1 and customerID<>lastCustomerID then
					response.write "<br><BR>"
					call showAlert ("<B>Œÿ«!  </B><BR>Â„Â ›«ﬂ Ê—Â« „—»Êÿ »Â Ìﬂ Õ”«» ‰Ì”  					<BR><BR><A HREF='payment.asp'>»—ê‘ </A>",CONST_MSG_ERROR)
					response.end
				end if
				LastCustomerID = customerID 
			else
				response.write "<br><BR>"
				call showAlert ("<B>Œÿ«!  </B><BR>ç‰Ì‰ ›«ﬂ Ê—Ì ÊÃÊœ ‰œ«—œ",CONST_MSG_ERROR)
				response.end
			end if
		next
	else
		if customerID = "" then
			response.redirect "cheq.asp"
		end if
	end if


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)
	customerName=RS1("AccountTitle")

	creationDate=	shamsiToday()
	creationTime=	currentTime10()
%>	
	<br>
	<div dir='rtl'>&nbsp;&nbsp;<B>ê«„ ”Ê„: Ê—Êœ «ÿ·«⁄«  çﬂ Â«</B>
	</div>
<!--#include File="../include_JS_InputMasks.asp"-->

<!-- Ê—Êœ «ÿ·«⁄«  œ—Ì«›  ÅÊ· -->
	<hr>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		function checkValidation(){
		  try{

			tmpBox=document.getElementsByName('PaymentDate')[0];
			if (tmpBox.value!='' && tmpBox.value!=null){
				if (!acceptDate(tmpBox))
					return false;
			}
			else{
				tmpBox.focus();
				alert(" «—ÌŒ Ê«—œ ‰‘œÂ")
				return false;
			}

			tmpBox=document.getElementsByName('TotalAmount')[0];
			if (txt2val(tmpBox.value)==0){
				alert("„ﬁœ«— Å—œ«Œ  ’›— «” ")
				return false;
			}

		  }catch(e){
				alert("Œÿ«Ì €Ì— „‰ Ÿ—Â");
				return false;
		  }

		}

	//-->

	var dialogActive=false;

	function maskReason(src){ 
		var theKey=event.keyCode;

		if (theKey==13){
			event.keyCode=9
			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì „⁄Ì‰:"
			var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('../accounting/dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
	//			dialogActive=false
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					document.all.accountName.value=Arguments[1];
				}
			}
		}
	}

	//-->
	</SCRIPT>
<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>

	<FORM METHOD=POST ACTION="?act=submitPayment" onsubmit="return checkValidation();">
		<input type="hidden" Name='tmpDlgArg' value=''>
		<table class="RcpMainTable" Cellspacing="1" Cellpadding="0" Width="500" align="center">
		<tr class="RcpMainTableTH">
			<td colspan="10"><TABLE width='550' align='center' cellpadding='5'>
			<TR>
				<TD bgcolor=white>”Ì” „: </TD>
<%
				mySQL="SELECT * FROM AXItemReasons WHERE Display=1 ORDER BY ID"
				set RS1=conn.execute(mySQL)
				while not RS1.eof
%>					<TD bgcolor=white>
						<INPUT TYPE="radio" NAME="Reason" value="<%=RS1("ID")%>" <% if reason=RS1("ID") then response.write "checked "%>><%=RS1("Name")%>
					</TD>
<%					RS1.movenext
				wend
%>		
			</TR></TABLE></td>
		</tr>
		<tr class="RcpMainTableTH">
			<td colspan="10"><TABLE class="RcpTable" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL"><TR>
				<TD align="left">Õ”«»:</TD>
				<TD align="right" width="5">
					<INPUT class="RcpGenInput" disabled TYPE="text" value="<%=customerID%>" maxlength="5" size="5" tabIndex="1" dir="LTR">
				</TD>
				<TD align="right" width="200">
					<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><%=CustomerName%>.
				</TD>
				<TD align="left"> «—ÌŒ:</TD>
				<TD>
					<table class="RcpTable">
					<tr>    
						<td dir="LTR">
							<input class="RcpGenInput" style="text-align:left;direction:LTR;" NAME="PaymentDate" TYPE="text" maxlength="10" size="10" value="" onblur="acceptDate(this)">
						</td>
						<td dir="RTL">«„—Ê“: <font color=white><%=weekdayname(weekday(date))%> <span dir=rtl><%=shamsiToday()%></span></font></TD>
					</tr>   
					</table></TD>
			</TR></TABLE></td>
		</tr>
		<tr class="RcpMainTableTR">
			<td colspan="10"><div>
			<TABLE class="RcpTable" Cellspacing="1" Cellpadding="0" Dir="RTL">
			<TR height="20">
				<TD colspan="15">
					<%
					if VouchIDS<>"" then
						response.write "„—»Êÿ »Â ›«ﬂ Ê—Â«Ì:"
						VouchIDarray = split(VouchIDS,",")
						for i=o to ubound(VouchIDarray) 
							VouchID = VouchIDarray(i)

							set RSV=Conn.Execute ("SELECT * FROM  Vouchers WHERE (id="& VouchID & ")" )
							set RSB=Conn.Execute ("SELECT * FROM  APItems WHERE (Type=6 and Link="& VouchID & ")" )
							APAmountOriginal = RSB("AmountOriginal")
							APRemainedAmount = RSB("RemainedAmount")
							VoucherTotalPrice = RSV("TotalPrice")
							SumAPAmountOriginal = cdbl(SumAPAmountOriginal) + cdbl(APAmountOriginal)
							SumAPRemainedAmount = cdbl(SumAPRemainedAmount) + cdbl(APRemainedAmount)
							SumVoucherTotalPrice = cdbl(SumVoucherTotalPrice) + cdbl(VoucherTotalPrice)
							customerID = RSV("VendorID")
							VoucherTitle = RSV("Title")
							%>
							<INPUT TYPE="hidden" name="VouchID" value="<%=VouchID%>">
							<INPUT TYPE="hidden" name="VoucherTotalPrice" value="<%=VoucherTotalPrice%>">
							<INPUT TYPE="hidden" name="APRemainedAmount" value="<%=APRemainedAmount%>">
							<li><B><%=VoucherTitle%></B>   »Â ‘„«—Â <A HREF="../AP/payment.asp?VouchID=<%=VouchID%>"><%=VouchID%></A>  (Ã„⁄ ›«ﬂ Ê—: <%=Separate(VoucherTotalPrice)%> —Ì«· - »«ﬁÌ„«‰œÂ: <%=Separate(APRemainedAmount)%> —Ì«·)
							<%
						next
					end if
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					%>
					<INPUT TYPE="hidden" name="SumAPAmountOriginal" value="<%=SumAPAmountOriginal%>">
					<INPUT TYPE="hidden" name="SumVoucherTotalPrice" value="<%=SumVoucherTotalPrice%>">
				</TD>
			</TR>
			</TABLE></div></td>
		</tr>
		<tr class="RcpMainTableTR">
			<td colspan="10"><div style="width:540;">
			<TABLE class="RcpTable" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TR height=20px>
				<TD class="RcpHeadInput" style="width:25;"> # </td>
				<TD class="RcpHeadInput" style="width:70;">‘„«—Â çﬂ</td>
				<TD class="RcpHeadInput" style="width:70;"> «—ÌŒ</td>
				<TD class="RcpHeadInput" style="width:150;">»«‰ﬂ</td>
				<TD class="RcpHeadInput" style="width:95;">„»·€</td>
				<TD class="RcpHeadInput" style="width:110;"> Ê÷ÌÕ</td>
			</TR>   
			</TABLE></div></td>
		</tr>
		<tr class="RcpMainTableTR">
			<td colspan="10" dir="RTL"><div style="overflow:auto; height:150px;width:100%;">
			<TABLE class="RcpTable" Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<Tbody id="ChequeLines">
			<TR bgcolor='#F0F0F0' onclick="currentRow=this.rowIndex;" >
				<TD align='center' width="25px">1</td>
				<TD dir="LTR"><INPUT class="RcpRowInput" TYPE="text" NAME="ChequeNos" size="10" maxLength="6" onKeyPress="return mask(this);"	onblur="check(this);"></td>
				<TD dir="LTR" ><INPUT class="RcpRowInput" TYPE="text" NAME="ChequeDates" maxlength="10" size="10"  onKeyPress="return maskDate(this);" onblur="acceptDate(this);"></td>
				<TD dir="RTL">
					<select name="Banks" class=inputBut style="width:150;">
					<option value="-1">«‰ Œ«» ﬂ‰Ìœ</option>
					<option value="-1">---------------------------------</option>
					<% set RSV=Conn.Execute ("SELECT * FROM Bankers WHERE (IsBankAccount = 1)") 
					Do while not RSV.eof
					%>
						<option value="<%=RSV("id")%>"><%=RSV("Name")%> </option>
					<%
					RSV.moveNext
					Loop
					RSV.close
					%>
					</select>
				
				</TD>
				<TD dir="LTR"><INPUT class="RcpRowInput" TYPE="text" NAME="Amounts" size="15" onKeyPress="return mask(this);" onBlur="setPrice(this);" ></td>
				<TD dir="RTL"><INPUT class="RcpRowInput" TYPE="text" NAME="Description" size="18" ></td>
			</TR>
			<TR bgcolor='#F0F0F0' onclick="currentRow=this.rowIndex;" >
				<TD colspan="15">
					<INPUT class="RcpGenInput" TYPE="button" value="«÷«›Â" onkeyDown="if(event.keyCode==9) return false;" onClick="addRow(this.parentNode.parentNode.rowIndex);">
				</TD>
			</TR>
			</Tbody></TABLE></div></td>
		</tr>
		<tr class="RcpMainTableTR">
			<td colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL">
			<TR>
				<TD  class="RcpHeadInput" align='center' > 
				<% if CDbl(VoucherTotalPrice) <> -1 then %>Ã„⁄ „«‰œÂ ›«ﬂ Ê—Â«:
				<%=Separate(SumAPRemainedAmount)%>
				<% end if %></td>
				<TD class="RcpHeadInput" align='center' width="25px"> </td>
				<TD class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="Ã„⁄:" size="10" tabindex="9999"></td>
				<TD class="RcpHeadInput"><INPUT class="RcpHeadInput3" readonly dir="LTR" TYPE="text" Name="TotalAmount"  size="15" tabindex="9999"></td>
			</TR>
			</TABLE></div></td>
		</tr>
		</table>
		<br>
		<TABLE class="RcpTable" Border="0" Cellspacing="5" Cellpadding="1" Dir="RTL">
		<TR>
			<TD align='center' bgcolor="#000000"><INPUT class="RcpGenInput" style="text-align:center" TYPE="submit" value="–ŒÌ—Â"></td>
			<TD align='center' bgcolor="#000000"><INPUT class="RcpGenInput" style="text-align:center" TYPE="button" value="«‰’—«›" onclick="window.location='?';"></td>
		</TR>
		</TABLE>
		</FORM>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			document.all.ChequeNos.focus();
		//-->
		</SCRIPT>
	<SCRIPT language="JavaScript">
	<!--
	//===================================================
	function delRow(rowNo){
		rowsCount=document.getElementsByName("ChequeNos").length
		if (rowsCount==1){
			alert("Õ–› «Ì‰ Œÿ «„ﬂ«‰ Å–Ì— ‰Ì” ");
			return false;
		}
		
		chqTable=document.getElementById("ChequeLines");
		theRow=chqTable.getElementsByTagName("tr")[rowNo];
		chqTable.removeChild(theRow);

		rowsCount--;
		for (rowNo=0; rowNo < rowsCount; rowNo++){
			chqTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0].innerText= rowNo+1;
		}
	}

	//===================================================
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
		tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="ChequeNos" size="10" maxLength="6" onKeyPress="return mask(this);" onblur="check(this);">'
	//	tempTD.innerHTML=chqTable.getElementsByTagName("tr")[0].getElementsByTagName("td")[1].innerHTML
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("dir", 'LTR');
		tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="ChequeDates" maxlength="10" size="10"  onKeyPress="return maskDate(this);" onblur="acceptDate(this);">'
	//	tempTD.innerHTML=chqTable.getElementsByTagName("tr")[0].getElementsByTagName("td")[2].innerHTML
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.innerHTML=chqTable.getElementsByTagName("tr")[0].getElementsByTagName("td")[3].innerHTML
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("dir", 'LTR');
		tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="Amounts" size="15" onKeyPress="return mask(this);" onBlur="setPrice(this);">'
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("dir", 'RTL');
		tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="Description" size="18" >'
		newRow.appendChild(tempTD);
					

		chqTable.insertBefore(newRow,theRow);

		chqTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
	}

	//===================================================
	function setPrice(src){
		myRow=src.parentNode.parentNode.rowIndex

		if (src.name=="Amounts" && document.getElementsByName("ChequeNos")[myRow].value==''){
			src.value=0;
		}
		else{
			src.value=val2txt(txt2val(src.value));
		}
		//cashAmount=parseInt(txt2val(document.getElementsByName("CashAmount")[0].value));
		//totalAmount = cashAmount;
		totalAmount = 0;

		for (rowNo=0; rowNo < document.getElementsByName("Amounts").length; rowNo++){
			totalAmount += parseInt(txt2val(document.getElementsByName("Amounts")[rowNo].value));
		}
		document.all.TotalAmount.value = val2txt(totalAmount);
	}

	var dialogActive=false;

	//===================================================
	function check(src){ 
		if (src.name=='ChequeNos'){
			if (src.value=='0'){
			if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ Œÿ çﬂ —« Õ–› ﬂ‰Ìœø\n\n"))
				delRow(src.parentNode.parentNode.rowIndex);
			}
		}
	}
	//===================================================
	function mask(src){ 
		var theKey=event.keyCode;
		if (theKey==13){
			return true;
		}
		else if (theKey < 48 || theKey > 57) { // 0-9 are acceptable
			return false;
		}
	}
	//-->
	</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------- Default
'-----------------------------------------------------------------------------------------------------
elseif request("act")="" then %>
<!-- Ã” ÃÊ »—«Ì ‰«„ Õ”«» -->
	<BR><BR>
	<FORM METHOD=POST ACTION="cheq.asp?act=submitsearch" onsubmit="if (document.all.CustomerNameSearchBox.value=='') return false;">
	<div dir='rtl'>&nbsp;&nbsp;<B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«»</B>
		<INPUT TYPE="text" NAME="CustomerNameSearchBox">&nbsp;
		<INPUT TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>
	<BR>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.CustomerNameSearchBox.focus();
	//-->
	</SCRIPT>
<%
end if
conn.Close
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
%>
</font>
<!--#include file="tah.asp" -->
