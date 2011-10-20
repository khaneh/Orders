<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "À»  ÕÊ«·Â »«‰ﬂ"
SubmenuItem=4
if not Auth("A" , 4) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>
	.GenInput { font-family:tahoma; font-size: 9pt; border: 1px solid black; text-align:right; }
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>

<BR><BR>
<%
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Submit Transaction
'-----------------------------------------------------------------------------------------------------
if request("act")="submit" then

	ON ERROR RESUME NEXT
		errorFound=	false
		CustomerID=		clng(request.form("CustomerID"))
		Amount=			cdbl(text2value(request.form("Amount")))
		bank =			cint(request.form("Bank"))
		ReceiptNo =		clng(request.form("ReceiptNo"))
		Reason=			cint(request.form("Reason"))

		if Err.Number<>0 then
			Err.clear
			errorFound=	true
			errorMsg=	Server.URLEncode("Œÿ«!")
		end if

		if NOT errorFound then
			effectiveDate=	sqlSafe(request.form("ReceiptDate"))

			'---- Checking wether CustomerID is valid 
			mySQL="SELECT ID FROM Accounts WHERE (ID="& CustomerID & ")"
			Set RS1=Conn.execute(mySQL)
			if RS1.eof then
				conn.close
				response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br>‘„«—Â Õ”«» ["& CustomerID & "] ÊÃÊœ ‰œ«—œ.")
			end if
			RS1.close
			'----

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
				errorMsg=Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
			end if 
			'----
			'-------- Finding SYS and firstGLAccount for the selected Reason
			mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
			Set RS1=Conn.execute(mySQL)
			if RS1.eof then
				errorFound=	true
				errorMsg=	Server.URLEncode("Œÿ«!")
			else
				Sys=			RS1("Acron")
				firstGLAccount=	RS1("GLAccount")
			end if
			RS1.close

			'-------- Finding GLAccount for the selected bank 
			'---- CheqStatus = 2 means 'vosool shode'
			mySQL = "SELECT BankerCheqStatusGLAccountRelation.GLAccount, Bankers.Name FROM BankerCheqStatusGLAccountRelation INNER JOIN Bankers ON BankerCheqStatusGLAccountRelation.Banker = Bankers.ID WHERE (Bankers.ID="& bank & ") AND (Bankers.IsBankAccount = 1) AND (BankerCheqStatusGLAccountRelation.GL = '"& openGL & "') AND (BankerCheqStatusGLAccountRelation.CheqStatus = 2)"
			Set RS1 = Conn.Execute(mySQL)
			if RS1.eof then
				conn.close
				response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br>Õ”«» »«‰ﬂ ÊÃÊœ ‰œ«—œ.")
			else
				BankName=	RS1("Name")
				GLAccount=	RS1("GLAccount")
			end if
			RS1.close
			'--------

		end if
		if Err.Number<>0 then
			Err.clear
			errorFound=	true
			errorMsg=	Server.URLEncode("Œÿ«!")
		end if
	ON ERROR GOTO 0
	if errorFound then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!")
	end if

	creationDate=	shamsiToday()
	comment=		sqlSafe(request.form("comment"))

	if comment <> "" then comment = " (" & comment & ")"

	if instr(BankName,"(")>0 then
		BankName=left(BankName,instr(BankName,"(")-1)
	end if

	theDescription = "Ê«—Ì“ ÊÃÂ ‰ﬁœ ÿÌ —”Ìœ "& ReceiptNo & " "& BankName & "  Ê”ÿ ﬂœ "& customerID & comment

	' ######################################################### 
	'					INSERTING RECEIPT  ...
	' #########################################################

	'	CHQAMOUNT ...
	CHQAMOUNT =0
	'	CHQQTTY must be Receipt Qtty of Cheques
	CHQQTTY =0
	DepositAmount = 0 

	mySQL="INSERT INTO Receipts (CreatedDate, sys, CreatedBy, EffectiveDate, Number, Customer, CashAmount, DepositAmount, ChequeAmount, ChequeQtty, TotalAmount) VALUES (N'" &_
	creationDate & "', '"& sys & "' , '"& session("ID") & "', N'" & effectiveDate & "', '"& ReceiptNo & "', '"& CustomerID & "', '"& amount & "', '"& DepositAmount & "', '"& CHQAMOUNT & "', '"& CHQQTTY & "', '"& amount & "');SELECT @@Identity AS NewReceipt"

	set RS1 = Conn.execute(mySQL).NextRecordSet
	ReceiptID=RS1("NewReceipt")
	RS1.close

	' ######################################################### 
	'					INSERTING CASH  ...
	' #########################################################
	mySQL="INSERT INTO ReceivedCash (Receipt, Amount, Banker, Description) VALUES ('"&_
			ReceiptID & "', '" & amount & "', '" & bank & "', N'"& theDescription & "')"
	conn.Execute(mySQL)

	chequeCount=0
	totalChequeAmount=0

	'**************************** Creating Item for Receipt  ****************
	'*** Type = 2 means Item is a Receipt
	ItemType =  2  

	mySQL="INSERT INTO "& sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, Link, Reason, CreatedDate, CreatedBy, Type, IsCredit, AmountOriginal, RemainedAmount) VALUES ('"&_
	GLAccount & "', '"& OpenGL & "', "& firstGLAccount & ", "& customerID & ", N'"& effectiveDate & "', "& ReceiptID & ", "& Reason & ", N'"& creationDate & "' , "& session("id") & ", "& ItemType & ", 1,  "& Amount & ","& Amount & ")"
	conn.Execute(mySQL)
	'***------------------------- Creating Item for Receipt  ----------------

	Conn.Execute ("UPDATE Accounts SET "& sys & "Balance="& sys & "Balance+"& Amount & "  WHERE (ID = "& customerID & ")")	

	Conn.Execute ("UPDATE Bankers SET CurrentBalance=CurrentBalance+"& amount & " WHERE ID="& bank ) 

	response.redirect "../"& sys & "/AccountReport.asp?act=showReceipt&sys="& sys & "&receipt=" & ReceiptID &"&msg="& Server.URLEncode("ÕÊ«·Â À»  ‘œ.<br><a href='../bank/submitDraft.asp'>»«“ê‘  »Â œ—Ì«›  ÕÊ«·Â</a>")

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------- Input a new GL Memo Doc
'-----------------------------------------------------------------------------------------------------
elseif request("act")="" then

%>
<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>
<FORM METHOD=POST ACTION="?act=submit" onsubmit="return isEmpty()">
<TABLE align=center style="border:1pt solid white">
<TR>
	<TD colspan='2'>
		<TABLE width='100%' cellpadding='5'>
		<TR>
			<TD bgcolor=white>”Ì” „: </TD>
<%
			mySQL="SELECT * FROM AXItemReasons WHERE Display=1 ORDER BY ID"
			set RS1=conn.execute(mySQL)
			while not RS1.eof
%>				<TD bgcolor=white><INPUT TYPE="radio" NAME="Reason" value="<%=RS1("ID")%>" <% if reason=RS1("ID") then response.write "checked "%>><%=RS1("Name")%><br><span style='font-size:7pt;width:100%;text-align:left;'> <%="[" & RS1("GLAccount") & "]"%> </span></TD>
<%				RS1.movenext
			wend
%>		
		</TR>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD>
		 «—ÌŒ:  
	</TD>
	<TD>
		<INPUT class="GenInput" style='direction:LTR;text-align:left;' NAME="ReceiptDate" TYPE="text" maxlength="10" size="10" value="" onblur="acceptDate(this)">
	</TD>
</TR>
<TR>
	<TD>
		‘„«—Â ÕÊ«·Â :
	</TD>
	<TD>
		<INPUT class="GenInput" style='direction:LTR;text-align:left;' NAME="ReceiptNo" TYPE="text" size="10" > 
	</TD>
</TR>
<TR>
	<TD>
		„»·€ ÕÊ«·Â :
	</TD>
	<TD>
		<INPUT TYPE="text" NAME="amount" class='GenInput' style='width:170px;direction:LTR;text-align:left;' onKeyPress="return maskNumber(this);" onBlur="this.value=val2txt(txt2val(this.value))"> —Ì«· 
	</TD>
</TR>
<TR>
	<TD>
		ﬂœ Õ”«» Å—œ«Œ  ﬂ‰‰œÂ:
	</TD>
	<TD>
		<table width='100%' border=0>
		<tr>
			<td width=75><INPUT TYPE="text" NAME="customerID" class='GenInput' style='width:70px;direction:LTR;text-align:left;' maxlength="10" onKeyPress='return mask(this);' onBlur='check(this);'></td>  
			<td><INPUT TYPE="text" NAME="accountName" size=28 readonly  value="<%=accountName%>" style="background-color:transparent"></td>  
		</tr>  
		</table>
	</TD>
</TR>
<TR>
	<TD>
		‰«„ »«‰ﬂ :
	</TD>
	<TD>
		<select name="Bank" class=inputBut style="width:250; ">
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
</TR>
<TR>
	<TD>
		 Ê÷ÌÕ: 
	</TD>
	<TD>
		<textarea NAME='comment' cols='30' class='GenInput' style='width:250px;'></textarea>
	</TD>
</TR>
<TR>
	<TD colspan='2' height='2' ><hr></TD>
</TR>
<TR>
	<TD colspan='2' align='center'>
		<INPUT TYPE="submit" value="À»  ÕÊ«·Â" class='GenButton'><br><br>
	</TD>
</TR>
</TABLE>
</FORM>
<BR><BR>

<SCRIPT LANGUAGE="JavaScript">
<!--

var dialogActive=false;

function mask(src){ 
	var theKey=event.keyCode;

	if (theKey==32){
		event.keyCode=9
		dialogActive=true
		document.all.tmpDlgArg.value="#"
		document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» :"
		window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		if (document.all.tmpDlgTxt.value !="") {
			window.showModalDialog('../AR/dialog_SelectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value), document.all.tmpDlgArg, 'dialogHeight:400px; dialogWidth:700px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
			dialogActive=false
			if (document.all.tmpDlgArg.value!="#"){
				Arguments=document.all.tmpDlgArg.value.split("#")
				src.value=Arguments[0];
				document.all.accountName.value=Arguments[1];
			}
		}
	}
}

function check(src){ 
	if (!dialogActive){
		badCode = false;
		if (window.XMLHttpRequest) {
			var objHTTP=new XMLHttpRequest();
		} else if (window.ActiveXObject) {
			var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
		}
		objHTTP.open('GET','../accounting/xml_CustomerAccount.asp?id='+src.value,false)
		objHTTP.send()
		tmpStr = unescape(objHTTP.responseText)
		//document.all['A1'].innerText= objHTTP.status
		//document.all['A2'].innerText= objHTTP.statusText
		//document.all['A3'].innerText= objHTTP.responseText
		document.all.accountName.value=tmpStr;
		}
}



function isEmpty()
{
	ReceiptNo = document.all.ReceiptNo.value
	CusID = document.all.customerID.value
	amount = document.all.amount.value
	Bank = document.all.Bank.value
	if(CusID=="" || amount=="" || Bank==-1)
		{
		alert("Œÿ«! ›—„ ﬂ«„· Å— ‰‘œÂ «” ")
		return false
		}
	else
		{
		str = "À»  ŒÊ«·Âø"
		return confirm(str)
		}
	
}
//-->
</SCRIPT>
<% 
end if
%>
<!--#include file="tah.asp" -->
