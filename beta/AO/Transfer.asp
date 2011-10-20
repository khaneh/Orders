<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'A0 (11 [=B])
PageTitle= "«‰ ﬁ«· «“ Õ”«» »Â Õ”«»"
SubmenuItem=3
if not Auth("B" , 3) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
'---------------------------------------------
'---------------------------- ShowErrorMessage
'---------------------------------------------
function ShowErrorMessage(msg)
	response.write "<table align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> Œÿ« ! <br>"& msg & "<br></td></tr></table><br>"
end function


'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------- Submit Transfer
'-----------------------------------------------------------------------------------------------------
if request("act")="submit" then

	ON ERROR RESUME NEXT

		customerIDfrom =	clng(request.form("customerIDfrom"))
		customerIDto =		clng(request.form("customerIDto"))
		customerNameFrom =	sqlSafe(request.form("customerNameFrom"))
		customerNameTo=		sqlSafe(request.form("customerNameTo"))

		GLAccFrom=			clng(request.form("GLAccFrom"))
		GLAccTo=			clng(request.form("GLAccTo"))
		GLAccNameFrom =		sqlSafe(request.form("GLAccNameFrom"))
		GLAccNameTo=		sqlSafe(request.form("GLAccNameTo"))

		amount =			cdbl(text2value(request.form("amount")))
		comment =			sqlSafe(request.form("comment"))
		memoDate=			sqlSafe(request.form("MemoDate"))

		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
		end if
	ON ERROR GOTO 0

	'---- Checking wether EffectiveDate is valid in current open GL
	if (memoDate < session("OpenGLStartDate")) OR (memoDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----


  '--- From ---
	mySQL = "SELECT ID FROM GLAccounts WHERE (HasAppendix = 1) AND (GL = "& openGL & ") AND (ID = "& GLAccFrom & ")"
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br>„⁄Ì‰ ÿ—› »œÂﬂ«— «‘ »«Â «” .")
	end if
	RS1.close

	mySQL = "SELECT ItemReason FROM AXItemReasonGLAccountRelations WHERE (GL = "& openGL & ") AND (GLAccount = "& GLAccFrom & ")"
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		'Using default reason (sys: AO, reason: Misc.)
		ReasonFrom=6
	else
		ReasonFrom=	cint(RS1("ItemReason"))
	end if
	RS1.close

	mySQL="SELECT * FROM AXItemReasons WHERE (ID="& ReasonFrom & ")"
	Set RS1=Conn.execute(mySQL)
	sysFrom=			RS1("Acron")
	ReasonNameFrom =	RS1("Name")
	RS1.close

  '---  To  ---
	mySQL = "SELECT ID FROM GLAccounts WHERE (GL = "& openGL & ") AND (HasAppendix = 1) AND (ID = "& GLAccTo & ")"
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br>„⁄Ì‰ ÿ—› »” «‰ﬂ«— «‘ »«Â «” .")
	end if
	RS1.close

	mySQL = "SELECT ItemReason FROM AXItemReasonGLAccountRelations WHERE (GL = "& openGL & ") AND (GLAccount = "& GLAccTo & ")"
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		'Using default reason (sys: AO, reason: Misc.)
		ReasonTo=6
	else
		ReasonTo=	cint(RS1("ItemReason"))
	end if
	RS1.close

	mySQL="SELECT * FROM AXItemReasons WHERE (ID="& ReasonTo & ")"
	Set RS1=Conn.execute(mySQL)
	sysTo=				RS1("Acron")
	ReasonNameTo =		RS1("Name")
	RS1.close
  '------------

	if customerIDto = customerIDfrom and ReasonFrom = ReasonTo then 
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br>Â„ ÿ—›Ì‰ Õ”«» Ê Â„ ”Ì” „ „»œ« Ê „ﬁ’œ ÌﬂÌ Â” ‰œ.")
	end if

	descriptionFrom =	"«‰ ﬁ«· „«‰œÂ Õ”«» »Â " & GLAccTo & 	" ["& ReasonNameTo & "] "
	descriptionTo =		"«‰ ﬁ«· „«‰œÂ Õ”«» «“ " & GLAccFrom &	" ["& ReasonNameFrom & "] "

	if customerIDto <> customerIDfrom then 
		descriptionFrom =	descriptionFrom & "Õ”«» " & customerIDTo &	" ["& customerNameTo & "] "
		descriptionTo =		descriptionTo & "Õ”«» " & customerIDFrom & " ["& customerNameFrom & "] "
	end if

	if comment<>"" then
		descriptionFrom =	descriptionFrom & "(" & comment & ")"
		descriptionTo =		descriptionTo  & "(" & comment & ")"
	end if
	
	creationDate=shamsiToday()
	CreationTime = currentTime10()

'*******************'*******************
'***--------------------  Do Transfer
'*******************'*******************
'** MemoType = 7 means Transfer 
	MemoType = 7		

	'-----------------------------------------------------------------------
	'---------------------------------------- From
	'------------------------------------- ( Debit Memo )
	mySQL="INSERT INTO "& sysFrom & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& creationDate & "' , "& session("id") & ", "& MemoType & ", "& customerIDfrom & ", 0 , "& Amount & ", N'"& descriptionFrom & "');SELECT @@Identity AS NewMemo"
	Set RS1 = conn.Execute(mySQL).NextRecordSet
	MemoID=RS1("NewMemo")
	RS1.close

	FromItemLink = MemoID

	'**************************** Creating Item for Debit Memo ****************
	'*** ItemType = 3 means Item is a Memo
	ItemType = 3 

	mySQL="INSERT INTO "& sysFrom & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, CreatedDate, CreationTime, CreatedBy, Type, Link, Reason, IsCredit, AmountOriginal, RemainedAmount) VALUES ("&_
		"NULL, "& OpenGL & ", "& GLAccFrom & ", "& customerIDfrom & ", N'"& memoDate & "', N'"& creationDate & "' , N'"& CreationTime & "', "& session("id") & ", "& ItemType & ", "& MemoID & ", "& ReasonFrom & ", 0,  "& Amount & ","& Amount & ")"
	conn.Execute(mySQL)

	Conn.Execute ("UPDATE Accounts SET "& sysFrom & "Balance="& sysFrom & "Balance-"& Amount & "  WHERE (ID = "& customerIDfrom & ")")	
	'***------------------------- Creating Item for Debit Memo ----------------

	'-----------------------------------------------------------------------
	'---------------------------------------- To
	'------------------------------------- ( Credit Memo )
	mySQL="INSERT INTO "& sysTo & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& creationDate & "' , "& session("id") & ", "& MemoType & ", "& customerIDto & ", 1 , "& Amount & ", N'"& descriptionTo & "');SELECT @@Identity AS NewMemo"
	Set RS1 = conn.Execute(mySQL).NextRecordSet
	MemoID=RS1("NewMemo")
	RS1.close

	ToItemLink = MemoID 

	'**************************** Creating Item for Credit Memo ****************
	'*** ItemType = 3 means Item is a Memo
	ItemType =  3  ' 3=SANAD
	mySQL="INSERT INTO "& sysTo & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, CreatedDate, CreationTime, CreatedBy, Type, Link, Reason, IsCredit, AmountOriginal, RemainedAmount) VALUES ("&_
		"NULL, "& OpenGL & ", "& GLAccTo  & ", "& customerIDto & ", N'"& memoDate & "', N'"& creationDate & "' , N'"& CreationTime & "', "& session("id") & ", "& ItemType & ", "& MemoID & ", "& ReasonTo & ", 1,  "& Amount & ","& Amount & ")"
	conn.Execute(mySQL)

	Conn.Execute ("UPDATE Accounts SET "& sysTo & "Balance="& sysTo & "Balance+"& Amount & "  WHERE (ID = "& customerIDto & ")")	
	'***------------------------- Creating Item for Credit Memo ----------------


'---------------------------------------------
'------------------------------------ Relation
'---------------------------------------------
	mySQL="INSERT INTO InterMemoRelation (FromItemType, FromItemLink, ToItemType, ToItemLink) VALUES ('"& sysFrom & "', "& FromItemLink & ", '"& sysTo & "', "& ToItemLink & ")"
	set RS1=conn.execute(mySQL)
'***-----------------------------------
'***-------------------- End Of Transfer
'***-----------------------------------

response.write "<br>" 
call showAlert ("œÊ «⁄·«„ÌÂ »« „Ê›ﬁÌ  «ÌÃ«œ ‘œ.",CONST_MSG_INFORM) 
response.write "<br>" 
response.write "<p align=center>"
response.write "<A Href='../"& sysFrom& "/AccountReport.asp?act=showMemo&sys="& sysFrom & "&memo="& FromItemLink & "' Target='_blank'><B>«⁄·«„ÌÂ »œÂﬂ«—</B></A> »—«Ì <A Href='../CRM/AccountInfo.asp?act=show&selectedCustomer="& customerIDFrom & "' Target='_blank'><B>" & customerNameFrom & " ["& customerIDFrom & "]</B></A> »Â œ·Ì· <b>'" & descriptionFrom &"'</b> "
response.write "<br><br>Ê<br><br>" 
response.write "<A Href='../"& sysTo& "/AccountReport.asp?act=showMemo&sys="& sysTo & "&memo="& ToItemLink & "' Target='_blank'><B>«⁄·«„ÌÂ »” «‰ﬂ«—</B></A> »—«Ì <A Href='../CRM/AccountInfo.asp?act=show&selectedCustomer="& customerIDTo & "' Target='_blank'><B>" & customerNameTo & " ["& customerIDTo & "]</B></A> »Â œ·Ì· <b>'" & descriptionTo &"'</b> "

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------- Transfer Form
'-----------------------------------------------------------------------------------------------------
elseif request("act")="" then
%>
	<style>
		Table { font-size: 9pt;}
		.InvRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
		.InvHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
		.InvRowInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
		.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
		.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right; direction: right-to-left;}
		.InvGenInput  { font-family:tahoma; font-size: 9pt; border: none; }
		.InvGenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
		.GLInput1 { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction:RTL;}
		.GLInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction:LTR; text-align:right; }
		.GLTable1 { border: none; direction:RTL; border:1px dashed red ;} 
		.GLTable1 tr {height:20; background-color: #CCCC88; }
		.GLTable2 { border: none; direction:RTL;} 
		.GLTable2 tr {height:20; text-align:center; background-color: #C3C300; }
		.GLTable2 td {border-bottom: 1px solid black; border-right: 1px solid black;}
		.GLTable3 tr {background-color: #F0F0F0}
		.GLTR1 { font-family:tahoma; font-size: 9pt; height:30; text-align:center; vertical-align:top; background-color: #C3C300; }
	'	.GLTR2 { height:20; text-align:center; background-color: #C3C300; }
		.GLTD1 { font-family:tahoma; font-size: 9pt; height:20; text-align:center; }
	</style>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	var okToProceed=false;
	var currentRow=0;
	var IsTaraz = false
	//-->
	</SCRIPT>

	<!-- Ê—Êœ «ÿ·«⁄«  ”‰œ -->
	<br>
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<FORM METHOD=POST ACTION="Transfer.asp?act=submit" onsubmit="return checkValidation();">
	<table class="GLTable1" align="center" Cellspacing="0" Cellpadding="0">
		<tr class="GLTR2">
			<TD>
			<TABLE class="GLTable2" Cellspacing="0" Cellpadding="0" >
			<tr>
				<td colspan=5 align=left>  «—ÌŒ: &nbsp;</td>
				<td class='InvRowInput2'><INPUT TYPE="text" class="InvRowInput2" style="width:80;text-align:Left;direction:LTR;" NAME="MemoDate" maxlength="10" onblur="acceptDate(this);"></td>
				<td> &nbsp; </td>
			</tr>
			<tr>
				<td style="width:26; border-right:none;"> # </td>
				<td style="width:50;">Õ”«»</td>
				<td style="width:100;">⁄‰Ê«‰</td>
				<td style="width:50;">„⁄Ì‰</td>
				<td style="width:100;">&nbsp;</td>
				<td style="width:80;">»œÂﬂ«—</td>
				<td style="width:80;">»” «‰ﬂ«—</td>
			</tr>
			</TABLE></TD>
		</tr>
		<tr>
			<TD><div style=" height:50px; width:*;">
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="GLTable3">
			<Tbody id="GLrows">
			<tr bgColor=#f0f0f0 onclick='document.getElementById("GLrows").getElementsByTagName("tr")[1].getAttribute("onclick")'>
				<td width='25' align=center>«“
				</td>
				<td dir='LTR'>
					<INPUT class='InvRowInput' TYPE='text' NAME='customerIDfrom' maxLength='6' onKeyPress='return mask(this);' style='width:50;' onBlur='check(this);'>
				</td>
				<td ><!-- FROM -->
					<INPUT class='InvRowInput2' TYPE='text' NAME='customerNameFrom'  readonly  tabIndex=9991 style='width:100;' onfocus='this.parentNode.parentNode.getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus()'>
				</td>
				<td dir='LTR'>
					<INPUT class='InvRowInput' TYPE='text' NAME='GLAccFrom' maxLength='5' onKeyPress='return mask(this);' style='width:50;' onBlur='check(this);'>
				</td>
				<td dir='RTL'>
					<INPUT class='InvRowInput2' TYPE='text' NAME='GLAccNameFrom'  readonly  tabIndex=9992 style='width:100;' onfocus='this.parentNode.parentNode.getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus()'>
				</td>
				<td dir='LTR'>
					<INPUT class='InvRowInput2' TYPE='text' NAME='amount2' id='amount2' size='2'  style='width:80;' onblur='setPrice(this)'  onKeyPress='return onlyNumber(this);'  onfocus='this.value=txt2val(this.value);this.select()'>
				</td>
				<td dir='LTR'>
					<INPUT class='InvRowInput2' TYPE='text' NAME='credits' size='2'  style='width:80;'  value="--------------------" disabled> 
				</td>
			</tr>
			<tr bgColor=#f0f0f0>
				<td width='25' align=center>»Â
				</td>
				<td dir='LTR'>
					<INPUT class='InvRowInput' TYPE='text' NAME='customerIDto' maxLength='6' onKeyPress='return mask(this);'  style='width:50;' onBlur='check(this);' >
				</td>
				<td ><!-- TO -->
					<INPUT class='InvRowInput2' TYPE='text' NAME='customerNameTo' readonly  tabIndex=9993 style='width:100;' onfocus='this.parentNode.parentNode.getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus()'>
				</td>
				<td dir='LTR'>
					<INPUT class='InvRowInput' TYPE='text' NAME='GLAccTo' maxLength='5' onKeyPress='return mask(this);' style='width:50;' onBlur='check(this);' >
				</td>
				<td dir='RTL'>
					<INPUT class='InvRowInput2' TYPE='text' NAME='GLAccNameTo' readonly  tabIndex=9994 style='width:100;' onfocus='this.parentNode.parentNode.getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus()'>
				</td>
				<td dir='LTR'>
					<INPUT class='InvRowInput2' TYPE='text' size='2'  style='width:80;' disabled value="--------------------">
				</td>
				<td dir='LTR'>
					<INPUT class='InvRowInput2' TYPE='text' NAME='amount' id='amount' size='2'  style='width:80;'  onblur='setPrice(this)' onKeyPress='return onlyNumber(this);'   onfocus='this.value=txt2val(this.value);this.select()'>
				</td>
			</tr>
			</Tbody>
			</TABLE></div>
			<BR>
			<CENTER>
			 Ê÷ÌÕ: <textarea class='InvRowInput2' style="width:300" NAME="comment"></textarea>
			<BR><BR>
			<INPUT class="InvGenButton" TYPE="submit" value="–ŒÌ—Â ">
			<br><br>
			</CENTER>
		</tr>
	</FORM>
	</table>

	<script language="JavaScript">
	<!--

	function setPrice(src){
		src.value=val2txt(txt2val(src.value));
		rowNo=src.parentNode.parentNode.rowIndex;
		if (src.name=="amount") 
			document.all.amount2.value = src.value
		if (src.name=="amount2") 
			document.all.amount.value = src.value
		if (src.value ==0) src.value=""
		if (src.value ==" ") src.value=""

	}

	var dialogActive=false;

	function selectCustomer(){
		var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		theSearch=document.all.tmpDlgTxt.value

		if (theSearch!='' && theSearch!=null ){
			theSpan=document.getElementById("customer");
			window.open ('../AR/Customers.asp?act=select&search='+theSearch,'','width=700, height=400, scrollbars=yes, resizable=yes');
		}
	}

	function mask(src){ 
		var theKey=event.keyCode;
		rowNo=src.parentNode.parentNode.rowIndex;
		invTable=document.getElementById("GLrows");
		theRow=invTable.getElementsByTagName("tr")[rowNo];

		if (src.name=="customerIDfrom" || src.name=="customerIDto"){
			if(theKey==32){
				event.keyCode=9;
				dialogActive=true;
				document.all.tmpDlgArg.value="#";
				document.all.tmpDlgTxt.value="‰«„ Õ”«»Ì —« ﬂÂ „Ì ŒÊ«ÂÌœ Ã” ÃÊ ﬂ‰Ìœ Ê«—œ ﬂ‰Ìœ:";
				window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				if (document.all.tmpDlgTxt.value !="") {

					window.showModalDialog('../AR/dialog_SelectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value), document.all.tmpDlgArg, 'dialogHeight:400px; dialogWidth:700px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#");
						src.value=Arguments[0];
						invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[2].getElementsByTagName("Input")[0].value=Arguments[1];
					}
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[3].getElementsByTagName("input")[0].focus();
				}
				dialogActive=false
			}
			else if (theKey >= 48 && theKey <= 57 )
				return true;
			else
				return false;
		}
		else if(src.name=="GLAccFrom" || src.name=="GLAccTo"){
			if(theKey==32){
				event.keyCode=0
				dialogActive=true
				document.all.tmpDlgArg.value="#"
				document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì œ› — ﬂ·:"
				var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				if (document.all.tmpDlgTxt.value !="") {
					var myTinyWindow = window.showModalDialog('../accounting/dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#")
						src.value=Arguments[0];
						invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[4].getElementsByTagName("Input")[0].value=Arguments[1];
					}
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[5].getElementsByTagName("Input")[0].focus();
				}
				dialogActive=false
			}
			else if (theKey >= 48 && theKey <= 57 )
				return true;
			else 
				return false;
		}
	}

	function onlyNumber(src){ 
		var theKey=event.keyCode;
		if (theKey==13){  // [Enter] 
			return true;
		}
		else if (theKey==45 && (src.value=='' || src.value=='0') ){ // [-] 
			src.value='';
			return true;
		}
		else if (theKey < 48 || theKey > 57) { // 0-9 are acceptible
			return false;
		}
	}

	function check(src){ 
		if (!dialogActive){
			if (src.name=="customerIDfrom" || src.name=="customerIDto"){
				badCode = false;
				if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
				objHTTP.open('GET','../accounting/xml_CustomerAccount.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape(objHTTP.responseText)
				if (src.name=="customerIDfrom"){
					document.all.customerNameFrom.value=tmpStr;
				}
				else{
					document.all.customerNameTo.value=tmpStr;
					if (document.all.customerIDto.value == ""){
						document.all.customerNameTo.value=document.all.customerNameFrom.value;
						document.all.customerIDto.value=document.all.customerIDfrom.value;
					}
				}
			}
			else if(src.name=="GLAccFrom" || src.name=="GLAccTo"){
				badCode = false;
				if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
				objHTTP.open('GET','../accounting/xml_GLAccount.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape( objHTTP.responseText)
				if (src.name=="GLAccFrom"){
					document.all.GLAccNameFrom.value=tmpStr;
				}
				else{
					document.all.GLAccNameTo.value=tmpStr;
				}
			}
		}
	}

	function checkValidation()
	{
	  try{
		document.all.customerIDto.focus();
		foundErr=false;
		if(document.all.customerNameFrom.value=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ"|| document.all.customerNameFrom.value==""){
			foundErr=true;
			errObj=document.all.customerIDfrom;
			errMsg="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ";
		}
		else if(document.all.GLAccNameFrom.value=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ" || document.all.GLAccNameFrom.value==""){
			foundErr=true;
			errObj=document.all.GLAccFrom;
			errMsg="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ";
		}
		else if(document.all.customerNameTo.value=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ" || document.all.customerNameTo.value==""){
			foundErr=true;
			errObj=document.all.customerIDto;
			errMsg="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ";
		}
		else if(document.all.GLAccNameTo.value=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ" || document.all.GLAccNameTo.value==""){
			foundErr=true;
			errObj=document.all.GLAccTo;
			errMsg="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ";
		}

		if(foundErr) {
			tmpCol=errObj.style.backgroundColor;
			errObj.style.backgroundColor="red";
			errObj.focus();
			alert("!Œÿ« \n\n"+errMsg);
			errObj.style.backgroundColor=tmpCol;
			return false;
		}
		memoDate=document.all.MemoDate;
		if (memoDate.value!='' && memoDate.value!=null){
			if (!acceptDate(memoDate))
				return false;
		}
		else{
			memoDate.focus();
			return false;
		}
		return true;
	  }
	  catch(e){
			alert(e.message)
			return false;
	  }
	}
	//-->

	</script>
<%
'-------------===================================---------------------
end if
%>

<!--#include file="tah.asp" -->
