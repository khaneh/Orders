<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle= "Ê—Êœ ”‰œ"
SubmenuItem=1
if not Auth(8 , 1) then NotAllowdToViewThisPage()

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

%>
<style>
	Table { font-size: 9pt;}
	.InvRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; direction:LTR; width:100%;}
	.InvHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.InvRowInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right; width:100%;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right; direction: right-to-left;}
	.InvGenInput  { font-family:tahoma; font-size: 9pt; border: none; }
	.InvGenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
	.GLInput1 { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction:RTL;}
	.GLInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction:LTR; text-align:right; }
	.GLTable1 { border: none; direction:RTL; border:1px dashed red ; width:685;} 
	.GLTable1 tr {height:20; background-color: #CCCC88; }
	.GLTable2 { border: none; direction:RTL;} 
	.GLTable2 tr {height:20; text-align:center; background-color: #F0F0F0; }
	.GLTable2 td {border-bottom: 1px solid black; border-right: 1px solid black;}
	.GLTable3 tr {background-color: #F0F0F0}
	.GLTR1 { font-family:tahoma; font-size: 9pt; height:30; text-align:center; vertical-align:top; background-color: #C3C300; }
	.GLTR2 { height:20; text-align:center; background-color: #C3C300; }
	.GLTD1 { font-family:tahoma; font-size: 9pt; height:20; text-align:center; }
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;
var currentRow=0;
var IsTaraz = false
//-->
</SCRIPT>
<%


'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------- Input a new GL Memo Doc
'-----------------------------------------------------------------------------------------------------
'xml_GLAccount.asp
if request("act")="" then
if session("IsClosed")="True" then 
	response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
end if
%>
<!-- Ê—Êœ «ÿ·«⁄«  ”‰œ -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<BR>
		<table class="GLTable1" align="center" Cellspacing="0" Cellpadding="0">
			<tr class="GLTR1" align="center" Cellspacing="1" Cellpadding="0">
			<TD colspan="10"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="2" Dir="RTL">
				<TR>
				<TD align="left">œ› — :</TD>
				<TD>	<%=OpenGLName%>
				</TD>
				<FORM METHOD=POST ACTION="GLMemoInput.asp?act=submitMemo" onsubmit="return checkValidation()">
				<TD align="left"> «—ÌŒ ”‰œ :</TD>
				<TD>	
					<INPUT class="GLInput2" NAME="GLMemoDate" TYPE="text" maxlength="10" size="10" value="<%=shamsiToday()%>" onblur="acceptDate(this)" >
				</TD>
				<TD align="left">‘„«—Â ”‰œ :</TD>
				<TD>	
<%				mySQL="SELECT ISNULL(MAX(GLDocID),100) AS LastMemo FROM GLDocs WHERE GL='"& OpenGL & "'"
				Set RS1=conn.Execute (mySQL)
				LastMemo = RS1("LastMemo")
%>
					<INPUT class="GLInput2" NAME="GLMemoNo" TYPE="text" maxlength="10" size="10" value="<%=LastMemo+1%>">
				</TD>
				</TR></TABLE></TD>
			</tr>
			<tr class="GLTR2">
				<TD colspan="10" align=right><div>
				<TABLE class="GLTable2" Cellspacing="0" Cellpadding="0" >
				<tr>
					<td style="width:26; border-right:none;"> # </td>
					<td style="width:50;">Õ”«»</td>
					<td style="width:170;">‰«„ Õ”«»</td>
					<td style="width:300;">‘—Õ</td>
					<td style="width:80;">»œÂﬂ«—</td>
					<td style="width:80;">»” «‰ﬂ«—</td>
				</tr>
				</TABLE></div></TD>
			</tr>
			<tr>
			<td colspan="10">
			<div style="overflow:auto; height:250px; width:*;">
				<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="GLTable3">
				<Tbody id="GLrows">
				<tr bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
					<td colspan="6">
						<INPUT class="InvGenButton" TYPE="button" value="«÷«›Â" onkeyDown="if(event.keyCode==9) {setCurrentRow(this.parentNode.parentNode.rowIndex); return false;};" onClick="addRow();">
					</td>
				</tr>
				</Tbody></TABLE>
			</div>
			</td>
			</tr>
			<tr bgcolor='#CCCC88'>
				<td align=left>
				<B><span id="tarazDiv"> </span></B>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				»œÂﬂ«—  : <input type=text name="totalDebit" style="border:none; background:none;" value=0>
				&nbsp;&nbsp;&nbsp;
				»” «‰ﬂ«— : <input type=text name="totalCredit" style="border:none; background:none;" value=0>
				</td>
			</tr>
			</table>
		<br> 
		<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
		<tr>
			<td align='center'><INPUT class="InvGenButton" TYPE="submit" name="submit" value="–ŒÌ—Â"></td>
			<td align='center'><INPUT class="InvGenButton" style="background-color: red;" TYPE="submit" name="submit" value="–ŒÌ—Â »Â ’Ê—  Ì«œœ«‘ " onclick="saveDraft()"> </td>
		</tr>
		</TABLE>
		</FORM>

<%
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Edit a GL Memo Doc
'-----------------------------------------------------------------------------------------------------
elseif request("act")="editDoc" then

	id=request("id")
	if id="" or not(isnumeric(id)) then
		response.write "<br>" 
		call showAlert ("Œÿ«! ‘„«—Â ”‰œ Ê«—œ ‰‘œÂ «” ", CONST_MSG_ERROR) 
		response.end
	end if
	id=clng(id)

	mySQL="SELECT * FROM GLDocs WHERE (deleted = 0) and (isRemoved=0) and (GLDocID = "& id & ")  AND (GL = "& OpenGL & ")"
	set RS1=conn.execute(mySQL)
	if RS1.eof then
		response.write "<br>" 
		call showAlert ("Œÿ«! ç‰Ì‰ ”‰œÌ ÊÃÊœ ‰œ«—œ", CONST_MSG_ERROR) 
		response.end
	end if 

	GLDoc = RS1("ID")
	GLDocDate = RS1("GLDocDate")
	GLDocID = RS1("GLDocID")
	DID = RS1("id")
	RS1.close

	mySQL="SELECT GLRows.*, Accounts.AccountTitle, GLAccounts.Name AS GLAccountTitle, glDocs.isCompound FROM GLRows INNER JOIN GLAccounts ON GLRows.GLAccount = GLAccounts.ID LEFT OUTER JOIN Accounts ON GLRows.Tafsil = Accounts.ID INNER JOIN glDocs on glRows.glDoc = glDocs.id WHERE (GLRows.deleted = 0) AND (GLRows.GLDoc = "& GLDoc & ") AND (GLAccounts.GL = "& OpenGL & ") ORDER BY GLRows.ID"
	set RS2=conn.execute(mySQL)
'response.write rs2("isCompound")
%>
<!-- Ê—Êœ «ÿ·«⁄«  ”‰œ -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
		<FORM METHOD=POST ACTION="GLMemoInput.asp?act=submitEditMemo" onsubmit="return checkValidation()">
		<table class="GLTable1" align="center" Cellspacing="0" Cellpadding="0">
			<tr class="GLTR1" align="center" Cellspacing="1" Cellpadding="0">
			<TD colspan="10"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="2" Dir="RTL">
				<TR>
				<TD align="left" >œ› — :</TD>
				<TD> <%=OpenGLName%>
				</TD>
				<TD align="left"> «—ÌŒ ”‰œ :</TD>
				<TD>	
					<INPUT class="GLInput2" style="text-align:left;" NAME="GLMemoDate" TYPE="text" maxlength="10" size="10" value="<%=GLDocDate%>" onblur="acceptDate(this)">
				</TD>
				<TD align="left">‘„«—Â ”‰œ :</TD>
				<TD>	
					<INPUT class="GLInput2" NAME="GLMemoNo" TYPE="text" maxlength="10" size="10" value="<%=id%>" readonly>
				</TD>
				</TR></TABLE></TD>
			</tr>
			<tr class="GLTR2">
				<TD colspan="10"><div>
				<TABLE class="GLTable2" Cellspacing="0" Cellpadding="0" >
				<tbody id="GLcols">
				<tr>
					<td style="width:26; border-right:none;"> # </td>
					<td style="width:60;"> ›’Ì·Ì</td>
					<td style="width:40;">„⁄Ì‰</td>
					<td style="width:250;">‘—Õ</td>
					<td style="width:85;">»œÂﬂ«—</td>
					<td style="width:85;">»” «‰ﬂ«—</td>
					<td style="width:50;">‘„«—Â çﬂ</td>
					<td style="width:70;"> «—ÌŒ çﬂ</td>
				</tr>
				</tbody>
				</TABLE></div></TD>
			</tr>
			<tr>
			<TD colspan="10"><div style="overflow:auto; height:250px; width:695;">
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="GLTable3" onclick='setCurrentCell()'>
			<Tbody id="GLrows">
<%
			Do while not RS2.eof
				i = i + 1
				GLAccount =		RS2("GLAccount")
				GLAccTitle =	RS2("GLAccountTitle")
				Tafsil =		RS2("Tafsil")
				AccTitle =		RS2("AccountTitle")
				Description =	RS2("Description")
				Ref1 =			RS2("Ref1")
				Ref2 =			RS2("Ref2")
				Amount =		RS2("Amount")
				IsCredit =		RS2("IsCredit")

				credit = ""
				debit = ""
				if IsCredit  then 
					credit = Amount 
					totalCredit = totalCredit + cdbl(Amount)
				else
					debit = Amount 
					totalDebit = totalDebit + cdbl(Amount)
				end if
%>
			<tr bgcolor='#F0F0F0'>
				<td style="width:25; border-right:none;" align=center> <%=i%> </td>
				<td width="60">
					<INPUT  class='InvRowInput' <%if not rs2("isCompound") then response.write "readonly"%> TYPE='text' NAME='Accounts' maxlength=6 value="<%=Tafsil%>" title="<%=AccTitle%>" onKeyPress='return mask(this);' onBlur='check(this);'>
				</td>
				<td width="40">
					<INPUT class='InvRowInput2' TYPE='text' NAME='GLAccounts' maxlength=5 value="<%=GLAccount%>" title="<%=GLAccTitle%>" onKeyPress='return mask(this);' onBlur='check(this);'>
				</td>
				<td width="250">
					<INPUT class='InvRowInput2' TYPE='text' NAME='Descriptions' value="<%=Description%>" onKeyPress='return mask(this);'>
				</td>
				<td width="85">
					<INPUT class='InvRowInput2' TYPE='text' NAME='debits' value="<%=Separate(debit)%>" onKeyPress='return mask(this);' onBlur='check(this);'>
				</td>
				<td width="85">
					<INPUT class='InvRowInput2' TYPE='text' NAME='credits' value="<%=Separate(credit)%>" onKeyPress='return mask(this);' onBlur='check(this);'>
				</td>
				<td width="50">
					<INPUT class='InvRowInput2' TYPE='text' NAME='Refs1' maxlength=20 value="<%=Ref1%>" onKeyPress='return mask(this);'>
				</td>
				<td width="70">
					<INPUT  class='InvRowInput' TYPE='text' NAME='Refs2' maxlength=20 value="<%=Ref2%>" onKeyPress='return mask(this);' onBlur='acceptDate(this);'>
				</td>
			</tr>
				
<%
				RS2.movenext
			Loop
%>
			<tr bgcolor='#F0F0F0'>
				<td colspan="15">
					<INPUT class="InvGenButton" TYPE="button" value="«÷«›Â" onClick="addRow();">
				</td>
			</tr>
			</Tbody>
			</TABLE></div>
			</TD>
			</tr>
			<tr class="GLTR2">
				<TD colspan="10"><div>
				<TABLE class="GLTable2" Cellspacing="0" Cellpadding="0" style="border-top:1 solid red;">
				<tr>
					<td colspan=3 style="width:256;">
						<span id="tarazDiv" style="font-weight:bold;"><%if totalCredit=totalDebit then response.write "<FONT COLOR='#008833'>”‰œ  —«“ «” </FONT>" else response.write "<FONT COLOR='#FF3300'>”‰œ  —«“ ‰Ì” </FONT>" end if%></span>&nbsp;</td>
					<td style="width:122; text-align:left;">Ã„⁄ »œÂﬂ«—:</td>
					<td style="width:85;"><input type=text name="totalDebit" id="totalDebit" style="width:85;border:none; background:none;" value=<%=Separate(totalCredit)%>></td>
					<td style="width:85;"><input type=text name="totalCredit" id="totalCredit" style="width:85;border:none; background:none;" value=<%=Separate(totalDebit)%>></td>
					<td colspan=2 style="width:122; text-align:right;">:Ã„⁄ »” «‰ﬂ«—</td>
				</tr>
				</TABLE></div></TD>
			</tr>
		</table><br> 
		<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
		<tr>
			<td align='center'>
			<INPUT TYPE="hidden" name="GLDoc" value="<%=GLDoc%>">
			<INPUT class="InvGenButton" style="background-color: red;" TYPE="submit" name="submit" value="–ŒÌ—Â »Â ’Ê—  Ì«œœ«‘ " onclick="saveDraft()">
			<INPUT class="InvGenButton" TYPE="submit" value="–ŒÌ—Â "></td>
			<td align='center'><INPUT class="InvGenButton" TYPE="button" value="«‰’—«›" onclick="history.back()"></td>
		</tr>
		</TABLE>
		</FORM>
	<script language="JavaScript">
	<!--
	var dialogActive=false;

	function mask(src){ 
		var theKey=event.keyCode;
		if (src.name=="Accounts"){
			if (theKey==32){
				event.keyCode=0;
				dialogActive=true;
				document.all.tmpDlgArg.value="#"
				document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì  ›’Ì·Ì:"
				var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				dialogActive=false;
				if (document.all.tmpDlgTxt.value !="") {
					var myTinyWindow = window.showModalDialog('../ar/dialog_selectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:300px; dialogWidth:600px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#")
						src.value=Arguments[0];
						src.title=Arguments[1];
					}
				}
//				src.parentNode.nextSibling.getElementsByTagName("INPUT")[0].focus();
			}
			else if (theKey >= 48 && theKey <= 57 ) 
				return true;
			else
				return false;
		}
		else if (src.name=="GLAccounts"){
			if (theKey==32){
				event.keyCode=0;
				dialogActive=true;
				document.all.tmpDlgArg.value="#"
				document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì œ› — ﬂ·:"
				var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				dialogActive=false;
				if (document.all.tmpDlgTxt.value !="") {
					var myTinyWindow = window.showModalDialog('dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#")
						src.value=Arguments[0];
						src.title=Arguments[1];
					}
				}
//				src.parentNode.nextSibling.getElementsByTagName("INPUT")[0].focus();
			}
			else if (theKey >= 48 && theKey <= 57 ) // [0]-[9] are acceptible
				return true;
			else
				return false;
		}
		else if (src.name=="debits" || src.name=="credits" || src.name=="Refs1"){
			if (theKey < 48 || theKey > 57) { // [0]-[9] are acceptible
				return false;
			}
			return true;
		}
		else if (src.name=="Descriptions" || src.name=="Refs2"){
			if (theKey==13){  // [Enter] 
				return false;
			}
			return true;
		}
	}

	function check(src){ 
		if (src.name=="Accounts"){
			if (!dialogActive){
					if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
				objHTTP.open('GET','xml_CustomerAccount.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape( objHTTP.responseText)
				src.title=tmpStr;
				if (tmpStr=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ"){
					src.value="";
				}
			}
		}
		else if(src.name=="GLAccounts"){
			if (!dialogActive){
					if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
				objHTTP.open('GET','xml_GLAccount.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape( objHTTP.responseText)
				src.title=tmpStr;
				if (tmpStr=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ"){
					src.value="";
				}
			}
		}
		else if(src.name=="debits" || src.name=="credits"){
			src.value=val2txt(txt2val(src.value));
			if (''+src.value=="NaN" || src.value=="0") src.value = ""
			rowNo=src.parentNode.parentNode.rowIndex;

			if (src.value!=""){
				if (src.name=="credits") 
					document.getElementsByName("debits")[rowNo].value = "";
				else
					document.getElementsByName("credits")[rowNo].value = "";
			}

			memoIsTaraz();
			
		}
	}

	function checkValidation(){ 
		emptySlot=-1;
		for (rowNo=0; rowNo < lastRow; rowNo++){
			if(document.getElementsByName("GLAccounts")[rowNo].value==''){
				emptySlot=rowNo;
			}
		}
		if (emptySlot != -1){
			errObj=document.getElementsByName("GLAccounts")[emptySlot];
			errObj.style.backgroundColor="red";
			document.getElementsByName("GLAccounts")[emptySlot].focus();
			alert('«Ì‰ ”ÿ— »Ì „⁄‰Ì «” \n\n«ê— „Ì ŒÊ«ÂÌœ ¬‰ —« Õ–› ﬂ‰Ìœ\n[Ctrl]+[Del]\n»“‰Ìœ.');
			errObj.style.backgroundColor="";
			return false;
		}

		if (IsTaraz==true)
			return true
		else{
			alert("Œÿ«! ”‰œ  —«“ ‰Ì” ")
			return false
		}
	}

	function saveDraft(){ 
		if (document.all.totalDebit.value != 0 || document.all.totalCredit.value != 0)
				IsTaraz = true
	}

	function memoIsTaraz(){ 
		var totalCredit = 0;
		var totalDebit = 0;
		for (rowNo=0; rowNo < lastRow; rowNo++){
			totalCredit += parseInt(txt2val(document.getElementsByName("credits")[rowNo].value));
			totalDebit += parseInt(txt2val(document.getElementsByName("debits")[rowNo].value));
		}
		document.all.totalCredit.value = val2txt(totalCredit);
		document.all.totalDebit.value = val2txt(totalDebit);
		if (totalDebit == totalCredit && totalCredit != 0){
			IsTaraz = true
			document.all.tarazDiv.innerHTML = "<FONT COLOR='#008833'>”‰œ  —«“ «” </FONT>"
		}
		else{
			IsTaraz = false
			document.all.tarazDiv.innerHTML = "<FONT COLOR='#FF3300'>”‰œ  —«“ ‰Ì” </FONT>"
		}
	}

	var currentRow=0;
	var lastRow=<%=i%>;
	var currentCol=1;
	var lastCol=7;

	function documentKeyDown() {
		var theKey = event.keyCode;
		/*	
			Arrow Keys' Codes:
			<	37	(Left)
			^	38	(Up)
			>	39	(Right)
			v	40	(Down)
		*/
		if (event.ctrlKey && !event.shiftKey)
			Ctrl=true;
		else
			Ctrl=false;

		if (theKey==9){
			Ctrl=true;
			if (event.shiftKey){
				theKey = 39;
			}
			else
				theKey = 37;
		}
		else if ((theKey == 45) && Ctrl){			// [Ins] Key + [ctrl]
			event.keyCode=36; // [Home] Key
			addRow();
		}
		else if ((theKey == 46) && Ctrl){			// [Del] Key + [ctrl]
			event.keyCode=36; // [Home] Key
			delRow()
		}

		if (((theKey == 37 || theKey == 39) && Ctrl) || theKey == 38 || theKey == 40){
			
			deactivateCurrentCell();

			switch (theKey){
			case 37:
				if (currentCol < lastCol) 
					currentCol++;
				else{
					currentCol=1;
					if(currentRow<lastRow)
						currentRow++;
				}
				break;
			case 38:
				if (currentRow > 0) 
					currentRow--;
				else
					currentRow=lastRow;
				break;
			case 39:
				if (currentCol > 1) 
					currentCol--;
				else{
					currentCol=lastCol;
					if(currentRow>0)
						currentRow--;
				}
				break;
			case 40:
				if (currentRow < lastRow) 
					currentRow++;
				else
					currentRow=0;
				break;
			default:
				break;
			}

			activateCurrentCell();

			event.keyCode=36; // [Home] Key
		}
	}

	function deactivateCurrentCell(){
		Rtable=document.getElementById("GLrows");
		Ctable=document.getElementById("GLcols");
		
		Rtable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0].setAttribute("bgColor", '#F0F0F0');
		Ctable.getElementsByTagName("tr")[0].getElementsByTagName("td")[currentCol].setAttribute("bgColor", '#F0F0F0');

		tmpCell=Rtable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[currentCol];
		if (tmpCell){
			tmpCell.setAttribute("bgColor", '');

		}
	}

	function activateCurrentCell(){
		Rtable=document.getElementById("GLrows");
		Ctable=document.getElementById("GLcols");
		
		Rtable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0].setAttribute("bgColor", '#FFB0B0');
		Ctable.getElementsByTagName("tr")[0].getElementsByTagName("td")[currentCol].setAttribute("bgColor", '#FFB0B0');

		tmpCell=Rtable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[currentCol];
		if (tmpCell){
			tmpCell.setAttribute("bgColor", '#FF0000');
			tmpCell.getElementsByTagName("input")[0].focus();
		}
		if(currentRow==lastRow){
			document.getElementById("GLrows").getElementsByTagName("tr")[currentRow].getElementsByTagName("INPUT")[0].focus();
		}
	}

	function setCurrentCell(){
		tmpRow=event.srcElement.parentNode.parentNode.rowIndex;
		tmpCol=event.srcElement.parentNode.cellIndex;
		if(tmpCol>0 && tmpRow>=0){
			deactivateCurrentCell();

			currentRow =tmpRow;
			currentCol=tmpCol;

			activateCurrentCell();
		}
		else if(event.srcElement.parentNode.rowIndex){
			deactivateCurrentCell();

			currentRow =event.srcElement.parentNode.rowIndex;
			currentCol=1;

			activateCurrentCell();
		}
	}

	function delRow(){
		if (currentRow==lastRow)
			return;
		deactivateCurrentCell();
		rowNo = currentRow
		Rtable=document.getElementById("GLrows");
		theRow=Rtable.getElementsByTagName("tr")[rowNo];
		Rtable.removeChild(theRow);

		lastRow--;

		for (rowNo=0; rowNo < lastRow ; rowNo++){
			tempTD=Rtable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
			tempTD.bgColor= '#F0F0F0';
			tempTD.innerText= rowNo+1;
		}
		activateCurrentCell();
	}
	function addRow(){

		deactivateCurrentCell();
		rowNo = currentRow
		Rtable=document.getElementById("GLrows");
		theRow=Rtable.getElementsByTagName("tr")[rowNo];
		newRow=document.createElement("tr");
		newRow.setAttribute("bgColor", '#f0f0f0');

		tempTD=document.createElement("td");
		tempTD.innerHTML=rowNo+1
		tempTD.setAttribute("align", 'center');
		tempTD.setAttribute("width", '25');
		newRow.appendChild(tempTD);


		tempTD=document.createElement("td");
		tempTD.setAttribute("width", '60');
		tempTD.innerHTML="<INPUT class='InvRowInput'  TYPE='text' NAME='Accounts' maxlength=6 onKeyPress='return mask(this);' onBlur='check(this);'>"
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("width", '40');
		tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='GLAccounts' maxlength=5 onKeyPress='return mask(this);' onBlur='check(this);'>"
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("width", '250');
		tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Descriptions' onKeyPress='return mask(this);'>"
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("width", '80');
		tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='debits' onKeyPress='return mask(this);' onBlur='check(this);'>"
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("width", '80');
		tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='credits' onKeyPress='return mask(this);' onBlur='check(this);'>"
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("width", '50');
		tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Refs1' maxlength=20 onKeyPress='return mask(this);'>"
		newRow.appendChild(tempTD);

		tempTD=document.createElement("td");
		tempTD.setAttribute("width", '70');
		tempTD.innerHTML="<INPUT class='InvRowInput' TYPE='text' NAME='Refs2' maxlength=20 onKeyPress='return mask(this);' onBlur='acceptDate(this);'>"
		newRow.appendChild(tempTD);


		Rtable.insertBefore(newRow,theRow);
		
		lastRow++;

		for (rowNo=0; rowNo < lastRow ; rowNo++){
			tempTD=Rtable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
			tempTD.bgColor= '#F0F0F0';
			tempTD.innerText= rowNo+1;
		}

		activateCurrentCell();

	}

	document.onkeydown = documentKeyDown;
	activateCurrentCell();
	memoIsTaraz();
	//-->
	</script>


<%

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------- Submit Edit GL Memo Doc
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submitEditMemo" then

	ON ERROR RESUME NEXT
		if request.form("submit")="–ŒÌ—Â »Â ’Ê—  Ì«œœ«‘ " then
			IsTemporary = 0
		else
			IsTemporary = 1
		end if

		GLMemoDate=		sqlSafe(request.form("GLMemoDate"))
		GLMemoNo=		clng(request.form("GLMemoNo"))
		totalDebit =	cdbl(text2value(request.form("totalDebit")))
		totalCredit=	cdbl(text2value(request.form("totalCredit")))
		GLDoc =			clng(request.form("GLDoc"))

		TotalItemCount	=	request.form("GLAccounts").count 

		ReDim Accounts(TotalItemCount)
		ReDim GLAccounts(TotalItemCount)
		ReDim Descriptions(TotalItemCount)
		ReDim Amounts(TotalItemCount)
		ReDim IsCredit(TotalItemCount)
		ReDim Refs1(TotalItemCount)
		ReDim Refs2(TotalItemCount)

		for i=1 to TotalItemCount
			Accounts(i) =		clng(text2value(request.form("Accounts")(i)))
			GLAccounts(i) =		clng(text2value(request.form("GLAccounts")(i)))
			Descriptions(i) =	sqlSafe(request.form("Descriptions")(i))
			Refs1(i) =			sqlSafe(request.form("Refs1")(i))
			Refs2(i) =			sqlSafe(request.form("Refs2")(i))
			credit		=		cdbl(text2value(request.form("credits")(i)))	
			debit		=		cdbl(text2value(request.form("debits")(i)))	
			if credit <> "" and credit <> "0" then 
				Amounts(i) = credit 
				IsCredit(i)= 1
			else
				Amounts(i) = debit 
				IsCredit(i)= 0
			end if
			if Accounts(i) = 0 then Accounts(i) = "NULL"
		next	

		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
		end if
	ON ERROR GOTO 0

	creationDate = shamsiToday()

	'---- Checking wether EffectiveDate is valid in current open GL
	if (GLMemoDate < session("OpenGLStartDate")) OR (GLMemoDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?act=editDoc&id="& GLMemoNo & "&errMsg=" & Server.URLEncode("Œÿ«!  «—ÌŒ Ê«—œ ‘œÂ „ ⁄·ﬁ »Â ”«· „«·Ì Ã«—Ì ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
'response.write "<br>I am going to submit!"
'response.end
	'---- Marking old GLDoc and its GLRows as DELETED
	conn.Execute("UPDATE GLRows SET deleted = 1 WHERE (GLDoc = "& GLDoc & ")")
	conn.Execute("UPDATE GLDocs SET deleted = 1 WHERE (ID = "& GLDoc & ")")
	'---- 
	
	'---- Creating a new GLDoc 
	mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, createdBy,  IsTemporary) VALUES ("& openGL & " , "& GLMemoNo & ", N'"& GLMemoDate & "' , N'"& creationDate & "', "& session("ID") & ", "& IsTemporary & ");SELECT @@Identity AS NewGLDoc"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	GLDoc = RS1 ("NewGLDoc")
	RS1.close
	'---- 

	'---- Inserting new GLRows 
	for i=1 to TotalItemCount
		mySQL="INSERT INTO GLRows ( GLDoc, GLAccount,  Tafsil, Amount, Description, Ref1, Ref2, IsCredit) VALUES ( "& GLDoc & ", "& GLAccounts(i) & ", "& Accounts(i) & ", "& Amounts(i) & ", N'"& Descriptions(i) & "', N'"& Refs1(i) & "', N'"& Refs2(i) & "', "& IsCredit(i) & ")"
		conn.Execute(mySQL)
	next	
	'---- 

	response.redirect "GLMemoDocShow.asp?id="& GLDoc & "&msg=" & Server.URLEncode(" €ÌÌ—«  À»  ‘œ.")

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Submit new GL Memo Doc
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submitMemo" then

	GLMemoDate=request.form("GLMemoDate")
	GLMemoNo=request.form("GLMemoNo")
	totalDebit = text2value(request.form("totalDebit"))
	totalCredit = text2value(request.form("totalCredit"))
	creationDate = shamsiToday()
	if request.form("submit")="–ŒÌ—Â »Â ’Ê—  Ì«œœ«‘ " then
		IsTemporary = 0
	else
		IsTemporary = 1
	end if

	
	if GLMemoNo="" or not(isnumeric(GLMemoNo)) then
		ShowErrorMessage("Œÿ«! ‘„«—Â ”‰œ Ê«—œ ‰‘œÂ «” ")
		response.end
	end if

	'---- Checking wether EffectiveDate is valid in current open GL
	if (GLMemoDate < session("OpenGLStartDate")) OR (GLMemoDate > session("OpenGLEndDate")) then
		Conn.close
		response.write "<BR><BR><BR><CENTER>Œÿ«! ”‰œ Ê«—œ ‘œÂ „ ⁄·ﬁ »Â ”«· „«·Ì Ã«—Ì ‰Ì” .</CENTER>"
		response.end
	end if 
	'----

'	Set RS3 = conn.Execute("SELECT * FROM  GLDocs WHERE (GLDocID = "& GLMemoNo & ") AND (GL = "& OpenGL & ")")
'	if not RS3.EOF then
'		OldGLMemoDate = RS3("GLDocDate")
'		response.write "<BR><BR><BR><CENTER>Œÿ«! «Ì‰ ”‰œ ﬁ»·« »«  «—ÌŒ <span dir=ltr> "& OldGLMemoDate & " </span> Ê«—œ ‘œÂ «” .  </CENTER>"
'		response.end
'	end if 

	WarningMsg=""
	Set RS3=Conn.Execute ("SELECT GLDocID FROM GLDocs WHERE (GLDocID='"& GLMemoNo & "') AND (GL='"& OpenGL & "')")
	if not RS3.eof then
		Set RS3= Conn.Execute("SELECT Max(GLDocID) AS MaxGLDocID FROM GLDocs WHERE (GL='"& OpenGL & "')")
		GLMemoNo=RS3("MaxGLDocID")+1
		WarningMsg="‘„«—Â ”‰œ  ﬂ—«—Ì »Êœ.<br>”‰œ »« ‘„«—Â <B>"& GLMemoNo & "</B> À»  ‘œ."
	end if
	RS3.Close
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----

	%>
	<br><br><br>
	<TABLE class="GLTable2" Cellspacing="0" Cellpadding="0" width=90% align=center>
	<tr>
		<td style="width:26; border-right:none;"> # </td>

		<td style="width:50; ">Õ”«»</td>
		<td style="width:170;">‰«„ Õ”«»</td>
		<td style="width:300;">‘—Õ</td>
		<td style="width:80;">»œÂﬂ«—</td>
		<td style="width:80;">»” «‰ﬂ«—</td>
	</tr>
	</table>
	<TABLE Border="0" width=90%  align=center Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="GLTable3">
	<Tbody id="GLrows">
	<%

	mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, createdBy,  IsTemporary) VALUES ("& OpenGL & " , "& GLMemoNo & ", N'"& GLMemoDate & "' , N'"& creationDate & "', "& session("id") & ", "& IsTemporary & ")"
	conn.Execute(mySQL)

	mySQL="SELECT max(ID) as GLDoc FROM GLDocs where GLDocID="& GLMemoNo
	set RS1=conn.execute(mySQL)
	GLDoc=RS1("GLDoc")


	for i=1 to request.form("Items").count 
		GLAccount = text2value(request.form("Items")(i))
		accTitle = request.form("accTitles")(i)
		theDescription = request.form("Descriptions")(i)
		debit = text2value(request.form("debits")(i))
		credit = text2value(request.form("credits")(i))

		if credit <> "" and credit <> "0" then 
			Amount = credit 
			IsCredit = 1
		else
			Amount = debit 
			IsCredit = 0
		end if

		if amount = "" then amount = 0

		mySQL="INSERT INTO GLRows ( GLDoc, GLAccount,  Amount, Description, IsCredit) VALUES ( "& GLDoc & ", "& GLAccount & ", "& Amount & ", N'"& theDescription & "', "& IsCredit & ")"
		
		conn.Execute(mySQL)

		%>
		<tr bgcolor='#F0F0F0' >
		<td style="width:25; border-right:none;"> <%=i%> </td>
		<td style="width:50;"><%=theItem%></td>
		<td style="width:170;"><%=accTitle%></td>
		<td style="width:300;"><%=theDescription%></td>
		<td style="width:80;"><%=debit%></td>
		<td style="width:80;"><%=credit%></td>
		</tr>
		<%
	next	
%>
</Tbody></TABLE>
<table  width="90%" align="center" Cellspacing="0" Cellpadding="0"  align=center>
<tr style="border:none;">
	<td style="width:26;"></td>
	<td style="width:50;"></td>
	<td style="width:170;"></td>
	<td style="width:300;">Ã„⁄</td>
	<td style="width:80;"><%=totalDebit%></td>
	<td style="width:80;"><%=totalCredit%></td>
</tr>
</TABLE>	
<CENTER>
<BR><BR>
”‰œ ›Êﬁ À»  ‘œ
(‘„«—Â ”‰œ: <A target=_blank  HREF="GLMemoDocShow.asp?id=<%=GLDoc%>"><%=GLMemoNo%></A>)
</CENTER>
<%
	'response.redirect "GLMemoDocShow.asp?id="& GLDoc &"&msg="& Server.URLEncode("”‰œ »« „Ê›ﬁÌ  «ÌÃ«œ ‘œ.") &"&errmsg="& Server.URLEncode(WarningMsg)
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------

end if
conn.Close
%>
</font>
<!-- <textarea id="TTT" rows="10" cols="50"></textarea> -->

<% if request("act")="" then%>

<script language="JavaScript">
<!--
function setCurrentRow(rowNo){
	if (rowNo == -1) rowNo=0;
	invTable=document.getElementById("GLrows");
	theTD=invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0];
	theTD.setAttribute("bgColor", '#F0F0F0');

	currentRow=rowNo;
	invTable=document.getElementById("GLrows");
	theTD=invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0];
	theTD.setAttribute("bgColor", '#FFB0B0');
}
function delRow(rowNo){
	invTable=document.getElementById("GLrows");
	theRow=invTable.getElementsByTagName("tr")[rowNo];
	invTable.removeChild(theRow);

	rowsCount=document.getElementsByName("Items").length;
	for (rowNo=0; rowNo < rowsCount ; rowNo++){
		tempTD=invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
		tempTD.bgColor= '#F0F0F0';
		tempTD.innerText= rowNo+1;
	}
}
function addRow(){
	rowNo = currentRow
	invTable=document.getElementById("GLrows");
	theRow=invTable.getElementsByTagName("tr")[rowNo];
	newRow=document.createElement("tr");
	newRow.setAttribute("bgColor", '#f0f0f0');
	newRow.setAttribute("onclick", theRow.getAttribute("onclick"));

	tempTD=document.createElement("td");
	tempTD.innerHTML=rowNo+1
	tempTD.setAttribute("align", 'center');
	tempTD.setAttribute("width", '26');
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.setAttribute("width", '50');
	tempTD.innerHTML="<INPUT class='InvRowInput' TYPE='text' NAME='Items' onKeyPress='return mask(this);' onBlur='check(this);' onfocus='setCurrentRow(this.parentNode.parentNode.rowIndex);' style='width:100%;' MAXLENGTH=5> "
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("width", '170');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='accTitles'  readonly  tabIndex=99999999 style='width:100%;' onfocus='this.parentNode.parentNode.getElementsByTagName(\"td\")[3].getElementsByTagName(\"Input\")[0].focus()'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("width", '300');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Descriptions' style='width:100%;'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.setAttribute("width", '80');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='debits' style='width:100%;' onblur='setPrice(this)'  onKeyPress='return onlyNumber(this);'  onfocus='this.value=txt2val(this.value);this.select()'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.setAttribute("width", '80');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='credits' style='width:100%;'  onblur='setPrice(this)' onKeyPress='return onlyNumber(this);'   onfocus='setCurrentRow(this.parentNode.parentNode.rowIndex);this.value=txt2val(this.value);this.select()'>"
	newRow.appendChild(tempTD);


	invTable.insertBefore(newRow,theRow);
	
	rowsCount=document.getElementsByName("Items").length;
	for (rowNo=0; rowNo < rowsCount ; rowNo++){
		tempTD=invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
		tempTD.bgColor= '#F0F0F0';
		tempTD.innerText= rowNo+1;
	}

	invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
//	document.all.dddd2.innerText=invTable.innerHTML
}

function setPrice(src){
	src.value=val2txt(txt2val(src.value));
	rowNo=src.parentNode.parentNode.rowIndex;
	if (src.name=="credits" && src.value!=0) 
		document.getElementsByName("debits")[rowNo].value = ""
	if (src.name=="debits" && src.value!=0) 
		document.getElementsByName("credits")[rowNo].value = ""
	if (src.value ==0) src.value=""
	if (src.value ==" ") src.value=""

	memoIsTaraz();

}


function saveDraft(){ 
	if (document.all.totalDebit.value != 0 || document.all.totalCredit.value != 0)
			IsTaraz = true
}

function checkValidation(){ 
	if (IsTaraz==true)
		return true
	else{
		alert("Œÿ«! ”‰œ  —«“ ‰Ì” ")
		return false
	}
}

function memoIsTaraz(){ 
	var totalCredit = 0;
	var totalDebit = 0;
	for (rowNo=0; rowNo < document.getElementsByName("debits").length; rowNo++){
		totalCredit += parseInt(txt2val(document.getElementsByName("credits")[rowNo].value));
		totalDebit += parseInt(txt2val(document.getElementsByName("debits")[rowNo].value));
	}
	document.all.totalCredit.value = val2txt(totalCredit);
	document.all.totalDebit.value = val2txt(totalDebit);
	if (totalDebit == totalCredit && totalCredit != 0){
		IsTaraz = true
		document.all.tarazDiv.innerHTML = "<FONT COLOR='#008833'>”‰œ  —«“ «” </FONT>"
	}
	else{
		IsTaraz = false
		document.all.tarazDiv.innerHTML = "<FONT COLOR='#FF3300'>”‰œ  —«“ ‰Ì” </FONT>"
	}

}

var dialogActive=false;

function mask(src){ 
	var theKey=event.keyCode;

	rowNo=src.parentNode.parentNode.rowIndex;
	invTable=document.getElementById("GLrows");
	theRow=invTable.getElementsByTagName("tr")[rowNo];

	if (src.name=="Items"){
		if (theKey==32){
			event.keyCode=9
			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì œ› — ﬂ·:"
			var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				dialogActive=false
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[2].getElementsByTagName("Input")[0].value=Arguments[1];
				}
				invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[3].getElementsByTagName("Input")[0].focus();
			}
		}
		else if (theKey >= 48 && theKey <= 57 ) { 
			//alert(theKey)
			//src.value=''
			return true;
		}
		else { 
			return false;
		}
	}
}
function check(src){ 
	if (src.name=="Items"){
		rowNo=src.parentNode.parentNode.rowIndex;
		rowsCount=document.getElementsByName("Items").length;
		if (!dialogActive){
			if (src.value=='0' || src.value==''){
				if (rowNo == 0 && rowsCount == 1) {
					src.focus();
				}
				else{
					if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ”ÿ— —« Õ–› ﬂ‰Ìœø")){
						delRow(rowNo);
						if (rowNo != rowsCount ){
							invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
						}else{
							invTable.getElementsByTagName("tr")[rowNo-1].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
						}
						return false;
					}
					else{
						src.focus();
					}
				}
			}
			else {
				if (window.XMLHttpRequest) {
					var objHTTP=new XMLHttpRequest();
				} else if (window.ActiveXObject) {
					var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
				}
				objHTTP.open('GET','xml_GLAccount.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape( objHTTP.responseText)
				invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[2].getElementsByTagName("Input")[0].value= tmpStr
				
				if (tmpStr=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ")
				{
					if (rowNo == 0 && rowsCount == 1) {
						src.focus();
					}
					else{
						if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ”ÿ— —« Õ–› ﬂ‰Ìœø")){
							delRow(rowNo);
							if (rowNo != rowsCount ){
								invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
							}else{
								invTable.getElementsByTagName("tr")[rowNo-1].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
							}
							return false;
						}
						else{
							src.focus();
							src.select()
						}
					}
				}
			}
		}
	}
}

function onlyNumber(src){ 
	var theKey=event.keyCode;
	if (theKey==13){  // [Enter] 
		return true;
	}
	else if (theKey < 48 || theKey > 57) { // 0-9 are acceptible
		return false;
	}
}

function areYouSureToExit()
{
	a= confirm("are you sure?")
	return a
}
setPrice(document.all.totalDebit)
setPrice(document.all.totalCredit)
if (parseInt(document.all.totalDebit.value)==0) IsTaraz=false;
//-->
</SCRIPT>

<%end if%>
<!--#include file="tah.asp" -->
