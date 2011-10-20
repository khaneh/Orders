<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'A0 (11 [=B])
PageTitle= "ÕﬁÊﬁ - Ê«„ Ê €Ì—Â"
SubmenuItem=2
if not Auth("B" , 2) then NotAllowdToViewThisPage()

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
	.InvHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.InvRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right; width:100%;}
	.InvRowInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; direction:RTL; width:100%;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right; direction: right-to-left;}
	.InvGenInput  { font-family:tahoma; font-size: 9pt; border: none; }
	.InvGenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
	.GLInput1 { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction:RTL;}
	.GLInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction:LTR;}
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
<%
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------- Input a new AO MEMO
'-----------------------------------------------------------------------------------------------------
if request("act")="" then
%>
<!-- Ê—Êœ «ÿ·«⁄«  ”‰œ -->
<script language="JavaScript">
<!--
function setCurrentRow(rowNo){
	if (rowNo == -1) rowNo=0;
	MemosTbl=document.getElementById("AOMemos");
	theTD=MemosTbl.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0];
	theTD.setAttribute("bgColor", '#F0F0F0');

	currentRow=rowNo;
	MemosTbl=document.getElementById("AOMemos");
	theTD=MemosTbl.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0];
	theTD.setAttribute("bgColor", '#FFB0B0');
}
function delRow(rowNo){
	rowsCount=document.getElementsByName("Accounts").length;
	if(rowsCount<2){
		alert("Õ–› Œÿ «„ﬂ«‰ Å–Ì— ‰Ì” ");
		return false;
	}
	
	MemosTbl=document.getElementById("AOMemos");
	theRow=MemosTbl.getElementsByTagName("tr")[rowNo];
	MemosTbl.removeChild(theRow);

	rowsCount=document.getElementsByName("Accounts").length;
	for (rowNo=0; rowNo < rowsCount ; rowNo++){
		tempTD=MemosTbl.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
		tempTD.bgColor= '#F0F0F0';
		tempTD.innerText= rowNo+1;
	}
	event.cancelBubble=true;
	setCurrentRow(currentRow);
	if (currentRow!=rowsCount){
		tempTD=MemosTbl.getElementsByTagName("tr")[currentRow].getElementsByTagName("INPUT")[0];
		tempTD.select();
		tempTD.focus();
	}


}
function addRow(){

	rowNo = currentRow;

	MemosTbl=document.getElementById("AOMemos");
	theRow=MemosTbl.getElementsByTagName("tr")[rowNo];
	theRow.getElementsByTagName("td")[0].bgColor= '#F0F0F0';

	rowsCount=document.getElementsByName("Accounts").length;
	if (currentRow==rowsCount)
		newRow=theRow.previousSibling.cloneNode(true);
	else
		newRow=theRow.cloneNode(true);

	MemosTbl.insertBefore(newRow,theRow);
	
	rowsCount=document.getElementsByName("Accounts").length;
	for (rowNo=0; rowNo < rowsCount ; rowNo++){
		tempTD=MemosTbl.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
		tempTD.innerText= rowNo+1;
	}
	event.cancelBubble=true;
	setCurrentRow(currentRow);
	tempTD=MemosTbl.getElementsByTagName("tr")[currentRow].getElementsByTagName("INPUT")[0]
	tempTD.select();
	tempTD.focus();
//	document.all.dddd2.innerText=MemosTbl.innerHTML
}

var dialogActive=false;
function rowMask(src){
	var theKey=event.keyCode;
	if (theKey == 45 && event.ctrlKey) {
		addRow();
	}
	else if (theKey == 46 && event.ctrlKey) {
		delRow(currentRow)
		setCurrentRow(currentRow);
	}
}
function mask(src){ 
	var theKey=event.keyCode;
	if (src.name=="Accounts"){
		if (theKey==32){
			event.keyCode=0;
			dialogActive=true;
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì  ›’Ì·Ì:"
			window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false;
			if (document.all.tmpDlgTxt.value !="") {
				dialogActive=true;
				window.showModalDialog('../AR/dialog_selectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:300px; dialogWidth:600px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				dialogActive=false;
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					src.title=Arguments[1];
				}
			}
		}
		else if (theKey >= 48 && theKey <= 57 ) 
			return true;
		else
			return false;
	}
	else if(src.name=="GLAccounts" || src.name=="firstGLAccounts"){
		if (theKey==32){
			event.keyCode=0;
			dialogActive=true;
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì œ› — ﬂ·:"
			window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false;
			if (document.all.tmpDlgTxt.value !="") {
				dialogActive=true;
				window.showModalDialog('../accounting/dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				dialogActive=false;
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					src.title=Arguments[1];
				}
			}
		}
		else if (theKey >= 48 && theKey <= 57 ) // [0]-[9] are acceptible
			return true;
		else
			return false;
	}
	else if(src.name=="Amounts"){
		if (theKey >= 48 && theKey <= 57 ) // [0]-[9] are acceptible
			return true;
		else
			return false;
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
			objHTTP.open('GET','../accounting/xml_CustomerAccount.asp?id='+src.value,false)
			objHTTP.send()
			tmpStr = unescape( objHTTP.responseText)
			src.title=tmpStr;
			if (tmpStr=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ"){
				src.value="";
			}
		}
	}
	else if(src.name=="GLAccounts" || src.name=="firstGLAccounts"){
		if (!dialogActive){
			if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
			objHTTP.open('GET','../accounting/xml_GLAccount.asp?id='+src.value,false)
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
	else if(src.name=="Amounts"){
		src.value=val2txt(txt2val(src.value));
	}
}

function isEmpty()
{
	if (document.all.Accounts.value=="" )
		{
		alert("›—„ Œ«·Ì «” ")
		return false
		}
	return true
}
//-->
</script>
	<br>
	<br>
<%
if session("id")=-1 then 
%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function importBulk(){
		//alert('asd?');
		src=escape(document.all.BulkInput.value)
		Rows = src.split("%0D%0A") // Line Break
		for (i=0;i<Rows.length;i++){
			if (Rows[i]!=''){
				Cols=Rows[i].split("%09"); // Tab


				setCurrentRow(currentRow+1);

				addRow();

				MemosTbl=document.getElementById("AOMemos");
				theRow=MemosTbl.getElementsByTagName("tr")[currentRow];

				theRow.getElementsByTagName("INPUT")[0].value= Cols[0];
				theRow.getElementsByTagName("INPUT")[1].value= Cols[1];

				theRow.getElementsByTagName("SELECT")[0].value= Cols[2];

				theRow.getElementsByTagName("INPUT")[2].value= Cols[3];
				theRow.getElementsByTagName("INPUT")[3].value= unescape(Cols[4]);
				theRow.getElementsByTagName("INPUT")[4].value= Cols[5];
				theRow.getElementsByTagName("INPUT")[5].value= Cols[6];


			}
		}
	}
	//-->
	</SCRIPT>
	<TABLE align=center style='border:2 dashed green;width:' width='300px;'>
	<TR>
		<TD>
			Ê—Êœ ÃœÊ·
		</TD>
	</TR>
	<TR>
		<TD>
			<TEXTAREA NAME="BulkInput" ROWS="5" COLS="50"></TEXTAREA>
		</TD>
	</TR>
	<TR>
		<TD align=center>
			<INPUT TYPE="button" value="Import" style='border:1 solid black;' onclick='importBulk();'>
		</TD>
	</TR>
	</TABLE>
	<br><br>	
<%
end if
%>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<table class="GLTable1" align="center" Cellspacing="0" Cellpadding="0">
	<FORM METHOD=POST ACTION="?act=submitMemo" onsubmit="return isEmpty();">
		<tr style="background-color:black;height:1px;">
			<TD></TD>
		</tr>
		<tr class="GLTR2"><TD>
			<TABLE class="GLTable2" Cellspacing="0" Cellpadding="0" >
			<tr>
				<td style="width:26; border-right:none;"> # </td>
				<td style="width:60;"> ›’Ì·Ì</td>
				<td style="width:40;">„⁄Ì‰</td>
				<td style="width:80;">»œ / »”</td>
				<td style="width:80;">„»·€</td>
				<td style="width:250;">‘—Õ</td>
				<td style="width:60;">ÿ—› „⁄Ì‰</td>
				<td style="width:70;"> «—ÌŒ</td>
				<td style="width:20;"></td>
			</tr>
			</TABLE></TD>
		</tr>
		<tr><TD>
			<div style="overflow:auto; height:250px; width:*;">
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="GLTable3">
			<Tbody id="AOMemos">
			<TR bgColor=#f0f0f0  onclick="setCurrentRow(this.rowIndex);" onKeyDown="rowMask(this);">
				<td width='25' align=center> 1 </td>
				<td width="60">
					<INPUT class='InvRowInput' TYPE='text' NAME='Accounts' maxlength=6 title="" onKeyPress='return mask(this);' onBlur='check(this);'></td>
				<td width="40">
					<INPUT class='InvRowInput2' TYPE='text' NAME='firstGLAccounts' maxlength=5 title="" onKeyPress='return mask(this);' onBlur='check(this);'></td>
				<td width="80">
					<SELECT class='InvRowInput2' NAME='isCredits'>
						<option Value="1">»” «‰ﬂ«—</option>
						<option Value="0">»œÂﬂ«—</option>
					</SELECT></td>
				<td width="80" dir="LTR">
					<INPUT class='InvRowInput2' TYPE='text' NAME='Amounts' onKeyPress='return mask(this);' onBlur='check(this)' onfocus='this.select()'></td>
				<td width="250">
					<INPUT class='InvRowInput2' TYPE='text' NAME='Descriptions'></td>
				<td width="60">
					<INPUT class='InvRowInput2' TYPE='text' NAME='GLAccounts' onKeyPress='return mask(this);' onBlur='check(this);' maxlength=5 ></td>
				<td width="70">
					<INPUT class='InvRowInput2' style='direction:LTR;text-align:left;width:70' TYPE='text' NAME='EffectiveDates' onblur="acceptDate(this)" maxlength=10></td>
			</TR>
			<TR bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
				<td colspan="8">
					<INPUT class="InvGenButton" TYPE="button" value="«÷«›Â" onkeyDown="if(event.keyCode==9) {setCurrentRow(this.parentNode.parentNode.rowIndex); return false;};" onClick="addRow();">
				</td>
			</TR>
			</Tbody>
			</TABLE></div>
			</TD>
		</tr>
	</table> 
	<TABLE Border="0" Cellspacing="5" Cellpadding="10" Dir="RTL" align='left'>
	<tr>
		<td align='center'><INPUT class="InvGenButton" TYPE="submit" value="–ŒÌ—Â "></td>
	</tr>
	</TABLE>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		setCurrentRow(0);
		document.getElementById("AOMemos").getElementsByTagName("INPUT")[0].focus();
	//-->
	</SCRIPT>

<%
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Submit new AO MEMO
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submitMemo" then

	creationDate=shamsiToday()
	errMsg=""
	successCount=0
	for i=1 to request.form("Accounts").count 
		errorFound=false
		ON ERROR RESUME NEXT
			Account =		clng(text2value(request.form("Accounts")(i)))
			firstGLAccount=	clng(text2value(request.form("firstGLAccounts")(i)))
			GLAccount =		clng(text2value(request.form("GLAccounts")(i)))
			Amount =		cdbl(text2value(request.form("Amounts")(i)))
			Description=	sqlSafe(request.form("Descriptions")(i))
			EffectiveDate =	sqlSafe(request.form("EffectiveDates")(i))

			if request.form("isCredits")(i) then
				isCredit = 1
			else
				isCredit = 0
			end if
			if Err.Number<>0 then
				Err.clear
				errorFound=True
			end if
			if NOT errorFound then
			' checking firstGLAccount is valid
				mySQL = "SELECT ID FROM GLAccounts WHERE (HasAppendix = 1) AND (GL = "& openGL & ") AND (ID = "& firstGLAccount & ")"
				Set RS1=Conn.execute(mySQL)
				if RS1.eof then
					errorFound=True
				end if
				RS1.close

				mySQL = "SELECT ItemReason FROM AXItemReasonGLAccountRelations WHERE (GL = "& openGL & ") AND (GLAccount = "& firstGLAccount & ")"
				Set RS1=Conn.execute(mySQL)
				if RS1.eof then
					'Using default reason (sys: AO, reason: Misc.)
					Reason=6
				else
					Reason=	cint(RS1("ItemReason"))
				end if
				RS1.close

				mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
				Set RS1=Conn.execute(mySQL)
				sys=			RS1("Acron")
				ReasonName =	RS1("Name")
				RS1.close
			' ---------

			'---- Checking wether EffectiveDate is valid in current open GL
				if (effectiveDate < session("OpenGLStartDate")) OR (effectiveDate > session("OpenGLEndDate")) then
					errorFound=True
				end if 
			'----
			end if
		ON ERROR GOTO 0
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		'Conn.close
		errMsg="Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ." & "<br>"
		errorFound=True
	end if 
	'----

		if errorFound then
			response.write "<br>" 
			errMsg= errMsg & "Œÿ«!<br>Œÿ »« „‘Œ’«  —Ì“ «÷«›Â ‰‘œ:<br>"
			errMsg=errMsg& "Õ”«»:" & firstGLAccount & "-" & Account & "<br>"
			errMsg=errMsg& "„»·€:" & Amount & "<br>"
			errMsg=errMsg& "‘—Õ:" & Description & "<br>"
			errMsg=errMsg& "ÿ—› „⁄Ì‰:" & GLAccount & "<br>"
			errMsg=errMsg& " «—ÌŒ :" & effectiveDate & "<br>"
			call showAlert (errMsg,CONST_MSG_ERROR) 
		else
			'**************************** Creating Memo ****************
			'MemoType = 1 means Misc. Memo
			mySQL="INSERT INTO "& sys & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& effectiveDate & "' , "& session("id") & ", 1, "& Account & ", "& IsCredit & ", "& Amount & ", N'"& Description & "');SELECT @@Identity AS NewMemo"
			Set RS1 = conn.Execute(mySQL).NextRecordSet
			MemoID=RS1("NewMemo")
			RS1.close

			'**************************** Creating an Item for Memo ****************
			'*** Type = 3 means Item is a Memo

			mySQL="INSERT INTO "& sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, Link, Reason, CreatedDate, CreatedBy, Type, IsCredit,  AmountOriginal, RemainedAmount) VALUES ("&_
			GLAccount & ", "& OpenGL & ", "& firstGLAccount & ", "& Account & ", N'"& effectiveDate & "', "& MemoID & ", "& Reason & ", N'"& creationDate & "' , "& session("id") & ", 3, "& IsCredit & ",  "& Amount & ","& Amount & ")"
			conn.Execute(mySQL)
		
			if IsCredit then
				mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance + '"& Amount & "' WHERE (ID='"& Account & "')"
			else
				mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance - '"& Amount & "' WHERE (ID='"& Account & "')"
			end if
			conn.Execute(mySQL)
			'***------------------------- Creating an Item for Memo ----------------
			successCount = successCount + 1
			msg="«⁄·«„ÌÂ <A Href='../"& sys & "/AccountReport.asp?act=showMemo&sys="& sys & "&memo=" & MemoID &"' target='_blank'>"& MemoID & "</a> À»  ‘œ."
			response.write "<br>" 
			call showAlert (msg,CONST_MSG_INFORM) 
		end if
	next	
	response.write "<br>" 
	response.write "<hr width='70%'>" 
	response.write "<br>" 
	call showAlert (successCount & " «⁄·«„ÌÂ »« „Ê›ﬁÌ  «ÌÃ«œ ‘œ.",CONST_MSG_INFORM) 
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
end if
conn.Close
%>
</font>
<!--#include file="tah.asp" -->
