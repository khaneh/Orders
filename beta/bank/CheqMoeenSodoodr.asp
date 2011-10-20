<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "çﬂ „⁄Ì‰"
SubmenuItem=7
if not Auth("A" , 7) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>
	.RcpTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; }
	.RcpMainTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; background-color: #558855; text-align:right; direction: right-to-left;}
	.RcpMainTableTH { background-color: #C3C300;}
	.RcpMainTableTR { background-color: #CCCC88; border: 0; }
	.RcpRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.RcpRowInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.RcpHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.RcpHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.RcpHeadInput3 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #D0E0FF; text-align:right; direction: right-to-left;}
	.RcpGenInput { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction: left-to-righ;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
	.CustGenTable  { font-family:tahoma; font-size: 9pt;}
	.CustGenInput { font-family:tahoma; font-size: 9pt;}
</STYLE>

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

		errorFound=false

		GLAccount =			clng(request.form("GLAccount"))
		GLAccountName =		sqlSafe(request.form("GLAccountName"))

		TotalAmount=		cdbl(text2value(request.form("TotalAmount")))

		if Err.Number<>0 then
			errorFound=	true
			errorMsg=	Err.Description
		end if

		if NOT errorFound then
			effectiveDate=	sqlSafe(request.form("GLMemoDate"))
			'---- Checking wether EffectiveDate is valid in current open GL
			if (effectiveDate < session("OpenGLStartDate")) OR (effectiveDate > session("OpenGLEndDate")) then
				errorFound=	true
				errorMsg=	Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ „⁄ »— ‰Ì” .")
			end if 
			'----
			'----- Check GL is closed
			if (session("IsClosed")="True") then
				Conn.close
				response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
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
				'CheqStatus = 21 «”‰«œ Å—œ«Œ ‰Ì Å«” ‰‘œÂ
				mySQL="SELECT 1 AS OK WHERE ("& Banker & " IN (SELECT Banker FROM BankerCheqStatusGLAccountRelation WHERE (CheqStatus = 21) AND (GL = "& OpenGL & ")))"
				set tmpRS= Conn.Execute(mySQL)
				if tmpRS.eof then
					errorFound=	true
					errorMsg=	Server.URLEncode("<B>Œÿ«!  </B><BR>Õ”«» «”‰«œ Å—œ«Œ „‰Ì »—«Ì »«‰ﬂ '"& Banker & "'  ⁄—Ì› ‰‘œÂ «” .<BR>")
				end if
				tmpRS.close

				Description = sqlSafe(request.form("Description")(i))
				Amount		= cdbl(text2value(request.form("Amounts")(i)))

				if ChequeNo=0 or ChequeDate="" or  Banker=-1 or Amount=0 then 
					errorFound=	true
					errorMsg=	Server.URLEncode("<B>Œÿ«!  </B><BR>«ÿ·«⁄«  çﬂ ‰«ﬁ’ «” <BR>")
				end if 
			next
		end if
		if Err.Number<>0 then
			errorFound=	true
			errorMsg=	Err.Description
		end if
	ON ERROR GOTO 0

	if errorFound then
		conn.close
		response.redirect "?act=getPayment&selectedCustomer="& CustomerID & "&Reason="& Reason & "&errMsg=" & errorMsg
	end if

	creationDate =shamsiToday()
	CreationTime = currentTime10()


	'************* Finding Last Memo Doc Number
	mySQL="SELECT ISNULL(MAX(GLDocID),0) AS LastMemo FROM GLDocs WHERE GL='"& OpenGL & "'"
	Set RS1=conn.Execute (mySQL)
	GLMemoNo = RS1("LastMemo") + 1

	'************* Inserting New Memo Doc
	mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, createdBy,  IsTemporary) VALUES ("& OpenGL& " , "& GLMemoNo & ", N'"& effectiveDate & "' , N'"& creationDate & "', "& session("id") & ", 1);SELECT @@Identity AS NewDoc"

	set RSE = Conn.execute(mySQL).NextRecordSet
	GLDoc = RSE ("NewDoc")
	RSE.close
	set RSE = Nothing


	'************* Inserting Cheques
	for i = 1 to CheqsCount
		ChequeNo	= clng(request.form("ChequeNos")(i))
		ChequeDate	= sqlSafe(request.form("ChequeDates")(i))
		Banker		= cint(request.form("Banks")(i))
		mySQL=		"SELECT GLAccount FROM BankerCheqStatusGLAccountRelation WHERE (CheqStatus = 21) AND (GL = "& OpenGL & ") AND (Banker = "& Banker & ")"
		set tmpRS = Conn.Execute(mySQL)
		DestGLAccount = tmpRS("GLAccount")
		tmpRS.close

		Amount		= cdbl(text2value(request.form("Amounts")(i)))

		theDescription = "çﬂ " & ChequeNo & " „Ê—Œ " & ChequeDate & " " & GLAccountName & " " & Description

		mySQL="INSERT INTO GLRows ( GLDoc, GLAccount,  Amount, Description, IsCredit, Ref1, Ref2) VALUES ( "& GLDoc & ", "& GLAccount & ", "& Amount & ", N'"& theDescription & " " & accTitle & "', 0, N'"& ChequeNo &"' , N'"& ChequeDate &"')"
		conn.Execute(mySQL)

		mySQL="INSERT INTO GLRows ( GLDoc, GLAccount,  Amount, Description, IsCredit, Ref1, Ref2) VALUES ( "& GLDoc & ", "& DestGLAccount & ", "& Amount & ", N'"& theDescription & "', 1, N'"& ChequeNo &"', N'"& ChequeDate &"')"
		conn.Execute(mySQL)
	next

	response.redirect "../accounting/GLMemoDocShow.asp?id="&GLDoc&"&msg=çﬂ À»  ‘œ"
	'response.write replace(request.form,"&", "<br>")
	response.end
	
elseif request("act")="" then

	GLAccount = ""
	GLAccountName=""

	creationDate=	shamsiToday()
	creationTime=	currentTime10()
%>	
	<br>
<!--#include File="../include_JS_InputMasks.asp"-->

<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>

	<FORM METHOD=POST ACTION="?act=submit" onsubmit="return checkValidation();">
		<input type="hidden" Name='tmpDlgArg' value=''>
		<table class="RcpMainTable" Cellspacing="1" Cellpadding="0" Width="500" align="center">
		<tr class="RcpMainTableTH">
			<td colspan="10"><TABLE class="RcpTable" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL"><TR>
				<TD align="left">„⁄Ì‰ »œÂﬂ«—:</TD>
				<TD align="right" width="5">
					<INPUT class="RcpGenInput" TYPE="text" NAME="GLAccount" value="<%=GLAccount%>" maxlength="5" size="5" tabIndex="1" dir="LTR" onblur="check(this);" onKeyPress="return mask(this);">
				</TD>
				<TD align="right" width="150">
					<TextArea NAME="GLAccountName" readonly rows=1 style="width:100%;font-family:tahoma;font-size:9pt;background:transparent; border:0;overflow:hidden;"><%=GLAccountName%></TextArea>
				</TD>
				<TD align="left"> «—ÌŒ:</TD>
				<TD>
					<table class="RcpTable">
					<tr>    
						<td dir="LTR">
							<input class="RcpGenInput" style="text-align:left;direction:LTR;" NAME="GLMemoDate" TYPE="text" maxlength="10" size="10" value="" tabindex="2" onblur="acceptDate(this)">
						</td>
						<td dir="RTL">«„—Ê“: <font color=white><%=weekdayname(weekday(date))%> <span dir=rtl><%=shamsiToday()%></span></font></TD>
					</tr>   
					</table></TD>
			</TR></TABLE></td>
		</tr>
		<tr class="RcpMainTableTR">
			<td colspan="10"><div style="width:550;">
			<TABLE class="RcpTable" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TR height=20px>
				<TD class="RcpHeadInput" style="width:25;"> # </td>
				<TD class="RcpHeadInput" style="width:70;">‘„«—Â çﬂ</td>
				<TD class="RcpHeadInput" style="width:70;"> «—ÌŒ</td>
				<TD class="RcpHeadInput" style="width:150;">»«‰ﬂ («”‰«œ Å—œ«Œ ‰Ì)</td>
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
			<td colspan="10"><div style="width:550;">
			<TABLE class="RcpTable" Cellspacing="0" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TR height=20px>
				<TD class="RcpHeadInput" style="width:318;text-align:left;"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="Ã„⁄:" size="10" tabindex="9999"></td>
				<TD class="RcpHeadInput" style="width:95;"><INPUT class="RcpHeadInput3" readonly dir="LTR" TYPE="text" Name="TotalAmount"  size="15" tabindex="9999"></td>
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
		function checkValidation(){
		  try{

			box=document.all.GLAccount;
			check(box);
			if (box.value==''){
				box.style.backgroundColor="red";
				alert("ÂÌç Õ”«»Ì «‰ Œ«» ‰‘œÂ");
				box.style.backgroundColor="";
				box.focus();
				return false;
			}

			tmpBox=document.getElementsByName('GLMemoDate')[0];
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
			badCode = false;
			if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
			if (src.name=='GLAccount') {
				objHTTP.open('GET','../accounting/xml_GLAccount.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape(objHTTP.responseText)
				document.all.GLAccountName.value=tmpStr;
				if (tmpStr == 'ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ')
					src.value='';
			}
			else if (src.name=='ChequeNos'){
				if (src.value=='0'){
					if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ Œÿ çﬂ —« Õ–› ﬂ‰Ìœø\n\n"))
						delRow(src.parentNode.parentNode.rowIndex);
				}
			}
		}
		//===================================================
		function mask(src){ 
			var theKey=event.keyCode;
			if (src.name=='GLAccount' && theKey==32) {
				event.keyCode=9
				document.all.tmpDlgArg.value="#"
				document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì „⁄Ì‰:"
				dialogActive=true
				window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				dialogActive=false
				if (document.all.tmpDlgTxt.value !="") {
					dialogActive=true
					window.showModalDialog('../accounting/dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					dialogActive=false
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#")
						src.value=Arguments[0];
						document.all.GLAccountName.value=Arguments[1];
					}
				}
			}
			else{
				if (theKey==13){
					return true;
				}
				else if (theKey < 48 || theKey > 57) { // 0-9 are acceptible
					return false;
				}
			}
		}
	//-->
	</SCRIPT>
<%
end if
%>

<!--#include file="tah.asp" -->
