<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "çﬂ „⁄Ì‰"
SubmenuItem=6
if not Auth("A" , 6) then NotAllowdToViewThisPage()

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
	accountID = request.form("accountID")
	accountName = request.form("accountName")
	GLMemoDate = request.form("GLMemoDate")
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	creationDate =shamsiToday()
	CreationTime = currentTime10()

	if GLMemoDate="" or not(isnumeric(accountID)) then 
		call showAlert( "Œÿ«! ›—„ ﬂ«„· Å— ‰‘œÂ «” " , CONST_MSG_ERROR )
		response.end
	end if 

	for cheques = 1 to request.form("ChequeNos").count 
		Item		= request.form("Items")(cheques)
		accTitle	= request.form("accTitles")(cheques)
		ChequeNo	= request.form("ChequeNos")(cheques)
		ChequeDate	= request.form("ChequeDates")(cheques)
		Bank		= request.form("Banks")(cheques)
		shobe		= request.form("shobe")(cheques)
		jaari		= request.form("jaari")(cheques)
		Description = request.form("Description")(cheques)
		Amount		= text2value(request.form("Amounts")(cheques))

		if ChequeNo="" or ChequeDate="" or Amount="0" or not(isnumeric(amount)) then 
			call showAlert( "Œÿ«! ›—„ ﬂ«„· Å— ‰‘œÂ «” " , CONST_MSG_ERROR )
			response.end
		end if 
	next

	mySQL="SELECT ISNULL(MAX(GLDocID),0) AS LastMemo FROM GLDocs WHERE GL='"& OpenGL & "'"
	Set RS1=conn.Execute (mySQL)
	GLMemoNo = RS1("LastMemo") + 1

	mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, createdBy,  IsTemporary) VALUES ("& OpenGL& " , "& GLMemoNo & ", N'"& GLMemoDate & "' , N'"& creationDate & "', "& session("id") & ", 1)"
	conn.Execute(mySQL)

	mySQL="SELECT max(ID) as GLDoc FROM GLDocs where GLDocID="& GLMemoNo
	set RS1=conn.execute(mySQL)
	GLDoc=RS1("GLDoc")

	for cheques = 1 to request.form("ChequeNos").count 
		Item		= request.form("Items")(cheques)
		accTitle	= request.form("accTitles")(cheques)
		ChequeNo	= request.form("ChequeNos")(cheques)
		ChequeDate	= request.form("ChequeDates")(cheques)
		Bank		= request.form("Banks")(cheques)
		shobe		= request.form("shobe")(cheques)
		jaari		= request.form("jaari")(cheques)
		Description = request.form("Description")(cheques)
		Amount		= text2value(request.form("Amounts")(cheques))
		

		if jaari<> "" then jaari="Ã " & jaari
		if Bank = -1 then Bank = ""

		theDescription = "çﬂ " & ChequeNo & " „Ê—Œ " & ChequeDate & " " & Bank & " " & shobe & " " & jaari & " " & Description

		mySQL="INSERT INTO GLRows ( GLDoc, GLAccount,  Amount, Description, IsCredit, Ref1, Ref2) VALUES ( "& GLDoc & ", "& Item & ", "& Amount & ", N'"& theDescription & "', 1, N'"& ChequeNo &"', N'"& ChequeDate &"')"
		conn.Execute(mySQL)

		mySQL="INSERT INTO GLRows ( GLDoc, GLAccount,  Amount, Description, IsCredit, Ref1, Ref2) VALUES ( "& GLDoc & ", "& accountID & ", "& Amount & ", N'"& theDescription & " " & accTitle & "', 0, N'"& ChequeNo &"' , N'"& ChequeDate &"')"
		conn.Execute(mySQL)
	next

	response.redirect "../accounting/GLMemoDocShow.asp?id="&GLDoc&"&msg=çﬂ À»  ‘œ"
	'response.write replace(request.form,"&", "<br>")
	response.end
	
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------- Search Form
'-----------------------------------------------------------------------------------------------------
elseif request("act")="" then
%>
	<BR><BR>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<FORM METHOD=POST ACTION="?act=submit">
		<table class="RcpMainTable" Cellspacing="1" Cellpadding="0" Width="690" align="center">
			<tr class="RcpMainTableTH">
			<TD colspan="10"><TABLE class="RcpTable" Border="0" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL"><TR>
				<TD align="left">„⁄Ì‰:</TD>
				<TD align="right" width="5">
					<INPUT class="RcpGenInput" TYPE="text"  maxlength="5" size="5" tabIndex="1" dir="LTR" name="accountID" onKeyPress='return maskTop(this);' onBlur='checkTop(this);' >
				</TD>
				<TD align="right">
					<INPUT TYPE="text" NAME="accountName" size=40 readonly  value="<%=accountName%>" style="background-color:transparent; border:0pt">
				</TD>
				<TD align="left"> «—ÌŒ:</TD>
				<TD><TABLE class="RcpTable">
					<TR>
						<TD dir="LTR">
							<INPUT class="RcpGenInput" NAME="GLMemoDate" tabIndex="1" TYPE="text" maxlength="10" size="10" value="<%=shamsiToday()%>" onblur="acceptDate(this)">
						</TD>
						<TD dir="RTL">«„—Ê“: <font color=white><%=weekdayname(weekday(date))%> <span dir=rtl><%=shamsiToday()%></span></font></TD>
					</TR>
					</TABLE></TD>
			</TR></TABLE></TD>
			</tr>
			<tr class="RcpMainTableTR">
			<TD colspan="10"><div>
			<TABLE class="RcpTable" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<tr>
				<td class="RcpHeadInput" align='center' width="25px"> # </td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value="Õ”«»" size="6" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value="⁄‰Ê«‰" size="14" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value="‘„«—Â çﬂ" size="10" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value=" «—ÌŒ çﬂ" size="10" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="»«‰ﬂ" size="8" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="‘⁄»Â" size="8" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="Ã«—Ì" size="8" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="‘—Õ" size="16" tabindex="9999" ></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="„»·€" size="10" tabindex="9999" ></td>
			</tr>
			</TABLE></div></TD>
			</TR>
			<tr class="RcpMainTableTR">
			<TD colspan="10" dir="RTL"><div style="overflow:auto; height:250px;width:690px;">

			<TABLE class="RcpTable" Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<Tbody id="ChequeLines">
			<%		
			for i=1 to 1
			%>
			<tr bgcolor='#F0F0F0' onclick="currentRow=this.rowIndex;" >
				<td align='center' width="25px"> <%=i%> </td>
				<td dir="LTR"><INPUT class="RcpRowInput" TYPE="text" NAME="Items" maxlength="5" size="5" onKeyPress="return mask(this);" tabIndex="2"  onBlur='check(this);'></td>
				<td dir="RTL"><INPUT class="RcpRowInput" TYPE="text" NAME="accTitles" size="14" onKeyPress="return mask(this);" tabIndex="9999" readonly></td>
				<td dir="LTR"><INPUT class="RcpRowInput" TYPE="text" NAME="ChequeNos" size="10" onKeyPress="return mask(this);" tabIndex="2" ></td>
				<td dir="LTR" ><INPUT class="RcpRowInput" TYPE="text" NAME="ChequeDates" maxlength="10" size="10"  onKeyPress="return maskDate(this);" onblur="acceptDate(this)" tabIndex="2"></td>
				<td dir="RTL">
					<select name="Banks" class=inputBut style="width:60; " tabIndex="2">
					<option value="-">«‰ Œ«» ﬂ‰Ìœ</option>
					<option value="-">---------------------------------</option>
						<option value="„·Ì"> „·Ì </option>
						<option value="„· "> „·  </option>
						<option value=" Ã«— ">  Ã«—  </option>
						<option value="”ÅÂ"> ”ÅÂ </option>
						<option value="’«œ—« "> ’«œ—«  </option>
						<option value="ﬂ‘«Ê—“Ì"> ﬂ‘«Ê—“Ì </option>
						<option value="„”ﬂ‰"> „”ﬂ‰ </option>
						<option value="—›«Â"> —›«Â </option>
						<option value="”«„«‰"> ”«„«‰ </option>
						<option value="”«„«‰"> ”«„«‰ </option>
						<option value="Å«—”Ì«‰"> Å«—”Ì«‰ </option>
						<option value="«ﬁ ’«œ‰ÊÌ‰"> «ﬁ ’«œ‰ÊÌ‰ </option>
					</select>
				
				</td>
				<td dir="RTL"><INPUT class="RcpRowInput" TYPE="text" NAME="shobe" size="8" tabIndex="2"></td>
				<td dir="RTL"><INPUT class="RcpRowInput" TYPE="text" NAME="jaari" size="8" tabIndex="2"></td>
				<td dir="RTL"><INPUT class="RcpRowInput" TYPE="text" NAME="Description" size="16" tabIndex="2"></td>
				<td dir="LTR"><INPUT class="RcpRowInput" TYPE="text" NAME="Amounts" size="10" onKeyPress="return mask(this);" onBlur="setPrice(this);" tabIndex="2"></td>
			</tr>
			<%
			next
			%>
			<tr bgcolor='#F0F0F0' onclick="currentRow=this.rowIndex;" >
				<td colspan="15">
					<INPUT class="RcpGenInput" TYPE="button" value="«÷«›Â" onkeyDown="if(event.keyCode==9) return false;" tabIndex="3" onClick="addRow(this.parentNode.parentNode.rowIndex);">
				</td>
			</tr>
			</Tbody></TABLE></div>
			<div ALIGN=LEFT>
			<TABLE Border="1" Cellspacing="1" Cellpadding="0" Dir="RTL">
			<tr>
				<td  class="RcpHeadInput" align='center' > 
				</td>
				<td class="RcpHeadInput" align='center' width="25px"> </td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="Ã„⁄:" size="10" tabindex="9999"></td>
				<td class="RcpHeadInput"><INPUT class="RcpHeadInput3" readonly dir="LTR" TYPE="text" Name="TotalAmount"  size="15" tabindex="9999"></td>
			</tr>
			</TABLE></div></TD>
			</TR>
		</table><br><CENTER>
		<TABLE class="RcpTable" Border="0" Cellspacing="5" Cellpadding="1" Dir="RTL">
		<tr>
			<td align='center' bgcolor="#000000"><INPUT class="RcpGenInput" style="text-align:center" TYPE="button" value="–ŒÌ—Â" onclick="if (document.getElementsByName('TotalAmount')[0].value!='0') CheckAndSubmit();"></td>
			<td align='center' bgcolor="#000000"><INPUT class="RcpGenInput" style="text-align:center" TYPE="button" value="«‰’—«›" onclick="document.close();"></td>
		</tr>
		</TABLE></CENTER>
	</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--
document.all.accountID.focus();

var dialogActive=false;

//===================================================
function maskTop(src){ 
	var theKey=event.keyCode;

	if (theKey==13){
		event.keyCode=9
		dialogActive=true
		document.all.tmpDlgArg.value="#"
		document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì „⁄Ì‰:"
		var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		if (document.all.tmpDlgTxt.value !="") {
			var myTinyWindow = window.showModalDialog('../accounting/dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
			dialogActive=false
			if (document.all.tmpDlgArg.value!="#"){
				Arguments=document.all.tmpDlgArg.value.split("#")
				src.value=Arguments[0];
				document.all.accountName.value=Arguments[1];
			}
		}
	}
}

//===================================================
function checkTop(src){ 
	if (!dialogActive){
		badCode = false;
		if (window.XMLHttpRequest) {
			var objHTTP=new XMLHttpRequest();
		} else if (window.ActiveXObject) {
			var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
		}
		objHTTP.open('GET','../accounting/xml_GLAccount.asp?id='+src.value,false)
		objHTTP.send()
		tmpStr = unescape( objHTTP.responseText)
		document.all.accountName.value=tmpStr;
		}
}

//===================================================
function mask(src){ 
	var theKey=event.keyCode;

	rowNo=src.parentNode.parentNode.rowIndex;
	invTable=document.getElementById("ChequeLines");
	theRow=invTable.getElementsByTagName("tr")[rowNo];

	if (src.name=="Items"){
		if (theKey==13){
			event.keyCode=9
			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì œ› — ﬂ·:"
			var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('../accounting/dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
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

//===================================================
function check(src){ 
	invTable=document.getElementById("ChequeLines");
	if (src.name=="Items"){
		rowNo=src.parentNode.parentNode.rowIndex;
		rowsCount=document.getElementsByName("Items").length;
		if (!dialogActive){
			if (src.value=='0' ){
				if (rowNo == 0 && rowsCount == 1) {
					src.focus();
				}
				else{
					if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ”ÿ— —« Õ–› ﬂ‰Ìœø")){
						delRow(rowNo);
						return false;
					}
					else{
						src.focus();
					}
				}
			}
			else {
				badCode = false;
				if (window.XMLHttpRequest) {
					var objHTTP=new XMLHttpRequest();
				} else if (window.ActiveXObject) {
					var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
				}
				objHTTP.open('GET','../accounting/xml_GLAccount.asp?id='+src.value,false)
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
							return false;
						}
						else{
							src.focus();
						}
					}
				}
			}
		}
	}
}

//===================================================
function delRow(rowNo){
	chqTable=document.getElementById("ChequeLines");
	theRow=chqTable.getElementsByTagName("tr")[rowNo];
	chqTable.removeChild(theRow);

	for (rowNo=0; rowNo < document.getElementsByName("ChequeNos").length; rowNo++){
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
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="Items" maxlength="5" size="5" onKeyPress="return mask(this);" tabIndex="2"  onBlur="check(this);">'
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'RTL');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="accTitles" size="14" onKeyPress="return mask(this);" tabIndex="9999" readonly>'
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="ChequeNos" size="10" onKeyPress="return mask(this);" tabIndex="2" >'
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="ChequeDates" maxlength="10" size="10"  onKeyPress="return maskDate(this);" onblur="acceptDate(this)" tabIndex="2">'
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'RTL');
	tempTD.innerHTML=chqTable.getElementsByTagName("tr")[0].getElementsByTagName("td")[5].innerHTML
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'RTL');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="shobe" size="8" tabIndex="2">'
	newRow.appendChild(tempTD);
				
	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'RTL');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="jaari" size="8" tabIndex="2">'
	newRow.appendChild(tempTD);
				
	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'RTL');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="Description" size="16" tabIndex="2">'
	newRow.appendChild(tempTD);
				
	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML='<INPUT class="RcpRowInput" TYPE="text" NAME="Amounts" size="10" onKeyPress="return mask(this);" onBlur="setPrice(this);" tabIndex="2">'
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

//===================================================
function CheckAndSubmit()
{
document.forms[0].submit()
invTable=document.getElementById("ChequeLines");

document.all.sss.value=document.forms[0].innerHTML
//document.forms[0].submit()

/*alert(joindocument.all.Items.value)
return false
	if (document.all.Items.value=="" )
		{
		alert("›—„ Œ«·Ì «” ")
		return false
		}
	else
	{
	document.forms[0].submit()
	return true
	}
	*/
}

//-->
</SCRIPT>
<%
end if
%>

<!--#include file="tah.asp" -->
