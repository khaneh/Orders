<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle= "ê“«—‘  ›’Ì·Ì"
SubmenuItem=6
if not Auth(8 , "B") then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>

	.TafsilTR1 { background-color: #eeeeee; }
	.TafsilTR1 td{ width:20pt;border-left:solid 1px black; border-bottom:solid 2 black; font-size:8pt; padding:2px; text-align:center;}

</style>
<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Search Form
'-----------------------------------------------------------------------------------------------------

account		=	sqlSafe(request("accountID"))
accountName	=	sqlSafe(request("accountName"))
FromDate	=	sqlSafe(request("FromDate"))
ToDate		=	sqlSafe(request("ToDate"))
moeenFrom	=	sqlSafe(request("moeenFrom"))
moeenTo		=	sqlSafe(request("moeenTo"))
Order		=	sqlSafe(request("Order"))

if request("ShowRemained")="on" then
	ShowRemained=true
else
	ShowRemained=false
end if

if account		= "" OR NOT IsNumeric(account)	then	account = "0"
if FromDate		= ""	then	FromDate = fiscalYear & "/01/01"
if ToDate		= ""	then	ToDate = shamsiToday()
if moeenFrom	= ""	then	moeenFrom = "00000"
if moeenTo		= ""	then	moeenTo = "99999"

%>
<BR><BR>
<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>

<FORM Name="SearchForm" METHOD=POST ACTION="?act=Show">
<TABLE style="border:1 solid #330066;" align=center cellpadding=3>
<TR>
	<TD align=left> ›’Ì·Ì </TD>
	<TD><INPUT dir="LTR" TYPE="text" NAME="accountID" maxlength="10" size="10"  value="<%=account%>" onKeyPress='return mask(this);' onBlur='check(this);'></TD>
	<TD colspan=2><INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent;"></TD>
</TR>
<TR>
	<TD align=left>«“  «—ÌŒ </TD>
	<TD><INPUT dir="LTR"  TYPE="text" NAME="FromDate" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=FromDate%>"></TD>
	<TD align=left> «  «—ÌŒ </TD>
	<TD><INPUT dir="LTR"  TYPE="text" NAME="ToDate" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=ToDate%>"></TD>
</TR>
<TR>
	<TD align=left>«“ „⁄Ì‰</TD>
	<TD><INPUT  dir="LTR"  TYPE="text" NAME="moeenFrom" maxlength="5" size="10" value="<%=moeenFrom%>"></TD>
	<TD align=left> « „⁄Ì‰</TD>
	<TD><INPUT  dir="LTR"  TYPE="text" NAME="moeenTo" maxlength="5" size="10" value="<%=moeenTo%>"></TD>
</TR>
<TR bgcolor="black" height="2">
	<TD colspan="4" style="padding:0;"></TD>
</TR>
<TR>
	<TD colspan=4><INPUT TYPE="checkbox" NAME="ShowRemained" <%if ShowRemained then response.write "checked"%>> „«‰œÂ ﬁ»· ‰„«Ì‘ œ«œÂ ‘Êœ.</TD>
</TR>
<TR>
	<TD colspan=4 align=center>
		<INPUT TYPE="submit" NAME="submit" value="„‘«ÂœÂ">
		<INPUT TYPE="hidden" Name="Order" Value="<%=request("Order")%>">
	</TD>
</TR>
</TABLE>
</FORM>

<%
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Search Action
'-----------------------------------------------------------------------------------------------------

if request("act")="Show" then
	Order=request("Order")
	select case Order
	case "1":
		order=" ORDER BY GLDocID "
	case "-1":
		order=" ORDER BY GLDocID DESC"
	case "2":
		order=" ORDER BY GLDocDate"
	case "-2":
		order=" ORDER BY GLDocDate DESC"
	case else:
		order=" ORDER BY GLAccount, GLDocDate"
	end select


	'mySQL="SELECT GLAccounts.Name AS Name, GLDocs.ID AS GLDoc, GLDocs.GLDocID, GLDocs.GLDocDate, GLRows.ID AS GLRowsID, GLRows.GLAccount AS GLAccount, GLRows.Amount, GLRows.IsCredit, GLRows.Description FROM GLRows INNER JOIN GLDocs ON GLDocs.ID = GLRows.GLDoc INNER JOIN GLAccounts ON GLRows.GLAccount = GLAccounts.ID WHERE ((GLDocs.IsTemporary=1 OR GLDocs.IsChecked=1 OR GLDocs.IsFinalized=1) AND GLDocs.deleted=0 AND GLRows.deleted=0 AND GLDocs.IsRemoved=0) AND ((GLAccounts.GL = "& OpenGL & ") AND (GLDocs.GL = "& OpenGL & ") AND (GLRows.GLAccount >= "& moeenFrom & ") AND (GLRows.GLAccount <= "& moeenTo & ") AND (GLRows.Tafsil = "& account & ") AND (GLDocs.GLDocDate <= N'"& ToDate & "') AND (GLDocs.GLDocDate >= N'"& FromDate & "')) "& Order
	'changed by kid 830809
	if ShowRemained then
		mySQL="SELECT DRV.GLAcc AS GLAccount, GLAccounts.Name, DRV.GLDoc, DRV.GLDocID, DRV.GLDocDate, DRV.Amount, DRV.IsCredit, DRV.Description, DRV.remainedCredit, DRV.remainedDebit FROM (SELECT ISNULL(Remains.GLAccount, EffGLRows.GLAccount) AS GLAcc, ISNULL(Remains.remCred, 0) AS remainedCredit, ISNULL(Remains.remDeb, 0) AS remainedDebit, EffGLRows.* FROM (SELECT isnull(SUM(IsCredit * Amount), 0) AS remCred, isnull(SUM(- ((IsCredit - 1) * Amount)), 0) AS remDeb, GLAccount FROM EffectiveGLRows WHERE (GL = "& OpenGL & ") AND (GLDocDate < N'"& FromDate & "') AND (Tafsil = "& account & ") GROUP BY GLAccount) Remains FULL OUTER JOIN (SELECT GLDoc, GLDocID, GLDocDate, Amount, IsCredit, Description, GLAccount FROM EffectiveGLRows WHERE (Tafsil = '"& account & "') AND (GL = "& OpenGL & ") AND (GLDocDate >= N'"& FromDate & "') AND (GLDocDate <= N'"& ToDate & "')) EffGLRows ON  Remains.GLAccount = EffGLRows.GLAccount) DRV INNER JOIN GLAccounts ON DRV.GLAcc = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLAccounts.ID >= "& moeenFrom & ") AND (GLAccounts.ID <= "& moeenTo & ")" & Order 
	else
		mySQL="SELECT EffGLRows.*, GLAccounts.Name FROM (SELECT GLDoc, GLDocID, GLDocDate, Amount, IsCredit, Description, GLAccount FROM EffectiveGLRows WHERE (Tafsil = '"& account & "') AND (GL = "& OpenGL & ") AND (GLDocDate >= N'"& FromDate & "') AND (GLDocDate <= N'"& ToDate & "')) EffGLRows INNER JOIN GLAccounts ON EffGLRows.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLAccounts.ID >= "& moeenFrom & ") AND (GLAccounts.ID <= "& moeenTo & ")" & Order 
	end if

	set RSS=Conn.Execute (mySQL)

	if RSS.eof then
		response.write "<br>" 
		call showAlert ("ÅÌœ« ‰‘œ.", CONST_MSG_ERROR) 
		response.end
	end if
%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function sortSubmit(num){
		document.SearchForm.Order.value=num;
		document.SearchForm.submit.click();
	}
	//-->
	</SCRIPT>
	<TABLE dir=rtl align=center width=640 cellspacing=0 cellpadding=2 style="border:1 solid black;">
	<%
	tmpCounter=0
	remainedAmount = 0
	totalRemainedAmount = 0
	LastGLAccount= ""

	remainedAmount = 0
	SumCredit = 0 
	SumDebit =	0

	tmpCounter =0
	remainedAmount = 0
	Do while not RSS.eof
		GLAccount=RSS("GLAccount")

		if LastGLAccount <> GLAccount then
			if tmpCounter > 0 then
				remainedAmount = SumDebit - SumCredit
%>
				<TR>
					<TD colspan=3 align=left style="border: 1 solid black; border-right:none;">Ã„⁄ ê—œ‘ „⁄Ì‰ <%=LastGLAccount%>&nbsp;</TD>
					<TD dir=ltr align=right style="border: 1 solid black; border-right:none;"><%=Separate(SumDebit)%>&nbsp;</TD> 
					<TD dir=ltr align=right style="border: 1 solid black; border-right:none;"><%=Separate(SumCredit)%>&nbsp;</TD>
					<TD align=right style="border: 1 solid black; border-right:none;">&nbsp;<% if remainedAmount > 0 then response.write "œ" else if remainedAmount < 0 then response.write "”"%></TD>
					<TD align=right dir=ltr style="border-top: 1 solid black;border-bottom: 1 solid black;" ><%=Separate(abs(remainedAmount))%></TD>
				</TR>
<% 
			end if
%>
			<TR>
			<td colspan=7 style="height:25px;background-color: #88AADD;">„⁄Ì‰ <%=GLAccount%> - <%=RSS("Name")%></td>
			</TR>
			<TR class="TafsilTR1">
				<TD style="width:60px;"><A HREF="javaScript:sortSubmit(1);">v</A>&nbsp;”‰œ&nbsp;<A HREF="javaScript:sortSubmit(-1);">^</A></TD>
				<TD style="width:80px;"><A HREF="javaScript:sortSubmit(2);">v</A>&nbsp; «—ÌŒ ”‰œ&nbsp;<A HREF="javaScript:sortSubmit(-2);">^</A></TD>
				<TD style="width:200px;">‘—Õ</TD>
				<TD style="width:100px;">»œÂﬂ«—</TD>
				<TD style="width:100px;">»” «‰ﬂ«—</TD>
				<TD colspan=2 style="width:100px;border-left:0;">„«‰œÂ</TD>
			</TR>
<%
			tmpCounter =1
			remainedAmount = 0
			SumCredit = 0 
			SumDebit =	0
		else
			tmpCounter = tmpCounter + 1
		end if

		if ShowRemained AND LastGLAccount <> GLAccount then
			credit = cdbl(RSS("remainedCredit"))
			debit =	 cdbl(RSS("remainedDebit"))

			remainedAmount =  debit - credit
			totalRemainedAmount = totalRemainedAmount + debit - credit
			SumCredit = SumCredit + credit
			SumDebit = SumDebit + debit
			totalCredit=totalCredit + credit
			totalDebit=totalDebit + debit
%>
			<TR bgcolor="#FFFFFF">
				<TD style="border-left:solid 1px black">&nbsp;</TD>
				<TD style="border-left:solid 1px black">&nbsp;</TD>
				<TD style="border-left:solid 1px black"><INPUT TYPE="text" value="<%="„«‰œÂ ﬁ»· «“ " & replace(FromDate,"/",".")%>" style="width=200pt; border:solid 0pt; font-size:8pt; background-color:transparent"></TD>
				<TD dir=ltr align=right style="border-left:solid 1px black"><%=Separate(debit)%>&nbsp;</TD>
				<TD dir=ltr align=right style="border-left:solid 1px black"><%=Separate(credit)%>&nbsp;</TD>
				<TD align=right style="border-left:solid 1px black">&nbsp;<% if remainedAmount > 0 then response.write "œ" else if remainedAmount < 0 then response.write "”"%></TD>
				<TD dir=ltr align=right><%=Separate(abs(remainedAmount))%></TD>
			</TR>
<%
			tmpCounter = tmpCounter + 1
		end if
		
		LastGLAccount = GLAccount

		if NOT isnull(RSS("Amount")) then
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
			Else
				tmpColor="#DDDDDD"
			End if 
			
			if RSS("IsCredit") then
				credit = cdbl(RSS("Amount"))
				debit = 0
			else
				credit = 0
				debit = cdbl(RSS("Amount"))
			end if

			remainedAmount = remainedAmount + debit - credit 
			totalRemainedAmount = totalRemainedAmount + debit - credit
			SumCredit = SumCredit + credit
			SumDebit = SumDebit + debit
			totalCredit=totalCredit + credit
			totalDebit=totalDebit + debit

%>
			<TR bgcolor="<%=tmpColor%>" >
				<TD style="border-left:solid 1px black"><A HREF="GLMemoDocShow.asp?id=<%=RSS("GLDoc")%>" target="_blank"><%=RSS("GLDocID")%></A></TD>
				<TD  align=center dir=ltr style="border-left:solid 1px black"><%=RSS("GLDocDate")%></TD>
				<TD  width=100 style="border-left:solid 1px black"><INPUT size="60" TYPE="text" value="<%=RSS("Description")%>" style="width=200pt; border:solid 0pt; font-size:8pt; background-color:transparent"></TD>
				<TD dir=ltr align=right style="border-left:solid 1px black"><% if debit<>"0" then %> <%=Separate(debit)%><% end if %>&nbsp;</TD>
				<TD dir=ltr align=right style="border-left:solid 1px black"><% if credit<>"0" then %> <%=Separate(credit)%><% end if %>&nbsp;</TD>
				<TD align=right style="border-left:solid 1px black">&nbsp;<% if remainedAmount > 0 then response.write "œ" else if remainedAmount < 0 then response.write "”"%></TD>
				<TD dir=ltr align=right><%=Separate(abs(remainedAmount))%></TD>
			</TR>
				  
<% 
		end if
		RSS.moveNext
	Loop
	remainedAmount = SumDebit - SumCredit
%>
	<TR>
		<TD colspan=3 align=left style="border: 1 solid black; border-right:none;">Ã„⁄ ê—œ‘ „⁄Ì‰ <%=LastGLAccount%>&nbsp;</TD>
		<TD dir=ltr align=right style="border: 1 solid black; border-right:none;"><%=Separate(SumDebit)%>&nbsp;</TD> 
		<TD dir=ltr align=right style="border: 1 solid black; border-right:none;"><%=Separate(SumCredit)%>&nbsp;</TD>
		<TD align=right style="border: 1 solid black; border-right:none;">&nbsp;<% if remainedAmount > 0 then response.write "œ" else if remainedAmount < 0 then response.write "”"%></TD>
		<TD align=right dir=ltr style="border-top: 1 solid black;border-bottom: 1 solid black;" ><%=Separate(abs(remainedAmount))%></TD>
	</TR>
	<TR bgcolor="black" height=1>
		<TD colspan=7></TD>
	</TR>
	<TR bgcolor="#88AADD">
		<TD colspan=3 align=left style="border-left:solid 1px black;">Ã„⁄ ê—œ‘  ›’Ì·Ì <%=account%>&nbsp;</TD>
		<TD dir=ltr align=right style="border-left:solid 1px black;"><%=Separate(totalDebit)%>&nbsp;</TD> 
		<TD dir=ltr align=right style="border-left:solid 1px black;"><%=Separate(totalCredit)%>&nbsp;</TD>
		<TD align=right style="border-left:solid 1px black;">&nbsp;<% if totalRemainedAmount > 0 then response.write "œ" else if totalRemainedAmount < 0 then response.write "”"%></TD>
		<TD align=right dir=ltr><%=Separate(abs(totalRemainedAmount))%></TD>
	</TR>

	</TABLE><br>
<% end if %>
<%

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------- Scripts
'-----------------------------------------------------------------------------------------------------

%>
<SCRIPT LANGUAGE="JavaScript">
<!--

var dialogActive=false;

function mask(src){ 
	var theKey=event.keyCode;

	if (theKey==13){
		event.keyCode=9
		dialogActive=true
		document.all.tmpDlgArg.value="#"
		document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì  ›’Ì·Ì:"
		var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		if (document.all.tmpDlgTxt.value !="") {
			var myTinyWindow = window.showModalDialog('../ar/dialog_selectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
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
		objHTTP.open('GET','xml_CustomerAccount.asp?id='+src.value,false);
		objHTTP.send();
		tmpStr = unescape(objHTTP.responseText);
		//document.all['A1'].innerText= objHTTP.status
		//document.all['A2'].innerText= objHTTP.statusText
		//document.all['A3'].innerText= objHTTP.responseText
		document.all.accountName.value=tmpStr;
		}
}

// On page load :
if (document.all.accountID) 
	check(document.all.accountID);

//-->
</SCRIPT>
<!--#include file="tah.asp" -->
