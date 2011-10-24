<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
Response.Buffer=false
'Accounting (8)
PageTitle= "œ› — ﬂ·"
SubmenuItem=5
if not Auth(8 , "A") then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>

	.MoeenTR1 { background-color: #eeeeee; }
	.MoeenTR1 td{ width:20pt;border-left:solid 1px black; border-bottom:solid 2 black; font-size:8pt; padding:2px; text-align:center;}

</style>
<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Search Form
'------------------------------------------------------------------------------------------		+
'--------------------------------------------------------------------------------------------- Scripts
'-----------------------------------------------------------------------------------------------------

account		=	request("accountID")
if account <> "" then accountName	= request("accountName")
ToAccount	=	request("ToAccount")
if ToAccount = "" then 
	ToAccount	=	account
	ToAccountName=	accountName
else
	ToAccount = request("ToAccount")
	ToAccountName=	request("ToAccountName")
end if


FromDate	=	request("FromDate")
ToDate		=	request("ToDate")
Order		=	request("Order")

if request("ShowRemained")="on" then
	ShowRemained=true
else
	ShowRemained=false
end if


'if account		= ""	then	account = "11011"
'if accountName	= ""	then	accountName = "’‰œÊﬁ"

if FromDate		= ""	then	FromDate = shamsiToday()
if ToDate		= ""	then	ToDate = shamsiToday()
%>
<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>
<BR>
<FORM Name="SearchForm" METHOD=POST ACTION="?act=Show" onSubmit="return checkValidation();"> 
<TABLE style="border:1 solid #330066;" align=center>
<TR>
	<TD align=left>„⁄Ì‰</TD>
	<TD><INPUT  dir="LTR"  TYPE="text" NAME="accountID" maxlength="5" size="10"  value="<%=account%>" onKeyPress='return mask(this);' onBlur='check(this);document.all.ToAccount.value=this.value;check(document.all.ToAccount);'></TD>
	<TD><INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent"></TD>
</TR>
<TR>
	<TD align=left> « „⁄Ì‰</TD>
	<TD><INPUT  dir="LTR"  TYPE="text" NAME="ToAccount" maxlength="5" size="10" value="<%=ToAccount%>" onKeyPress='return mask(this);' onBlur='check(this);'></TD>
	<TD><INPUT TYPE="text" NAME="ToAccountName" size=30 readonly  value="<%=ToAccountName%>" style="background-color:transparent"></TD>
</TR>
<TR>
	<TD align=left style="cursor:hand" title="ﬂ·Ìﬂ ﬂ‰Ìœ  «  «—ÌŒ Ìﬂ —Ê“ »Â ⁄ﬁ» »—Êœ" onclick="subtractOneDay();">«“  «—ÌŒ</TD>
	<TD><INPUT dir="LTR"  TYPE="text" NAME="FromDate" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=FromDate%>"></TD>
	<TD>
		<INPUT TYPE="checkbox" NAME="ShowRemained" <%if ShowRemained then response.write "checked"%>> „«‰œÂ ﬁ»· ‰„«Ì‘ œ«œÂ ‘Êœ.
	</TD>
</TR>
<TR>
	<TD align=left style="cursor:hand" title="ﬂ·Ìﬂ ﬂ‰Ìœ  «  «—ÌŒ Ìﬂ —Ê“ »Â Ã·Ê »—Êœ" onclick="addOneDay();"> «  «—ÌŒ </TD>
	<TD><INPUT dir="LTR" TYPE="text" NAME="ToDate" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=ToDate%>"></TD>
	<TD align=center>
		<INPUT TYPE="submit" NAME="submit" value="„‘«ÂœÂ">
		<INPUT TYPE="hidden" Name="Order" Value="">
	</TD>
</TR>
</TABLE>
</FORM>

<SCRIPT LANGUAGE="JavaScript">
<!--
var dialogActive=false;
function mask(src){ 
	var theKey=event.keyCode;

	if (theKey==13 || theKey==32){
		event.keyCode=9
		dialogActive=true
		document.all.tmpDlgArg.value="#"
		document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì „⁄Ì‰:"
		window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		dialogActive=false
		if (document.all.tmpDlgTxt.value !="") {
			window.showModalDialog('dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
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
		if (window.XMLHttpRequest) {
			var objHTTP=new XMLHttpRequest();
		} else if (window.ActiveXObject) {
			var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
		}
		objHTTP.open('GET','xml_GLAccount.asp?id='+src.value,false);
		objHTTP.send();
		tmpStr = unescape( objHTTP.responseText);
		if (tmpStr == 'ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ')
			src.value='';
		if (src.name=='accountID'){
			document.all.accountName.value=tmpStr;
		}
		else if (src.name=='ToAccount'){
			document.all.ToAccountName.value=tmpStr;
		}
	}
}

function checkValidation(){
  try{
	box=document.all.accountID;
	check(box);
	if (box.value==''){
		box.style.backgroundColor="red";
		alert("ÂÌç Õ”«»Ì «‰ Œ«» ‰‘œÂ");
		box.style.backgroundColor="";
		box.focus();
		return false;
	}
	box=document.all.FromDate;
	if (box.value!='' && box.value!=null){
		if (!acceptDate(box))
			return false;
	}
	box=document.all.ToDate;
	if (box.value!='' && box.value!=null){
		if (!acceptDate(box))
			return false;
	}
  }catch(e){
		alert("Œÿ«Ì €Ì— „‰ Ÿ—Â");
		return false;
  }
}

function addOneDay(){
  try{
	box=document.all.FromDate;
	if (box.value!='' && box.value!=null){
		if (acceptDate(box)){
			year=parseInt(box.value.substring(0,4));
			month=parseInt(box.value.substr(5,2).replace(/^0/g,""));
			day=parseInt(box.value.substr(8,2).replace(/^0/g,""));
			day++;
			if (day>31){
				day=1;
				month++;
			}
			if (month>12){
				month=1;
				year++;
			}
			box.value=''+year+'/'+month+'/'+day;
			acceptDate(box);
		}
	}
	box=document.all.ToDate;
	if (box.value!='' && box.value!=null){
		if (acceptDate(box)){
			year=parseInt(box.value.substring(0,4));
			month=parseInt(box.value.substr(5,2).replace(/^0/g,""));
			day=parseInt(box.value.substr(8,2).replace(/^0/g,""));
			day++;
			if (day>31){
				day=1;
				month++;
			}
			if (month>12){
				month=1;
				year++;
			}
			box.value=''+year+'/'+month+'/'+day;
			acceptDate(box);
		}
	document.SearchForm.submit.click();
	}
  }catch(e){
		alert("Œÿ«Ì €Ì— „‰ Ÿ—Â");
		return false;
  }
}

function subtractOneDay(){
  try{
	box=document.all.FromDate;
	if (box.value!='' && box.value!=null){
		if (acceptDate(box)){
			year=parseInt(box.value.substring(0,4));
			month=parseInt(box.value.substr(5,2).replace(/^0/g,""));
			day=parseInt(box.value.substr(8,2).replace(/^0/g,""));
			day--;
			if (day<1){
				day=31;
				month--;
			}
			if (month<1){
				month=12;
				year--;
			}
			box.value=''+year+'/'+month+'/'+day;
			acceptDate(box);
		}
	}
	box=document.all.ToDate;
	if (box.value!='' && box.value!=null){
		if (acceptDate(box)){
			year=parseInt(box.value.substring(0,4));
			month=parseInt(box.value.substr(5,2).replace(/^0/g,""));
			day=parseInt(box.value.substr(8,2).replace(/^0/g,""));
			day--;
			if (day<1){
				day=31;
				month--;
			}
			if (month<1){
				month=12;
				year--;
			}
			box.value=''+year+'/'+month+'/'+day;
			acceptDate(box);
		}
	document.SearchForm.submit.click();
	}
  }catch(e){
		alert("Œÿ«Ì €Ì— „‰ Ÿ—Â");
		return false;
  }
}
//-->
</SCRIPT>

<SCRIPT LANGUAGE="JavaScript">
<!--
function sortSubmit(num){
	document.SearchForm.Order.value=num;
	document.SearchForm.submit.click();
}
//-->
</SCRIPT>

<%
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Search Action
'-----------------------------------------------------------------------------------------------------

if request("act")="Show" then

' Major Changes By kid 830820

	Order=request("Order")
	select case Order
	case "1":
		order=", GLDocID "
	case "-1":
		order=", GLDocID DESC"
	case "2":
		order=", GLDocDate"
	case "-2":
		order=", GLDocDate DESC"
	case else:
		order=", GLDocDate, IsCredit DESC, Amount"
	end select

'	if ShowRemained then
		mySQL="SELECT GLAccounts.ID, GLAccounts.Name AS AccountName, GLAccountGroups.ID AS GroupID, GLAccountGroups.Name AS GroupName,  GLAccountSuperGroups.ID AS SuperGroupID, GLAccountSuperGroups.Name AS SuperGroupName, GLs.ID AS GLID, GLs.Name AS GLname FROM GLs INNER JOIN GLAccountSuperGroups ON GLs.ID = GLAccountSuperGroups.GL INNER JOIN GLAccountGroups ON GLs.ID = GLAccountGroups.GL AND GLAccountSuperGroups.ID = GLAccountGroups.GLSuperGroup INNER JOIN GLAccounts ON GLs.ID = GLAccounts.GL AND GLAccountGroups.ID = GLAccounts.GLGroup WHERE (GLs.ID = "& OpenGL & ") AND (GLAccounts.ID >= "& account & ") AND (GLAccounts.ID <= "& ToAccount & ") ORDER BY GLAccounts.ID"
'	else
		'mySQL="SELECT GLAccounts.ID, GLAccounts.Name AS AccountName, GLAccountGroups.ID AS GroupID, GLAccountGroups.Name AS GroupName,  GLAccountSuperGroups.ID AS SuperGroupID, GLAccountSuperGroups.Name AS SuperGroupName, GLs.ID AS GLID, GLs.Name AS GLname FROM GLs INNER JOIN GLAccountSuperGroups ON GLs.ID = GLAccountSuperGroups.GL INNER JOIN GLAccountGroups ON GLs.ID = GLAccountGroups.GL AND GLAccountSuperGroups.ID = GLAccountGroups.GLSuperGroup INNER JOIN GLAccounts ON GLs.ID = GLAccounts.GL AND GLAccountGroups.ID = GLAccounts.GLGroup INNER JOIN (SELECT DISTINCT GLAccount AS ID FROM EffectiveGLRows WHERE (GL = "& OpenGL & ") AND (GLAccount >= "& account & ") AND (GLAccount <= "& ToAccount & ") AND (GLDocDate >= N'"& FromDate & "') AND (GLDocDate <= N'"& ToDate & "')) UsedAccs ON UsedAccs.ID = GLAccounts.ID WHERE (GLs.ID = "& OpenGL & ") ORDER BY GLAccounts.ID"
	'end if

	set rsGLAccs=Conn.Execute (mySQL)

	Do While not rsGLAccs.EOF 

		GLAccount=rsGLAccs("ID")

		AccountInfoParams = "&DateFrom=" & FromDate & "&DateTo=" & ToDate
		if ShowRemained then AccountInfoParams = AccountInfoParams & "&ShowRemained=on"
%>
		<TABLE dir=rtl align=center width=640 cellspacing=0 cellpadding=0 style="border:2 solid #330066;">
		<TR bgcolor="#CCCCEE" height="30">
			<TD colspan=7>
				<A HREF="AccountInfo.asp?OpenGL=<%=rsGLAccs("GLID")&AccountInfoParams%>" Target="_blank"><%=rsGLAccs("GLname")%></A>
				> <A HREF="AccountInfo.asp?act=groups&id=<%=rsGLAccs("SuperGroupID")&AccountInfoParams%>" Target="_blank"><%=rsGLAccs("SuperGroupName")%></A>
				> <A HREF="AccountInfo.asp?act=account&id=<%=rsGLAccs("GroupID")&AccountInfoParams%>" Target="_blank"><%=rsGLAccs("GroupName")%></A>
				> <%=rsGLAccs("AccountName")%>
				[<%=GLAccount%>]
			</TD>
		</TR>
		<TR bgcolor="black" height="2">
			<TD colspan="7" style="padding:0;"></TD>
		</TR>
		<TR class="MoeenTR1">
				<TD style="width:60px;"><A HREF="javaScript:sortSubmit(1);">v</A>&nbsp;”‰œ&nbsp;<A HREF="javaScript:sortSubmit(-1);">^</A></TD>
				<TD style="width:80px;"><A HREF="javaScript:sortSubmit(2);">v</A>&nbsp; «—ÌŒ ”‰œ&nbsp;<A HREF="javaScript:sortSubmit(-2);">^</A></TD>
				<TD style="width:200px;">‘—Õ</TD>
				<TD style="width:100px;">»œÂﬂ«—</TD>
				<TD style="width:100px;">»” «‰ﬂ«—</TD>
				<TD style="width:10px;">-</TD>
				<TD style="width:100px;border-left:0;">„«‰œÂ</TD>
		</TR>

<%		
		if ShowRemained then

			mySQL="SELECT SUM(IsCredit * Amount) AS totalCredit, SUM(- ((IsCredit - 1) * Amount)) AS totalDebit FROM EffectiveGLRows WHERE (GL = "& OpenGL & ") AND (GLAccount = "& GLAccount & ") AND (GLDocDate < N'"& FromDate & "') GROUP BY GLAccount"

			set RSS=Conn.Execute (mySQL)

			if RSS.eof then
				debit = 0
				credit = 0
			else
				debit = cdbl(RSS("totalDebit"))
				credit = cdbl(RSS("totalCredit"))
			end if

			RSS.close

			remainedAmount = debit - credit
			SumCredit = credit 
			SumDebit =	debit 

%>
			<TR bgcolor="#DDDDDD" height=30>
				<TD colspan=3 align=left style="border-left:solid 1px black"><INPUT TYPE="text" style="text-align:left;" value="„«‰œÂ ﬁ»· «“  «—ÌŒ <%=replace(FromDate,"/",".")%>" style="width=200pt; border:solid 0pt; font-size:8pt; background-color:transparent"></TD>
				<TD dir=ltr align=right style="border-left:solid 1px black"><%=Separate(debit)%>&nbsp;</TD>
				<TD dir=ltr align=right style="border-left:solid 1px black"><%=Separate(credit)%>&nbsp;</TD>
				<TD dir=ltr align=center style="border-left:solid 1px black"><% if remainedAmount > 0 then %>œ<% else %>”<% end if %></TD>
				<TD dir=ltr align=right><%=Separate(remainedAmount)%> &nbsp;</TD>
			</TR>
			<TR bgcolor="black" height="2">
				<TD colspan="7" style="padding:0;"></TD>
			</TR>
<% 
		else
			remainedAmount = 0
			SumCredit = 0 
			SumDebit =	0
		end if

		mySQL="SELECT * FROM EffectiveGLRows WHERE (GL = "& OpenGL & ") AND (GLAccount = "& GLAccount & ") AND (GLDocDate >= N'"& FromDate & "') AND (GLDocDate <= N'"& ToDate & "') ORDER BY GLAccount" & order
		set RSS=Conn.Execute (mySQL)	

		tmpCounter=0

		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
			
			debit = 0
			credit = 0
			if RSS("IsCredit") then
				credit = cdbl(RSS("Amount"))
			else
				debit = cdbl(RSS("Amount"))
			end if
			remainedAmount = remainedAmount - credit + debit

			SumCredit = SumCredit + credit 
			SumDebit =	SumDebit + debit 

%>
			<TR bgcolor="<%=tmpColor%>" height=30>
				<TD style="border-left:solid 1px black"><A HREF="javascript:void(0);" onclick="window.open('GLMemoDocShow.asp?id=<%=RSS("GLDoc")%>')"><%=RSS("GLDocID")%></A></TD>
				<TD  align=center dir=ltr style="border-left:solid 1px black"><%=RSS("GLDocDate")%></TD>
				<TD  width=100 style="border-left:solid 1px black">
					<INPUt size="70" TYPE="text" value="<%=RSS("Description")%>" style="width=200pt; border:solid 0pt; font-size:8pt; background-color:transparent">
				</TD>
				<TD dir=ltr align=right style="border-left:solid 1px black"><% if debit<>"0" then %> <%=Separate(debit)%><% end if %> &nbsp;</TD>
				<TD dir=ltr align=right style="border-left:solid 1px black"><% if credit<>"0" then %> <%=Separate(credit)%><% end if %> &nbsp;</TD>
				<TD dir=ltr align=center style="border-left:solid 1px black"><% if remainedAmount > 0 then %>œ<% else %>”<% end if %></TD>
				<TD dir=ltr align=right><%=Separate(remainedAmount)%> &nbsp;</TD>
			</TR>
				  
<% 
			RSS.moveNext
		Loop
		RSS.close

		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 
		%>
		<TR bgcolor="black" height="2">
			<TD colspan="7" style="padding:0;"></TD>
		</TR>
		<TR bgcolor="<%=tmpColor%>" height=30>
			<TD colspan=3 align=left style="border-left:solid 1px black"><INPUT TYPE="text" style="text-align:left;" value="Ã„⁄ ê—œ‘ <%=GLAccount%>  «  «—ÌŒ <%=replace(ToDate,"/",".")%>" style="width=200pt; border:solid 0pt; font-size:8pt; background-color:transparent"></TD>
			<TD dir=ltr align=right style="border-left:solid 1px black"><% if SumDebit<>"0" then response.write Separate(SumDebit)%> &nbsp;</TD>
			<TD dir=ltr align=right style="border-left:solid 1px black"><% if SumCredit<>"0" then response.write Separate(SumCredit)%> &nbsp;</TD>
			<TD dir=ltr align=center style="border-left:solid 1px black"><% if remainedAmount > 0 then %>œ<% else %>”<% end if %></TD>
			<TD dir=ltr align=right><%=Separate(remainedAmount)%> &nbsp;</TD>
		</TR>
		</TABLE><br><br>
<% 

		rsGLAccs.moveNext
	Loop
	rsGLAccs.close
	Set rsGLAccs = Nothing

end if %>
<!--#include file="tah.asp" -->
