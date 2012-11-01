<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="ورود سند"
SubmenuItem=2
if not Auth(6 , 2) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/orderEdit.asp?e=n&radif="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	.MmoTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; direction: right-to-left;}
	.MmoMainTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; background-color: #558855; text-align:right; direction: RTL;}
	.MmoMainTableTH { background-color: #C3C300;}
	.MmoMainTableTR { background-color: #CCCC88; border: 0; }
	.MmoRowInput { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.MmoRowInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.MmoHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.MmoHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.MmoHeadInput3 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #D0E0FF; text-align:right; direction: RTL;}
	.MmoGenInput { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction: LTR;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<TITLE>ورود سند</TITLE>
</HEAD>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;
var currentRow=null;
//-->
</SCRIPT>
<BODY bgcolor="#F0F0F0">
<font face="tahoma">
<%

if request("act")="submitsearch" then
	if request("CustomerNameSearchBox") <> "" then
		SA_TitleOrName=request("CustomerNameSearchBox")
		SA_Action="return true;"
		SA_SearchAgainURL="MemoInput.asp"
		SA_StepText="گام دوم : انتخاب حساب"
%>
		<FORM METHOD=POST ACTION="MemoInput.asp?act=getMemo">
		<!--#include File="include_SelectAccount.asp"-->
		</FORM>
<%
	end if
elseif request("act")="selectOrder" then
	if request("selectedCustomer") <> "" then
		SO_Customer=request("selectedCustomer")
		SO_Action="return true;"
		SO_StepText="گام سوم :"
%>
		<FORM METHOD=POST ACTION="MemoInput.asp?act=getMemo">
		<!--#include File="include_SelectOrder.asp"-->
		</FORM>
<%
	end if
elseif request("act")="getMemo" then
	customerID=request("selectedCustomer")

	mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)
	AccountNo=RS1("ID")
	customerName=RS1("AccountTitle")
	
	creationDate=shamsiToday()
	creationTime=time
'	creationTime=Hour(creationTime)&":"&Minute(creationTime)
'	if instr(creationTime,":")<3 then creationTime="0" & creationTime
'	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)

%><BR><BR><BR>
	<div align=center dir='rtl'><B>گام چهارم : ورود اطلاعات سند</B>
	</div>
<!--  ورود اطلاعات سند -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
		<table class="MmoMainTable" Cellspacing="1" Cellpadding="5" Width="500" align="center">
		<FORM METHOD=POST ACTION="MemoInput.asp?act=submitMemo" onsubmit="if (document.all.AccountTitle.value=='') return false;">
			<tr class="MmoMainTableTH">
			<TD colspan="10"><TABLE class="MmoTable" Border="0" Width="100%" Cellspacing="1" Cellpadding="0"><TR>
				<TD align="left">حساب:</TD>
				<TD align="right" width="5">
					<INPUT class="MmoGenInput" disabled TYPE="text" value="<%=customerID%>" maxlength="5" size="5" tabIndex="1">
				</TD>
				<TD align="right">
					<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><%=CustomerName%>.
				</TD>
				<TD align="left">تاريخ:</TD>
				<TD><TABLE class="MmoTable">
					<TR>
						<TD dir="LTR">
							<INPUT class="MmoGenInput" style="text-align:left;" NAME="MemoDate" TYPE="text" maxlength="10" size="10" value="<%=CreationDate%>" onblur="acceptDate(this)">
						</TD>
						<TD dir="RTL"><%=weekdayname(weekday(date))%></TD>
					</TR>
					</TABLE></TD>
				</TR></TABLE></TD>
			</tr>
			<tr class="MmoMainTableTH" height="25">
				<TD> نوع سند </TD>
				<TD> بدهكار/بستانكار </TD>
				<TD> شرح </TD>
				<TD> مقدار </TD>
			</tr>
			<tr class="MmoMainTableTR">
				<TD valign="top">
				<SELECT class="MmoRowInput" NAME="MemoType">
				<%
				mySQL="SELECT * FROM AXMemoTypes WHERE Display=1"
				Set RS1=conn.execute(mySQL)
				while not RS1.eof
				%>
					<OPTION Value="<%=RS1("ID")%>"><%=RS1("Name")%></OPTION>
				<%
				RS1.moveNext
				wend
				%>
				</SELECT>
				</TD>
				<TD valign="top">
				<SELECT class="MmoRowInput" NAME="IsCredit">
					<OPTION Value="0">بدهكار</OPTION>
					<OPTION Value="1" selected >بستانكار</OPTION>
				</SELECT>
				</TD>
				<TD valign="top"><TEXTAREA class="MmoRowInput" dir="RTL" NAME="Description" ROWS="3" COLS="30"></TEXTAREA></TD>
				<TD valign="top"><INPUT class="MmoRowInput" Dir="LTR" NAME="Amount" TYPE="text" size="15" onblur="setPrice(this);" onKeyPress="return mask(this);"></TD>
			</tr>
		</table><br>
		<TABLE class="MmoTable" align=center Border="0" Cellspacing="5" Cellpadding="1">
		<tr>
			<td align='center' bgcolor="#000000"><INPUT class="MmoGenInput" style="text-align:center" TYPE="button" value="ذخيره" onclick="submit();"></td>
			<td align='center' bgcolor="#000000"><INPUT class="MmoGenInput" style="text-align:center" TYPE="button" value="انصراف" onclick="window.location='MemoInput.asp';"></td>
		</tr>
		</TABLE>
		</FORM>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
//			document.all.CashAmount.focus();
		//-->
		</SCRIPT>

<%elseif request("act")="submitMemo" then

	ON ERROR RESUME NEXT
		MemoDate=	sqlSafe(request.form("MemoDate"))
		CustomerID=	clng(request.form("CustomerID"))
		Amount=		cdbl(text2value(request.form("Amount")))
		Description=sqlSafe(request.form("Description"))
		MemoType=	cint(request.form("MemoType"))

		IsCredit=	cbool(request.form("IsCredit"))
		if IsCredit then IsCredit=1 else IsCredit=0 end if
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("خطا!")
		end if
	ON ERROR GOTO 0

	'---- Checking wether EffectiveDate is valid in current open GL
	if (MemoDate < session("OpenGLStartDate")) OR (MemoDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?act=getMemo&selectedCustomer="& CustomerID & "&errMsg=" & Server.URLEncode("خطا!<br>تاريخ وارد شده در محدوده سال مالي جاري نيست.")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("خطا! سال مالي جاري بسته شده و شما قادر به تغيير در آن نيستيد.")
	end if 
	'----
	if (trim(Description)="") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("خطا!بدون شرح ثبت نمي‌كنيم.")
	end if 
	creationDate=	shamsiToday()

	GLAccount=		"91001"	'This must be changed... (Sales A)
	firstGLAccount=	"13003"	'This must be changed... (Business Debitors)

	mySQL="INSERT INTO ARMemo (CreatedDate, CreatedBy, Account, Type, IsCredit, Description, Amount) VALUES (N'" &_
	MemoDate & "', '"& session("ID") & "', '"& CustomerID & "', '"& MemoType & "', '"& IsCredit & "', N'"& Description & "', '"& Amount & "');SELECT @@Identity AS NewMemo"
	Set RS1 = conn.Execute(mySQL).NextRecordSet
	MemoID=RS1("NewMemo")
	RS1.close

	'**************************** Creating ARItem for Memo  ****************
	'*** Type = 3 means ARItem is a Memo
	
	mySQL="INSERT INTO ARItems (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount) VALUES ('" &_
	GLAccount & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& CustomerID & "', N'"& MemoDate & "', '"& IsCredit & "', '3', '"& MemoID & "', '"& Amount & "', N'"& creationDate & "', '"& session("ID") & "', '"& Amount & "')"	
	conn.Execute(mySQL)
	'***------------------------- Creating ARItem for Memo  ----------------
	if IsCredit then
		mySQL="UPDATE Accounts SET ARBalance = ARBalance + '"& Amount & "' WHERE (ID='"& CustomerID & "')"
	else
		mySQL="UPDATE Accounts SET ARBalance = ARBalance - '"& Amount & "' WHERE (ID='"& CustomerID & "')"
	end if
	conn.Execute(mySQL)
	conn.close
	response.redirect "AccountReport.asp?act=showMemo&sys=AR&memo=" & MemoID &"&msg=" & Server.URLEncode("اعلاميه ثبت شد.")
else%>
<!-- جستجو براي نام حساب -->
	<FORM METHOD=POST ACTION="MemoInput.asp?act=submitsearch" onsubmit="if (document.all.CustomerNameSearchBox.value=='') return false;"><BR><BR>
	<div dir='rtl'>&nbsp;<B>گام اول : جستجو براي نام حساب</B>
		<INPUT TYPE="text" NAME="CustomerNameSearchBox">&nbsp;
		<INPUT class="GenButton" TYPE="submit" value="جستجو"><br>
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
<% if request("act")="getMemo" then %>
<script language="JavaScript">
<!--
function delRow(rowNo){
	chqTable=document.getElementById("ChequeLines");
	theRow=chqTable.getElementsByTagName("tr")[rowNo];
	chqTable.removeChild(theRow);

	for (rowNo=0; rowNo < document.getElementsByName("ChequeNos").length; rowNo++){
		chqTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0].innerText= rowNo+1;
	}
}

function setPrice(src){
	if (echoNum(getNum($(src).val()))=="NaN")
		$(src).val(0);
	else
		$(src).val(echoNum(getNum($(src).val())));
}

function mask(src){ 
	var theKey=event.keyCode;
	if (theKey==13){
		return true;
	}
	else if (theKey < 48 || theKey > 57) { // 0-9 are acceptible
		return false;
	}
}
//-->
</script>
<%end if%>
