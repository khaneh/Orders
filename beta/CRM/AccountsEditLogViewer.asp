<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CRM (1)
PageTitle="مشتريان"
SubmenuItem=1
if not Auth(1 , 1) then NotAllowdToViewThisPage()
%>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<!--#include file="../config.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable1 {font-family:tahoma; width:100%; height:100%; border:2px solid black; direction: RTL; background-color:#CCCCCC; }
	.CustTable1 a {text-decoration:none;color:#000088}
	.CustTable1 a:hover {text-decoration:underline;}
	.CustTable2 {font-family:tahoma; border:none; direction: RTL; width:100%; height:100%;}
	.CustTable3 {font-family:tahoma; border:1px solid black; direction: RTL; background-color:black; }
	.CustTable3 td {padding:5px;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; height:100%; background-color:#C3DBEB;}
	.CustTableMenu {width:100%; border:none; direction: RTL;}
	.CustTableMenu td {border-bottom:1 solid black; height:25px; padding:5px;}
	.CustTableMenuSelected {background-color:#C3DBEB;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
</STYLE>
<%
AutoKey = clng(request("AutoKey"))

set RS1=Conn.execute("SELECT AccountsEditLog.*, Users.RealName AS CSRName FROM AccountsEditLog LEFT OUTER JOIN  Users ON AccountsEditLog.CSR = Users.ID where AccountsEditLog.AutoKey = " & AutoKey )



%>
<TABLE topmargin=0 leftmargin=0 width="100%">
<TR>
	<TD>
<table class="CustTable3" cellspacing='1' cellspacing='1'  width="100%">
<tr>
	<td class="CusTD1" width="80">عنوان حساب : </td>
	<td class="CusTD2"><span dir="RTL"><%=RS1("AccountTitle")%></span></td>
</tr> 
<tr>
	<td class="CusTD1">شماره حساب :</td>
	<td class="CusTD2"><%=RS1("ID")%></td>
</tr>
<tr>
	<td class="CusTD1">نام شركت : </td>
<%		if RS1("IsPersonal") then %>
		<td class="CusTD2"><font color="Gray">حساب شخصي است</font></td>
<%		else%>
		<td class="CusTD2"><%=RS1("CompanyName")%></td>
<%		end if%>
</tr>
<tr>
	<td class="CusTD1">تاريخ ايجاد : </td>
	<td class="CusTD2"><%=RS1("CreatedDate")%></td>
</tr>
<tr>
	<td class="CusTD1">مسوول پيگيري : </td>
	<td class="CusTD2"><%=RS1("CSRName")%></td>
</tr>
<tr>
	<td class="CusTD1">مانده فروش : </td>
	<td class="CusTD2"><%=Separate(RS1("ARBalance"))%></td>
</tr>
<tr>
	<td class="CusTD1">مانده خريد: </td>
	<td class="CusTD2"><%=Separate(RS1("APBalance"))%></td>
</tr>
<%if Auth("B" , 0) then %>
<tr>
	<td class="CusTD1">مانده ساير:</td>
	<td class="CusTD2"><%=Separate(RS1("AOBalance"))%></td>
</tr>
<%end if %>
<tr>
	<td class="CusTD1">سقف اعتبار : </td>
	<td class="CusTD2"><%=RS1("CreditLimit")%></td>
</tr>
<tr>
	<td class="CusTD1" valign="top">رابط اصلي : </td>
	<td class="CusTD2"><TABLE class="CustContactTable">
		<TR>
			<TD><%=RS1("Dear1")%> - <%=RS1("FirstName1")%> - <%=RS1("LastName1")%></TD>
			<TD width="140" style="border-right:1 solid black;">تلفن : <%=RS1("Tel1")%></TD>
		</TR>
		<TR>
			<TD rowspan="2" valign="top">آدرس : <%=RS1("City1")%> - <%=RS1("Address1")%></TD>
			<TD style="border-right:1 solid black;">فاكس : <%=RS1("Fax1")%></TD>
		</TR>
		<TR>
			<TD style="border-right:1 solid black;">ايميل : <%=RS1("Email1")%></TD>
		</TR>
		<TR>
			<TD>کدپستي: <%=RS1("PostCode1")%></TD>
			<TD style="border-right:1 solid black;">موبايل: <%=RS1("Mobile1")%></TD>
		</TR>
		</TABLE>
	</td>
</tr>
<tr>
	<td class="CusTD1" valign="top">رابط دوم : </td>
	<td class="CusTD2"><TABLE class="CustContactTable">
		<TR>
			<TD><%=RS1("Dear2")%> - <%=RS1("FirstName2")%> - <%=RS1("LastName2")%></TD>
			<TD style="border-right:1 solid black;" width="140">تلفن : <%=RS1("Tel2")%></TD>
		</TR>
		<TR>
			<TD rowspan="2" valign="top">آدرس : <%=RS1("City2")%> - <%=RS1("Address2")%></TD>
			<TD style="border-right:1 solid black;">فاكس : <%=RS1("Fax2")%></TD>
		</TR>
		<TR>
			<TD style="border-right:1 solid black;">ايميل : <%=RS1("Email2")%></TD>
		</TR>
		<TR>
			<TD>کدپستي: <%=RS1("PostCode2")%></TD>
			<TD style="border-right:1 solid black;">موبايل: <%=RS1("Mobile2")%></TD>
		</TR>
		</TABLE>
	</td>
</tr>
</table>
	</TD>
</TR>
</TABLE>

<!--#include file="tah.asp" -->