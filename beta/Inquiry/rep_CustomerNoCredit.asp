<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ „‘ —Ì«‰Ì òÂ «⁄ »«— ¬‰Â« ’›— ‰Ì” "
SubmenuItem=11
if not Auth("C" , 9) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; height:100%; background-color:#C3DBEB;}
</STYLE>
<%
if request("act")="show" then

	'ON ERROR RESUME Next
	Ord =			request("Ord")
	If Ord="" Then Ord = 0
	select case Ord
	case "1":
		order="accounts.id"
	case "-1":
		order="accounts.id DESC"
	case "2":
		order="AccountTitle"
	case "-2":
		order="AccountTitle DESC"
	case "3":
		order="creditLimit"
	case "-3":
		order="creditLimit DESC"
	case "4":
		order="arBalance"
	case "-4":
		order="arBalance DESC"
	case "5":
		order="apBalance"
	case "-5":
		order="apBalance DESC"
	case "6":
		order="aoBalance"
	case "-6":
		order="aoBalance DESC"
	case "7":
		order="realName"
	case "-7":
		order="realName DESC"
	
	case else:
		order="accounts.id"
		Ord=1
	end select
%>
	<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
	<TR bgcolor="eeeeee" style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
		<TD width=50 onclick='go2Page(1,1);' style="<%if abs(ord)=1 then response.write style%>">‘„«—Â „‘ —Ì <%if abs(ord)=1 then response.write arrow%></TD>
		<td onclick='go2Page(1,2);' style="<%if abs(ord)=2 then response.write style%>">‰«„  <%if abs(ord)=2 then response.write arrow%></td>
		<td onclick='go2Page(1,3);' style="<%if abs(ord)=3 then response.write style%>">«⁄ »«— <%if abs(ord)=3 then response.write arrow%></td>
		<td onclick='go2Page(1,4);' style="<%if abs(ord)=4 then response.write style%>">„«‰œÂ ›—Ê‘ <%if abs(ord)=4 then response.write arrow%></td>
		<td onclick='go2Page(1,5);' style="<%if abs(ord)=5 then response.write style%>">„«‰œÂ Œ—Ìœ <%if abs(ord)=5 then response.write arrow%></td>
		<td onclick='go2Page(1,6);' style="<%if abs(ord)=6 then response.write style%>">„«‰œÂ ”«Ì— <%if abs(ord)=6 then response.write arrow%></td>
		<td onclick='go2Page(1,7);' style="<%if abs(ord)=7 then response.write style%>">„”ÊÊ· ÅÌêÌ—Ì <%if abs(ord)=7 then response.write arrow%></td>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=7 height=2 bgcolor=0></TD>
	</TR>
	
<%
	mySQL="select accounts.*,users.realName from accounts left outer join users on accounts.csr=users.id where creditLimit>0 ORDER BY " & order
	
	Set rs=Server.CreateObject("ADODB.Recordset")'Conn.Execute(mySQL)

	PageSize = 50
	rs.PageSize = PageSize

	rs.CursorLocation=3 'in ADOVBS_INC adUseClient=3
	rs.Open mySQL ,Conn,3
	TotalPages = rs.PageCount

	CurrentPage=1

	if isnumeric(Request.QueryString("p")) then
		pp=clng(Request.QueryString("p"))
		if pp <= TotalPages AND pp > 0 then
			CurrentPage = pp
		end if
	end if

	if not rs.eof then
		rs.AbsolutePage=CurrentPage
	end if

	if rs.eof then
%>		<tr>
			<td bgcolor="#BBBBBB" height="30" colspan="7" align=center><b>ÂÌç .</b></td>
		</tr>
<%	else
		Do While NOT rs.eof AND (rs.AbsolutePage = CurrentPage)
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
	%>
			<TR bgcolor="<%=tmpColor%>" >
				<TD dir=ltr align=right><A href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rs("ID")%>"><%=rs("id")%></A></TD>
				<TD><%=rs("accountTitle")%></TD>
				<TD><%=rs("creditLimit")%></TD>
				<TD><%=rs("arBalance")%></TD>
				<TD><%=rs("apBalance")%></TD>
				<TD><%=rs("aoBalance")%></TD>
				<TD><%=rs("realName")%></TD>
			</TR>
			  
	<% 
		rs.moveNext
		Loop

		if TotalPages > 1 then
			pageCols=20
%>			
			<TR bgcolor="eeeeee" >
				<TD colspan=7 height=2 bgcolor=0></TD>
			</TR>
			<TR class="RepTableTitle">
				<TD bgcolor="#CCCCEE" height="30" colspan="7">
				<table width=100% cellspacing=0 style="cursor:hand;color:gray;">
				<tr>
					<td style="height:25;border-bottom:1 solid black;" colspan=<%=pagecols%>>
						<b>’›ÕÂ <%=CurrentPage%> «“ <%=TotalPages%></b>
						&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>,0);">’›ÕÂ »⁄œ &gt;</a>
					</td>
				</tr>
				<tr>
<%				for i=1 to TotalPages 
					if i = CurrentPage then 
%>						<td style="color:black;"><b>[<%=i%>]</b></td>
<%					else
%>						<td onclick="go2Page(<%=i%>,0);"><%=i%></td>
<%					end if 
					if i mod pageCols = 0 then response.write "</tr><tr>" 
				next 

%>				</tr>
				</table>
				</TD>
			</TR>
<%		end if
%>
		</TABLE>
		<br>

		<SCRIPT LANGUAGE="JavaScript">
		function go2Page(p,ord) {
			if(ord==0){
				ord=<%=Ord%>;
			}
			else if(ord==<%=Ord%>){
				ord= 0-ord;
			}
			str = '?act=show&FromDate=' + escape('<%=FromDate%>') + '&ToDate=' + escape('<%=ToDate%>') + '&Ord=' + escape(ord) + '&p=' + escape(p);  
			window.location = str;
		}
		</SCRIPT>
<%	
	End if
End If

%>
<!--#include file="tah.asp" -->
