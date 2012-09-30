<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ „œÌ—Ì "
SubmenuItem=6
if not Auth(2 , 7) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
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
	.CusTD3 {background-color: #DDDDDD; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; height:100%; background-color:#C3DBEB;}
	.GenInput { border: 1 solid black; font-family:tahoma; font-size: 9pt;}
	.searchTable {background-color: #330099; border: 2 dashed #330099}
	.searchTable th {background-color: #BBBBBB; font-weight:normal; font-size:9pt;}
	.searchTable td {background-color: #EEEEEE;}
</STYLE>
<%

ResultsInPage =	request("ResultsInPage")
FromDate	=	request("FromDate")
ToDate		=	request("ToDate")
AllowedPercent=	request("AllowedPercent")
Ord			=	request("Ord")
CurrentPage =	request("p")


if FromDate		= ""	then	FromDate = shamsiToday()
if ToDate		= ""	then	ToDate = shamsiToday()

if request("checkDate")="on" OR request("act")="" then checkDate=True

if request("CopyShop")="" then
	CopyShop=	-1
else
	CopyShop=	cint(request("CopyShop"))
end if

if request("approve")="" then
	approve=	-1
else
	approve =	cint(request("approve"))
end if

if request("issue")="" then
	issue=	0
else
	issue =		cint(request("issue"))
end if

if request("void")="" then
	void=	0
else
	void =		cint(request("void"))
end if

if request("creator")="" then
	creator=	0
else
	creator =	cint(request("creator"))
end if


if ResultsInPage="" then
	ResultsInPage=	50
else
	ResultsInPage =	cint(ResultsInPage)
end if


if AllowedPercent="" then
	AllowedPercent=	10
else
	AllowedPercent =	cint(AllowedPercent)
end if


%>
<br>
<FORM METHOD=POST ACTION="?act=show">
<INPUT TYPE="hidden" Name="p" Value="<%=CurrentPage%>">
<INPUT TYPE="hidden" Name="ord" Value="<%=Ord%>">
<TABLE align=center border=0 bgcolor=#0 cellspacing=1 cellpadding=2 class="searchTable">
<TR>
	<TH colspan=3 height=25 align=center bgcolor=#CCCCCC>Ã” ÃÊ</TD>
</TR>
<TR>
	<TD title=" «—ÌŒ «ÌÃ«œ ›«ﬂ Ê—"><INPUT <%if checkDate then response.write "checked"%> TYPE="checkbox" NAME="checkDate"> «—ÌŒ</TD>
	<TD>«“  <INPUT class="GenInput" TYPE="text" NAME="FromDate" dir="LTR" value="<%=FromDate%>" size="10" onKeyPress="return maskDate(this);" onBlur="acceptDate(this);" maxlength="10"></TD>
	<TD> ‹« <INPUT class="GenInput" TYPE="text" NAME="ToDate" dir="LTR" value="<%=ToDate%>" size="10" onKeyPress="return maskDate(this);" onBlur="acceptDate(this);" maxlength="10"></TD>
</TR>
<TR>
	<TD> Œ›Ì› „Ã«“ &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;</TD>
	<TD colspan=2>
		<INPUT class="GenInput" TYPE="text" NAME="AllowedPercent" dir="LTR" value="<%=AllowedPercent%>" size="2" onKeyPress="return maskNumber(this);" maxlength="2">
		%
	</TD>
</TR>
<%
	if Auth(2 , 8) then ' Can Change the conditions
		checked1 = ""
		checked2 = ""
		checked3 = ""
		if CopyShop = 1 then
			checked1 = " checked"
		elseif CopyShop = 0 then
			checked2 = " checked"
		else
			checked3 = " checked"
		end if
%>
		<TR>
			<TD><INPUT TYPE="radio" NAME="CopyShop" Value="1" <%=checked1%>>ﬂÅÌ ‘«Å</TD>
			<TD><INPUT TYPE="radio" NAME="CopyShop" Value="0" <%=checked2%>>€Ì— ﬂÅÌ ‘«Å</TD>
			<TD><INPUT TYPE="radio" NAME="CopyShop" Value="-1"<%=checked3%>>Â—œÊ</TD>
		</TR>
<%
		checked1 = ""
		checked2 = ""
		checked3 = ""
		if approve = 0 then
			checked1 = " checked"
		elseif approve = 1 then
			checked2 = " checked"
		else
			checked3 = " checked"
		end if
%>
		<TR>
			<TD><INPUT TYPE="radio" NAME="approve" Value="0" <%=checked1%>> «ÌÌœ ‰‘œÂ</TD>
			<TD><INPUT TYPE="radio" NAME="approve" Value="1" <%=checked2%>> «ÌÌœ ‘œÂ</TD>
			<TD><INPUT TYPE="radio" NAME="approve" Value="-1"<%=checked3%>>Â—œÊ</TD>
		</TR>
<%
		checked1 = ""
		checked2 = ""
		checked3 = ""
		if issue = 0 then
			checked1 = " checked"
		elseif issue = 1 then
			checked2 = " checked"
		else
			checked3 = " checked"
		end if
%>
		<TR>
			<TD><INPUT TYPE="radio" NAME="issue" Value="0" <%=checked1%>>’«œ— ‰‘œÂ</TD>
			<TD><INPUT TYPE="radio" NAME="issue" Value="1" <%=checked2%>>’«œ— ‘œÂ</TD>
			<TD><INPUT TYPE="radio" NAME="issue" Value="-1"<%=checked3%>>Â—œÊ</TD>
		</TR>
<%
		checked1 = ""
		checked2 = ""
		checked3 = ""
		if void = 0 then
			checked1 = " checked"
		elseif void = 1 then
			checked2 = " checked"
		else
			checked3 = " checked"
		end if
%>
		<TR>
			<TD><INPUT TYPE="radio" NAME="void" Value="0" <%=checked1%>>«»ÿ«· ‰‘œÂ</TD>
			<TD><INPUT TYPE="radio" NAME="void" Value="1" <%=checked2%>>«»ÿ«· ‘œÂ</TD>
			<TD><INPUT TYPE="radio" NAME="void" Value="-1"<%=checked3%>>Â—œÊ</TD>
		</TR>
<%
		checked1 = ""
		checked2 = ""
		checked3 = ""
		if void = 0 then
			checked1 = " checked"
		elseif void = 1 then
			checked2 = " checked"
		else
			checked3 = " checked"
		end if
%>
		<TR>
			<TD>«ÌÃ«œ ﬂ‰‰œÂ</TD>
			<TD colspan="2">
			<SELECT NAME="creator" style='font-family: tahoma,arial ; font-size: 8pt; width: 140px'>
				<option value="0" style="color:blue;">*Â„Â*</option>
<%				set RS_TEMP=Conn.Execute ("SELECT ID, RealName FROM Users WHERE Display=1 ORDER BY RealName") 
				Do while not RS_TEMP.eof
%>
					<option value="<%=RS_TEMP("ID")%>" <%if RS_TEMP("ID")=creator then response.write " selected "%> ><%=RS_TEMP("RealName")%></option>
<%
				RS_TEMP.moveNext
				Loop
				RS_TEMP.close
%>
			</SELECT></TD>
		</TR>
<%
	end if
%>
<TR>
	<TD> ⁄œ«œ œ— ’›ÕÂ</TD>
	<TD colspan=2>
		<INPUT class="GenInput" TYPE="text" NAME="ResultsInPage" dir="LTR" value="<%=ResultsInPage%>" size="2" onKeyPress="return maskNumber(this);" maxlength="4">
	</TD>
</TR>
<TR>
	<TD colspan=3 align=center><INPUT class="genButton" TYPE="submit" value="‰„«Ì‘"></TD>
</TR>
</TABLE>
</FORM>
<% 

'if request("act")="show" then

	select case Ord
	case "1":
		order="Invoices.ID"
	case "-1":
		order="Invoices.ID DESC"
	case "2":
		order="Users.RealName"
	case "-2":
		order="Users.RealName DESC"
	case "3":
		order="Invoices.CreatedDate"
	case "-3":
		order="Invoices.CreatedDate DESC"
	case "4":
		order="InvoiceOrderRelations.[Order]"
	case "-4":
		order="InvoiceOrderRelations.[Order] DESC"
	case "5":
		order="InvStatus"
	case "-5":
		order="InvStatus DESC"
	case "6":
		order="TotalDiscount"
	case "-6":
		order="TotalDiscount DESC"
	case "7":
		order="TotalReverse"
	case "-7":
		order="TotalReverse DESC"
	case "8":
		order="TotalReceivable"
	case "-8":
		order="TotalReceivable DESC"
	case else:
		order="Invoices.CreatedDate DESC"
		Ord=-3
	end select

	criteria=""
	writeAND=""
	if checkDate then
		criteria = "(Invoices.CreatedDate >= '"& FromDate & "') AND (Invoices.CreatedDate <= '"& ToDate & "') "
		writeAND=" AND "
	end if

	if CopyShop <> -1 then
		if CopyShop = 1 then
			tmp="(InvoiceOrderRelations.[Order] IS NULL) "
		else
			tmp="(InvoiceOrderRelations.[Order] IS NOT NULL) "
		end if
		criteria = criteria & writeAND & tmp
		writeAND=" AND "
	end if

	if approve <> -1 then
		criteria = criteria & writeAND & "(Invoices.Approved = '"& approve & "') "
		writeAND=" AND "
	end if

	if issue <> -1 then
		criteria = criteria & writeAND & "(Invoices.Issued = '"& issue & "') "
		writeAND=" AND "
	end if

	if void <> -1 then
		criteria = criteria & writeAND & "(Invoices.Voided = '"& void & "') "
		writeAND=" AND "
	end if

	if creator <> 0 then
		criteria = criteria & writeAND & "(Invoices.CreatedBy = '"& creator & "') "
		writeAND=" AND "
	end if

	if writeAND="" then
		criteria = "(1=1) "
	end if

%>
	  <TaBlE class="CustTable4" cellspacing="2" cellspacing="2">
	  <Tr>
		<Td colspan="2" valign="top" align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="9" class="CusTableHeader" style="text-align:right;">›«ﬂ Ê— Â«</td>
				</tr>
				<%
				mySQL="SELECT Invoices.*, Users.RealName AS Creator, InvoiceOrderRelations.[Order], orderStatus.name as vazyat, " &_
					"orderSteps.name as marhale, CONVERT(int, Invoices.Approved) + CONVERT(int, Invoices.Issued) * 2 AS InvStatus " &_
					"FROM orders inner join orderStatus on orders.status = orderStatus.id "&_
					"inner join orderSteps on orders.step = orderSteps.id RIGHT OUTER JOIN " &_
					" InvoiceOrderRelations ON orders.id = InvoiceOrderRelations.[Order] RIGHT OUTER JOIN " &_
					" Invoices INNER JOIN " &_
					" Users ON Invoices.CreatedBy = Users.ID ON InvoiceOrderRelations.Invoice = Invoices.ID " &_
					"WHERE " & criteria &_
					"ORDER BY " & order
'response.write mySQL
				if ord<0 then
					style="background-color: #33CC99;"
					arrow="<br><span style='font-family:webdings'>6 6 6</span>"
				else
					style="background-color: #33CC99;"
					arrow="<br><span style='font-family:webdings'>5 5 5</span>"
				end if

				Set RS1 = Server.CreateObject("ADODB.Recordset")

				PageSize = ResultsInPage
				RS1.PageSize = PageSize 

				RS1.CursorLocation=3 'in ADOVBS_INC adUseClient=3
				RS1.Open mySQL ,Conn,3
				TotalPages = RS1.PageCount

				if isnumeric(CurrentPage) then
					CurrentPage=clng(CurrentPage)
					if CurrentPage > TotalPages OR CurrentPage <= 0 then
						CurrentPage = 1
					end if
				else
					CurrentPage=1
				end if

				if not RS1.eof then
					RS1.AbsolutePage=CurrentPage
				end if

				if RS1.eof then
%>
					<tr>
						<td colspan="9" class="CusTD3">ÂÌç</td>
					</tr>
<%
				else
%>					<tr	class="CusTD3" style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
						<td >#</td>
						<TD onclick='go2Page(1,1);'  style="<%if abs(ord)=1 then response.write style%>"># ›«ﬂ Ê—	<%if abs(ord)=1 then response.write arrow%></TD>
						<TD onclick='go2Page(1,2);'  style="<%if abs(ord)=2 then response.write style%>">«ÌÃ«œ ﬂ‰‰œÂ<%if abs(ord)=2 then response.write arrow%></TD>
						<TD onclick='go2Page(1,3);'  style="<%if abs(ord)=3 then response.write style%>"> «—ÌŒ		<%if abs(ord)=3 then response.write arrow%></TD>
						<TD onclick='go2Page(1,4);'  style="<%if abs(ord)=4 then response.write style%>"># ”›«—‘	<%if abs(ord)=4 then response.write arrow%></TD>
						<TD onclick='go2Page(1,5);'  style="<%if abs(ord)=5 then response.write style%>">Ê÷⁄Ì 		<%if abs(ord)=5 then response.write arrow%></TD>
						<TD onclick='go2Page(1,-6);' style="<%if abs(ord)=6 then response.write style%>"> Œ›Ì›		<%if abs(ord)=6 then response.write arrow%></TD>
						<TD onclick='go2Page(1,-7);' style="<%if abs(ord)=7 then response.write style%>">»—ê‘ 		<%if abs(ord)=7 then response.write arrow%></TD>
						<TD onclick='go2Page(1,-8);' style="<%if abs(ord)=8 then response.write style%>">„»·€		<%if abs(ord)=8 then response.write arrow%></TD>

					</tr>
<%					tmpCounter = 0
					SumDiscount = 0
					SumReverse = 0
					SumReceivable = 0
					AlertColor="bgcolor=#FFAAAA"
					Do while NOT RS1.eof AND RS1.AbsolutePage = CurrentPage
						tmpCounter = tmpCounter + 1
						if tmpCounter mod 2 = 1 then
							tmpColor="#FFFFFF"
							tmpColor2="#FFFFBB"
						Else
							tmpColor="#DDDDDD"
							tmpColor2="#EEEEBB"
						End if 

						TotalPrice =		cdbl(RS1("TotalPrice"))
						TotalDiscount =		cdbl(RS1("TotalDiscount"))
						TotalReverse =		cdbl(RS1("TotalReverse"))
						TotalReceivable =	cdbl(RS1("TotalReceivable"))
						
						if TotalPrice<>0 then
							DiscountPercent =	cint(TotalDiscount * 100/TotalPrice)
							ReversePercent  =	cint(TotalReverse * 100/TotalPrice)
							ReceivablePercent=	100 - cint(TotalReceivable * 100/TotalPrice)
						else
							DiscountPercent = 0
							ReversePercent  = 0
							ReceivablePercent=0
						end if

						if DiscountPercent > AllowedPercent then
							DiscountAlert=AlertColor
						else
							DiscountAlert=""
						end if

						if ReversePercent > AllowedPercent then
							ReverseAlert=AlertColor
						else
							ReverseAlert=""
						end if

						if ReceivablePercent > AllowedPercent then
							ReceivableAlert=AlertColor
						else
							ReceivableAlert=""
						end if

						if RS1("IsReverse") then
							tmpColor="#FF9966"
							SumDiscount =	SumDiscount - TotalDiscount
							SumReverse =	SumReverse - TotalReverse
							SumReceivable = SumReceivable - TotalReceivable
						else
							SumDiscount =	SumDiscount + TotalDiscount
							SumReverse =	SumReverse + TotalReverse
							SumReceivable = SumReceivable + TotalReceivable
						end if

						OrderNo=	RS1("Order")
						OrderTxt=	RS1("vazyat") & " - " & RS1("Marhale")
						if isnull(OrderNo) then 
							OrderNo=	"<FONT COLOR='gray'>»œÊ‰ ”›«—‘</FONT>"
							OrderTxt=	"»œÊ‰ ”›«—‘"
						end if
						InvStatus=	cint(RS1("InvStatus"))
						InvStatusTxt=""
					  '--
						if InvStatus Mod 2 = 1 then
							InvStatusTxt="<FONT COLOR='Green'> «ÌÌœ ‘œÂ</FONT> - "
						else
							InvStatusTxt="<FONT COLOR='red'> «ÌÌœ ‰‘œÂ</FONT> - "
						end if
						InvStatus=InvStatus \2
						if InvStatus Mod 2 = 1  then
							InvStatusTxt=InvStatusTxt & "<FONT COLOR='Green'>’«œ— ‘œÂ</FONT> "
						else
							InvStatusTxt=InvStatusTxt & "<FONT COLOR='red'>’«œ— ‰‘œÂ</FONT> "
						end if
					  '--
%>
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD><%=RS1("Creator")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
							<TD dir="LTR" align='right' title="<%=OrderTxt%>"><%=OrderNo%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=InvStatusTxt%>&nbsp;</TD>
							<TD dir="LTR" align='right' title="<%=DiscountPercent%>%  Œ›Ì›" <%=DiscountAlert%>><%=Separate(TotalDiscount)%>&nbsp;</TD>
							<TD dir="LTR" align='right' title="<%=ReversePercent%>% »—ê‘ " <%=ReverseAlert%>><%=Separate(TotalReverse)%>&nbsp;</TD>
							<TD dir="LTR" align='right' title="<%=ReceivablePercent%>% ﬂ”— ‘œÂ" <%=ReceivableAlert%>><%=Separate(TotalReceivable)%>&nbsp;</TD>
						</TR>
<%						RS1.moveNext
					Loop

					if TotalPages = 1 then

						if SumPrice<>0 then
							DiscountPercent =	cint(SumDiscount * 100/SumPrice)
							ReversePercent  =	cint(SumReverse * 100/SumPrice)
							ReceivablePercent=	100 - cint(SumReceivable * 100/SumPrice)
						else
							DiscountPercent = 0
							ReversePercent  = 0
							ReceivablePercent=0
						end if

						if DiscountPercent > AllowedPercent then
							DiscountAlert=AlertColor
						else
							DiscountAlert=""
						end if

						if ReversePercent > AllowedPercent then
							ReverseAlert=AlertColor
						else
							ReverseAlert=""
						end if

						if ReceivablePercent > AllowedPercent then
							ReceivableAlert=AlertColor
						else
							ReceivableAlert=""
						end if

%>
						<TR bgcolor="#BBBBBB">
							<TD style="height:30px;" colspan="6" align="left">Ã„⁄:</TD>
							<TD dir="LTR" align='right' title="<%=DiscountPercent%>%  Œ›Ì›" <%=DiscountAlert%>><%=Separate(SumDiscount)%>&nbsp;</TD>
							<TD dir="LTR" align='right' title="<%=ReversePercent%>% »—ê‘ " <%=ReverseAlert%>><%=Separate(SumReverse)%>&nbsp;</TD>
							<TD dir="LTR" align='right' title="<%=ReceivablePercent%>% ﬂ”— ‘œÂ" <%=ReceivableAlert%>><%=Separate(SumReceivable)%>&nbsp;</TD>
						</TR>
<%
					else
						pageCols=20
						'----------------------------------------SAM----------------------------------
						if SumPrice<>0 then
							DiscountPercent =	cint(SumDiscount * 100/SumPrice)
							ReversePercent  =	cint(SumReverse * 100/SumPrice)
							ReceivablePercent=	100 - cint(SumReceivable * 100/SumPrice)
						else
							DiscountPercent = 0
							ReversePercent  = 0
							ReceivablePercent=0
						end if

%>						<TR bgcolor='#BBBBBB'>
							<TD style="height:30px;" colspan="6" align="left">Ã„⁄ «Ì‰ ’›ÕÂ:</TD>
							<TD dir="LTR" align='right' title="<%=DiscountPercent%>%  Œ›Ì›"><%=Separate(SumDiscount)%>&nbsp;</TD>
							<TD dir="LTR" align='right' title="<%=ReversePercent%>% »—ê‘ " ><%=Separate(SumReverse)%>&nbsp;</TD>
							<TD dir="LTR" align='right' title="<%=ReceivablePercent%>% ﬂ”— ‘œÂ" ><%=Separate(SumReceivable)%>&nbsp;</TD>
						</TR>
						<TR class="RepTableTitle">
							<TD bgcolor='#33AACC' height="30" colspan="9">
							<table width=100% cellspacing=0 style="cursor:hand;color:#444444">
							<tr>
								<td style="height:25;border-bottom:1 solid black;" colspan=<%=pagecols%>>
									<b>’›ÕÂ <%=CurrentPage%> «“ <%=TotalPages%></b>
									&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>,0);">’›ÕÂ »⁄œ &gt;</a>
								</td>
							</tr>
							<tr>
<%							for i=1 to TotalPages 
								if i = CurrentPage then 
%>									<td style="color:black;"><b>[<%=i%>]</b></td>
<%								else
%>									<td onclick="go2Page(<%=i%>,0);"><%=i%></td>
<%								end if 
								if i mod pageCols = 0 then response.write "</tr><tr>" 
							next 

%>							</tr>
							</table>
							</TD>
						</TR>
<%					end if
				end if
%>
			</table>
			<SCRIPT LANGUAGE="JavaScript">
			<!--
			function go2Page(p,ord) {
				if(ord==0){
					ord=<%=Ord%>;
				}
				else if(ord==<%=Ord%>){
					ord= 0-ord;
				}
//				str='?act=show&FromDate='+escape('<%=FromDate%>')+'&ToDate='+escape('<%=ToDate%>')+'&Ord='+escape(ord)+'&p='+escape(p)
//				window.location=str;
				document.all.ord.value=ord;
				document.all.p.value=p;
				document.forms[0].submit();
			}
			//-->
			</SCRIPT>

		</Td>
	  </Tr>
	</TaBlE>
	<br>
<%
'end if
%>

<!--#include file="tah.asp" -->