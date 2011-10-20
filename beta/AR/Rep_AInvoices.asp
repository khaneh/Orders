<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ ›—Ê‘"
SubmenuItem=11
if not Auth("C" , 4) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
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

	ON ERROR RESUME NEXT

		ResultsInPage =	cint(request("ResultsInPage"))

		FromDate =		sqlSafe(request("FromDate"))
		ToDate =		sqlSafe(request("ToDate"))
		
		if FromDate="" AND ToDate="" then
			pageTitle="»Â ’Ê—  ﬂ«„·"
		elseif FromDate="" then
			pageTitle="«“ «» œ«  «  «—ÌŒ " & replace (ToDate,"/",".")
		elseif ToDate="" then
			pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  « «‰ Â« "
		else
			pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  «  «—ÌŒ " & replace (ToDate,"/",".")
		end if
	
		if ToDate = "" then ToDate = "9999/99/99"
		
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "OtherReports.asp?errMsg=" & Server.URLEncode("Œÿ« œ— Ê—ÊœÌ.")
		end if
	ON ERROR GOTO 0
%>
<br>
<br>
	<table class="CustTable" cellspacing='1' style='width:680;' align='center'>
<%
	mySumSQL="SELECT SUM(CONVERT(bigint, Invoices.TotalReceivable)) AS SumAmount FROM Invoices INNER JOIN ARItems ON Invoices.ID = ARItems.Link WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IsA = 1) AND (ARItems.GL = "& openGL & ") AND (ARItems.Type = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') "
	Set RS1=Conn.Execute(mySumSQL)
	if NOT RS1.eof then
		SumAmount=RS1("SumAmount")
%>	
			<tr>
				<td colspan="2" class="CusTableHeader" style="text-align:right;height:35;">›«ﬂ Ê— Â«Ì «·› (<%=pageTitle%>)</td>
			</tr>
			<tr>
				<td width="580px" class="CusTD3" style="text-align:left">Ã„⁄ „»·€</td>
				<td width="100px" bgcolor="white" dir=LTR align=center><%=Separate(SumAmount)%></TD>
			</tr>		
		<br><br>

<%		
	end if
	RS1.close
	Set RS1= Nothing

	mySumSQL="SELECT SUM(CONVERT(bigint, ARItems.RemainedAmount)) AS SumRemainedAmount FROM Invoices INNER JOIN ARItems ON Invoices.ID = ARItems.Link WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IsA = 1) AND (ARItems.GL = "& openGL & ") AND (ARItems.Type = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') "
	Set RS1=Conn.Execute(mySumSQL)
	if NOT RS1.eof then
		SumRemainedAmount = RS1("SumRemainedAmount")
%>
			<tr>
				<td width="580px" class="CusTD3" style="text-align:left">Ã„⁄ „«‰œÂ Â«</td>
				<td width="100px" bgcolor="white" dir=LTR align=center><%=Separate(SumRemainedAmount)%></TD>
			</tr>
		<br><br>
<%		
	end if
	RS1.close
	Set RS1= Nothing
%>
	</table>
	<br>
	<table class="CustTable" cellspacing='1' style='width:90%;' align='center'>
		<tr>
			<td colspan="8" class="CusTableHeader" style="text-align:right;height:35;">›«ﬂ Ê— Â«Ì «·› (<%=pageTitle%>)</td>
		</tr>
<%
		mySQL="SELECT Invoices.ID, Invoices.IssuedDate, Invoices.Number, Invoices.TotalReceivable, Invoices.Customer, Accounts.AccountTitle, ARItems.RemainedAmount FROM Invoices INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IsA = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') AND (ARItems.GL = "& openGL & ") AND (ARItems.Type = 1) ORDER BY CONVERT(int, Invoices.Number)"


		Set RS1 = Server.CreateObject("ADODB.Recordset")

		PageSize = ResultsInPage
		RS1.PageSize = PageSize 

		RS1.CursorLocation=3 'in ADOVBS_INC adUseClient=3
		RS1.Open mySQL ,Conn,3
		TotalPages = RS1.PageCount

		CurrentPage=1

		if isnumeric(Request.QueryString("p")) then
			pp=clng(Request.QueryString("p"))
			if pp <= TotalPages AND pp > 0 then
				CurrentPage = pp
			end if
		end if

		if not RS1.eof then
			RS1.AbsolutePage=CurrentPage
		end if


		if RS1.eof then
%>			<tr>
				<td colspan="8" class="CusTD3" style='direction:RTL;'>ÂÌç ›«ﬂ Ê—Ì ÅÌœ« ‰‘œ.</td>
			</tr>
<%		else
%>			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3" style='direction:RTL;'># ›«ﬂ Ê—</td>
				<td class="CusTD3">⁄‰Ê«‰ Õ”«»</td>
				<td class="CusTD3">‘„«—Â Õ”«»</td>
				<td class="CusTD3">‘„«—Â «·›</td>
				<td class="CusTD3"> «—ÌŒ ’œÊ—</td>
				<td class="CusTD3">„»·€</td>
				<td class="CusTD3">„«‰œÂ</td>
			</tr>
<%			tmpCounter=(CurrentPage - 1) * PageSize
			LastNumber=RS1("Number")
			Do while NOT RS1.eof AND (RS1.AbsolutePage = CurrentPage)
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
				Number=	RS1("Number")
				if trim("" & Number) = "" then
					tmpColor3="BGcolor='#FFBBBB'"
					LastNumber=0
				elseif clng(Number) > clng(LastNumber+1) then
					tmpColor3="BGcolor='#FFBBBB'"
					LastNumber=Number
				else
					tmpColor3=""
					LastNumber=Number
				end if
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: pointer;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD style="height:30px;"><%=RS1("ID")%></TD>
					<TD><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD><%=RS1("Customer")%>&nbsp;</TD>
					<TD dir="LTR" align='right' <%=tmpColor3%> ><%=Number%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("IssuedDate")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
					<TD><%=Separate(RS1("RemainedAmount"))%>&nbsp;</TD>
				</TR>
<%				RS1.moveNext
			Loop

			if ToDate="9999/99/99" then ToDate="" 

			if TotalPages > 1 then
				pageCols=20
%>			
				<TR class="RepTableTitle">
					<TD bgcolor='#33AACC' height="30" colspan="8">
					<table width=100% cellspacing=0 style="cursor:hand;color:#444444">
					<tr>
						<td style="height:25;border-bottom:1 solid black;" colspan=<%=pagecols%>>
							<b>’›ÕÂ <%=CurrentPage%> «“ <%=TotalPages%></b>
							&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>);">’›ÕÂ »⁄œ &gt;</a>
						</td>
					</tr>
					<tr>
<%					for i=1 to TotalPages 
						if i = CurrentPage then 
%>							<td style="color:black;"><b>[<%=i%>]</b></td>
<%						else
%>							<td onclick="go2Page(<%=i%>);"><%=i%></td>
<%						end if 
						if i mod pageCols = 0 then response.write "</tr><tr>" 
					next 

%>					</tr>
					</table>

					<SCRIPT LANGUAGE="JavaScript">
					<!--
					function go2Page(p) {
						window.location="?act=show&ResultsInPage=<%=ResultsInPage%>&p="+p+"&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>";
					}
					//-->
					</SCRIPT>

					</TD>
				</TR>
<%			end if
		end if
		RS1.close
		Set RS1 = Nothing
%>
	</table>
	<br>
<%
elseif request("act")="showByDay" then

	ON ERROR RESUME NEXT

		ResultsInPage =	cint(request("ResultsInPage"))

		FromDate =		sqlSafe(request("FromDate"))
		ToDate =		sqlSafe(request("ToDate"))
		
		if FromDate="" AND ToDate="" then
			pageTitle="»Â ’Ê—  ﬂ«„·"
		elseif FromDate="" then
			pageTitle="«“ «» œ«  «  «—ÌŒ " & replace (ToDate,"/",".")
		elseif ToDate="" then
			pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  « «‰ Â« "
		else
			pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  «  «—ÌŒ " & replace (ToDate,"/",".")
		end if
	
		if ToDate = "" then ToDate = "9999/99/99"
		
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "OtherReports.asp?errMsg=" & Server.URLEncode("Œÿ« œ— Ê—ÊœÌ.")
		end if
	ON ERROR GOTO 0
%>
<br>
<br>

<%
	mySumSQL="SELECT SUM(CONVERT(bigint, Invoices.TotalReceivable)) AS SumAmount FROM Invoices INNER JOIN ARItems ON Invoices.ID = ARItems.Link WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IsA = 1) AND (ARItems.GL = "& openGL & ") AND (ARItems.Type = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') "
	Set RS1=Conn.Execute(mySumSQL)
	if NOT RS1.eof then
		SumAmount=RS1("SumAmount")
%>
		<table class="CustTable" cellspacing='1' style='width:680;' align='center'>
			<tr>
				<td colspan="2" class="CusTableHeader" style="text-align:right;height:35;">›«ﬂ Ê— Â«Ì «·› (<%=pageTitle%>)</td>
			</tr>
			<tr>
				<td width="580px" class="CusTD3" style="text-align:left">Ã„⁄ „»·€</td>
				<td width="100px" bgcolor="white" dir=LTR align=center><%=Separate(SumAmount)%></TD>
			</tr>
		</table>
		<br><br>

<%		
	end if
	RS1.close
	Set RS1= Nothing
%>
	<br>
	<table class="CustTable" cellspacing='1' style='width:90%;' align='center'>
		<tr>
			<td colspan="7" class="CusTableHeader" style="text-align:right;height:35;">›«ﬂ Ê— Â«Ì «·› (<%=pageTitle%>)</td>
		</tr>
<%
		mySQL="SELECT Invoices.IssuedDate, COUNT(Invoices.ID) AS CNT_Invoices, SUM(Invoices.TotalReceivable) AS SumPrice, COUNT(DISTINCT Invoices.Customer) AS CNT_Customers FROM Invoices INNER JOIN ARItems ON Invoices.ID = ARItems.Link WHERE (ARItems.GL = '"& openGL & "') AND (ARItems.Type = 1) AND (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IsA = 1) GROUP BY Invoices.IssuedDate HAVING (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') ORDER BY Invoices.IssuedDate"

'response.write mySQL
'response.end

		Set RS1 = Server.CreateObject("ADODB.Recordset")

		PageSize = ResultsInPage
		RS1.PageSize = PageSize 

		RS1.CursorLocation=3 'in ADOVBS_INC adUseClient=3
		RS1.Open mySQL ,Conn,3
		TotalPages = RS1.PageCount

		CurrentPage=1

		if isnumeric(Request.QueryString("p")) then
			pp=clng(Request.QueryString("p"))
			if pp <= TotalPages AND pp > 0 then
				CurrentPage = pp
			end if
		end if

		if not RS1.eof then
			RS1.AbsolutePage=CurrentPage
		end if


		if RS1.eof then
%>			<tr>
				<td colspan="7" class="CusTD3" style='direction:RTL;'>ÂÌç ›«ﬂ Ê—Ì ÅÌœ« ‰‘œ.</td>
			</tr>
<%		else
%>			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3">—Ê“</td>
				<td class="CusTD3"> ⁄œ«œ ›«ﬂ Ê— «·›</td>
				<td class="CusTD3"> ⁄œ«œ „‘ —Ì«‰</td>
				<td class="CusTD3">Ã„⁄ „»·€</td>
			</tr>
<%			tmpCounter=(CurrentPage - 1) * PageSize
			Do while NOT RS1.eof AND (RS1.AbsolutePage = CurrentPage)
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: pointer;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="drill2date('<%=RS1("IssuedDate")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD dir="LTR" align='center'><%=RS1("IssuedDate")%>&nbsp;</TD>
					<TD align='center'><%=RS1("CNT_Invoices")%>&nbsp;</TD>
					<TD align='center'><%=RS1("CNT_Customers")%>&nbsp;</TD>
					<TD><%=Separate(RS1("SumPrice"))%>&nbsp;</TD>
				</TR>
<%				RS1.moveNext
			Loop

			if ToDate="9999/99/99" then ToDate="" 

			if TotalPages > 1 then
				pageCols=20
%>			
				<TR class="RepTableTitle">
					<TD bgcolor='#33AACC' height="30" colspan="7">
					<table width=100% cellspacing=0 style="cursor:hand;color:#444444">
					<tr>
						<td style="height:25;border-bottom:1 solid black;" colspan=<%=pagecols%>>
							<b>’›ÕÂ <%=CurrentPage%> «“ <%=TotalPages%></b>
							&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>);">’›ÕÂ »⁄œ &gt;</a>
						</td>
					</tr>
					<tr>
<%					for i=1 to TotalPages 
						if i = CurrentPage then 
%>							<td style="color:black;"><b>[<%=i%>]</b></td>
<%						else
%>							<td onclick="go2Page(<%=i%>);"><%=i%></td>
<%						end if 
						if i mod pageCols = 0 then response.write "</tr><tr>" 
					next 

%>					</tr>
					</table>

					<SCRIPT LANGUAGE="JavaScript">
					<!--
					function go2Page(p) {
						window.location="?act=showByDay&ResultsInPage=<%=ResultsInPage%>&p="+p+"&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>";
					}
					//-->
					</SCRIPT>

					</TD>
				</TR>
<%			end if
		end if
		RS1.close
		Set RS1 = Nothing
'Rep_AInvoices.asp?act=show&ResultsInPage=50&p=2&FromDate=1384/01/01&ToDate=
%>
	</table>
	<br>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function drill2date(date) {
		window.open('?act=show&ResultsInPage=50&p=1&FromDate='+date+'&ToDate='+date);
	}
	//-->
	</SCRIPT>
<%
end if%>
<!--#include file="tah.asp" -->
