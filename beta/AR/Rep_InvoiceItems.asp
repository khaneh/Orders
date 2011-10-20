<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ ›—Ê‘"
SubmenuItem=11
if not Auth("C" , 3) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable th {font-size:9pt; padding:5px; background-color:#0080C0;height:25px;}
	.RepTable td {height:25px;}
	.RepTable input {font-family:tahoma; font-size:9pt; border:1 solid black;}
	.RepTable select {font-family:tahoma; font-size:9pt; border:1 solid black;}
</STYLE>
<%
if request("act")="show" then

	ON ERROR RESUME NEXT

		Category =		cint(request("Category"))

		isA =			clng(request("isA"))

		FromDate =		sqlSafe(request("FromDate"))
		ToDate =		sqlSafe(request("ToDate"))

		ResultsInPage =	cint(request("ResultsInPage"))
		
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
		
		SalesAction =	request("SalesAction")
		SalesPerson =	request("SalesPerson")

		if SalesAction="" then 
			SalesAction = 2 ' «ÌÌœ
		else
			SalesAction = cint(SalesAction)
		end if
		if SalesPerson="" then 
			SalesPerson = 0 'Â„Â
		else
			SalesPerson = cint(SalesPerson)
		end if

		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "OtherReports.asp?errMsg=" & Server.URLEncode("Œÿ« œ— Ê—ÊœÌ.")
		end if
	ON ERROR GOTO 0

%>
<br>
	<FORM METHOD=POST ACTION="?act=show">
	<table class="RepTable" id="AInvoices" width="300" align=center>
	<tr>
		<th colspan="4">ê“«—‘ ¬Ì „ Â«Ì ›—Ê‘</td>
	</tr>
	<tr>
		<td align=left>«“&nbsp; «—ÌŒ</td>
		<td><INPUT TYPE="text" NAME="FromDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="acceptDate(this);" Value="<%=FromDate%>"></td>
		<td align=left> «&nbsp; «—ÌŒ</td>
		<td><INPUT TYPE="text" NAME="ToDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="acceptDate(this);" <%if ToDate<>"9999/99/99" then response.write "Value='" & ToDate & "'" %>></td>
	</tr>
	<tr>
		<td align=left>ê—ÊÂ</td>
		<td align=left>
			<SELECT NAME="Category" style="width:150px;">
				<option value="">--- «‰ Œ«» ﬂ‰Ìœ ---</option>
				<option value="0" <%if Category=0 then response.write "selected " : GroupName="[ Â„Â ] "%>>Â„Â</option>
<%
			mySQL = "SELECT * FROM InvoiceItemCategories"
			Set RS1 = Conn.Execute(mySQL)
			Do Until RS1.eof
%>				
				<option value="<%=RS1("ID")%>" <%if RS1("ID")=Category then response.write "selected " : GroupName="[" & RS1("Name") & "] "%>><%=RS1("Name")%></option>
<%
				RS1.moveNext
			Loop
			RS1.close
			Set RS1 = Nothing

			otherCriteria=""
			if isA = 1 then
				isAName = "»«  Ìﬂ «·›"
				otherCriteria = " AND (isA=1) "
			elseif isA = 2 then
				isAName = "»œÊ‰  Ìﬂ «·›"
				otherCriteria = " AND (isA=0) "
			end if

			if Auth("C" , 6) then 'ê“«—‘ ¬Ì „ Â«Ì ›—Ê‘ »—Õ”» ›—Ê‘‰œÂ
				if SalesPerson <> 0 then
					select case SalesAction
					case 0:
						otherCriteria = otherCriteria & " AND ((Invoices.CreatedBy=" & SalesPerson & ") OR (Invoices.ApprovedBy=" & SalesPerson & ") OR (Invoices.IssuedBy=" & SalesPerson & "))"
					case 1:
						otherCriteria = otherCriteria & " AND (Invoices.CreatedBy=" & SalesPerson & ")"
					case 2:
						otherCriteria = otherCriteria & " AND (Invoices.ApprovedBy=" & SalesPerson & ")"
					case 3:
						otherCriteria = otherCriteria & " AND (Invoices.IssuedBy=" & SalesPerson & ")"
					end select
				end if
			end if

			if Category = 0 then
				GroupName= "(Â„Â)"
'	S A M ------------------------------- THIS FILE HAS BEEN CHANGED BY SAM
				mySQL_Sum="SELECT Count(*) AS CNT, SUM(A4Qtty * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS A4Qtty, SUM(SumAppQtty * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS SumAppQtty, SUM(SumPrice * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS SumPrice, SUM(SumDiscount * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS SumDiscount, SUM(SumReverse * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS SumReverse, SUM(SumVat * - (2 * CONVERT(tinyint, IsReverse) -1)) AS SumVat FROM (SELECT InvoiceItems.ID, InvoiceItems.Name, SUM(InvoiceLines.Qtty * InvoiceLines.Sets * ROUND(InvoiceLines.Length * InvoiceLines.Width / 630 + .49, 0)) AS A4Qtty, SUM(InvoiceLines.AppQtty) AS SumAppQtty, SUM(CONVERT(bigint,InvoiceLines.Price)) AS SumPrice, SUM(InvoiceLines.Discount) AS SumDiscount, SUM(InvoiceLines.Reverse) AS SumReverse, SUM(InvoiceLines.Vat) AS SumVat, Invoices.IsReverse FROM InvoiceItems INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') " & otherCriteria & " GROUP BY Invoices.IsReverse, InvoiceItems.ID, InvoiceItems.Name) DRV"
				'mySQL="SELECT InvoiceItems.ID, InvoiceItems.Name, SUM(InvoiceLines.Qtty * InvoiceLines.Sets * ROUND(InvoiceLines.Length * InvoiceLines.Width / 630 + .49, 0)) AS A4Qtty, SUM(InvoiceLines.AppQtty) AS SumAppQtty, SUM(CONVERT(bigint,InvoiceLines.Price)) AS SumPrice, SUM(InvoiceLines.Discount) AS SumDiscount, SUM(InvoiceLines.Reverse) AS SumReverse, SUM(InvoiceLines.Vat) AS SumVat, Invoices.IsReverse, MAX(InvoiceItemCategories.Name) as CategoryName FROM InvoiceItemCategories INNER JOIN InvoiceItemCategoryRelations ON InvoiceItemCategories.ID = InvoiceItemCategoryRelations.InvoiceItemCategory INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') " & otherCriteria & "  GROUP BY Invoices.IsReverse, InvoiceItems.ID, InvoiceItems.Name ORDER BY InvoiceItems.ID, Invoices.IsReverse"
				'mySQL="SELECT InvoiceItemCategories.ID, InvoiceItemCategories.Name, SUM(InvoiceLines.Qtty * InvoiceLines.Sets * ROUND(InvoiceLines.Length * InvoiceLines.Width / 630 + .49, 0)) AS A4Qtty, SUM(InvoiceLines.AppQtty) AS SumAppQtty, SUM(CONVERT(bigint,InvoiceLines.Price)) AS SumPrice, SUM(InvoiceLines.Discount) AS SumDiscount, SUM(InvoiceLines.Reverse) AS SumReverse, SUM(InvoiceLines.Vat) AS SumVat, InvoiceItemCategories.Name as CategoryName FROM InvoiceItemCategories INNER JOIN InvoiceItemCategoryRelations ON InvoiceItemCategories.ID = InvoiceItemCategoryRelations.InvoiceItemCategory INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate BETWEEN N'"& FromDate & "' AND  N'"& ToDate & "') " & otherCriteria & " GROUP BY InvoiceItemCategories.ID, InvoiceItemCategories.Name ORDER BY InvoiceItemCategories.ID "
				mySQL="SELECT InvoiceItemCategories.ID, InvoiceItemCategories.Name, ISNULL(drv.A4Qtty,0) as A4Qtty, ISNULL(drv.SumAppQtty,0) as SumAppQtty, ISNULL(drv.SumPrice,0) as SumPrice, ISNULL(drv.SumDiscount,0) as SumDiscount, ISNULL(drv.SumReverse,0) as SumReverse, ISNULL(drv.SumVat,0) as SumVat, ISNULL(drv.CategoryName,'-') as CategoryName FROM InvoiceItemCategories LEFT OUTER JOIN (SELECT InvoiceItemCategories.ID, SUM(InvoiceLines.Qtty * InvoiceLines.Sets * ROUND(InvoiceLines.Length * InvoiceLines.Width / 630 + .49, 0)) AS A4Qtty, SUM(InvoiceLines.AppQtty) AS SumAppQtty, SUM(CONVERT(bigint,InvoiceLines.Price)) AS SumPrice, SUM(InvoiceLines.Discount) AS SumDiscount, SUM(InvoiceLines.Reverse) AS SumReverse, SUM(InvoiceLines.Vat) AS SumVat, MAX(InvoiceItemCategories.Name) as CategoryName FROM InvoiceItemCategories INNER JOIN InvoiceItemCategoryRelations ON InvoiceItemCategories.ID = InvoiceItemCategoryRelations.InvoiceItemCategory INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate BETWEEN N'"& FromDate & "' AND  N'"& ToDate & "') " & otherCriteria & " GROUP BY InvoiceItemCategories.ID) drv ON InvoiceItemCategories.ID = drv.ID ORDER BY InvoiceItemCategories.ID"
			else
				mySQL_Sum="SELECT Count(*) AS CNT, SUM(A4Qtty * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS A4Qtty, SUM(SumAppQtty * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS SumAppQtty, SUM(SumPrice * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS SumPrice, SUM(SumDiscount * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS SumDiscount, SUM(SumReverse * - (2 * CONVERT(tinyint, IsReverse) - 1)) AS SumReverse, SUM(SumVat * -(2 * CONVERT(tinyint, IsReverse) - 1)) AS SumVat FROM (SELECT InvoiceItems.ID, InvoiceItems.Name, SUM(InvoiceLines.Qtty * InvoiceLines.Sets * ROUND(InvoiceLines.Length * InvoiceLines.Width / 630 + .49, 0)) AS A4Qtty, SUM(InvoiceLines.AppQtty) AS SumAppQtty, SUM(CONVERT(bigint,InvoiceLines.Price)) AS SumPrice, SUM(InvoiceLines.Discount) AS SumDiscount, SUM(InvoiceLines.Reverse) AS SumReverse, SUM(InvoiceLines.Vat) AS SumVat, Invoices.IsReverse FROM InvoiceItemCategories INNER JOIN InvoiceItemCategoryRelations ON InvoiceItemCategories.ID = InvoiceItemCategoryRelations.InvoiceItemCategory INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (InvoiceItemCategories.ID = "& Category & ") AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') " & otherCriteria & " GROUP BY Invoices.IsReverse, InvoiceItems.ID, InvoiceItems.Name) DRV"
				mySQL="SELECT InvoiceItems.ID, InvoiceItems.Name, SUM(InvoiceLines.Qtty * InvoiceLines.Sets * ROUND(InvoiceLines.Length * InvoiceLines.Width / 630 + .49, 0)) AS A4Qtty, SUM(InvoiceLines.AppQtty) AS SumAppQtty, SUM(CONVERT(bigint,InvoiceLines.Price)) AS SumPrice, SUM(InvoiceLines.Discount) AS SumDiscount, SUM(InvoiceLines.Reverse) AS SumReverse, SUM(InvoiceLines.Vat) AS SumVat , Invoices.IsReverse, MAX(InvoiceItemCategories.Name) as CategoryName FROM InvoiceItemCategories INNER JOIN InvoiceItemCategoryRelations ON InvoiceItemCategories.ID = InvoiceItemCategoryRelations.InvoiceItemCategory INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (InvoiceItemCategories.ID = "& Category & ") AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') " & otherCriteria & " GROUP BY Invoices.IsReverse, InvoiceItems.ID, InvoiceItems.Name ORDER BY InvoiceItems.ID,Invoices.IsReverse"
			end if

%>
			</SELECT>
		</td>
		<td align=left> ⁄œ«œ</td>
		<td><INPUT TYPE="text" NAME="ResultsInPage" style="width:75px;direction:LTR;text-align:right;" maxlength=5 value="<%=ResultsInPage%>"></td>
	</tr>
	<tr>
		<td align=left>›«ﬂ Ê—Â«</td>
		<td align=left>
			<SELECT NAME="isA" style="width:150px;">
				<option value="0" <%if isA=0 then response.write "selected " %>>Â„Â </option>
				<option value="1" <%if isA=1 then response.write "selected " %>>»«  Ìﬂ «·›</option>
				<option value="2" <%if isA=2 then response.write "selected " %>>»œÊ‰  Ìﬂ «·›</option>
			</SELECT>
		</td>
	</tr>
<%
'response.write mySQL
'response.end
	if true Or Auth("C" , 6) then 'ê“«—‘ ¬Ì „ Â«Ì ›—Ê‘ »—Õ”» ›—Ê‘‰œÂ
%>
	<tr>
		<td align=left>CSR</td>
		<td align=left>
			<SELECT NAME="SalesAction" style="width:100px;">
				<option value="1" <%if SalesAction=1 then response.write "selected " %>>À»  ﬂ‰‰œÂ</option>
				<option value="2" <%if SalesAction=2 then response.write "selected " %>> «ÌÌœ ﬂ‰‰œÂ</option>
				<option value="3" <%if SalesAction=3 then response.write "selected " %>>’«œ— ﬂ‰‰œÂ</option>
				<option value="0" <%if SalesAction=0 then response.write "selected " %>>Â—ﬂœ«„ «“ „Ê«—œ</option>
			</SELECT>
		</td>
		<td align=left>
			<SELECT NAME="SalesPerson" style="width:100px;">
				<option value="0" <%if SalesPerson=0 then response.write "selected " %>>Â„Â </option>
<%
			set RS=Conn.Execute ("SELECT * FROM Users WHERE ID<>0 ORDER BY RealName") 

			Do while not RS.eof
%>
				<option value="<%=RS("ID")%>" <%if SalesPerson=RS("ID") then response.write "selected " %>><%=RS("RealName")%></option>
<%
				RS.moveNext
			Loop
			RS.close
%>
			</SELECT>
		</td>
	</tr>
<%
	end if
%>
	<tr>
		<td colspan=4 align=center><INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ "></td>
	</tr>
	</table>
	</FORM>
<br>
<%

	Set RS1=Conn.Execute(mySQL_Sum)
	if NOT RS1.eof then

		CNT =			cdbl(RS1("CNT"))
		if CNT >0 then
			A4Qtty =		cdbl(RS1("A4Qtty"))
			SumAppQtty =	cdbl(RS1("SumAppQtty"))
			SumPrice =		cdbl(RS1("SumPrice"))
			SumDiscount =	cdbl(RS1("SumDiscount"))
			SumReverse =	cdbl(RS1("SumReverse"))
			SumVat =		cdbl(RS1("SumVat"))
		else
			A4Qtty =		0
			SumAppQtty =	0
			SumPrice =		0
			SumDiscount =	0
			SumReverse =	0
			SumVat =		0
		end if
		SumSales =	SumPrice - SumDiscount - SumReverse '+ SumVat
%>
		<table class="CustTable" cellspacing='1' style='width:90%;' align='center'>
			<tr>
				<td colspan="8" class="CusTableHeader" style="text-align:right;height:35;">Ã„⁄ ¬Ì „ Â«Ì ›—Ê‘ <%=isAName %> &nbsp; <%=GroupName%>  (<%=pageTitle%>)</td>
			</tr>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3" style="direction:RTL;"> ⁄œ«œ A4</td>
				<td class="CusTD3"> ⁄œ«œ „ÊÀ—</td>
				<td class="CusTD3">Ã„⁄ „»·€</td>
				<td class="CusTD3">Ã„⁄  Œ›Ì›</td>
				<td class="CusTD3">Ã„⁄ »—ê‘ </td>
				<td class="CusTD3">Ã„⁄ „«·Ì« </td>
				<td class="CusTD3">Ã„⁄ Œ«·’ ›—Ê‘</td>
			</tr>
			<tr bgcolor="white">
				<TD dir=LTR align=right><%=Separate(CNT)%></TD>
				<TD dir=LTR align=right><%=Separate(A4Qtty)%></TD>
				<TD dir=LTR align=right><%=Separate(SumAppQtty)%></TD>
				<TD dir=LTR align=right><%=Separate(SumPrice)%></TD>
				<TD dir=LTR align=right><%=Separate(SumDiscount)%></TD>
				<TD dir=LTR align=right><%=Separate(SumReverse)%></TD>
				<TD dir=LTR align=right><%=Separate(SumVat)%></TD>
				<TD dir=LTR align=right><%=Separate(SumSales)%></TD>
			</tr>
		</table>
		<br><br>
<%
	end if
	RS1.close
%>
	<table class="CustTable" cellspacing='1' style='width:90%;' align='center'>
		<tr>
			<td colspan="10" class="CusTableHeader" style="text-align:right;height:35;">¬Ì „ Â«Ì ›—Ê‘ <%=isAName %>&nbsp;<%=GroupName%>  (<%=pageTitle%>)</td>
		</tr>
<%
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
				<td colspan="10" class="CusTD3" style='direction:RTL;'>ÂÌç ¬Ì „Ì ÅÌœ« ‰‘œ.</td>
			</tr>
<%		else
%>			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3" style='direction:RTL;'>‰«„ ¬Ì „</td>
				<td class="CusTD3" style="direction:RTL;"> ⁄œ«œ A4</td>
				<td class="CusTD3"> ⁄œ«œ „ÊÀ—</td>
				<td class="CusTD3">Ã„⁄ „»·€</td>
				<td class="CusTD3">Ã„⁄  Œ›Ì›</td>
				<td class="CusTD3">Ã„⁄ »—ê‘ </td>
				<td class="CusTD3">Ã„⁄ „«·Ì« </td>
				<td class="CusTD3">Ã„⁄ Œ«·’ ›—Ê‘</td>
				<td class="CusTD3">ê—ÊÂ</td>
			</tr>
			<SCRIPT LANGUAGE="JavaScript">
			<!--
			function drill(item) {
				<%if Category=0 then 
					response.write "window.open('?act=show&category='+item+'&FromDate="& FromDate & "&ToDate=" & ToDate & "&isA=" & isA & "&SalesAction=" & SalesAction & "&SalesPerson=" & SalesPerson & "&ResultsInPage="&ResultsInPage&"');"
				else
					response.write "window.open('?act=showItemDetails&Item='+item+'&FromDate="& FromDate & "&ToDate=" & ToDate & "&isA=" & isA & "&SalesAction=" & SalesAction & "&SalesPerson=" & SalesPerson & "');"
				end if
				%>
				
			}
			function drillGroup(cat) {
				window.location('../AR/Rep_InvoiceItems.asp?act=show&category='+cat+"&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>&isA=<%=isA%>&SalesAction=<%=SalesAction%>&SalesPerson=<%=SalesPerson%>");
			}
			//-->
			</SCRIPT>
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
				if Category=0 then 
					A4Qtty =		cdbl(RS1("A4Qtty"))
					SumAppQtty =	cdbl(RS1("SumAppQtty"))
					SumPrice =		cdbl(RS1("SumPrice"))
					SumDiscount =	cdbl(RS1("SumDiscount"))
					SumReverse =	cdbl(RS1("SumReverse"))
					SumVat =		cdbl(RS1("SumVat"))
				else
					if RS1("IsReverse") then
					
						A4Qtty =		-cdbl(RS1("A4Qtty"))
						SumAppQtty =	-cdbl(RS1("SumAppQtty"))
						SumPrice =		-cdbl(RS1("SumPrice"))
						SumDiscount =	-cdbl(RS1("SumDiscount"))
						SumReverse =	-cdbl(RS1("SumReverse"))
						SumVat =		-cdbl(RS1("SumVat"))

						tmpColor="#FF9966"
					else
						A4Qtty =		cdbl(RS1("A4Qtty"))
						SumAppQtty =	cdbl(RS1("SumAppQtty"))
						SumPrice =		cdbl(RS1("SumPrice"))
						SumDiscount =	cdbl(RS1("SumDiscount"))
						SumReverse =	cdbl(RS1("SumReverse"))
						SumVat =		cdbl(RS1("SumVat"))
					end if
				end if
				SumSales =	SumPrice - SumDiscount - SumReverse '+ SumVat

%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand; height:30px;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="drill('<%=RS1("ID")%>');">
					<TD><%=tmpCounter%></TD>
					<TD><%=RS1("Name")%></TD>
					<TD dir=LTR align=right><%=Separate(A4Qtty)%></TD>
					<TD dir=LTR align=right><%=Separate(SumAppQtty)%></TD>
					<TD dir=LTR align=right><%=Separate(SumPrice)%></TD>
					<TD dir=LTR align=right><%=Separate(SumDiscount)%></TD>
					<TD dir=LTR align=right><%=Separate(SumReverse)%></TD>
					<TD dir=LTR align=right><%=Separate(SumVat)%></TD>
					<TD dir=LTR align=right><%=Separate(SumSales)%></TD>
					<td dir=LTR align=center><%=RS1("CategoryName")%></td>
				</TR>
<%				RS1.moveNext
			Loop

			if ToDate="9999/99/99" then ToDate="" 

			if TotalPages > 1 then
				pageCols=20
%>			
				<TR>
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
						window.location="?act=show&ResultsInPage=<%=ResultsInPage%>&p="+p+"&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>&Category=<%=Category%>";
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
elseif request("act")="showItemDetails" then
	ON ERROR RESUME NEXT

		Item =			clng(request("Item"))

		FromDate =		sqlSafe(request("FromDate"))
		ToDate =		sqlSafe(request("ToDate"))

		isA =			cint(request("isA"))

		ResultsInPage =	50
		
		if ToDate = "" then ToDate = "9999/99/99"

		if FromDate="" AND ToDate="9999/99/99" then
			pageTitle="»Â ’Ê—  ﬂ«„·"
		elseif FromDate="" then
			pageTitle="«“ «» œ«  «  «—ÌŒ " & replace (ToDate,"/",".")
		elseif ToDate="9999/99/99" then
			pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  « «‰ Â« "
		else
			pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  «  «—ÌŒ " & replace (ToDate,"/",".")
		end if
	
		SalesAction =	request("SalesAction")
		SalesPerson =	request("SalesPerson")

		if SalesAction="" then 
			SalesAction = 2 ' «ÌÌœ
		else
			SalesAction = cint(SalesAction)
		end if
		if SalesPerson="" then 
			SalesPerson = 0 'Â„Â
		else
			SalesPerson = cint(SalesPerson)
		end if

		if Err.Number<>0 then
			Err.clear

			conn.close
			response.redirect "OtherReports.asp?errMsg=" & Server.URLEncode("Œÿ« œ— Ê—ÊœÌ.")
		end if
	ON ERROR GOTO 0

	mySQL= "SELECT Name FROM InvoiceItems WHERE (ID = "& Item & ")"
	Set RS1 = Conn.Execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "OtherReports.asp?errMsg=" & Server.URLEncode("ç‰Ì‰ çÌ“Ì ÊÃÊœ ‰œ«—œ.")
	else
		ItemName=RS1("Name")
	end if
	RS1.close
	Set RS1 = Nothing
	otherCriteria = ""
	if isA = 1 then
		isAName = "»«  Ìﬂ «·›"
		otherCriteria = " AND (Invoices.isA=1) "
	elseif isA = 2 then
		isAName = "»œÊ‰  Ìﬂ «·›"
		otherCriteria = " AND (Invoices.isA=0) "
	end if

	if Auth("C" , 6) then 'ê“«—‘ ¬Ì „ Â«Ì ›—Ê‘ »—Õ”» ›—Ê‘‰œÂ
		if SalesPerson <> 0 then
			select case SalesAction
			case 0:
				otherCriteria = otherCriteria & " AND ((Invoices.CreatedBy=" & SalesPerson & ") OR (Invoices.ApprovedBy=" & SalesPerson & ") OR (Invoices.IssuedBy=" & SalesPerson & "))"
			case 1:
				otherCriteria = otherCriteria & " AND (Invoices.CreatedBy=" & SalesPerson & ")"
			case 2:
				otherCriteria = otherCriteria & " AND (Invoices.ApprovedBy=" & SalesPerson & ")"
			case 3:
				otherCriteria = otherCriteria & " AND (Invoices.IssuedBy=" & SalesPerson & ")"
			end select
		end if
	end if


%>
<br>
	<table class="CustTable" cellspacing='1' style='width:90%;' align='center'>
		<tr>
			<td colspan="10" class="CusTableHeader" style="text-align:right;height:35;">›«ﬂ Ê—   Â«Ì <%=isAName%> ‘«„·  '<%=ItemName%>' (<%=pageTitle%>)</td>
		</tr>
<%
		mySQL="SELECT Invoices.ID, Invoices.IssuedDate, Invoices.TotalReceivable, Invoices.IsReverse, Accounts.AccountTitle, SUM(InvoiceLines.Qtty * InvoiceLines.Sets * ROUND(InvoiceLines.Length * InvoiceLines.Width / 630 + .49, 0)) AS A4Qtty, SUM(InvoiceLines.AppQtty) AS SumAppQtty, SUM(InvoiceLines.Price - InvoiceLines.Discount - InvoiceLines.Reverse) AS SumReceivable, dbo.isInvoiceHasPaper(Invoices.ID) AS isPaper, dbo.isInvoiceHasHavale(Invoices.ID) AS isHavale FROM Invoices INNER JOIN InvoiceLines ON Invoices.ID = InvoiceLines.Invoice INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (InvoiceLines.Item = "& Item & ") AND (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'" & FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') " & otherCriteria & " GROUP BY Accounts.AccountTitle, Invoices.ID, Invoices.TotalReceivable, Invoices.IsReverse, Invoices.IssuedDate, Invoices.IssuedDate ORDER BY Invoices.IssuedDate DESC"
'response.write mySQL
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
				<td colspan="10" class="CusTD3" style='direction:RTL;'>ÂÌç ›«ﬂ Ê—Ì ÅÌœ« ‰‘œ.</td>
			</tr>
<%		else
%>			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3" style="direction:RTL;"># ›«ﬂ Ê—</td>
				<td class="CusTD3"> «—ÌŒ ’œÊ—</td>
				<td class="CusTD3">⁄‰Ê«‰ Õ”«»</td>
				<td class="CusTD3" style="direction:RTL;"> ⁄œ«œ A4</td>
				<td class="CusTD3"> ⁄œ«œ „ÊÀ—</td>
				<td class="CusTD3">„»·€ Œ«·’ ¬Ì „</td>
				<td class="CusTD3">„»·€ ›«ﬂ Ê—</td>
				<td class="CusTD3" width="5px">I</td>
				<td class="CusTD3" width="5px">II</td>
			</tr>
			<SCRIPT LANGUAGE="JavaScript">
			<!--
			function drill(inv) {
				window.open('../AR/AccountReport.asp?act=showInvoice&invoice='+inv);
			}
			
			//-->
			</SCRIPT>
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

				if RS1("IsReverse") then
					A4Qtty =			-cdbl(RS1("A4Qtty"))
					SumAppQtty =		-cdbl(RS1("SumAppQtty"))
					SumReceivable =		-cdbl(RS1("SumReceivable"))
					TotalReceivable =	-cdbl(RS1("TotalReceivable"))

					tmpColor="#FF9966"
				else
					A4Qtty =			cdbl(RS1("A4Qtty"))
					SumAppQtty =		cdbl(RS1("SumAppQtty"))
					SumReceivable =		cdbl(RS1("SumReceivable"))
					TotalReceivable =	cdbl(RS1("TotalReceivable"))
				end if

%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand; height:30px;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="drill('<%=RS1("ID")%>');">
					<TD><%=tmpCounter%></TD>
					<TD><%=RS1("ID")%></TD>
					<TD dir=LTR align=right><%=RS1("IssuedDate")%></TD>
					<TD><%=RS1("AccountTitle")%></TD>
					<TD dir=LTR align=right><%=Separate(A4Qtty)%></TD>

					<TD dir=LTR align=right><%=Separate(SumAppQtty)%></TD>
					<TD dir=LTR align=right><%=Separate(SumReceivable)%></TD>
					<TD dir=LTR align=right><%=Separate(TotalReceivable)%></TD>
					<td dir="ltr" align="center"><%if cbool(RS1("isPaper")) then response.write("*")%></td>
					<td dir="ltr" align="center"><%if cbool(RS1("isHavale")) then response.write("*")%></td>
				</TR>
<%				RS1.moveNext
			Loop

			if ToDate="9999/99/99" then ToDate="" 

			if TotalPages > 1 then
				pageCols=20
%>			
				<TR class="RepTableTitle">
					<TD bgcolor='#33AACC' height="30" colspan="10">
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
<div>
I: ›«ò Ê— ‘«„· ¬&#1740; „ ò«€–
</div>
<div>
II: ›«ò Ê— ‘«„· ÕÊ«·Â ò«€–
</div>
					<SCRIPT LANGUAGE="JavaScript">
					<!--
					function go2Page(p) {
						window.location="?act=showItemDetails&Item=<%=Item%>&ResultsInPage=<%=ResultsInPage%>&p="+p+"&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>&isA=<%=isA%>&SalesAction=<%=SalesAction%>&SalesPerson=<%=SalesPerson%>";
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
end if%>
<!--#include file="tah.asp" -->
