<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%if request("act")="" then%>
	<!--#include File="../include_farsiDateHandling.asp"-->
	<HTML>
	<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<!-- 	<meta http-equiv="Content-Language" content="fa"> -->
	<style>
		Table { font-family:tahoma; font-size: 9pt;}
	</style>
	<TITLE><%if request("act")="show" then response.write "گزارش فروش مشتريان بر حسب آيتم هاي فروش در گروه هاي مختلف از تاريخ "& replace(FromDate,"/",".") & " تا "& replace(ToDate,"/",".") else response.write "Kid Invoice Items Categories"%></TITLE>
	</HEAD>
	<BODY>
	<!--#include File="../include_JS_InputMasks.asp"-->
	<FORM METHOD=POST ACTION="?act=show">
	<TABLE border=1 align=center dir="RTL">
	<TR>
		<TD align=left>از تاريخ</TD>
		<TD><INPUT dir="LTR" class="GenInput" NAME="FromDate" tabIndex="3" TYPE="text" maxlength="10" size="10" value="<%=request("FromDate")%>" onBlur="acceptDate(this);" style="width:100px;"></TD>
	</TR>
	<TR>
		<TD align=left>تا تاريخ</TD>
		<TD><INPUT dir="LTR" class="GenInput" NAME="ToDate" tabIndex="4" TYPE="text" maxlength="10" size="10" value="<%=request("ToDate")%>" onBlur="acceptDate(this)" style="width:100px;"></TD>
	</TR>
	<TR>
		<TD align=center colspan=2><INPUT style="font-family:tahoma;" TYPE="Submit" Value=" ادامه " tabindex=6></TD>
	</TR>
	</FORM>
	</TABLE>
	<br>
<%
elseif request("act")="show" then
	FromDate=	request("FromDate")
	ToDate=		request("ToDate")
	reportTitle = "گزارش فروش مشتريان بر حسب آيتم هاي فروش در گروه هاي مختلف"
	%>
	<html>
	<head>
	<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1256">
	<title><%=reportTitle%></title>
	</head>
	<body>
	<%
	function Separate(inputTxt)

	if not isnumeric(inputTxt) or "" & inputTxt="" then 
		Separate=inputTxt
	else
		myMinus=""
		input=inputTxt
		t=instr(input, ".")
		if t>0 then 
			expPart = mid(input, t+1, 2)
			input = left(input, t-1)
		end if
		if left(input,1)="-" then
			myMinus="-"
			input=right(input,len(input)-1)
		end if
		if len(input) > 3 then
			tmpr=right(input ,3)
			tmpl=left(input , len(input) - 3 )
			result = tmpr
			while len(tmpl) > 3
				tmpr=right(tmpl,3)
				result = tmpr & "," & result 
				tmpl=left(tmpl , len(tmpl) - 3 )
			wend
			if len(tmpl) > 0 then
				result = tmpl & "," & result
			end if 
		else
			result = input
		end if 
		if t>0 then 
			result = result & "." & expPart
		end if

		Separate=myMinus & result
	end if
	end function

'---------------------

'	conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open conStr

	'mySQL="SELECT Kid.AccID,Kid.AccountTitle,Kid.Price,Kid.Discount,Kid.Reverse,Kid.CatID,Kid.CatName,Kid.InvQtty,COUNT(*) AS TotalInvQtty, SUM(Invoices.TotalReceivable) AS TotalReceivable,SUM(Invoices.TotalDiscount) AS TotalDiscount,SUM(Invoices.TotalReverse)  AS TotalReverse FROM (SELECT Accounts.ID AS AccID,Accounts.AccountTitle,SUM(InvoiceLines.Price) AS Price,SUM(InvoiceLines.Discount) AS Discount, SUM(InvoiceLines.Reverse) AS Reverse,InvoiceItemCategories.ID AS CatID,InvoiceItemCategories.Name AS CatName,COUNT(*) AS InvQtty FROM InvoiceItemCategoryRelations INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceItemCategories ON InvoiceItemCategoryRelations.InvoiceItemCategory = InvoiceItemCategories.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) GROUP BY Accounts.AccountTitle,InvoiceItemCategories.ID,InvoiceItemCategories.Name,Accounts.ID HAVING (InvoiceItemCategories.ID > 0)) Kid INNER JOIN Invoices ON Kid.AccID = Invoices.Customer GROUP BY Kid.AccID,Kid.AccountTitle,Kid.Price,Kid.Discount,Kid.Reverse,Kid.CatID,Kid.CatName,Kid.InvQtty,Invoices.Issued,Invoices.Voided HAVING (Invoices.Issued = 1) AND (Invoices.Voided = 0) ORDER BY Kid.AccID,Kid.CatID"

	'mySQL="SELECT Kid.AccID,Kid.AccountTitle,Kid.Price,Kid.Discount,Kid.Reverse,Kid.CatID,Kid.CatName,Kid.InvQtty,COUNT(*) AS TotalInvQtty, SUM(Invoices.TotalReceivable) AS TotalReceivable,SUM(Invoices.TotalDiscount) AS TotalDiscount,SUM(Invoices.TotalReverse)  AS TotalReverse FROM (SELECT Accounts.ID AS AccID,Accounts.AccountTitle,SUM(InvoiceLines.Price) AS Price,SUM(InvoiceLines.Discount) AS Discount, SUM(InvoiceLines.Reverse) AS Reverse,InvoiceItemCategories.ID AS CatID,InvoiceItemCategories.Name AS CatName,COUNT(DISTINCT Invoices.ID) AS InvQtty FROM InvoiceItemCategoryRelations INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceItemCategories ON InvoiceItemCategoryRelations.InvoiceItemCategory = InvoiceItemCategories.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') GROUP BY Accounts.AccountTitle,InvoiceItemCategories.ID,InvoiceItemCategories.Name,Accounts.ID HAVING (InvoiceItemCategories.ID > 0)) Kid INNER JOIN Invoices ON Kid.AccID = Invoices.Customer WHERE (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') GROUP BY Kid.AccID,Kid.AccountTitle,Kid.Price,Kid.Discount,Kid.Reverse,Kid.CatID,Kid.CatName,Kid.InvQtty,Invoices.Issued,Invoices.Voided HAVING (Invoices.Issued = 1) AND (Invoices.Voided = 0) ORDER BY Kid.AccID,Kid.CatID"
	' coment by Alix - 84-10-10 (add CSRname column to query)

	mySQL="SELECT Kid.AccID, Kid.AccountTitle, Kid.Price, Kid.Discount, Kid.Reverse, Kid.CatID, Kid.CatName, Kid.InvQtty, COUNT(*) AS TotalInvQtty, SUM(Invoices.TotalReceivable) AS TotalReceivable, SUM(Invoices.TotalDiscount) AS TotalDiscount, SUM(Invoices.TotalReverse) AS TotalReverse, Kid.CSRname FROM (SELECT Accounts.ID AS AccID, Accounts.AccountTitle, SUM(InvoiceLines.Price) AS Price, SUM(InvoiceLines.Discount) AS Discount, SUM(InvoiceLines.Reverse) AS Reverse, InvoiceItemCategories.ID AS CatID, InvoiceItemCategories.Name AS CatName, COUNT(DISTINCT Invoices.ID) AS InvQtty, Users.RealName AS CSRname FROM InvoiceItemCategoryRelations INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceItemCategories ON InvoiceItemCategoryRelations.InvoiceItemCategory = InvoiceItemCategories.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID INNER JOIN Users ON Accounts.CSR = Users.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') GROUP BY Accounts.AccountTitle, InvoiceItemCategories.ID, InvoiceItemCategories.Name, Accounts.ID, Users.RealName HAVING (InvoiceItemCategories.ID > 0)) Kid INNER JOIN Invoices ON Kid.AccID = Invoices.Customer WHERE (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') GROUP BY Kid.AccID, Kid.AccountTitle, Kid.Price, Kid.Discount, Kid.Reverse, Kid.CatID, Kid.CatName, Kid.InvQtty, Invoices.Issued, Invoices.Voided, Kid.CSRname HAVING (Invoices.Issued = 1) AND (Invoices.Voided = 0) ORDER BY Kid.AccID, Kid.CatID"

'------------------------------------------------------------------------------------------------------------------------------------

	Set RS1 = conn.Execute(mySQL)

	cols=20
	Dim Col1(20)
	Dim Col2(20)
	for i=1 to cols
		Col1(i)=""
		Col2(i)=""
	next
	LastAcc=0
	tmpCounter=0

	%>
	<style>
		.resTable1 { Font-family:tahoma; Font-Size:9pt; Background-color:#336699; border: 1 solid #336699;}
		.resTable1 th {Background-color:#6699CC;}
		.resRow1 {Background-color:#F8F8FF;}
		.resRow0 {Background-color:#CCCCDD;}
	</style>

	<CENTER><H3><%=reportTitle%></H3></CENTER>
	<table class="resTable1" Cellspacing="1" Cellpadding="2" align=center>
		<th>شماره حساب</th>
		<th>عنوان حساب</th>
		<th>گروه 1 (خدمات بيروني)</th>
		<th>تعداد</th>
		<th>گروه 2 (كاغذ افست)</th>
		<th>تعداد</th>
		<th>گروه 3 (افست)</th>
		<th>تعداد</th>
		<th>گروه 4 (ديجيتال)</th>
		<th>تعداد</th>
		<th>گروه 5 (ليتو گرافي)</th>
		<th>تعداد</th>
		<th>گروه 6 (پيرامون)</th>
		<th>تعداد</th>
		<th>گروه 7 (عمليات تكميلي)</th>
		<th>تعداد</th>
		<th>گروه 8 (طراحي)</th>
		<th>تعداد</th>
		<th>گروه 9 (آي تي)</th>
		<th>تعداد</th>
		<th>گروه 10 (پلات)</th>
		<th>تعداد</th>
		<th>گروه 11 (ساير خدمات )</th>
		<th>تعداد</th>
		<th>گروه 13 (ديجيتال سياه و سفيد)</th>
		<th>تعداد</th>
		<th>گروه 14 (عمليات تكميلي ديجيتال)</th>
		<th>تعداد</th>
		<th>تعداد كل فاكتور ها</th>
		<th>جمع كل قابل پرداخت</th>
		<th>جمع كل تخفيف</th>
		<th>جمع كل برگشت</th>
		<th>مسئول پيگيري</th>

		</tr>
	<%
		if RS1.EOF then
	%>
		<tr>
			<td width="100%" class="resRow1" colspan="33" align="center">
				جوابي يافت نشد
			</td>
		</tr>
	<%
		else
			Do While NOT RS1.EOF 'and tmpCounter < 10	
				AccID			=RS1("AccID")
				if AccID<>LastAcc AND LastAcc <> 0 then
					'write last row
					tmpCounter=tmpCounter+1
				%>
				<tr class="resRow<%=tmpCounter mod 2%>">
					<td><%=LastAcc%></td>
					<td><%=Server.HTMLEncode(AccountTitle)%></td>
					<td><%=Separate(Col1(1))%></td>
					<td><%=Separate(Col2(1))%></td>
					<td><%=Separate(Col1(2))%></td>
					<td><%=Separate(Col2(2))%></td>
					<td><%=Separate(Col1(3))%></td>
					<td><%=Separate(Col2(3))%></td>
					<td><%=Separate(Col1(4))%></td>
					<td><%=Separate(Col2(4))%></td>
					<td><%=Separate(Col1(5))%></td>
					<td><%=Separate(Col2(5))%></td>
					<td><%=Separate(Col1(6))%></td>
					<td><%=Separate(Col2(6))%></td>
					<td><%=Separate(Col1(7))%></td>
					<td><%=Separate(Col2(7))%></td>
					<td><%=Separate(Col1(8))%></td>
					<td><%=Separate(Col2(8))%></td>
					<td><%=Separate(Col1(9))%></td>
					<td><%=Separate(Col2(9))%></td>
					<td><%=Separate(Col1(10))%></td>
					<td><%=Separate(Col2(10))%></td>
					<td><%=Separate(Col1(11))%></td>
					<td><%=Separate(Col2(11))%></td>
					<td><%=Separate(Col1(13))%></td>
					<td><%=Separate(Col2(13))%></td>
					<td><%=Separate(Col1(14))%></td>
					<td><%=Separate(Col2(14))%></td>
					<td><%=Separate(TotalInvQtty)%></td>
					<td><%=Separate(TotalReceivable)%></td>
					<td><%=Separate(TotalDiscount)%></td>
					<td><%=Separate(TotalReverse)%></td>
					<td><%=Server.HTMLEncode(CSRname)%></td>	
				</tr>
		<%
					for i=1 to cols
						Col1(i)=""
						Col2(i)=""
					next
				end if

				AccountTitle	=RS1("AccountTitle")
				Price			=cdbl(RS1("Price"))
				Discount		=cdbl(RS1("Discount"))
				Reverse			=cdbl(RS1("Reverse"))
				CatID			=cint(RS1("CatID"))
				CatName			=RS1("CatName")
				InvQtty			=clng(RS1("InvQtty"))
				TotalInvQtty	=clng(RS1("TotalInvQtty"))
				TotalReceivable	=cdbl(RS1("TotalReceivable"))
				TotalDiscount	=cdbl(RS1("TotalDiscount"))
				TotalReverse	=cdbl(RS1("TotalReverse"))
				CSRname			=RS1("CSRname")

				Col1(CatID)=Price-Discount-Reverse
				Col2(CatID)=InvQtty

				LastAcc=AccID
				RS1.moveNext
			Loop
			tmpCounter=tmpCounter+1
			%>
				<tr class="resRow<%=tmpCounter mod 2%>">
					<td><%=LastAcc%></td>
					<td><%=Server.HTMLEncode(AccountTitle)%></td>
					<td><%=Separate(Col1(1))%></td>
					<td><%=Separate(Col2(1))%></td>
					<td><%=Separate(Col1(2))%></td>
					<td><%=Separate(Col2(2))%></td>
					<td><%=Separate(Col1(3))%></td>
					<td><%=Separate(Col2(3))%></td>
					<td><%=Separate(Col1(4))%></td>
					<td><%=Separate(Col2(4))%></td>
					<td><%=Separate(Col1(5))%></td>
					<td><%=Separate(Col2(5))%></td>
					<td><%=Separate(Col1(6))%></td>
					<td><%=Separate(Col2(6))%></td>
					<td><%=Separate(Col1(7))%></td>
					<td><%=Separate(Col2(7))%></td>
					<td><%=Separate(Col1(8))%></td>
					<td><%=Separate(Col2(8))%></td>
					<td><%=Separate(Col1(9))%></td>
					<td><%=Separate(Col2(9))%></td>
					<td><%=Separate(Col1(10))%></td>
					<td><%=Separate(Col2(10))%></td>
					<td><%=Separate(Col1(11))%></td>
					<td><%=Separate(Col2(11))%></td>
					<td><%=Separate(Col1(13))%></td>
					<td><%=Separate(Col2(13))%></td>
					<td><%=Separate(Col1(14))%></td>
					<td><%=Separate(Col2(14))%></td>
					<td><%=Separate(TotalInvQtty)%></td>
					<td><%=Separate(TotalReceivable)%></td>
					<td><%=Separate(TotalDiscount)%></td>
					<td><%=Separate(TotalReverse)%></td>
					<td><%=Server.HTMLEncode(CSRname)%></td>
				</tr>
	</table>
<%

end if

'------------------------------------------------------------------------------------------------------------------------------------
end if
%>
</body>
</html>