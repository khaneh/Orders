<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%if request("act")="" then%>
	<!--#include File="include_farsiDateHandling.asp"-->
	<HTML>
	<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
	<meta http-equiv="Content-Language" content="fa">
	<style>
		Table { font-family:tahoma; font-size: 9pt;}
	</style>
	<TITLE><%if request("act")="show" then response.write "گزارش فروش مشتريان بر حسب آيتم هاي فروش در گروه هاي مختلف از تاريخ "& replace(FromDate,"/",".") & " تا "& replace(ToDate,"/",".") else response.write "Kid Invoice Items Categories"%></TITLE>
	</HEAD>
	<BODY>
	<!--#include File="include_JS_InputMasks.asp"-->
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

'	Response.ContentType = "application/msexcel"
'	Response.AddHeader "Content-Disposition","attachment; filename=Customers.txt"

	Separator=chr(9)

	%>شماره حساب<%=Separator%>عنوان حساب<%=Separator%>گروه 1 (خدمات بيروني)<%=Separator%>تعداد<%=Separator%>گروه 2 (كاغذ افست)<%=Separator%>تعداد<%=Separator%>گروه 3 (افست)<%=Separator%>تعداد<%=Separator%>گروه 4 (ديجيتال)<%=Separator%>تعداد<%=Separator%>گروه 5 (ليتو گرافي)<%=Separator%>تعداد<%=Separator%>گروه 6 (پيرامون)<%=Separator%>تعداد<%=Separator%>گروه 7 (عمليات تكميلي)<%=Separator%>تعداد<%=Separator%>گروه 8 (طراحي)<%=Separator%>تعداد<%=Separator%>گروه 9 (آي تي)<%=Separator%>تعداد<%=Separator%>گروه 10 (پلات)<%=Separator%>تعداد<%=Separator%>گروه 11 (ساير خدمات )<%=Separator%>تعداد<%=Separator%>گروه 13 (ديجيتال سياه و سفيد)<%=Separator%>تعداد<%=Separator%>تعداد كل فاكتور ها<%=Separator%>جمع كل قابل پرداخت<%=Separator%>جمع كل تخفيف<%=Separator%>جمع كل برگشت<%=vbCrLf%>
<%

'	conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open conStr

	'mySQL="SELECT Kid.AccID,Kid.AccountTitle,Kid.Price,Kid.Discount,Kid.Reverse,Kid.CatID,Kid.CatName,Kid.InvQtty,COUNT(*) AS TotalInvQtty, SUM(Invoices.TotalReceivable) AS TotalReceivable,SUM(Invoices.TotalDiscount) AS TotalDiscount,SUM(Invoices.TotalReverse)  AS TotalReverse FROM (SELECT Accounts.ID AS AccID,Accounts.AccountTitle,SUM(InvoiceLines.Price) AS Price,SUM(InvoiceLines.Discount) AS Discount, SUM(InvoiceLines.Reverse) AS Reverse,InvoiceItemCategories.ID AS CatID,InvoiceItemCategories.Name AS CatName,COUNT(*) AS InvQtty FROM InvoiceItemCategoryRelations INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceItemCategories ON InvoiceItemCategoryRelations.InvoiceItemCategory = InvoiceItemCategories.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) GROUP BY Accounts.AccountTitle,InvoiceItemCategories.ID,InvoiceItemCategories.Name,Accounts.ID HAVING (InvoiceItemCategories.ID > 0)) Kid INNER JOIN Invoices ON Kid.AccID = Invoices.Customer GROUP BY Kid.AccID,Kid.AccountTitle,Kid.Price,Kid.Discount,Kid.Reverse,Kid.CatID,Kid.CatName,Kid.InvQtty,Invoices.Issued,Invoices.Voided HAVING (Invoices.Issued = 1) AND (Invoices.Voided = 0) ORDER BY Kid.AccID,Kid.CatID"

	mySQL="SELECT Kid.AccID,Kid.AccountTitle,Kid.Price,Kid.Discount,Kid.Reverse,Kid.CatID,Kid.CatName,Kid.InvQtty,COUNT(*) AS TotalInvQtty, SUM(Invoices.TotalReceivable) AS TotalReceivable,SUM(Invoices.TotalDiscount) AS TotalDiscount,SUM(Invoices.TotalReverse)  AS TotalReverse FROM (SELECT Accounts.ID AS AccID,Accounts.AccountTitle,SUM(InvoiceLines.Price) AS Price,SUM(InvoiceLines.Discount) AS Discount, SUM(InvoiceLines.Reverse) AS Reverse,InvoiceItemCategories.ID AS CatID,InvoiceItemCategories.Name AS CatName,COUNT(DISTINCT Invoices.ID) AS InvQtty FROM InvoiceItemCategoryRelations INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceItemCategories ON InvoiceItemCategoryRelations.InvoiceItemCategory = InvoiceItemCategories.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') GROUP BY Accounts.AccountTitle,InvoiceItemCategories.ID,InvoiceItemCategories.Name,Accounts.ID HAVING (InvoiceItemCategories.ID > 0)) Kid INNER JOIN Invoices ON Kid.AccID = Invoices.Customer WHERE (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') GROUP BY Kid.AccID,Kid.AccountTitle,Kid.Price,Kid.Discount,Kid.Reverse,Kid.CatID,Kid.CatName,Kid.InvQtty,Invoices.Issued,Invoices.Voided HAVING (Invoices.Issued = 1) AND (Invoices.Voided = 0) ORDER BY Kid.AccID,Kid.CatID"

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

	Do While NOT RS1.EOF 'and tmpCounter < 10
		tmpCounter=tmpCounter+1
		AccID			=RS1("AccID")
		if AccID<>LastAcc AND LastAcc <> 0 then
			'write last row
			%><%=LastAcc%><%=Separator%><%=AccountTitle%><%=Separator%><%=Col1(1)%><%=Separator%><%=Col2(1)%><%=Separator%><%=Col1(2)%><%=Separator%><%=Col2(2)%><%=Separator%><%=Col1(3)%><%=Separator%><%=Col2(3)%><%=Separator%><%=Col1(4)%><%=Separator%><%=Col2(4)%><%=Separator%><%=Col1(5)%><%=Separator%><%=Col2(5)%><%=Separator%><%=Col1(6)%><%=Separator%><%=Col2(6)%><%=Separator%><%=Col1(7)%><%=Separator%><%=Col2(7)%><%=Separator%><%=Col1(8)%><%=Separator%><%=Col2(8)%><%=Separator%><%=Col1(9)%><%=Separator%><%=Col2(9)%><%=Separator%><%=Col1(10)%><%=Separator%><%=Col2(10)%><%=Separator%><%=Col1(11)%><%=Separator%><%=Col2(11)%><%=Separator%><%=Col1(13)%><%=Separator%><%=Col2(13)%><%=Separator%><%=TotalInvQtty%><%=Separator%><%=TotalReceivable%><%=Separator%><%=TotalDiscount%><%=Separator%><%=TotalReverse%><%=vbCrLf%>
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

		Col1(CatID)=Price-Discount-Reverse
		Col2(CatID)=InvQtty

		LastAcc=AccID
		RS1.moveNext
	Loop
			%><%=LastAcc%><%=Separator%><%=AccountTitle%><%=Separator%><%=Col1(1)%><%=Separator%><%=Col2(1)%><%=Separator%><%=Col1(2)%><%=Separator%><%=Col2(2)%><%=Separator%><%=Col1(3)%><%=Separator%><%=Col2(3)%><%=Separator%><%=Col1(4)%><%=Separator%><%=Col2(4)%><%=Separator%><%=Col1(5)%><%=Separator%><%=Col2(5)%><%=Separator%><%=Col1(6)%><%=Separator%><%=Col2(6)%><%=Separator%><%=Col1(7)%><%=Separator%><%=Col2(7)%><%=Separator%><%=Col1(8)%><%=Separator%><%=Col2(8)%><%=Separator%><%=Col1(9)%><%=Separator%><%=Col2(9)%><%=Separator%><%=Col1(10)%><%=Separator%><%=Col2(10)%><%=Separator%><%=Col1(11)%><%=Separator%><%=Col2(11)%><%=Separator%><%=Col1(13)%><%=Separator%><%=Col2(13)%><%=Separator%><%=TotalInvQtty%><%=Separator%><%=TotalReceivable%><%=Separator%><%=TotalDiscount%><%=Separator%><%=TotalReverse%><%=vbCrLf%>
<%
end if
%>