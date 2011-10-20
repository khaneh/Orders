<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%
Response.CodePage = 65001
Response.CharSet = "utf-8"

reportTitle = "گزارش انواع فروش بر حسب مشتريان"
%>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include File="reports_top.asp"-->

<%
	FromDate=	request("FromDate")
	ToDate=		request("ToDate")

	mySQL="SELECT Kid.AccID as [Account Code], Kid.AccountTitle as [Account Name], Kid.Price as [Price], Kid.Discount as [تخفیف], Kid.Reverse as [Returns], Kid.CatID as [Category ID], Kid.CatName as [Category Name], Kid.InvQtty as [Invoice Qtty], COUNT(*) AS [Total Invoice Qtty], SUM(Invoices.TotalReceivable) AS [Total Recivable], SUM(Invoices.TotalDiscount) AS [Total Discount], SUM(Invoices.TotalReverse) AS [Total Returns], Kid.CSRname FROM (SELECT Accounts.ID AS AccID, Accounts.AccountTitle, SUM(InvoiceLines.Price) AS Price, SUM(InvoiceLines.Discount) AS Discount, SUM(InvoiceLines.Reverse) AS Reverse, InvoiceItemCategories.ID AS CatID, InvoiceItemCategories.Name AS CatName, COUNT(DISTINCT Invoices.ID) AS InvQtty, Users.RealName AS CSRname FROM InvoiceItemCategoryRelations INNER JOIN InvoiceItems ON InvoiceItemCategoryRelations.InvoiceItem = InvoiceItems.ID INNER JOIN InvoiceItemCategories ON InvoiceItemCategoryRelations.InvoiceItemCategory = InvoiceItemCategories.ID INNER JOIN InvoiceLines ON InvoiceItems.ID = InvoiceLines.Item INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID INNER JOIN Users ON Accounts.CSR = Users.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') GROUP BY Accounts.AccountTitle, InvoiceItemCategories.ID, InvoiceItemCategories.Name, Accounts.ID, Users.RealName HAVING (InvoiceItemCategories.ID > 0)) Kid INNER JOIN Invoices ON Kid.AccID = Invoices.Customer WHERE (Invoices.IssuedDate >= N'"& FromDate & "') AND (Invoices.IssuedDate <= N'"& ToDate & "') GROUP BY Kid.AccID, Kid.AccountTitle, Kid.Price, Kid.Discount, Kid.Reverse, Kid.CatID, Kid.CatName, Kid.InvQtty, Invoices.Issued, Invoices.Voided, Kid.CSRname HAVING (Invoices.Issued = 1) AND (Invoices.Voided = 0) ORDER BY Kid.AccID, Kid.CatID"
%>
<!--#include File="reports_tah.asp"-->
