<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%
'Response.Buffer=false
Response.CodePage = 65001
Response.CharSet = "utf-8"

reportTitle = "گزارش جمع فروش و برگشت الف بر حسب مشتري"
%>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->


<!--#include File="reports_top.asp"-->
<%

	dateFrom = request("dateFrom")
	dateTo = request("dateTo")


	mySQL = "SELECT Accounts.ID AS كد, Accounts.AccountTitle AS [عنوان حساب], Accounts.City1 AS شهر, Accounts.Address1 AS آدرس, Accounts.EconomicalCode AS كداقتصادي, SUM(Invoices.IsReverse * Invoices.TotalReceivable) AS SumReverse, SUM((1 - Invoices.IsReverse) * Invoices.TotalReceivable) AS [جمع برگشت], COUNT(Accounts.AccountTitle) AS InvoiceCount FROM Accounts INNER JOIN Invoices ON Accounts.ID = Invoices.Customer WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.IssuedDate >= N'" & dateFrom & "') AND (Invoices.IssuedDate <= N'" & dateTo & "') AND (Invoices.IsA = 1) GROUP BY Accounts.AccountTitle, Accounts.EconomicalCode, Accounts.ID, Accounts.City1, Accounts.Address1, Accounts.Tel1, Accounts.Mobile1 ORDER BY Accounts.ID"
	
	'mySQL = "SELECT Accounts.ID AS كد, Accounts.AccountTitle AS [عنوان حساب], Accounts.City1 AS شهر, Accounts.Address1 AS آدرس, Accounts.Tel1 AS تلفن, Accounts.Mobile1 AS موبايل, Accounts.EconomicalCode AS كداقتصادي, SUM(Invoices.IsReverse * Invoices.TotalReceivable) AS [جمع برگشت], SUM((1 - Invoices.IsReverse) * Invoices.TotalReceivable) AS [جمع فاكتور], COUNT(Accounts.AccountTitle) AS [تعداد فاكتورها] FROM Accounts INNER JOIN Invoices ON Accounts.ID = Invoices.Customer WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.IssuedDate >= N'" & dateFrom & "') AND (Invoices.IssuedDate <= N'" & dateTo & "') AND (Invoices.IsA = 1) GROUP BY Accounts.AccountTitle, Accounts.EconomicalCode, Accounts.ID, Accounts.City1, Accounts.Address1, Accounts.Tel1, Accounts.Mobile1 ORDER BY Accounts.ID"
	
%>
<!--#include File="reports_tah.asp"-->
