<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%
Response.CodePage = 65001
Response.CharSet = "utf-8"

'Response.Buffer=false
reportTitle = "گزارش آيتم هاي فروش بر حسب مشتري"
%>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include File="reports_top.asp"-->
<%
	dateFrom = request("dateFrom")
	dateTo = request("dateTo")
	salses_items = request("salses_items")

	mySQL = "SELECT InvoiceLines.Item AS [كد آيتم فروش], SUM(InvoiceLines.AppQtty) AS [تعداد فروش], COUNT(InvoiceLines.AppQtty) AS [تعداد فاكتور], SUM(InvoiceLines.Price)  AS قيمت, SUM(InvoiceLines.Discount) AS تخفيف, SUM(InvoiceLines.Reverse) AS برگشت,SUM(InvoiceLines.Vat) AS ماليات , Accounts.ID AS [حساب مشتري], Accounts.AccountTitle AS [نام مشتري],  Users.RealName AS [مسئول حساب] FROM InvoiceLines INNER JOIN Invoices ON InvoiceLines.Invoice = Invoices.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID INNER JOIN Users ON Accounts.CSR = Users.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (InvoiceLines.Item IN (" & salses_items & ")) AND (Invoices.IssuedDate >= N'" & dateFrom & "') AND (Invoices.IssuedDate <= N'" & dateTo & "') GROUP BY InvoiceLines.Item, Accounts.ID, Accounts.AccountTitle, Users.RealName ORDER BY Accounts.AccountTitle"

%>
<!--#include File="reports_tah.asp"-->


