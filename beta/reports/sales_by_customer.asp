<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Response.Buffer=false
reportTitle = "ÒÇÑÔ ÝÑæÔ ÈÑ ÍÓÈ ãÔÊÑí"
%>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include File="reports_top.asp"-->
<%
	dateFrom = request("dateFrom")
	dateTo = request("dateTo")

'	mySQL = "SELECT dbo.Accounts.ID AS [Account ID], SUM(dbo.Invoices.TotalPrice) AS TotalPrice, SUM(dbo.Invoices.TotalDiscount) AS TotalDiscount, SUM(dbo.Invoices.TotalReverse) AS TotalReverse, SUM(dbo.Invoices.TotalReceivable) AS TotalReceivable, dbo.Users.RealName AS CSR ,dbo.Accounts.AccountTitle AS [Account Title], dbo.Accounts.CompanyName, dbo.Accounts.Dear1, dbo.Accounts.FirstName1, dbo.Accounts.LastName1, dbo.Accounts.City1, dbo.Accounts.Address1, dbo.Accounts.PostCode1 FROM dbo.Invoices INNER JOIN dbo.Accounts ON dbo.Invoices.Customer = dbo.Accounts.ID INNER JOIN dbo.Users ON dbo.Accounts.CSR = dbo.Users.ID WHERE (dbo.Invoices.Voided = 0) AND (dbo.Invoices.Issued = 1) AND (dbo.Invoices.IssuedDate >= N'" & dateFrom & "') AND (dbo.Invoices.IssuedDate <= N'" & dateTo & "') and postable1 = 1 GROUP BY dbo.Accounts.ID, dbo.Accounts.AccountTitle, dbo.Users.RealName, dbo.Accounts.CompanyName, dbo.Accounts.Dear1, dbo.Accounts.FirstName1, dbo.Accounts.LastName1, dbo.Accounts.City1, dbo.Accounts.Address1, dbo.Accounts.PostCode1, dbo.Accounts.Dear2, dbo.Accounts.FirstName2, dbo.Accounts.LastName2, dbo.Accounts.City2, dbo.Accounts.Address2, dbo.Accounts.PostCode2 ORDER BY dbo.Accounts.AccountTitle"

'	mySQL = "SELECT Accounts.ID AS [Account ID], SUM(Invoices.TotalPrice) AS TotalPrice, SUM(Invoices.TotalDiscount) AS TotalDiscount, SUM(Invoices.TotalReverse) AS TotalReverse, SUM(Invoices.TotalReceivable) AS TotalReceivable, ARItems.RemainedAmount, Users.RealName AS CSR, Accounts.AccountTitle AS [Account Title], Accounts.CompanyName, Accounts.Dear1, Accounts.FirstName1, Accounts.LastName1, Accounts.City1, Accounts.Address1, Accounts.PostCode1 FROM Invoices INNER JOIN Accounts ON Invoices.Customer = Accounts.ID INNER JOIN Users ON Accounts.CSR = Users.ID INNER JOIN ARItems ON Invoices.ID = ARItems.Link WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'" & dateFrom & "') AND (Invoices.IssuedDate <= N'" & dateTo & "') AND (Accounts.Postable1 = 1) AND (ARItems.Type = 1) GROUP BY Accounts.ID, Accounts.AccountTitle, Users.RealName, Accounts.CompanyName, Accounts.Dear1, Accounts.FirstName1, Accounts.LastName1, Accounts.City1, Accounts.Address1, Accounts.PostCode1, Accounts.Dear2, Accounts.FirstName2, Accounts.LastName2, Accounts.City2, Accounts.Address2, Accounts.PostCode2, ARItems.RemainedAmount ORDER BY Accounts.AccountTitle"
'S A M
	mySQL = "SELECT Accounts.ID AS [Account ID], SUM(Invoices.TotalPrice * - (2 * CONVERT(tinyint, Invoices.IsReverse) - 1)) AS TotalPrice, SUM(Invoices.TotalDiscount * - (2 * CONVERT(tinyint, Invoices.IsReverse) - 1)) AS TotalDiscount, SUM(Invoices.TotalReverse * - (2 * CONVERT(tinyint, Invoices.IsReverse) - 1)) AS TotalReverse, SUM(Invoices.TotalReceivable * - (2 * CONVERT(tinyint, Invoices.IsReverse) - 1)) AS TotalReceivable, SUM(Invoices.TotalVat * - (2 * CONVERT(tinyint,Invoices.IsReverse) - 1)) AS TotalVat, SUM(ARItems.RemainedAmount * - (2 * CONVERT(tinyint, Invoices.IsReverse) - 1)) AS SumRemainedAmount, Users.RealName AS CSR, Accounts.AccountTitle AS [Account Title], Accounts.CompanyName, Accounts.Dear1, Accounts.FirstName1, Accounts.LastName1, Accounts.City1, Accounts.Address1, Accounts.PostCode1 FROM Invoices INNER JOIN Accounts ON Invoices.Customer = Accounts.ID INNER JOIN Users ON Accounts.CSR = Users.ID INNER JOIN ARItems ON Invoices.ID = ARItems.Link WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedDate >= N'" & dateFrom & "') AND (Invoices.IssuedDate <= N'" & dateTo & "') AND (ARItems.Type = 1 OR ARItems.Type = 4) GROUP BY Accounts.ID, Accounts.AccountTitle, Users.RealName, Accounts.CompanyName, Accounts.Dear1, Accounts.FirstName1, Accounts.LastName1, Accounts.City1, Accounts.Address1, Accounts.PostCode1, Accounts.Dear2, Accounts.FirstName2, Accounts.LastName2, Accounts.City2, Accounts.Address2, Accounts.PostCode2 ORDER BY Accounts.AccountTitle"







%>
<!--#include File="reports_tah.asp"-->



