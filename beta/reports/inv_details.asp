<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%
'Response.Buffer=false
Response.CodePage = 65001
Response.CharSet = "utf-8"

reportTitle = "گزارش جزييات خروج از انبار"
%>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include File="reports_top.asp"-->
<%

	dateFrom = request("dateFrom")
	dateTo = request("dateTo")

	mySQL = "SELECT dbo.InventoryItems.OldItemID AS [كد كالا], dbo.InventoryItems.Name AS [نام كالا], dbo.InventoryLog.Qtty AS تعداد, dbo.InventoryLog.logDate AS [تاريخ خروج], dbo.InventoryLog.owner AS مالك, dbo.InventoryLog.comments AS توضيحات, dbo.InventoryLog.RelatedInvoiceID AS [فاكتور فروش], dbo.Users.RealName AS گيرنده, Users_1.RealName AS [مسئول انبار] FROM dbo.InventoryLog INNER Join dbo.InventoryPickuplists ON dbo.InventoryLog.RelatedID = dbo.InventoryPickuplists.id INNER Join dbo.Users ON dbo.InventoryPickuplists.GiveTo = dbo.Users.ID INNER Join dbo.InventoryItems ON dbo.InventoryLog.ItemID = dbo.InventoryItems.ID INNER Join dbo.Users Users_1 ON dbo.InventoryLog.CreatedBy = Users_1.ID WHERE (dbo.InventoryLog.IsInput = 0) AND (dbo.InventoryLog.Voided = 0) AND (dbo.InventoryLog.logDate >= N'" & dateFrom & "') AND (dbo.InventoryLog.logDate <= N'" & dateTo & "')"
	
%>
<!--#include File="reports_tah.asp"-->
