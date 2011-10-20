<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%
'Response.Buffer=false
Response.CodePage = 65001
Response.CharSet = "utf-8"
reportTitle = "گزارش تاخير سفارش ها"
%>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include File="reports_top.asp"-->
<%
	dateFrom = request("dateFrom")
	dateTo = request("dateTo")

	mySQL = "SELECT orders_trace.radif_sefareshat AS [شماره سفارش], orders_trace.salesperson AS [سفارش گيرنده], orders_trace.order_date AS [تاريخ سفارش], orders_trace.return_date AS [موعد تحويل], OrderTraceLog.InsertedDate AS [تاريخ آماده تحويل], dbo.diffDate(OrderTraceLog.InsertedDate,  orders_trace.return_date) AS تاخير, orders_trace.order_kind AS [نوع سفارش], orders_trace.Type AS [كد نوع] FROM OrderTraceLog RIGHT OUTER JOIN orders_trace ON OrderTraceLog.[Order] = orders_trace.radif_sefareshat WHERE (OrderTraceLog.Step = 10) AND (orders_trace.order_date >= N'" & dateFrom & "') AND (orders_trace.order_date <= N'" & dateTo & "')"
	
%>
<!--#include File="reports_tah.asp"-->


