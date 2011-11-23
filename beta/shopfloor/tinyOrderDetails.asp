<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
response.buffer= true


%>
<!--#include file="../config.asp" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	Table { font-size: 9pt;}
</style>
<TITLE>وضعيت سفارش</TITLE></TITLE>
</HEAD>
<SCRIPT LANGUAGE="JavaScript">
<!--
function documentKeyDown() {
	var theKey = event.keyCode;
	if (theKey == 27) { 
		window.close();
	}
	else if (theKey >= 37 && theKey <= 40) { 
		event.keyCode = 9;
	}
}

document.onkeydown = documentKeyDown;
//-->
</SCRIPT>
<BODY >
<font face="tahoma">
<%
if request("act")="state" then
	if request("sefaresh")<>"" then
		myCriteria= "radif_sefareshat = N'"& clng(request("sefaresh")) & "'"
	%>
	<br>
	<center>
	<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#888855"  width="800">
	<%	
	set RS1=Conn.Execute ("SELECT * FROM orders_trace WHERE ("& myCriteria & ")")
	set RS2=Conn.Execute ("SELECT * FROM OrderTraceSteps WHERE (ID = "& clng(request("marhale_box")) & ")")
		if not RS1.eof then
			tmpCounter=0
	%>
			<TR bgcolor="#EEFFCC">
				<TD width="44"># سفارش</TD>
				<TD width="64">تاريخ سفارش</TD>
				<TD width="64">تاريخ تحويل</TD>
				<TD width="124">نام شركت</TD>
				<TD width="124">نام مشتري</TD>
				<TD width="84">عنوان كار</TD>
				<TD width="44">نوع</TD>
				<TD width="64">وضعيت</TD>
				<TD width="74">مرحله</TD>
				<TD width="66">سفارش گيرنده</TD>
			</TR>
			<TR bgcolor="#FFFFFF">
				<TD width="40"><%=RS1("radif_sefareshat")%></TD>
				<TD width="60" DIR="LTR"><%=RS1("order_date")%></TD>
				<TD width="60" DIR="LTR"><%=RS1("return_date")%></TD>
				<TD width="120"><%=RS1("company_name")%>&nbsp;</TD>
				<TD width="120"><%=RS1("customer_name")%>&nbsp;</TD>
				<TD width="80"><%=RS1("order_title")%>&nbsp;</TD>
				<TD width="40"><%=RS1("order_kind")%></TD>
				<TD width="60"><%=RS1("vazyat")%></TD>
				<TD width="140">( <%=RS1("marhale")%>) --&gt;<BR><I><B><%=RS2("Name")%></B></I></TD>
				<TD width="60"><%=RS1("salesperson")%>&nbsp;</TD>
			</TR>
			<TR bgcolor="#EEFFCC">
				<TD colspan="10" align="center"><!-- آيا مشخصات سفارش صحيح است؟ -->
					<INPUT TYPE="button" Value="تاييد" style="font-family:tahoma,arial; font-size:9pt;width:70px;" onClick="window.dialogArguments.all.clearToGo.value='OK';window.close();">&nbsp; &nbsp; &nbsp;
					<INPUT TYPE="button" Value="انصراف" style="font-family:tahoma,arial; font-size:9pt;width:70px;" onClick="window.dialogArguments.all.clearToGo.value='';window.close();">
				</TD>
			</TR>
	
	<%
		else
	%>		<TR>
				<TD align="center" style="font-size:20pt;color:red">چنين سفارشي وجود ندارد.</TD>
			</TR>
	<%
		end if 
		RS1.close
	%>
	</TABLE>
	</center>
	<%
	end if
elseif request("act")="cost" then 
	if request("order")<>"" and request("driver")<>"" and request("value")<>"" then
	%>
	<br>
	
	<center>
	
	<%	
	set RS0=Conn.Execute("SELECT * FROM orders_trace WHERE radif_sefareshat=" & request("order"))
	set RS1=Conn.Execute ("SELECT costs.*,orderTraceLog.statusText,orderTraceLog.stepText,users.realName, cost_drivers.name as driverName,cost_centers.name as costCenterName, cost_unitSizes.name as unitName FROM costs left outer join orderTraceLog on costs.orderTraceLog_id=orderTraceLog.id inner join cost_drivers on costs.driver_id=cost_drivers.id inner join users on costs.user_id=users.id inner join cost_centers on cost_centers.id=cost_drivers.cost_center_id inner join cost_unitSizes on cost_unitSizes.id=cost_drivers.unitsize WHERE costs.[order] ="& clng(request("order")))
	set RS2=Conn.Execute ("SELECT cost_centers.name as costCenterName, cost_drivers.name as driverName, cost_unitSizes.name as unitName FROM cost_drivers inner join cost_centers on cost_drivers.cost_center_id=cost_centers.id inner join cost_unitSizes on cost_unitSizes.id=cost_drivers.unitsize WHERE cost_drivers.id = "&request("driver"))
		if not RS0.eof then
			tmpCounter=0
	%>
	<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#888855"  width="800">
			<tr>
				<td align="center" colspan="8"><b>مشخصات سفارش</b></td>
			</tr>
			<TR bgcolor="#EEFFCC">
				<TD width="44"># سفارش</TD>
				<TD width="64">تاريخ سفارش</TD>
				<TD width="64">تاريخ تحويل</TD>
				<TD width="124">نام شركت</TD>
				<TD width="124">نام مشتري</TD>
				<TD width="84">عنوان كار</TD>
				<TD width="44">نوع</TD>
				<TD width="66">سفارش گيرنده</TD>
			</TR>
			<TR bgcolor="#FFFFFF">
				<TD width="40"><%=RS0("radif_sefareshat")%></TD>
				<TD width="60" DIR="LTR"><%=RS0("order_date")%></TD>
				<TD width="60" DIR="LTR"><%=RS0("return_date")%></TD>
				<TD width="120"><%=RS0("company_name")%>&nbsp;</TD>
				<TD width="120"><%=RS0("customer_name")%>&nbsp;</TD>
				<TD width="80"><%=RS0("order_title")%>&nbsp;</TD>
				<TD width="40"><%=RS0("order_kind")%></TD>
				<TD width="60"><%=RS0("salesperson")%>&nbsp;</TD>
			</TR>
		</table>
		&nbsp;&nbsp;&nbsp;
		<br>
		<table border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#888855"  width="800">
			<tr>
				<td>تاريخ ثبت</td>
				<td>مركز</td>
				<td>عملگر</td>
				<td>مقدار</td>
				<td>كاربر</td>
				<td>وضعيت</td>
				<td>مرحله</td>
			</tr>
		<%
				while not rs1.eof
		%>
			<tr>
				<td><%=rs1("start_time")%></td>
				<td><%=rs1("costCenterName")%></td>
				<td><%=rs1("driverName")%></td>
				<td><%=rs1("value") & " (" & rs1("unitName") & ")"%></td>
				<td><%=rs1("realName")%></td>
				<td><%=rs1("statusText")%></td>
				<td><%=rs1("stepText")%></td>
			</tr>
		<%
					
					rs1.moveNext
				wend
			
		%>
			<tr bgcolor="#33FFFF">
				<td>هم اكنون</td>
				<td><%=rs2("costCenterName")%></td>
				<td><%=rs2("driverName")%></td>
				<td><%=request("value") & " (" & rs2("unitName") & ")"%></td>
				<td><%=session("CSRName")%></td>
				<td><%=rs0("vazyat")%></td>
				<td><%=rs0("marhale")%></td>
			</tr>
			<TR bgcolor="#EEFFCC">
				<TD colspan="7" align="center"><!-- آيا مشخصات سفارش صحيح است؟ -->
					<INPUT TYPE="button" Value="تاييد" style="font-family:tahoma,arial; font-size:9pt;width:70px;" onClick="window.dialogArguments.all.clearToGo.value='OK';window.close();">&nbsp; &nbsp; &nbsp;
					<INPUT TYPE="button" Value="انصراف" style="font-family:tahoma,arial; font-size:9pt;width:70px;" onClick="window.dialogArguments.all.clearToGo.value='';window.close();">
				</TD>
			</TR>
		</table>
	
	<%
		rs0.close
		rs1.close
		rs2.close
		else
	%>
	<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#888855"  width="800">
			<TR>
				<TD align="center" style="font-size:20pt;color:red">چنين سفارشي وجود ندارد.</TD>
			</TR>
	</table>
	<%
		end if 
		'RS1.close
	%>
	
	</center>
	<%
	end if
end if
Conn.Close
%>
</font>
</BODY>
</HTML>