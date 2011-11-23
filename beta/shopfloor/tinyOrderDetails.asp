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
<TITLE>tiny Order Details</TITLE>
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

Conn.Close
%>
</font>
</BODY>
</HTML>