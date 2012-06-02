<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="æÑæÏ ãÑÇÍá"
SubmenuItem=1
if not Auth(3 , 1) then NotAllowdToViewThisPage()

'myCriteria="vazyat= N'ÏÑ ÌÑíÇä' AND (marhale= N'Ç ÏíÌíÊÇá' OR marhale= N'ÏÑ Õİ ÔÑæÚ') AND order_kind = N'ÏíÌíÊÇá'"
if session("ID")="102" then 
	' MASHAD
	myCriteria="vazyat= N'ÏÑ ÌÑíÇä' AND order_kind = N'ÏíÌíÊÇá' AND (marhale <> N'ÊÍæíá ÔÏ' AND marhale <> N'ÂãÇÏå ÊÍæíá' AND marhale <> N'ÊÓæíå ÔÏ')"
elseif session("ID")="103" then
	'MONZAVI
	myCriteria="vazyat= N'ÏÑ ÌÑíÇä' AND (order_kind <> N'ÏíÌíÊÇá' AND order_kind <> N'ØÑÇÍí') AND (marhale <> N'ÊÍæíá ÔÏ' AND marhale <> N'ÂãÇÏå ÊÍæíá' AND marhale <> N'ÊÓæíå ÔÏ')"
else
	myCriteria="vazyat= N'ÏÑ ÌÑíÇä' " 'AND marhale <> N'ÊÍæíá ÔÏ'"
end if
'myCriteria="1=1"

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->


<font face="tahoma">
<%
if request("act")="setStatus" then
	myStatus=request.form("Status")
	set RS1=server.CreateObject("ADODB.Recordset")
	set RS2=Conn.Execute ("SELECT * FROM OrderTraceSteps WHERE (ID = "& request("marhale_box") & ")")
	tmpCounter=0
	%>
	<CENTER>
	<%
	if request.form("checkOrderID").count <> 0  or request.form("sefaresh_box")<>"" then
	%><br>
	<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#558855"  align=center width="700">
		<TR bgcolor="#CCFFCC">
			<TD width="45">-</TD>
			<TD width="44"># ÓİÇÑÔ</TD>
			<TD width="64">ÊÇÑíÎ ÓİÇÑÔ</TD>
			<TD width="64">ÊÇÑíÎ ÊÍæíá</TD>
			<TD width="124">äÇã ÔÑßÊ</TD>
			<TD width="124">äÇã ãÔÊÑí</TD>
			<TD width="84">ÚäæÇä ßÇÑ</TD>
			<TD width="44">äæÚ</TD>
			<TD width="64">æÖÚíÊ</TD>
			<TD width="74">ãÑÍáå</TD>
			<TD width="66">ÓİÇÑÔ íÑäÏå</TD>
		</TR>
<%	for i=1 to request.form("checkOrderID").count
		RS1.ActiveConnection = Conn
		RS1.CursorType = 3 'adOpenStatic = 3
		RS1.LockType = 3 'adLockOptimistic = 3
		RS1.Source = "SELECT * FROM orders_trace WHERE (radif_sefareshat = '"& request.form("checkOrderID")(i)&"')"
		RS1.Open
		set RS3=Conn.Execute ("SELECT * FROM Orders WHERE (ID = "& request.form("checkOrderID")(i) & ")")
		if not RS1.eof then
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
			Else
				tmpColor="#DDDDDD"
			End if 

%>
			<TR bgcolor="<%=tmpColor%>">
				<TD><%=tmpCounter%></TD>
				<TD width="40"><a href="manageOrder.asp?radif=<%=RS1("radif_sefareshat")%>"><%=RS1("radif_sefareshat")%></a></TD>
				<TD width="60" DIR="LTR"><%=RS1("order_date")%></TD>
				<TD width="60" DIR="LTR"><%=RS1("return_date")%></TD>
				<TD width="120"><%=RS1("company_name")%>&nbsp;</TD>
				<TD width="120"><%=RS1("customer_name")%>&nbsp;</TD>
				<TD width="80"><%=RS1("order_title")%>&nbsp;</TD>
				<TD width="40"><%=RS1("order_kind")%></TD>
				<TD width="60"><%=RS1("vazyat")%></TD>
				<TD width="140">( <%=RS1("marhale")%>) --&gt;<BR><I><B><%=RS2("Name")%></B></I></TD>
				<TD width="30"><%=RS1("salesperson")%>&nbsp;</TD>
			</TR>
		<%

			'RS1("marhale")=RS2("Name")
			'RS1("Step") = request.form("marhale_box")
			'RS1("LastUpdatedTime") = currentTime10()
			'RS1("LastUpdatedDate") = shamsitoday()
			'RS1("LastUpdatedBy") = session("ID")
			'RS1.update
	
			mySql="UPDATE orders_trace SET  step= "& request.form("marhale_box") & ",  marhale= N'"& RS2("Name") & "',  LastUpdatedDate=N'"& shamsitoday() & "' , LastUpdatedTime=N'"& currentTime10() & "', LastUpdatedBy=N'"& session("ID")& "'  WHERE (radif_sefareshat= N'"& request.form("checkOrderID")(i) & "')"	
	
			if request.form("marhale_box")="10" then
				call InformCSRorderIsReady(request.form("checkOrderID")(i) , cint(RS3("CreatedBy")))
			end if

			conn.Execute mySql
		else
%>			<TR bgcolor="<%=tmpColor%>">
				<TD><%=tmpCounter%></TD>
				<TD colspan="10">No such Order&nbsp;</TD>
			</TR>
<%		end if 
		RS1.close
	next
	if request.form("sefaresh_box")<>"" then	
		RS1.ActiveConnection = Conn
		RS1.CursorType = 3 'adOpenStatic = 3	
		RS1.LockType = 3 'adLockOptimistic = 3
		RS1.Source = "SELECT * FROM orders_trace WHERE (radif_sefareshat = '"& request.form("sefaresh_box") &"')"
		RS1.Open
		set RS3=Conn.Execute ("SELECT * FROM Orders WHERE (ID = "& request.form("sefaresh_box") & ")")
		if not RS1.eof then
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
			Else
				tmpColor="#DDDDDD"
			End if 
%>
			<TR bgcolor="<%=tmpColor%>">
				<TD><%=tmpCounter%></TD>
				<TD width="40"><a href="manageOrder.asp?radif=<%=RS1("radif_sefareshat")%>"><%=RS1("radif_sefareshat")%></a></TD>
				<TD width="60" DIR="LTR"><%=RS1("order_date")%></TD>
				<TD width="60" DIR="LTR"><%=RS1("return_date")%></TD>
				<TD width="120"><%=RS1("company_name")%>&nbsp;</TD>
				<TD width="120"><%=RS1("customer_name")%>&nbsp;</TD>
				<TD width="80"><%=RS1("order_title")%>&nbsp;</TD>
				<TD width="40"><%=RS1("order_kind")%></TD>
				<TD width="60"><%=RS1("vazyat")%></TD>
				<TD width="140">( <%=RS1("marhale")%>) --&gt;<BR><I><B><%=RS2("Name")%></B></I></TD>
				<TD width="30"><%=RS1("salesperson")%>&nbsp;</TD>
			</TR>
		<%
			RS1("marhale")=RS2("Name")
			RS1("Step") = request.form("marhale_box")
			RS1("LastUpdatedTime") = currentTime10()
			RS1("LastUpdatedDate") = shamsitoday()
			RS1("LastUpdatedBy") = session("ID")
			RS1.update

			if request.form("marhale_box")="10" then
				call InformCSRorderIsReady(request.form("sefaresh_box") , cint(RS3("CreatedBy")))
			end if


		else
%>			<TR bgcolor="<%=tmpColor%>">
				<TD><%=tmpCounter%></TD>
				<TD colspan="10">No such Order&nbsp;<%="'"&request.form("sefaresh_box")&"'"%></TD>
			</TR>
<%		end if 
		RS1.close
	end if
%>
	</TABLE>
	</CENTER>
<Br>

<%
end if
end if 

if true then

'		response.write "myCriteria:'"& myCriteria& "'<BR>" 
'		set RS1=Conn.Execute ("SELECT * FROM orders_trace WHERE ("& myCriteria & ") ORDER BY radif_sefareshat DESC")
'SELECT * FROM orders_trace LEFT OUTER JOIN u81DAY ON orders_trace.radif_sefareshat = u81DAY.Field3 WHERE (u81DAY.Field2 IS NULL) AND (orders_trace.order_date > N'1381/00')

%>

	<FORM METHOD=POST ACTION="default.asp?act=setStatus" onSubmit="return checkValidation();">
	<INPUT TYPE="hidden" NAME="clearToGo">
	<br>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" width="700"  align=center >
		<TR bgcolor="#CCCCFF">
			<TD width="20%">ÔãÇÑå ÓİÇÑÔ: <INPUT TYPE="text" NAME="sefaresh_box" size="6" maxLength="7" onKeyPress="return maskNumber(this);" value="<%=request("orderNum")%>">&nbsp;&nbsp;&nbsp;</TD>
			<TD valign="bottom">
			<table  align="center"><TR>
				<TD><SELECT NAME="marhale_box" style='font-family: tahoma,arial ; font-size: 8pt; width: 160px' <%	if request.form("showMyTasks")="on" then response.write "onChange='document.all.re_marhale_box.value=this.value;'" %>>
					<%
					set RS_STEP=Conn.Execute ("SELECT * FROM OrderTraceSteps WHERE (IsActive=1) order by ord")
					Do while not RS_STEP.eof	
					%>
						<OPTION value="<%=RS_STEP("ID")%>" <%if cint(request("marhale_box"))=cint(RS_STEP("ID")) then response.write "selected" %> ><%=RS_STEP("name")%></option>
						<%
						RS_STEP.moveNext
					loop
					'RS_STEP.close
					'RS_STEP = nothing
					%>

					
				</SELECT></TD>
				<TD><INPUT TYPE="submit" Name="Submit" Value="ÊÇííÏ" style="font-family:tahoma,arial; font-size:9pt;width:70px;"></TD></TR></table>
			</TD>
			<TD align="left"><!--ßÇÑåÇí ãä ÑÇ äÔÇä ÈÏå<INPUT TYPE="checkbox" NAME="showMyTasks" <%if request.form("showMyTasks") ="on" then response.write "checked"%>>-->
			</TD>
		</TR>
	</TABLE>
	<br>
<%	if request.form("showMyTasks")="on" then
	set RS1=Conn.Execute ("SELECT * FROM orders_trace LEFT OUTER JOIN u81DAY ON orders_trace.radif_sefareshat = u81DAY.Field3 WHERE (u81DAY.Field2 IS NULL) AND (orders_trace.order_date > N'1381/00') AND ("& myCriteria & ") ORDER BY radif_sefareshat DESC")
	if not RS1.eof then
		tmpCounter=0
%>
		<center>
		<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#555588" width="700"  class="t7pt">
			<TR bgcolor="#CCCCFF">
				<TD width="5">-</TD>
				<TD width="10"><INPUT TYPE="checkbox" NAME="checkAll" onClick="selectAll();"></TD>
				<TD width="40"># ÓİÇÑÔ</TD>
				<TD width="60">ÊÇÑíÎ ÓİÇÑÔ</TD>
				<TD width="60">ÊÇÑíÎ ÊÍæíá</TD>
				<TD width="100">äÇã ÔÑßÊ</TD>
				<TD width="100">äÇã ãÔÊÑí</TD>
				<TD width="80">ÚäæÇä ßÇÑ</TD>
				<TD width="40">äæÚ</TD>
				<TD width="60">æÖÚíÊ</TD>
				<TD width="70">ãÑÍáå</TD>
				<TD width="30">ÓİÇÑÔ íÑäÏå</TD>

			</TR>
<%		Do while not RS1.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
			Else
				tmpColor="#DDDDDD"
			End if 
%>
			<TR bgcolor="<%=tmpColor%>">
				<TD width="5"><%=tmpCounter%></TD>
				<TD width="10"><INPUT TYPE="checkbox" NAME="checkOrderID" value="<%=RS1("radif_sefareshat")%>"></TD>
				<TD width="40"><%=RS1("radif_sefareshat")%></TD>
				<TD width="60" DIR="LTR"><%=RS1("order_date")%></TD>
				<TD width="60" DIR="LTR"><%=RS1("return_date")%></TD>
				<TD width="100"><%=RS1("company_name")%>&nbsp;</TD>
				<TD width="100"><%=RS1("customer_name")%>&nbsp;</TD>
				<TD width="80"><%=RS1("order_title")%>&nbsp;</TD>
				<TD width="40"><%=RS1("order_kind")%></TD>
				<TD width="60"><%=RS1("vazyat")%></TD>
				<TD width="70"><%=RS1("marhale")%></TD>
				<TD width="30"><%=RS1("salesperson")%>&nbsp;</TD>
			</TR>
<%			RS1.moveNext
		Loop
%>
			<TR bgcolor="#CCCCFF">
				<TD colspan="12">
				<table border=0 align="center"><TR>
						<TD><SELECT NAME="re_marhale_box" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 140px' onChange="document.all.marhale_box.value=this.value;">
						<%
						set RS_STEP=Conn.Execute ("SELECT * FROM OrderTraceSteps WHERE (IsActive=1)")
						Do while not RS_STEP.eof	
						%>
							<OPTION value="<%=RS_STEP("ID")%>" <%if request.form("marhale_box")=RS_STEP("ID") then response.write "selected" %> ><%=RS_STEP("name")%></option>
							<%
							RS_STEP.moveNext
						loop
						'RS_STEP.close
						'RS_STEP = nothing
						%>

						
				</SELECT></TD>
					<TD><INPUT TYPE="submit" Name="Submit" Value="ÊÇííÏ" style="font-family:tahoma,arial; font-size:9pt;width:70px;"></TD>
				</TR></table>
				</TD>
			</TR>
		</TABLE>
		</center>
<%		
	End if
	End if
%>
	</FORM>
<div>
<div>ÊÛííÑ ãÑÇÍá ÒÇÑÔ äÔÏå</div>
<%
	mySQL="declare @date varchar(10);set @date =dbo.udf_date_dateToSolarWithSlash(dateadd(day,-7,getdate()));select * from OrderTraceLog where insertedDate > @date"
	set rs=conn.Execute(mySQL)
	while not rs.eof
		response.write rs("stepText")
		rs.moveNext
	wend
%>
</div>
<%
end if

Conn.Close
%>
</font>
</BODY>
<SCRIPT LANGUAGE="JavaScript">
<!--
function selectAll(){
	if (document.all.checkAll.checked)
		if (document.all.checkOrderID.length > 0)
			for(i=0;i<document.all.checkOrderID.length;i++){
				document.all.checkOrderID[i].checked=true;
			}
		else
			document.all.checkOrderID.checked=true
	else
		if (document.all.checkOrderID.length > 0)
			for(i=0;i<document.all.checkOrderID.length;i++){
				document.all.checkOrderID[i].checked=false;
			}
		else
			document.all.checkOrderID.checked=false
	return true;
}

function nothingChecked(){
<% if request.form("showMyTasks")="on" then %>
	var result=true;
	if (document.all.checkOrderID.length > 0)
		for(i=0;i<document.all.checkOrderID.length;i++){
			if (document.all.checkOrderID[i].checked)
				result=false;
		}
	else
		result=!document.all.checkOrderID.checked;
	return result;
<%else%>
	return false;
<%end if%>
}

function checkValidation(){
	if(document.all.sefaresh_box.value != null && document.all.sefaresh_box.value != ""){
		var myTinyWindow = window.showModalDialog('tinyOrderDetails.asp?sefaresh='+document.all.sefaresh_box.value + "&marhale_box=" + document.all.marhale_box.value,document,'dialogHeight:170px; dialogWidth:850px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
		if (document.all.clearToGo.value=='OK')
			return true;
		else
			return false;
	}
	else if (nothingChecked())
		return false;
	else 
		return true;

	
}
document.all.sefaresh_box.focus();
//-->
</SCRIPT>

<!--#include file="tah.asp" -->
