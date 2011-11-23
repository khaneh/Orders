<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="Ê—Êœ „—«Õ·"
SubmenuItem=1
if not Auth(3 , 1) then NotAllowdToViewThisPage()

'myCriteria="vazyat= N'œ— Ã—Ì«‰' AND (marhale= N'ç«Å œÌÃÌ «·' OR marhale= N'œ— ’› ‘—Ê⁄') AND order_kind = N'œÌÃÌ «·'"
if session("ID")="102" then 
	' MASHAD
	myCriteria="vazyat= N'œ— Ã—Ì«‰' AND order_kind = N'œÌÃÌ «·' AND (marhale <> N' ÕÊÌ· ‘œ' AND marhale <> N'¬„«œÂ  ÕÊÌ·' AND marhale <> N' ”ÊÌÂ ‘œ')"
elseif session("ID")="103" then
	'MONZAVI
	myCriteria="vazyat= N'œ— Ã—Ì«‰' AND (order_kind <> N'œÌÃÌ «·' AND order_kind <> N'ÿ—«ÕÌ') AND (marhale <> N' ÕÊÌ· ‘œ' AND marhale <> N'¬„«œÂ  ÕÊÌ·' AND marhale <> N' ”ÊÌÂ ‘œ')"
else
	myCriteria="vazyat= N'œ— Ã—Ì«‰' " 'AND marhale <> N' ÕÊÌ· ‘œ'"
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
			<TD width="44"># ”›«—‘</TD>
			<TD width="64"> «—ÌŒ ”›«—‘</TD>
			<TD width="64"> «—ÌŒ  ÕÊÌ·</TD>
			<TD width="124">‰«„ ‘—ﬂ </TD>
			<TD width="124">‰«„ „‘ —Ì</TD>
			<TD width="84">⁄‰Ê«‰ ﬂ«—</TD>
			<TD width="44">‰Ê⁄</TD>
			<TD width="64">Ê÷⁄Ì </TD>
			<TD width="74">„—Õ·Â</TD>
			<TD width="66">”›«—‘ êÌ—‰œÂ</TD>
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
elseif request("act")="setCost" then
	set rs1=Conn.Execute("SELECT * FROM cost_drivers where id =" & request("driver"))
	'response.write "SELECT TOP 1 * FROM orderTraceLog WHERE [order]=" & request("order_id") & " order by id desc"
	set rs2=Conn.Execute("SELECT TOP 1 * FROM orderTraceLog WHERE [order]=" & request("order_id") & " order by id desc")
	if not rs1.eof then 
		conn.Execute("INSERT INTO costs (driver_id, value, rate, start_time, end_time, [order], orderTraceLog_id, user_id) VALUES (" & request("driver") & ", " & request("driverValue") & ", " & rs1("rate") & ", getdate(), getdate(), " & request("order_id") & ", " & rs2("id") & ", " & session("id") & ") ")
	end if
	rs1.close
	rs2.close
	set RS0=Conn.Execute("SELECT * FROM orders_trace WHERE radif_sefareshat=" & request("order_id"))
	set RS1=Conn.Execute ("SELECT costs.*,orderTraceLog.statusText,orderTraceLog.stepText,users.realName, cost_drivers.name as driverName,cost_centers.name as costCenterName, cost_unitSizes.name as unitName FROM costs left outer join orderTraceLog on costs.orderTraceLog_id=orderTraceLog.id inner join cost_drivers on costs.driver_id=cost_drivers.id inner join users on costs.user_id=users.id inner join cost_centers on cost_centers.id=cost_drivers.cost_center_id inner join cost_unitSizes on cost_unitSizes.id=cost_drivers.unitsize WHERE costs.[order] ="& clng(request("order_id")))
	set RS2=Conn.Execute ("SELECT cost_centers.name as costCenterName, cost_drivers.name as driverName, cost_unitSizes.name as unitName FROM cost_drivers inner join cost_centers on cost_drivers.cost_center_id=cost_centers.id inner join cost_unitSizes on cost_unitSizes.id=cost_drivers.unitsize WHERE cost_drivers.id = "&request("driver"))
		if not RS0.eof then
			tmpCounter=0
	%>
	<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#888855"  width="800">
			<tr>
				<td align="center" colspan="8"><b>„‘Œ’«  ”›«—‘</b></td>
			</tr>
			<TR bgcolor="#EEFFCC">
				<TD width="44"># ”›«—‘</TD>
				<TD width="64"> «—ÌŒ ”›«—‘</TD>
				<TD width="64"> «—ÌŒ  ÕÊÌ·</TD>
				<TD width="124">‰«„ ‘—ﬂ </TD>
				<TD width="124">‰«„ „‘ —Ì</TD>
				<TD width="84">⁄‰Ê«‰ ﬂ«—</TD>
				<TD width="44">‰Ê⁄</TD>
				<TD width="66">”›«—‘ êÌ—‰œÂ</TD>
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
				<td> «—ÌŒ À» </td>
				<td>„—ﬂ“</td>
				<td>⁄„·ê—</td>
				<td>„ﬁœ«—</td>
				<td>ﬂ«—»—</td>
				<td>Ê÷⁄Ì </td>
				<td>„—Õ·Â</td>
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
		</table>
	<% 
	end if
end if 

if true then

'		response.write "myCriteria:'"& myCriteria& "'<BR>" 
'		set RS1=Conn.Execute ("SELECT * FROM orders_trace WHERE ("& myCriteria & ") ORDER BY radif_sefareshat DESC")
'SELECT * FROM orders_trace LEFT OUTER JOIN u81DAY ON orders_trace.radif_sefareshat = u81DAY.Field3 WHERE (u81DAY.Field2 IS NULL) AND (orders_trace.order_date > N'1381/00')

%>
	<FORM METHOD=POST ACTION="default.asp?act=setStatus" onSubmit="return checkValidation('state');">
	<INPUT TYPE="hidden" NAME="clearToGo">
	<br>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" width="700"  align=center >
		<TR bgcolor="#CCCCFF">
			<TD width="20%">‘„«—Â ”›«—‘: <INPUT TYPE="text" NAME="sefaresh_box" size="6" maxLength="7" onKeyPress="return maskNumber(this);" value="<%=request("orderNum")%>">&nbsp;&nbsp;&nbsp;</TD>
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
				<TD><INPUT TYPE="submit" Name="Submit" Value=" «ÌÌœ" style="font-family:tahoma,arial; font-size:9pt;width:70px;"></TD></TR>
			</table>
			</TD>
			<TD align="left"><!--ﬂ«—Â«Ì „‰ —« ‰‘«‰ »œÂ<INPUT TYPE="checkbox" NAME="showMyTasks" <%if request.form("showMyTasks") ="on" then response.write "checked"%>>-->
			</TD>
		</TR>
	</TABLE>
	</FORM>
	<br>
	<%
	set rs=Conn.Execute("Select cost_drivers.*,cost_unitSizes.name as unitsizeName from cost_drivers inner join cost_unitSizes on cost_drivers.unitsize=cost_unitSizes.id where cost_drivers.has_order=1 and cost_drivers.id in (select driver_id from cost_user_relations where user_id="&session("ID")&")")
	if not rs.eof then 
	%>
	<b>À»  ⁄„·ﬂ—œ:</b>
	<form method="post" action="?act=setCost" onSubmit="return checkValidation('cost');">
		<table border="0" cellspacing="2" dir="rtl" width="700" align="center">
			<tr bgcolor="#CCCCFF">
				<td width="20%">‘„«—Â ”›«—‘: 
					<input type="text" name="order_id" size="6" maxlength="7" onkeypress="return maskNumber(this);" value="<%=request("orderNum")%>">
					&nbsp;&nbsp;&nbsp;
				</td>
				<td valign="bottom">
					<table align="center">
						<tr>
							<td>
								<select name="driver">
								<%
								while not rs.eof
									%>
									<option value="<%=rs("id")%>"><%=rs("name") & " (" & rs("unitsizeName") & ")"%></option>
									<%
									rs.moveNext
								wend
								%>
								</select>
							</td>
							<td>
								<input type="text" name="driverValue" size="5" maxlength="50">
							</td>
							<td><%%></td>
							<td>
								<input type="submit" value="«÷«›Â ‘Êœ">
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>
<%	end if
	if request.form("showMyTasks")="on" then
	set RS1=Conn.Execute ("SELECT * FROM orders_trace LEFT OUTER JOIN u81DAY ON orders_trace.radif_sefareshat = u81DAY.Field3 WHERE (u81DAY.Field2 IS NULL) AND (orders_trace.order_date > N'1381/00') AND ("& myCriteria & ") ORDER BY radif_sefareshat DESC")
	if not RS1.eof then
		tmpCounter=0
%>
		<center>
		<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#555588" width="700"  class="t7pt">
			<TR bgcolor="#CCCCFF">
				<TD width="5">-</TD>
				<TD width="10"><INPUT TYPE="checkbox" NAME="checkAll" onClick="selectAll();"></TD>
				<TD width="40"># ”›«—‘</TD>
				<TD width="60"> «—ÌŒ ”›«—‘</TD>
				<TD width="60"> «—ÌŒ  ÕÊÌ·</TD>
				<TD width="100">‰«„ ‘—ﬂ </TD>
				<TD width="100">‰«„ „‘ —Ì</TD>
				<TD width="80">⁄‰Ê«‰ ﬂ«—</TD>
				<TD width="40">‰Ê⁄</TD>
				<TD width="60">Ê÷⁄Ì </TD>
				<TD width="70">„—Õ·Â</TD>
				<TD width="30">”›«—‘ êÌ—‰œÂ</TD>

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
					<TD><INPUT TYPE="submit" Name="Submit" Value=" «ÌÌœ" style="font-family:tahoma,arial; font-size:9pt;width:70px;"></TD>
				</TR></table>
				</TD>
			</TR>
		</TABLE>
		</center>
<%		
	End if
	End if
%>
	

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

function checkValidation(act){
	if (act=='state') {
		if(document.all.sefaresh_box.value != null && document.all.sefaresh_box.value != ""){
			var myTinyWindow = window.showModalDialog('tinyOrderDetails.asp?act='+act+'&sefaresh='+document.all.sefaresh_box.value + "&marhale_box=" + document.all.marhale_box.value,document,'dialogHeight:170px; dialogWidth:850px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
			if (document.all.clearToGo.value=='OK')
				return true;
			else
				return false;
		}
		else if (nothingChecked())
			return false;
		else 
			return true;
	} else if (act=='cost'){
		if(document.all.order_id.value != null && document.all.order_id.value != "") {
			var myTinyWindow = window.showModalDialog('tinyOrderDetails.asp?act='+act+'&order='+document.all.order_id.value + "&driver=" + document.all.driver.value + "&value=" + document.all.driverValue.value,document,'dialogHeight:270px; dialogWidth:850px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
			if (document.all.clearToGo.value=='OK')
				return true;
			else
				return false; 
		}
	}
	
}
document.all.sefaresh_box.focus();
//-->
</SCRIPT>

<!--#include file="tah.asp" -->
