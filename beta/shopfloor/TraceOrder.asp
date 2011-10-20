<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="ÅÌêÌ—Ì ”›«—‘"
SubmenuItem=3
if not Auth(3 , 3) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<%
Server.ScriptTimeout = 3600
%>
<SCRIPT LANGUAGE='JavaScript'>
<!--
function checkValidation(){
	if (document.all.search_box.value != ''){
		return true;
	}
	else{
		document.all.search_box.focus();
		return false;
	}
}
//-->
</SCRIPT>
<font face="tahoma">
<%
if request("act")="search" then
%>
	<br><br>
	<TABLE border="4" cellspacing="0" cellpadding="0" width="600" align="center" bordercolor="#555599">
	<FORM METHOD=POST ACTION="TraceOrder.asp?act=search" onSubmit="return checkValidation();">
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="5" dir="RTL" width="100%">
		<TR bgcolor="#AAAAEE">
			<TD>‰«„ „‘ —Ì Ì« ‘—ﬂ  Ì« ‘„«—Â ”›«—‘:</TD>
			<TD><INPUT TYPE="text" NAME="search_box" value="<%=request.form("search_box")%>"></TD>
			<TD><INPUT TYPE="submit" NAME="SubmitB" Value="Ã” ÃÊ" style="font-family:tahoma,arial; font-size:10pt;width:100px;"></TD>
			<TD align="left"><% if Auth(3 , 4) then %><A HREF="TraceOrder.asp?act=advancedSearch">Ã” ÃÊÌ ÅÌ‘—› Â</A><% end if %></TD>
		</TR>
		</TABLE>
	</TD></TR>
	</FORM>
	</TABLE>
	<script language="JavaScript">
	<!--
		document.all.search_box.focus();
	//-->
	</script>
	<br>
<%
	searchText=request("search_box")
	if isNumeric(searchText) then
		myCriteria= "radif_sefareshat = '"& clng(searchText) & "'"
	else
		searchText=sqlSafe(searchText)
		myCriteria= "REPLACE([company_name], ' ', '') LIKE REPLACE(N'%"& searchText & "%', ' ', '') OR REPLACE([customer_name], ' ', '') LIKE REPLACE(N'%"& searchText & "%', ' ', '')"
	end if
	mySQL="SELECT     orders_trace.*, OrderTraceStatus.Icon FROM Orders INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN OrderTraceStatus ON orders_trace.status = OrderTraceStatus.ID WHERE ("& myCriteria & ") AND (Orders.Closed=0) ORDER BY radif_sefareshat DESC"	
	set RS1=Conn.Execute (mySQL)
	if not RS1.eof then
		tmpCounter=0
%>
	<div align="left">
	<table border="1" cellspacing="0" cellpadding="0" dir="RTL">
		<TR bgcolor="#CCCCFF">
			<TD width="44"># ”›«—‘</TD>
			<TD width="64"> «—ÌŒ ”›«—‘</TD>
			<TD width="64"> «—ÌŒ  ÕÊÌ·</TD>
			<TD width="122">‰«„ ‘—ﬂ </TD>
			<TD width="122">‰«„ „‘ —Ì</TD>
			<TD width="84">⁄‰Ê«‰ ﬂ«—</TD>
			<TD width="44">‰Ê⁄</TD>
			<TD width="73">„—Õ·Â</TD>
			<TD width="63">”›«—‘ êÌ—‰œÂ</TD>
			<TD width="24">Ê÷⁄</TD>
		</TR>
	</table>
	</div>
	<div id="results" align="left" style="overflow:auto; width:752; height:330;border:0px;" dir="LTR" >
	<TABLE border="1" cellspacing="0" cellpadding="2" borderColor="#555588" dir="RTL">
<%		Do while not RS1.eof
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
		Else
			tmpColor="#DDDDDD"
		End if 
%>
		<TR bgcolor="<%=tmpColor%>">
<!-- 				<TD><%=tmpCounter%></TD> -->
			<TD width="40" DIR="LTR"><A HREF="manageOrder.asp?radif=<%=RS1("radif_sefareshat")%>" target="_blank"><%=RS1("radif_sefareshat")%></A></TD>
			<TD width="60" DIR="LTR"><%=RS1("order_date")%></TD>
			<TD width="60" DIR="LTR"><%=RS1("return_date") & " ("& RS1("return_time") & ")"%></TD>
			<TD width="118"><%=RS1("company_name") & "<br> ·›‰:("& RS1("telephone")& ")"%>&nbsp;</TD>
			<TD width="118"><%=RS1("customer_name")%>&nbsp;</TD>
			<TD width="80"><%=RS1("order_title")%>&nbsp;</TD>
			<TD width="40"><%=RS1("order_kind")%></TD>
			<TD width="69"><%=RS1("marhale")%></TD>
			<TD width="59"><%=RS1("salesperson")%>&nbsp;</TD>
			<TD width="20"><IMG SRC="<%=RS1("Icon")%>" WIDTH="20" HEIGHT="20" BORDER=0 ALT="œ—Ã—Ì«‰"></TD>
		</TR>
		<TR bgcolor="#FFFFFF">
			<TD colspan="10" style="height:10px"></TD>
		</TR>
<%			RS1.moveNext
		Loop
%>		<TR bgcolor="#ccccFF">
			<TD colspan="10"> ⁄œ«œ ‰ «ÌÃ Ã” ÃÊ: <%=tmpCounter%></TD>
		</TR>
	</TABLE>
	</div>
	<HR>
<%	elseif request("search_box")<>"" then
%>	<TABLE border="1" cellspacing="0" cellpadding="0" dir="RTL" align="center" width="600">
		<TR bgcolor="#FFFFDD">
			<TD align="center" style="height:40px;font-size:12pt;font-weight:bold;color:red">«Ì‰ ”›«—‘ ﬁ»·« ›«ﬂ Ê— ‘œÂ Ì« «’·« ÊÃÊœ ‰œ«—œ.<br>»—«Ì «ÿ„Ì«‰« »Â <A HREF="orderTraceEdit.asp">«’·«Õ Ê ‰„«Ì‘</A> „—«Ã⁄Â ﬂ‰Ìœ.</TD>
		</TR>
	</TABLE>
<%		End if
elseif request("act")="advancedSearch" then
'------  Advanced Search 
%>
<!--#include File="../include_JS_InputMasks.asp"-->
<%
	Server.ScriptTimeout = 3600
	tmpTime=time
	tmpTime=Hour(tmpTime)&":"&Minute(tmpTime)
	if instr(tmpTime,":")<3 then tmpTime="0" & tmpTime
	if len(tmpTime)<5 then tmpTime=Left(tmpTime,3) & "0" & Right(tmpTime,1)
%>
	<br><br>
	<TABLE border="4" cellspacing="0" cellpadding="0" width="700" align="center" bordercolor="#555599">
	<FORM METHOD=POST ACTION="TraceOrder.asp?act=advancedSearch" onSubmit="return checkValidation();">
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="1" dir="RTL" width="100%" bgcolor="white">
		<TR bgcolor="#EEEEEE">
			<TD><INPUT TYPE="checkbox" NAME="check_sefaresh" onclick="check_sefaresh_Click()" checked></TD>
			<TD>‘„«—Â ”›«—‘</TD>
			<TD><INPUT TYPE="text" NAME="az_sefaresh" dir="LTR" value="<%=request.form("az_sefaresh")%>" size="8" maxlength="6" onKeyPress="return maskNumber(this);"></TD>
			<TD> «</TD>
			<TD><INPUT TYPE="text" NAME="ta_sefaresh" dir="LTR" value="<%=request.form("ta_sefaresh")%>" size="8" maxlength="6" onKeyPress="return maskNumber(this);" ></TD>
			<td rowspan="11" style="width:1px" bgcolor="#555599"></td>
			<TD><INPUT TYPE="checkbox" NAME="check_kind" onclick="check_kind_Click()" checked></TD>
			<TD>‰Ê⁄ ”›«—‘</TD>
			<TD colspan="3"><SELECT NAME="order_kind_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
				<OPTION value="«›” " <%if request.form("order_kind_box")="«›” " then response.write "selected" %> >«›” </option>
				<OPTION value="œÌÃÌ «·" <%if request.form("order_kind_box")="œÌÃÌ «·" then response.write "selected" %> >œÌÃÌ «·</option>
				<OPTION value="”Ì«Â Ê ”›Ìœ" <%if request.form("order_kind_box")="”Ì«Â Ê ”›Ìœ" then response.write "selected" %> >”Ì«Â Ê ”›Ìœ</option>
				<OPTION value="ÿ—«ÕÌ" <%if request.form("order_kind_box")="ÿ—«ÕÌ" then response.write "selected" %> >ÿ—«ÕÌ</option>
				<OPTION value="’Õ«›Ì" <%if request.form("order_kind_box")="’Õ«›Ì" then response.write "selected" %> >’Õ«›Ì</option>
				<OPTION value="›Ì·„" <%if request.form("order_kind_box")="›Ì·„" then response.write "selected" %> >›Ì·„</option>
				<OPTION value="“Ì‰ﬂ" <%if request.form("order_kind_box")="“Ì‰ﬂ" then response.write "selected" %> >“Ì‰ﬂ</option>
				<OPTION value="·„Ì‰ " <%if request.form("order_kind_box")="·„Ì‰ " then response.write "selected" %> >·„Ì‰ </option>
			</SELECT></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD><INPUT TYPE="checkbox" NAME="check_tarikh_sefaresh" onclick="check_tarikh_sefaresh_Click()" checked></TD>
			<TD> «—ÌŒ ”›«—‘</TD>
			<TD><INPUT TYPE="text" NAME="az_tarikh_sefaresh" dir="LTR" value="<%=request.form("az_tarikh_sefaresh")%>" size="10" onKeyPress="return maskDate(this);" onBlur="if(acceptDate(this))document.all.ta_tarikh_sefaresh.value=this.value;" maxlength="10"></TD>
			<TD> «</TD>
			<TD><INPUT TYPE="text" NAME="ta_tarikh_sefaresh" dir="LTR" value="<%=request.form("ta_tarikh_sefaresh")%>" size="10" onKeyPress="return maskDate(this);"  onblur="acceptDate(this)" maxlength="10"></TD>
			<TD><INPUT TYPE="checkbox" NAME="check_marhale" onclick="check_marhale_Click()" checked></TD>
			<TD>„—Õ·Â</TD>
			<TD><SELECT NAME="marhale_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
					<%
					set RS_STEP=Conn.Execute ("SELECT * FROM OrderTraceSteps WHERE (IsActive=1)")
					Do while not RS_STEP.eof	
					%>
						<OPTION value="<%=RS_STEP("ID")%>" <%if cint(request.form("marhale_box"))=cint(RS_STEP("ID")) then response.write "selected" %> ><%=RS_STEP("name")%></option>
						<%
						RS_STEP.moveNext
					loop
					%>
			</SELECT></TD>
			<TD><span id="marhale_not_check_label" style='font-weight:bold;color:red'>‰»«‘œ</span></TD>
			<TD><INPUT TYPE="checkbox" NAME="marhale_not_check" onclick="marhale_not_check_Click();" checked></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#EEEEEE">
			<TD><INPUT TYPE="checkbox" NAME="check_tarikh_tahvil" onclick="check_tarikh_tahvil_Click()" checked></TD>
			<TD> «—ÌŒ  ÕÊÌ·</TD>
			<TD><INPUT TYPE="text" NAME="az_tarikh_tahvil" dir="LTR" value="<%=request.form("az_tarikh_tahvil")%>" size="10" onKeyPress="return maskDate(this);"  onblur="acceptDate(this)" maxlength="10"></TD>
			<TD> «</TD>
			<TD><INPUT TYPE="text" NAME="ta_tarikh_tahvil" dir="LTR" value="<%=request.form("ta_tarikh_tahvil")%>" size="10" onKeyPress="return maskDate(this);"  onblur="acceptDate(this)" maxlength="10"></TD>
			<TD><INPUT TYPE="checkbox" NAME="check_vazyat" onclick="check_vazyat_Click()" checked></TD>
			<TD>Ê÷⁄Ì </TD>
			<TD><SELECT NAME="vazyat_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
				<%
				set RS_STATUS=Conn.Execute ("SELECT * FROM OrderTraceStatus WHERE (IsActive=1)")
				Do while not RS_STATUS.eof	
				%>
					<OPTION value="<%=RS_STATUS("ID")%>" <%if cint(request.form("vazyat_box"))=cint(RS_STATUS("ID")) then response.write "selected" %> ><%=RS_STATUS("Name")%></option>
					<%
					RS_STATUS.moveNext
				loop
				%>
			</SELECT></TD>
			<TD><span id="vazyat_not_check_label" style='font-weight:bold;color:red'>‰»«‘œ</span></TD>
			<TD><INPUT TYPE="checkbox" NAME="vazyat_not_check" onclick="vazyat_not_check_Click();" checked></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD colspan="5">&nbsp;</TD>
			<TD><INPUT TYPE="checkbox" NAME="check_salesperson" onclick="check_salesperson_Click()" checked></TD>
			<TD>”›«—‘ êÌ—‰œÂ</TD>
			<TD colspan="3"><SELECT NAME="salesperson_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
				<OPTION value="»Ìœ¬»«œÌ" <%if request.form("salesperson_box")="»Ìœ¬»«œÌ" then response.write "selected" %> >»Ìœ¬»«œÌ</option>
				<OPTION value="Õ”‰ÌÂ" <%if request.form("salesperson_box")="Õ”‰ÌÂ" then response.write "selected" %> >Õ”‰ÌÂ</option> 
				<OPTION value="ç·«ÃÊ—" <%if request.form("salesperson_box")="ç·«ÃÊ—" then response.write "selected" %> >ç·«ÃÊ—</option>
				<OPTION value="“—ê—" <%if request.form("salesperson_box")="“—ê—" then response.write "selected" %> >“—ê—</option>
				<OPTION value="ﬂÊ›Ì" <%if request.form("salesperson_box")="ﬂÊ›Ì" then response.write "selected" %> >ﬂÊ›Ì</option>
				<OPTION value="Ê«÷ÕÌ" <%if request.form("salesperson_box")="Ê«÷ÕÌ" then response.write "selected" %> >Ê«÷ÕÌ</option>
				<OPTION value="„Õﬁﬁ" <%if request.form("salesperson_box")="„Õﬁﬁ" then response.write "selected" %> >„Õﬁﬁ</option>
				<OPTION value="„⁄ „œ  »«—" <%if request.form("salesperson_box")="„⁄ „œ  »«—" then response.write "selected" %> >„⁄ „œ  »«—</option>
				<OPTION value="Ã·«·Ì" <%if request.form("salesperson_box")="Ã·«·Ì" then response.write "selected" %> >Ã·«·Ì</option>
				<OPTION value="ÃÂ«‰Ì«‰" <%if request.form("salesperson_box")="ÃÂ«‰Ì«‰" then response.write "selected" %> >ÃÂ«‰Ì«‰</option>
				<OPTION value="ﬂÊ«ﬂ»" <%if request.form("salesperson_box")="ﬂÊ«ﬂ»" then response.write "selected" %> >ﬂÊ«ﬂ»</option>
			</SELECT></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#EEEEEE">
			<TD><INPUT TYPE="checkbox" NAME="check_company_name" onclick="check_company_name_Click()" checked></TD>
			<TD>‰«„ ‘—ﬂ </TD>
			<TD colspan="3"><INPUT TYPE="text" NAME="company_name_box" value="<%=request.form("company_name_box")%>"></TD>
			<TD><INPUT TYPE="checkbox" NAME="check_telephone" onclick="check_telephone_Click()" checked></TD>
			<TD> ·›‰</TD>
			<TD colspan="3"><INPUT TYPE="text" NAME="telephone_box" value="<%=request.form("telephone_box")%>"></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD><INPUT TYPE="checkbox" NAME="check_customer_name" onclick="check_customer_name_Click()" checked></TD>
			<TD>‰«„ „‘ —Ì</TD>
			<TD colspan="3"><INPUT TYPE="text" NAME="customer_name_box" value="<%=request.form("customer_name_box")%>"></TD>
			<TD><INPUT TYPE="checkbox" NAME="check_order_title" onclick="check_order_title_Click()" checked></TD>
			<TD>⁄‰Ê«‰ ”›«—‘</TD>
			<TD colspan="3"><INPUT TYPE="text" NAME="order_title_box" value="<%=request.form("order_title_box")%>"></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#AAAAEE">
			<td colspan="11" style="height:30px">
			<TABLE align="center" width="50%">
			<TR>
				<TD><INPUT TYPE="submit" Name="Submit" Value=" «ÌÌœ" style="font-family:tahoma,arial; font-size:10pt;width:100px;"></TD>
				<TD>&nbsp;<INPUT disabled TYPE="button" onclick="if(document.selection.createRange().text != ''){document.execCommand('Delete');};"></TD>
				<TD align="left"><INPUT TYPE="button" Name="Cancel" Value="Å«ﬂ ﬂ‰" style="font-family:tahoma,arial; font-size:10pt;width:100px;" onclick="window.location='advancedsearch.asp';"></TD>
			</TR>
			</TABLE>
			</td>
		</TR>
		</TABLE>
	</TD></TR>
	</FORM>
	</TABLE>
	<br>
	<SCRIPT LANGUAGE='JavaScript'>
	<!--
	function check_sefaresh_Click(){
		if ( document.all.check_sefaresh.checked ) {
			document.all.az_sefaresh.style.visibility = "visible";
			document.all.ta_sefaresh.style.visibility = "visible";
			document.all.az_sefaresh.focus();
		}
		else{
			document.all.az_sefaresh.style.visibility = "hidden";
			document.all.ta_sefaresh.style.visibility = "hidden";
		}
	}

	function check_tarikh_sefaresh_Click(){
		if ( document.all.check_tarikh_sefaresh.checked ) {
			document.all.az_tarikh_sefaresh.style.visibility = "visible";
			document.all.ta_tarikh_sefaresh.style.visibility = "visible";
			document.all.az_tarikh_sefaresh.focus();
		}
		else{
			document.all.az_tarikh_sefaresh.style.visibility = "hidden";
			document.all.ta_tarikh_sefaresh.style.visibility = "hidden";
		}
	}

	function check_tarikh_tahvil_Click(){
		if ( document.all.check_tarikh_tahvil.checked ) {
			document.all.az_tarikh_tahvil.style.visibility = "visible";
			document.all.ta_tarikh_tahvil.style.visibility = "visible";
			document.all.az_tarikh_tahvil.focus();
		}
		else{
			document.all.az_tarikh_tahvil.style.visibility = "hidden";
			document.all.ta_tarikh_tahvil.style.visibility = "hidden";
		}
	}

	function check_company_name_Click(){
		if (document.all.check_company_name.checked) {
			document.all.company_name_box.style.visibility = "visible";
			document.all.company_name_box.focus();
		}
		else{
			document.all.company_name_box.style.visibility = "hidden";
		}
	}

	function check_customer_name_Click(){
		if (document.all.check_customer_name.checked) {
			document.all.customer_name_box.style.visibility = "visible";
			document.all.customer_name_box.focus();
		}
		else{
			document.all.customer_name_box.style.visibility = "hidden";
		}
	}

	function check_kind_Click(){
		if (document.all.check_kind.checked) {
			document.all.order_kind_box.style.visibility = "visible";
			document.all.order_kind_box.focus();
		}
		else{
			document.all.order_kind_box.style.visibility = "hidden";
		}
	}

	function check_marhale_Click(){
		if (document.all.check_marhale.checked) {
			document.all.marhale_box.style.visibility = "visible";
			document.all.marhale_box.focus();
			document.all.marhale_not_check.style.visibility = "visible";
			marhale_not_check_Click();	
		}
		else{
			document.all.marhale_box.style.visibility = "hidden";
			document.all.marhale_not_check.style.visibility = "hidden";
			document.all.marhale_not_check_label.style.color='#BBBBBB'
		}
	}

	function check_vazyat_Click(){
		if (document.all.check_vazyat.checked) {
			document.all.vazyat_box.style.visibility = "visible";
			document.all.vazyat_box.focus();
			document.all.vazyat_not_check.style.visibility = "visible";
			vazyat_not_check_Click();
		}
		else{
			document.all.vazyat_box.style.visibility = "hidden";
			document.all.vazyat_not_check.style.visibility = "hidden";
			document.all.vazyat_not_check_label.style.color='#BBBBBB'
		}
	}

	function check_salesperson_Click(){
		if (document.all.check_salesperson.checked) {
			document.all.salesperson_box.style.visibility = "visible";
			document.all.salesperson_box.focus();
		}
		else{
			document.all.salesperson_box.style.visibility = "hidden";
		}
	}

	function check_telephone_Click(){
		if (document.all.check_telephone.checked) {
			document.all.telephone_box.style.visibility = "visible";
			document.all.telephone_box.focus();
		}
		else{
			document.all.telephone_box.style.visibility = "hidden";
		}
	}

	function check_order_title_Click(){
		if (document.all.check_order_title.checked) {
			document.all.order_title_box.style.visibility = "visible";
			document.all.order_title_box.focus();
		}
		else{
			document.all.order_title_box.style.visibility = "hidden";
		}
	}

	function marhale_not_check_Click(){
		if (document.all.marhale_not_check.checked) {
			document.all.marhale_not_check_label.style.color='red'
		}
		else{
			document.all.marhale_not_check_label.style.color='#BBBBBB'
		}
	}

	function vazyat_not_check_Click(){
		if (document.all.vazyat_not_check.checked) {
			document.all.vazyat_not_check_label.style.color='red'
		}
		else{
			document.all.vazyat_not_check_label.style.color='#BBBBBB'
		}
	}

	function Form_Load(){
	<%
	maybeAND = ""
	myCriteria = ""
	If request.form("check_sefaresh") = "on" Then
		if request.form("az_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "radif_sefareshat >= " & request.form("az_sefaresh")
			maybeAND=" AND "
		end if
		if request.form("ta_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "radif_sefareshat <= " & request.form("ta_sefaresh")
			maybeAND=" AND "
		end if
		If (request.form("az_sefaresh") = "") AND (request.form("ta_sefaresh") = "") then
			response.write "document.all.check_sefaresh.checked = false;" & vbCrLf
			response.write "document.all.az_sefaresh.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_sefaresh.style.visibility = 'hidden';" & vbCrLf
		End if
	Else
		response.write "document.all.check_sefaresh.checked = false;" & vbCrLf
		response.write "document.all.az_sefaresh.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.ta_sefaresh.style.visibility = 'hidden';" & vbCrLf

	End If

	If request.form("check_tarikh_sefaresh") = "on" Then
		if request.form("az_tarikh_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "order_date >= '" & request.form("az_tarikh_sefaresh") & "'"
			maybeAND=" AND "
		end if
		if request.form("ta_tarikh_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "order_date <= '" & request.form("ta_tarikh_sefaresh") & "'"
			maybeAND=" AND "
		end if 
		If (request.form("az_tarikh_sefaresh") = "") AND (request.form("ta_tarikh_sefaresh") = "") then
			response.write "document.all.check_tarikh_sefaresh.checked = false;" & vbCrLf
			response.write "document.all.az_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
		End if
	Else
		response.write "document.all.check_tarikh_sefaresh.checked = false;" & vbCrLf
		response.write "document.all.az_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.ta_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
	End If

	If request.form("check_tarikh_tahvil") = "on" Then
		if request.form("az_tarikh_tahvil") <> "" then
			myCriteria = myCriteria & maybeAND & "return_date >= '" & request.form("az_tarikh_tahvil") & "'"
			maybeAND=" AND "
		end if
		if request.form("ta_tarikh_tahvil") <> "" then
			myCriteria = myCriteria & maybeAND & "return_date <= '" & request.form("ta_tarikh_tahvil") & "'"
			maybeAND=" AND "
		end if
		If (request.form("az_tarikh_tahvil") = "") AND (request.form("ta_tarikh_tahvil") = "") then
			response.write "document.all.check_tarikh_tahvil.checked = false;" & vbCrLf
			response.write "document.all.az_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
		End if
	Else
		response.write "document.all.check_tarikh_tahvil.checked = false;" & vbCrLf
		response.write "document.all.az_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.ta_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
	End If

	If (request.form("check_company_name") = "on" AND request.form("company_name_box") <> "" ) then
		myCriteria = myCriteria & maybeAND & "company_name Like N'%" & request.form("company_name_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_company_name.checked = false;" & vbCrLf
		response.write "document.all.company_name_box.style.visibility = 'hidden';" & vbCrLf
	End if

	If (request.form("check_customer_name") = "on" AND request.form("customer_name_box") <> "")then
		myCriteria = myCriteria & maybeAND & "customer_name Like N'%" & request.form("customer_name_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_customer_name.checked = false;" & vbCrLf
		response.write "document.all.customer_name_box.style.visibility = 'hidden';" & vbCrLf
	End if

	If request.form("check_kind") = "on" then
		myCriteria = myCriteria & maybeAND & "order_kind = N'" & request.form("order_kind_box") & "'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_kind.checked = false;" & vbCrLf
		response.write "document.all.order_kind_box.style.visibility = 'hidden';" & vbCrLf
	End if

	If request.form("check_marhale") = "on" then
		If request.form("marhale_not_check") = "on" then
			myCriteria = myCriteria & maybeAND & "NOT(step = " & request.form("marhale_box") & ")"
		Else
			myCriteria = myCriteria & maybeAND & "step = " & request.form("marhale_box") 
			response.write "document.all.marhale_not_check.checked = false;" & vbCrLf
			response.write "document.all.marhale_not_check_label.style.color='#BBBBBB';"& vbCrLf
		End If
		maybeAND=" AND "
	Else
		response.write "document.all.check_marhale.checked = false;" & vbCrLf
		response.write "document.all.marhale_box.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.marhale_not_check.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.marhale_not_check_label.style.color='#BBBBBB';"& vbCrLf
	End if

	If request.form("check_vazyat") = "on" then
		If request.form("vazyat_not_check") = "on" then
			myCriteria = myCriteria & maybeAND & "NOT(status = " & request.form("vazyat_box") & ")"
		Else
			myCriteria = myCriteria & maybeAND & "status = " & request.form("vazyat_box") 
			response.write "document.all.vazyat_not_check.checked = false;" & vbCrLf
			response.write "document.all.vazyat_not_check_label.style.color='#BBBBBB';"& vbCrLf
		End If
		maybeAND=" AND "
	Else
		response.write "document.all.check_vazyat.checked = false;" & vbCrLf
		response.write "document.all.vazyat_box.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.vazyat_not_check.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.vazyat_not_check_label.style.color='#BBBBBB';"& vbCrLf
	End if

	If request.form("check_salesperson") = "on" then
		myCriteria = myCriteria & maybeAND & "salesperson = N'" & request.form("salesperson_box") & "'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_salesperson.checked = false;" & vbCrLf
		response.write "document.all.salesperson_box.style.visibility = 'hidden';" & vbCrLf
	End if

	If (request.form("check_order_title") = "on" AND request.form("order_title_box") <> "")then
		myCriteria = myCriteria & maybeAND & "order_title Like N'%" & request.form("order_title_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_order_title.checked = false;" & vbCrLf
		response.write "document.all.order_title_box.style.visibility = 'hidden';" & vbCrLf
	End if

	If (request.form("check_telephone") = "on" AND request.form("telephone_box") <> "")then
		myCriteria = myCriteria & maybeAND & "telephone Like N'%" & request.form("telephone_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_telephone.checked = false;" & vbCrLf
		response.write "document.all.telephone_box.style.visibility = 'hidden';" & vbCrLf
	End if

	%>
	}
	function checkValidation(){
		return true;
	}

	Form_Load();

	//-->
	</SCRIPT>
<%
	if request("Submit")=" «ÌÌœ" then
		IF maybeAND <> " AND " THEN
			response.write "Nothing !!!!!!!!!!"
		ELSE
			mySQL="SELECT orders_trace.*, OrderTraceStatus.Icon FROM Orders INNER JOIN  orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN  OrderTraceStatus ON orders_trace.status = OrderTraceStatus.ID WHERE ("& myCriteria & ") AND (Orders.Closed=0) ORDER BY radif_sefareshat DESC"	
			set RS1=Conn.Execute (mySQL)
			if not RS1.eof then
				tmpCounter=0
%>
			<div align="left">
			<table border="1" cellspacing="0" cellpadding="0" dir="RTL">
				<TR bgcolor="#CCCCFF">
					<TD width="44"># ”›«—‘</TD>
					<TD width="64"> «—ÌŒ ”›«—‘</TD>
					<TD width="64"> «—ÌŒ  ÕÊÌ·</TD>
					<TD width="122">‰«„ ‘—ﬂ </TD>
					<TD width="122">‰«„ „‘ —Ì</TD>
					<TD width="84">⁄‰Ê«‰ ﬂ«—</TD>
					<TD width="44">‰Ê⁄</TD>
					<TD width="73">„—Õ·Â</TD>
					<TD width="63">”›«—‘ êÌ—‰œÂ</TD>
					<TD width="24">Ê÷⁄</TD>
				</TR>
			</table>
			</div>
			<div id="results" align="left" style="overflow:auto; width:752; height:330;border:0px;" dir="LTR" >
			<TABLE border="1" cellspacing="0" cellpadding="2" borderColor="#555588" dir="RTL">
<%				Do while not RS1.eof
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
				Else
					tmpColor="#DDDDDD"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>">
					<TD width="40" DIR="LTR"><A HREF="manageOrder.asp?radif=<%=RS1("radif_sefareshat")%>" target="_blank"><%=RS1("radif_sefareshat")%></A></TD>
					<TD width="60" DIR="LTR"><%=RS1("order_date")%></TD>
					<TD width="60" DIR="LTR"><%=RS1("return_date") & " ("& RS1("return_time") & ")"%></TD>
					<TD width="118"><%=RS1("company_name") & "<br> ·›‰:("& RS1("telephone")& ")"%>&nbsp;</TD>
					<TD width="118"><%=RS1("customer_name")%>&nbsp;</TD>
					<TD width="80"><%=RS1("order_title")%>&nbsp;</TD>
					<TD width="40"><%=RS1("order_kind")%></TD>
					<TD width="69"><%=RS1("marhale")%></TD>
					<TD width="59"><%=RS1("salesperson")%>&nbsp;</TD>
					<TD width="20"><IMG SRC="<%=RS1("icon")%>" WIDTH="20" HEIGHT="20" BORDER=0 ALT="œ—Ã—Ì«‰">
					</TD>
				</TR>
				<TR bgcolor="#FFFFFF">
					<TD colspan="10" style="height:10px"></TD>
				</TR>
<%					RS1.moveNext
				Loop
%>				<TR bgcolor="#ccccFF">
					<TD colspan="10"> ⁄œ«œ ‰ «ÌÃ Ã” ÃÊ: <%=tmpCounter%></TD>
				</TR>
			</TABLE>
			</div>
			<BR>
<%			else
%>			<TABLE border="1" cellspacing="0" cellpadding="0" dir="RTL" align="center" width="600">
				<TR bgcolor="#FFFFDD">
					<TD align="center" style="height:40px;font-size:12pt;font-weight:bold;color:red">ÂÌç ÃÊ«»Ì ‰œ«—Ì„ »Â ‘„« »œÂÌ„.</TD>
				</TR>
			</TABLE>
			<hr>
<%			end if
		END IF
	end if
end if

Conn.Close
%>
</font>
<!--#include file="tah.asp" -->