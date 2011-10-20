<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="ÃœÊ·  Ê·Ìœ"
SubmenuItem=6
if not Auth(3 , 6) then NotAllowdToViewThisPage()

'OutService Page Request
'By Alix - Last changed: 81/01/13
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; height:100%; background-color:#C3DBEB;}
	.repHeader td{text-align: center; font-size:7pt;}
</STYLE>
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------ Edit Order Desk Items Action
'-----------------------------------------------------------------------------------------------------
if request("act")="doEdit" then
	
	'response.write "<br>" & replace(request.form,"&","<br>")
	'response.end

	for i = 1 to request.form("id").count
		radif_sefareshat = request.form("radif_sefareshat")(i)
		size = request.form("size")(i)
		forms = request.form("forms")(i)
		tirag = request.form("tirag")(i)
		zinc = request.form("zinc")(i)
		is2roo = request.form("is2roo")(i)
		colorStatus = request.form("colorStatus")(i)
		lastEditedBy = session("id")
		lastEditedDate = shamsiToday()
		orderDeskItemID = request.form("id")(i)

		mySQL2="UPDATE orderDesk SET [size] = '" & size & "', forms =" & forms & ", Tirag ='" & Tirag & "', Zinc ='" & Zinc & "', is2roo ='" & is2roo & "', colorStatus ='" & colorStatus & "', lastEditedBy =" & lastEditedBy & ", lastEditedDate ='" & lastEditedDate & "' where id=" & orderDeskItemID
		
		conn.execute(mySQL2) 
	next
	response.redirect "orderDesk.asp"

end if

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------- Add to Order Desk Table
'-----------------------------------------------------------------------------------------------------
if request("submit")="»Ì«›“«" then
	'response.write "==" & replace(request.form, "&", "<br>")
	radif_sefareshat = request.form("radif_sefareshat")
	size = request.form("size")
	forms = request.form("forms")
	tirag = request.form("tirag")
	zinc = request.form("zinc")
	is2roo = request.form("is2roo")
	colorStatus = request.form("colorStatus")
	createdBy = session("id")
	createdDate = shamsiToday()

	mySQL="INSERT INTO orderDesk (order_ID, [size], forms, Tirag, Zinc, is2roo, colorStatus, createdBy, createdDate) VALUES (" & radif_sefareshat & ", '" & size & "', '" & forms & "', '" & tirag & "', '" & zinc & "', '" & is2roo & "', '" & colorStatus & "'," & createdBy & ", '" & createdDate & "')"
	conn.execute(mySQL) 
	response.redirect "orderDesk.asp"
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Print Order Desk Items
'-----------------------------------------------------------------------------------------------------
if request("submit")="„‘«ÂœÂ ¬Œ—Ì‰ »—‰«„Â" then
%>
	<BR>	<BR>
	<CENTER>
	<% 	ReportLogRow = PrepareReport ("orderDesk.rpt", "alak", ".", "/beta/dialog_printManager.asp?act=Fin") %>
	<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

	</CENTER>

	<BR>	
	<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no></iframe>
<%
	response.end

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Print Order Desk Items
'-----------------------------------------------------------------------------------------------------
elseif request("submit")="ç«Å »—‰«„Â —Ê“«‰Â" then

	mySQL2="UPDATE orderDesk SET selected4print = 0 WHERE (selected4print = 1)"
	conn.execute(mySQL2) 

	for i = 1 to request.form("id").count
		orderDeskItemID = request.form("id")(i)

		mySQL2="UPDATE orderDesk SET selected4print = 1 WHERE id=" & orderDeskItemID
		conn.execute(mySQL2) 
	next
	'response.redirect "orderDesk.asp"
%>
	<BR>	<BR>
	<CENTER>
	<% 	ReportLogRow = PrepareReport ("orderDesk.rpt", "alak", ".", "/beta/dialog_printManager.asp?act=Fin") %>
	<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

	</CENTER>

	<BR>	
	<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no></iframe>
<%
	response.end
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------- Edit Order Desk Items
'-----------------------------------------------------------------------------------------------------
elseif request("submit")="ÊÌ—«Ì‘" then
%>
	<FORM METHOD=POST ACTION="?act=doEdit"><BR><BR>
	<table class="CustTable" cellspacing='1' style='width:95%;' align=center>
		<tr>
			<td colspan="10" class="CusTableHeader" style="text-align:right;">ÊÌ—«Ì‘ ÃœÊ· »—‰«„Â —Ì“Ì  Ê·Ìœ</td>
		</tr>
		<tr>
			<td class="CusTD3"><small> «—ÌŒ  ÕÊÌ·</small></td>
			<td class="CusTD3"><small>”›«—‘ œÂ‰œÂ \<font color="aaaaaa"> ⁄‰Ê«‰ ﬂ«—</small></font></td>
			<td class="CusTD3"><small> ”›«—‘</small></td>
			<td class="CusTD3"><small>”«Ì“</small></td>
			<td class="CusTD3"><small> ⁄œ«œ ›—„</small></td>
			<td class="CusTD3"><small> Ì—«é Â— ›—„</small></td>
			<td class="CusTD3"><small>“Ì‰ﬂ</small></td>
			<td class="CusTD3"><small>—Ê</small></td>
			<td class="CusTD3"><small>—‰ê</small></td>
		</tr>
<%
for i = 1 to request.form("id").count
	orderDeskItemID = request.form("id")(i)
	
	if not isNumeric(orderDeskItemID) then
		call showAlert("Œÿ«! ‘„«—Â ”›«—‘ »«Ìœ ⁄œœ »«‘œ.",CONST_MSG_ERROR)
		response.end
	end if

	mySQL="SELECT orderDesk.*, orders_trace.return_date, orders_trace.company_name,  orders_trace.customer_name, orders_trace.order_title FROM orderDesk INNER JOIN orders_trace ON orderDesk.order_ID = orders_trace.radif_sefareshat where id=" & orderDeskItemID
	
	Set RS1 = conn.execute(mySQL)
	if RS1.eof then
		call showAlert("Œÿ«! «Ì‰ ‘„«—Â ”›«—‘ ÊÃÊœ ‰œ«—œ.",CONST_MSG_ERROR)
		response.end
	end if
%>		

		<TR bgcolor="#FFFFFF"><INPUT TYPE="hidden" name="id" value="<%=RS1("id")%>">
			<TD dir="LTR" align='right'><%=RS1("return_date")%></TD>
			<TD><A HREF="../shopfloor/manageOrder.asp?radif=<%=RS1("order_id")%>" target="_blank"><%=RS1("customer_name")%> - <%=RS1("company_name")%><BR><BR><font color="aaaaaa"><%=RS1("order_title")%></font></A></TD>
			<TD dir="LTR" align='center'><INPUT TYPE="hidden" NAME="radif_sefareshat" value="<%=RS1("order_id")%>" size=3><%=RS1("order_id")%></TD>
			<TD dir="LTR" align='center'><INPUT TYPE="text" NAME="size" size=5 value="<%=trim(RS1("size"))%>"></TD>
			<TD dir="LTR" align='center'><INPUT TYPE="text" NAME="forms" size=3 value="<%=trim(RS1("forms"))%>"></TD>
			<TD dir="LTR" align='center'><INPUT TYPE="text" NAME="tirag" size=3 value="<%=trim(RS1("tirag"))%>"></TD>
			<TD dir="LTR" align='center'><INPUT TYPE="text" NAME="zinc" size=3 value="<%=trim(RS1("zinc"))%>"></TD>
			<TD dir="RTL" align='center'>
				<SELECT NAME="is2roo" style="font-family:tahoma; font-size:8pt">
					<option value="0" <% if RS1("is2roo")="0" then %>selected<% end if %>>Ìﬂ —Ê</option>
					<option value="1" <% if RS1("is2roo")="1" then %>selected<% end if %>>”ﬂÂ</option>
					<option value="2" <% if RS1("is2roo")="2" then %>selected<% end if %>>‰‘«‰</option>
					<option value="3" <% if RS1("is2roo")="3" then %>selected<% end if %>>œÊ”—Ì</option>
				</SELECT>
			</TD>
			<TD dir="RTL" align='center'>
				<SELECT NAME="colorStatus" style="font-family:tahoma; font-size:8pt">
					<option value="0" <% if RS1("colorStatus")="0" then %>selected<% end if %>>«” «‰œ«—œ</option>
					<option value="1" <% if RS1("colorStatus")="1" then %>selected<% end if %>>‰„Ê‰Â</option>
					<option value="2" <% if RS1("colorStatus")="2" then %>selected<% end if %>>„‘ —Ì</option>
				</SELECT>
			</TD>
		</TR>
	<%
next
%>
	</table>
	<CENTER><BR>
	<INPUT TYPE="submit" value=" «ÌÌœ"> <INPUT TYPE="button" value="«‰’—«›" onclick="javascript:history.go(-1)">
	</CENTER>
	</FORM>
<%

response.end

end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------- Show Order Desk Table
'-----------------------------------------------------------------------------------------------------
%>
<TABLE class="RepTable" width='100%' align='center'	>
<TR>
		<td align=left>Ã” ÃÊ:</td>
		<td align=right>
			<INPUT NAME="SearchBox" TYPE="Text" style="border:1 solid black;width:150px;" Value="" onKeyPress="return handleSearch();"></td>
			
		<td align=left><A HREF="#openOrders"> ·Ì”  ”›«—‘ Â«Ì »«“ <FONT face="Wingdings" size=8>H</FONT></A></td>
</TR>
</TABLE>

<TaBlE class="CustTable4" cellspacing="2" cellspacing="2">
<Tr>
<Td colspan="2" valign="top" align="center">
<FORM METHOD=POST ACTION="">
<table class="CustTable" cellspacing='1' style='width:98%;'>
	<tr>
		<td colspan="13" class="CusTableHeader" style="text-align:left;">
			<TABLE width=100% cellspacing=0 cellpadding=0>
			<TR>
				<TD><B>ÃœÊ· »—‰«„Â —Ì“Ì  Ê·Ìœ</B></TD>
				<TD align=left>
				<INPUT TYPE="submit" name="submit" value="„‘«ÂœÂ ¬Œ—Ì‰ »—‰«„Â"> 
				<INPUT TYPE="submit" name="submit" value="ç«Å »—‰«„Â —Ê“«‰Â"> 
				<INPUT TYPE="submit" name="submit" value="ÊÌ—«Ì‘"></TD>
			</TR>
			</TABLE>
		</td>
	</tr>
	<%
	Ord=request("Ord")

	select case Ord
	case "1":
		order= " ORDER BY orders_trace.return_date"
	case "-1":
		order= " ORDER BY orders_trace.return_date DESC"
	case "2":
		order= " ORDER BY orders_trace.customer_name"
	case "-2":
		order= " ORDER BY orders_trace.customer_name DESC"
	case "3":
		order= " ORDER BY orders_trace.radif_sefareshat"
	case "-3":
		order= " ORDER BY orders_trace.radif_sefareshat DESC"
	case "4":
		order= " ORDER BY orderDesk.[size]"
	case "-4":
		order= " ORDER BY orderDesk.[size] DESC"
	case "5":
		order= " ORDER BY convert(int,orderDesk.forms)* convert(int,orderDesk.Tirag), orders_trace.return_date"
	case "-5":
		order= " ORDER BY convert(int,orderDesk.forms)* convert(int,orderDesk.Tirag) DESC, orders_trace.return_date"
	case "6":
		order= " ORDER BY convert(int,orderDesk.Zinc), orders_trace.return_date"
	case "-6":
		order= " ORDER BY convert(int,orderDesk.Zinc) DESC , orders_trace.return_date"
	case "7":
		order= " ORDER BY orderDesk.is2roo"
	case "-7":
		order= " ORDER BY orderDesk.is2roo DESC"
	case "8":
		order= " ORDER BY orderDesk.colorStatus"
	case "-8":
		order= " ORDER BY orderDesk.colorStatus DESC"
	case "9":
		order= " ORDER BY orders_trace.marhale"
	case "-9":
		order= " ORDER BY orders_trace.marhale DESC"
	case else:
		order= " ORDER BY orders_trace.return_date"
		Ord=1
	end select


	mySQL="SELECT orders_trace.return_date, orders_trace.marhale, orders_trace.customer_name, orders_trace.company_name, orders_trace.radif_sefareshat, orders_trace.order_title, orderDesk.id, orderDesk.[size], orderDesk.Tirag, orderDesk.forms, orderDesk.Zinc, orderDesk.is2roo, orderDesk.colorStatus FROM orderDesk INNER JOIN orders_trace ON orderDesk.order_ID = orders_trace.radif_sefareshat INNER JOIN Orders ON orders_trace.radif_sefareshat = Orders.ID where (orders_trace.Type = 2) AND (orders_trace.status = 1) AND (orders_trace.step NOT IN (11, 13, 10)) AND (Orders.Closed = 0)" & order 
	Set RS1 = conn.execute(mySQL)
	if RS1.eof then
%>
		<tr>
			<td colspan="13" class="CusTD3">Œ«·Ì</td>
		</tr>
<%
	else

		if ord<0 then
			style="background-color: #33CC99;"
			arrow="<br><span style='font-family:webdings'>6</span>"
		else
			style="background-color: #33CC99;"
			arrow="<br><span style='font-family:webdings'>5</span>"
		end if
%>

		<tr class="repHeader" bgcolor="eeeeee" style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
			<td>#</td>
			<td onclick='go2Page(1,1);' style="<%if abs(ord)=1 then response.write style%>"> «—ÌŒ  ÕÊÌ·	<%if abs(ord)=1 then response.write arrow%></TD>
			<td onclick='go2Page(1,2);' style="<%if abs(ord)=2 then response.write style%>">”›«—‘ œÂ‰œÂ \ ⁄‰Ê«‰ ﬂ«—<%if abs(ord)=2 then response.write arrow%></TD>
			<td onclick='go2Page(1,9);' style="<%if abs(ord)=9 then response.write style%>">„—Õ·Â<%if abs(ord)=9 then response.write arrow%></TD>
			<td onclick='go2Page(1,3);' style="<%if abs(ord)=3 then response.write style%>"> ”›«—‘		<%if abs(ord)=3 then response.write arrow%></TD>
			<td onclick='go2Page(1,4);' style="<%if abs(ord)=4 then response.write style%>">”«Ì“		<%if abs(ord)=4 then response.write arrow%></TD>
			<td onclick='go2Page(1,5);' style="<%if abs(ord)=5 then response.write style%>"> ⁄œ«œ		<%if abs(ord)=5 then response.write arrow%></TD>
			<td onclick='go2Page(1,6);' style="<%if abs(ord)=6 then response.write style%>">“Ì‰ﬂ		<%if abs(ord)=6 then response.write arrow%></TD>
			<td onclick='go2Page(1,7);' style="<%if abs(ord)=7 then response.write style%>">—Ê			<%if abs(ord)=7 then response.write arrow%></TD>
			<td onclick='go2Page(1,8);' style="<%if abs(ord)=8 then response.write style%>">—‰ê			<%if abs(ord)=8 then response.write arrow%></TD>
			<td title=''>ﬂ«€– Ê ﬂ«— »Ì—Ê‰</TD>
		</tr>
		<TR bgcolor="black" height="1">
			<TD colspan="10" style="padding:0;"></TD>
		</TR>

<%
		tmpCounter=0
		Do while not RS1.eof 
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
			'alert(this.getElementByTagName('td').items(0).data);
%>
			<TR bgcolor="<%=tmpColor%>" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>';this.nextSibling.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>';this.nextSibling.style.backgroundColor='<%=tmpColor%>'">
				<TD rowspan=2 style="height:30px;" align=center><%=tmpCounter%><BR><INPUT TYPE="checkbox" NAME="id" value="<%=RS1("id")%>"></TD>
				<TD rowspan=2 dir="LTR" align='center'><%=RS1("return_date")%>
				</TD>
				<TD rowspan=2 colspan=2><A target="_blank"  HREF="../shopfloor/manageOrder.asp?radif=<%=RS1("radif_sefareshat")%>"><%=RS1("customer_name")%> - <%=RS1("company_name")%><BR><BR><font color="aaaaaa"><%=RS1("order_title")%></A></font> 
				
				</TD>
				<TD dir="LTR" align='center'><%=RS1("radif_sefareshat")%></TD>
				<TD dir="LTR" align='center'><%=RS1("size")%></TD>
				<TD dir="LTR" align='center'><%=RS1("forms")%>*<%=RS1("Tirag")%></TD>
				<TD dir="LTR" align='center'> <%=RS1("Zinc")%></TD>
				<TD><%
				is2roo = ""

				if RS1("is2roo")="0" then 
					is2roo="Ìﬂ —Ê"
				elseif RS1("is2roo")="1" then 
					is2roo="”ﬂÂ"
				elseif RS1("is2roo")="2" then 
					is2roo="‰‘«‰"
				elseif RS1("is2roo")="3" then 
					is2roo="œÊ”—Ì"
				end if
				response.write is2roo
				
				%>&nbsp;</TD>
				<TD><%
				colorStatus = ""

				if RS1("colorStatus")="0" then 
					colorStatus="«” «‰œ«—œ"
				elseif RS1("colorStatus")="1" then 
					colorStatus="‰„Ê‰Â"
				elseif RS1("colorStatus")="2" then 
					colorStatus="„‘ —Ì"
				end if
				response.write colorStatus
				
				%>&nbsp;</TD>
				<TD rowspan=2 >
				<%
				mySQL="SELECT PurchaseOrders.ID, PurchaseRequests.TypeName, PurchaseOrders.Status FROM PurchaseRequests INNER JOIN PurchaseRequestOrderRelations ON PurchaseRequests.ID = PurchaseRequestOrderRelations.Req_ID INNER JOIN PurchaseOrders ON PurchaseRequestOrderRelations.Ord_ID = PurchaseOrders.ID WHERE PurchaseRequests.Order_ID = " & RS1("radif_sefareshat")
				Set RS2 = conn.execute(mySQL)
				Do while not RS2.eof 
					if RS2("status")="OK" then
						response.write "<IMG SRC='../images/yes.gif' BORDER=0> " 
					elseif RS2("status")="CANCEL" then
						response.write "<IMG SRC='../images/no.gif' BORDER=0> " 
					else
						response.write "<IMG SRC='../images/hum.gif' BORDER=0> " 
					end if
					response.write "<A TARGET='_blank' HREF='../purchase/outServiceTrace.asp?od=" & RS2("ID") & "'>" & RS2("TypeName") & "</A><BR>"
					RS2.moveNext
				Loop
				RS2.close
				set RS2=nothing
				%><BR>							
				<%
				mySQL="SELECT ItemName FROM InventoryItemRequests where Status <> 'del' and Order_ID=" & RS1("radif_sefareshat")
				Set RS3 = conn.execute(mySQL)
				Do while not RS3.eof 
					response.write "<li>" & RS3("ItemName") & "<BR>"
					RS3.moveNext
				Loop
				RS3.close
				set RS3=nothing
				%>							
				&nbsp;</TD>
			</TR>
			<TR bgcolor="<%=tmpColor%>" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>';this.previousSibling.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>';this.previousSibling.style.backgroundColor='<%=tmpColor%>'">
				<TD colspan=6 align=center><%=RS1("marhale")%>
				</TD>
			</TR>
<%						RS1.moveNext
		Loop
	end if
	%>
	<tr>
		<td colspan="12" class="CusTableHeader" style="text-align:left;">
		<INPUT TYPE="submit" name="submit" value="„‘«ÂœÂ ¬Œ—Ì‰ »—‰«„Â"> 
		<INPUT TYPE="submit" name="submit" value="ç«Å »—‰«„Â —Ê“«‰Â"> <INPUT TYPE="submit" name="submit" value="ÊÌ—«Ì‘"></td>
	</tr>
</table><BR><BR><BR><BR>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--
function go2Page(p,ord) {
	if(ord==0){
		ord=<%=Ord%>;
	}
	else if(ord==<%=Ord%>){
		ord= 0-ord;
	}
	str='?Ord='+escape(ord)+'&p='+escape(p) 
	window.location=str;
}
//-->
</SCRIPT>

<SCRIPT LANGUAGE="JavaScript">
function documentKeyDown() {
	var theKey = window.event.keyCode;
	var obj = window.event.srcElement;
	if (theKey == 114) { 
		if (obj.name=="SearchBox"){document.all.SearchBox.value="";};
		window.event.keyCode=0;
		document.all.SearchBox.focus();
		return false;
	}
}

document.onkeydown = documentKeyDown;

var lastFund = 0;
var lastSrch = "";
function handleSearch(){
	var theKey=event.keyCode;
	if (theKey==13){
		event.keyCode=0;
		srch=document.all.SearchBox.value;
		if (srch == '') {
			return;
		}
		if (srch!=lastSrch){
			lastFund = 0;
			lastSrch=srch;
		}
		var found = false;
		var text = document.body.createTextRange();
		found=text.findText(srch)
		for (var i=0; i<=lastFund && found ; i++) {
			found=text.findText(srch)
			text.moveStart("character", 1);
			text.moveEnd("textedit");
		}
		if (found) {
			text.moveStart("character", -1);
			text.findText(srch);
			text.select();
			lastFund++;
			theRow=text.parentElement();
			while(theRow.nodeName!='TR'){
				theRow=theRow.parentNode;
			}
			theRow.scrollIntoView();

		}
		else{
			if (lastFund == '0'){
				alert('⁄»«—  "' + srch +'" œ— «Ì‰ ’›ÕÂ ÅÌœ« ‰‘œ.');
			}
			else{
				alert('ÂÌç Ã«Ì œÌê—Ì ⁄»«—  "' + srch +'" ÅÌœ« ‰‘œ.');
			}
			lastFund=0;
			lastSrch="";
		}
	}
}
</script>


<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------- Show Order Desk Table
'-----------------------------------------------------------------------------------------------------
%>
<A name="openOrders"></A>
<TaBlE class="CustTable4" cellspacing="2" cellspacing="2">
<Tr>
<Td colspan="2" valign="top" align="center">
	<table class="CustTable" cellspacing='1' style='width:95%;'>
		<tr>
			<td colspan="12" class="CusTableHeader" style="text-align:right;">›Â—”  ﬂ«—Â«Ì »«“ «›” </td>
		</tr>
<%
		mySQL="SELECT *, orders_trace.radif_sefareshat FROM orders_trace INNER JOIN Orders ON orders_trace.radif_sefareshat = Orders.ID LEFT OUTER JOIN orderDesk ON Orders.ID = orderDesk.order_ID WHERE (orders_trace.Type = 2) AND (orders_trace.status = 1) AND (orders_trace.step NOT IN (11, 13, 10)) AND (Orders.Closed = 0) AND (orderDesk.id IS NULL) ORDER BY orders_trace.radif_sefareshat"
		Set RS1 = conn.execute(mySQL)
		if RS1.eof then
%>
			<tr>
				<td colspan="12" class="CusTD3">ÂÌç</td>
			</tr>
<%
		else
%>
			<tr>
				<td class="CusTD3"><small> «—ÌŒ  ÕÊÌ·</small></td>
				<td class="CusTD3"><small>”›«—‘ œÂ‰œÂ \<font color="aaaaaa"> ⁄‰Ê«‰ ﬂ«—</small></font></td>
				<td class="CusTD3"><small> ”›«—‘</small></td>
				<td class="CusTD3"><small>”«Ì“</small></td>
				<td class="CusTD3"><small> ⁄œ«œ ›—„</small></td>
				<td class="CusTD3"><small> Ì—«é Â— ›—„</small></td>
				<td class="CusTD3"><small>“Ì‰ﬂ</small></td>
				<td class="CusTD3"><small>—Ê</small></td>
				<td class="CusTD3"><small>—‰ê</small></td>
				<td class="CusTD3"><small> &nbsp; </small></td>
			</tr>
<%					tmpCounter=0
			Do while not RS1.eof 

				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
				'alert(this.getElementByTagName('td').items(0).data);
%>				<FORM METHOD=POST ACTION="" onsubmit="document.getElementById('addForm').disabled=true;">
				<TR bgcolor="<%=tmpColor%>" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'">
					<TD dir="LTR" align='right'><%=RS1("return_date")%></TD>
					<TD><A HREF="../shopfloor/manageOrder.asp?radif=<%=RS1("radif_sefareshat")%>" target="_blank"><%=RS1("customer_name")%> - <%=RS1("company_name")%><BR><BR><font color="aaaaaa"><%=RS1("order_title")%></font></A></TD>
					<TD dir="LTR" align='center'><INPUT TYPE="hidden" NAME="radif_sefareshat" value="<%=RS1("radif_sefareshat")%>" size=3><%=RS1("radif_sefareshat")%></TD>
					<TD dir="LTR" align='center'><SMALL>”«Ì“</SMALL><BR><INPUT TYPE="text" NAME="size" size=5></TD>
					<TD dir="LTR" align='center'><SMALL>›—„</SMALL><BR><INPUT TYPE="text" NAME="forms" size=3></TD>
					<TD dir="LTR" align='center'><SMALL> Ì—«é</SMALL><BR><INPUT TYPE="text" NAME="tirag" size=3></TD>
					<TD dir="LTR" align='center'><SMALL>“Ì‰ﬂ</SMALL><BR><INPUT TYPE="text" NAME="zinc" size=3></TD>
					<TD dir="RTL" align='center'><BR>
						<SELECT NAME="is2roo" style="font-family:tahoma; font-size:8pt">
							<option value="0">Ìﬂ —Ê</option>
							<option value="1">”ﬂÂ</option>
							<option value="2">‰‘«‰</option>
							<option value="3">œÊ”—Ì</option>
						</SELECT>
					</TD>
					<TD dir="RTL" align='center'><BR>
						<SELECT NAME="colorStatus" style="font-family:tahoma; font-size:8pt">
							<option value="0">«” «‰œ«—œ</option>
							<option value="1">‰„Ê‰Â</option>
							<option value="2">„‘ —Ì</option>
						</SELECT>
					</TD>
					<TD dir="LTR" align='center'>
					
					<!INPUT TYPE="image" src="../images/add.gif" NAME="submit" value="add" size=3>
					<input type="submit" id="addForm" NAME="submit" value="»Ì«›“«">					
					</TD>
				</TR></FORM>
<%						RS1.moveNext
			Loop
		end if
		%>
	</table>


<!--#include file="tah.asp" -->
