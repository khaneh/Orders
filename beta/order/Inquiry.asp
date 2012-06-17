<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="«” ⁄·«„"
SubmenuItem=9
if not Auth(2 , 9) then NotAllowdToViewThisPage() '«” ⁄·«„ 
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
'Server.ScriptTimeout = 3600
%>
<STYLE>
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.mySection{border: 1px #F90 dashed;margin: 15px 10px 0 15px;padding: 5px 0 5px 0;}
	.myRow{border: 2px #F05 dashed;margin: 10px 0 10px 0;padding: 0 3px 5px 0;}
	.exteraArea{border: 1px #33F dotted;margin: 5px 0 0 5px;padding: 0 3px 5px 0;}
	.myLabel {margin: 0px 3px 0 0px;white-space: nowrap;padding: 5px 0 5px 0;}
	.myProp {font-weight: bold;color: #40F; margin: 0px 3px 0 0px;padding: 5px 0 5px 0;}
	div.btn label{background-color:yellow;color: blue;padding: 3px 30px 3px 30px;cursor: pointer;}
	div.btn{margin: -5px 250px 0px 5px;}
	div.btn img{margin: 0px 20px -5px 0;cursor: pointer;}
	span.price{background-color: #FED;}
</STYLE>

<%
'-----------------------------
' Trace Quote
'-----------------------------
if Request.QueryString("act")="" then
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
	<hr>
	<TABLE border="4" cellspacing="0" cellpadding="0" width="600" align="center" bordercolor="#555599">
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="5" dir="RTL" width="100%">
	<!-- Trace Quote -->
		<FORM METHOD=POST ACTION="?act=search" onSubmit="return checkValidation();">
		<TR bgcolor="#AAAAEE"><TD colspan="4">
			<FONT SIZE="" COLOR="#555599" ><B>ÅÌêÌ—Ì «” ⁄·«„:</B></FONT>
		</TD></TR>
		<TR bgcolor="#AAAAEE">
			<TD>‰«„ „‘ —Ì Ì« ‘—ﬂ  Ì« ‘„«—Â «” ⁄·«„:</TD>
			<TD><INPUT TYPE="text" NAME="search_box" value="<%=request.form("search_box")%>"></TD>
			<TD><INPUT TYPE="submit" NAME="SubmitB" Value="Ã” ÃÊ" style="font-family:tahoma,arial; font-size:10pt;width:100px;"></TD>
			<TD align="left"><% if Auth(2 , 5) then %><A HREF="?act=advancedSearch">Ã” ÃÊÌ ÅÌ‘—› Â</A><% End If %></TD>
		</TR>
		</FORM>
		</TABLE>
	</TD></TR>
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="5" dir="RTL" width="100%">
	<!-- Input Quote -->
		<FORM METHOD=POST ACTION="?act=quoteInpCustSearch" onsubmit="if (document.all.CustomerNameSearchBox.value=='') return false;">
			<TR bgcolor="#AAAAEE">
				<TD>
					<FONT SIZE="" COLOR="#555599" ><B>Ê—Êœ «” ⁄·«„ ÃœÌœ:</B></FONT>
				</TD>
			</TR>
			<TR bgcolor="#AAAAEE">
				<TD>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«» &nbsp;
					<INPUT TYPE="text" NAME="CustomerNameSearchBox">&nbsp;
					<INPUT TYPE="submit" value="Ã” ÃÊ"><br>

			<SCRIPT LANGUAGE="JavaScript">
				document.all.CustomerNameSearchBox.focus();
			</SCRIPT>
				</TD>
			</TR>
		</FORM>
		</TABLE>
	</TD></TR>
	</TABLE>
	<script language="JavaScript">
		document.all.search_box.focus();
	</script>
	<hr>
<%
	'
elseif Request.QueryString("act")="search" then
%>
	<hr>
	<TABLE border="4" cellspacing="0" cellpadding="0" width="600" align="center" bordercolor="#555599">
	<FORM METHOD=POST ACTION="?act=search" onSubmit="return checkValidation();">
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="5" dir="RTL" width="100%">
		<TR bgcolor="#AAAAEE">
			<TD>‰«„ „‘ —Ì Ì« ‘—ﬂ  Ì« ‘„«—Â «” ⁄·«„:</TD>
			<TD><INPUT TYPE="text" NAME="search_box" value="<%=request.form("search_box")%>"></TD>
			<TD><INPUT TYPE="submit" NAME="SubmitB" Value="Ã” ÃÊ" style="font-family:tahoma,arial; font-size:10pt;width:100px;"></TD>
			<TD align="left"><A HREF="?act=advancedSearch">Ã” ÃÊÌ ÅÌ‘—› Â</A></TD>
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
	<hr>
<%
	
	search=request("search_box")
	if search="" then
		'By Default show Open Quotes of Current User
		myCriteria= "Quotes.CreatedBy = " & session("ID")
	elseif isNumeric(search) then
		search=clng(search)
		myCriteria= "Quotes.ID = '"& search & "'"
	else
		search=sqlSafe(search)
		myCriteria= "REPLACE([company_name], ' ', '') LIKE REPLACE(N'%"& search & "%', ' ', '') OR REPLACE([customer_name], ' ', '') LIKE REPLACE(N'%"& search & "%', ' ', '')"
	End If

	mySQL=	"SELECT Quotes.*, OrderTraceStatus.Name AS StatusName, OrderTraceStatus.Icon , Invoices.ID AS InvoiceID, Invoices.Approved, Invoices.Voided, Invoices.Issued "&_ 
			" FROM Quotes INNER JOIN "&_ 
			" OrderTraceStatus ON OrderTraceStatus.ID = Quotes.status LEFT OUTER JOIN "&_ 
			" Invoices INNER JOIN "&_ 
			" InvoiceQuoteRelations ON Invoices.ID = InvoiceQuoteRelations.Invoice ON Quotes.ID = InvoiceQuoteRelations.Quote "&_ 
			" WHERE ("& myCriteria & ")  "&_ 
			" ORDER BY Quotes.order_date DESC, Quotes.ID DESC "

	set RS1=Conn.Execute (mySQL)
	if not RS1.eof then
		tmpCounter=0
%>
	<div align="center" dir="LTR" >
	<table border="1" cellspacing="0" cellpadding="2" dir="RTL"  borderColor="#555588" >
		<TR valign=top bgcolor="#CCCCFF">
			<TD width="41"># «” ⁄·«„</TD>
			<TD width="65"> «—ÌŒ «” ⁄·«„<br> «—ÌŒ «⁄ »«—</TD>
			<TD width="130">‰«„ ‘—ﬂ  - „‘ —Ì</TD>
			<TD width="110">⁄‰Ê«‰ ﬂ«—</TD>
			<TD width="36">‰Ê⁄</TD>
			<TD width="76">„—Õ·Â</TD>
			<TD width="56">«” ⁄·«„ êÌ—‰œÂ</TD>
			<TD width="40">›«ﬂ Ê—</TD>
		</TR>
<%		Do while not RS1.eof

		if isnull(RS1("InvoiceID")) then
			InvoiceStatus="<span style='color:red;'><b>‰œ«—œ</b></span>"
		else
			if RS1("Voided") then
				style="style='color:Red' Title='»«ÿ· ‘œÂ'"
			elseif RS1("Issued") then
				style="style='color:Red' Title='’«œ— ‘œÂ'"
			elseif RS1("Approved") then
				style="style='color:Green' Title=' «ÌÌœ ‘œÂ'"
			else
				style="style='color:#3399FF' Title=' «ÌÌœ ‰‘œÂ'"
			end if
			InvoiceStatus="<A " & style & " HREF='../AR/AccountReport.asp?act=showInvoice&invoice=" & RS1("InvoiceID")& "' Target='_blank'>" & RS1("InvoiceID") & "</A>"
		end if
			
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
		Else
			tmpColor="#EEEEEE"
		End If 
%>
		<TR valign=top bgcolor="<%=tmpColor%>">
			<TD width="40" DIR="LTR"><A HREF="?act=show&quote=<%=RS1("ID")%>" target="_blank"><%=RS1("ID")%></A></TD>
			<TD width="65" DIR="LTR"><%=RS1("order_date")%><br><%=RS1("return_date") & " ("& RS1("return_time") & ")"%></TD>
			<TD width="130"><%=RS1("company_name")%><br><span style='color:gray'><%=RS1("customer_name")%></span><br> ·›‰:(<%=RS1("telephone")%>)&nbsp;</TD>
			<TD width="80"><%=RS1("order_title")%>&nbsp;</TD>
			<TD width="40"><%=RS1("order_kind")%></TD>
			<TD width="85"><%=RS1("marhale")%></TD>
			<TD width="50"><%=RS1("salesperson")%>&nbsp;</TD>
			<TD width="30"><%=InvoiceStatus%>&nbsp;</TD>
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
<%	
	elseif request("search_box")<>"" Then
	
%>	<TABLE border="1" cellspacing="0" cellpadding="0" dir="RTL" align="center" width="600">
		<TR bgcolor="#FFFFDD">
			<TD align="center" style="height:40px;font-size:12pt;font-weight:bold;color:red;padding:5px;">«Ì‰ «” ⁄·«„ ﬁ«»· ÅÌêÌ—Ì ‰Ì” <br>
			(Ì« ﬁ»·« ›«ﬂ Ê— ‘œÂ Ì« ﬂ‰”· ‘œÂ Ê Ì« «’·« ÊÃÊœ ‰œ«—œ)<br><br>
			»—«Ì «ÿ„Ì‰«‰ »Â <A HREF="?act=show&quote=<%=request("search_box")%>" style="color:blue;">«’·«Õ Ê ‰„«Ì‘</A> „—«Ã⁄Â ﬂ‰Ìœ.</TD>
		</TR>
	</TABLE>
<%	
	End If

elseif Request.QueryString("act")="show" then
  if isnumeric(request("quote")) then
	quote=request("quote")
	'mySQL="SELECT Accounts.ID AS AccID, Accounts.AccountTitle, Quotes.* FROM Accounts INNER JOIN Quotes ON Accounts.ID = Quotes.Customer WHERE (Quotes.ID='"& quote & "')"

	mySQL=	"SELECT Quotes.*, Accounts.ID AS AccID, Accounts.AccountTitle, "&_ 
			"Invoices.ID AS InvoiceID, Invoices.Approved, Invoices.Voided, Invoices.Issued "&_ 
			"FROM Quotes INNER JOIN "&_ 
			"Accounts ON Accounts.ID = Quotes.Customer LEFT OUTER JOIN "&_ 
			"InvoiceQuoteRelations INNER JOIN "&_ 
			"Invoices ON InvoiceQuoteRelations.Invoice = Invoices.ID ON InvoiceQuoteRelations.Quote = Quotes.ID "&_ 
			"WHERE (Quotes.ID='"& quote & "')"
'response.write mySQL
	set RS1=conn.execute (mySQL)
	If RS1.EOF then
		response.write "<BR><BR><BR><BR><CENTER>‘„«—Â «” ⁄·«„ „⁄ »— ‰Ì” </CENTER>"
		response.end
	End If

	If RS1("Step")=4 then
		stamp="<div style='border:2 dashed red;width:150px; text-align:center; padding: 10px;color:red;font-size:15pt;font-weight:bold;'>—œ ‘œÂ</div>"
	End If

%>
	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr height="10">
			<td colspan=2></td>
		</tr>
		<tr height="10">
			<td width="150"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="20">
			<td colspan=2></td>
		</tr>
	</table>

	<CENTER>
		<input type="button" value="ÊÌ—«Ì‘" Class="GenButton" onclick="window.location='?act=editQuote&quote=<%=quote%>';">&nbsp;
		<%' 	ReportLogRow = PrepareReport ("OrderForm.rpt", "Order_ID", quote, "/beta/dialog_printManager.asp?act=Fin") %>
		<!--INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<% 'ReportLogRow%>);"-->
	<% If isnull(RS1("InvoiceID")) Then %>
		<input type="button" value="ÅÌ‘ ‰ÊÌ” ﬁÌ„ " Class="GenButton" onclick="window.location='../AR/InvoiceInput.asp?act=getInvoice&selectedCustomer=<%=RS1("AccID")%>&selectedQuotes=<%=quote%>';">
	<% Else %>
		<input type="button" value="ÅÌ‘ ‰ÊÌ” ﬁÌ„  (<%=RS1("InvoiceID")%>)" Class="GenButton" style="color:#3399FF;" onclick="window.location='../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("InvoiceID")%>';">
	<% End If %>

	<% If RS1("Step")<>5 Then %>
		<input type="button" value=" »œÌ· »Â ”›«—‘" Class="GenButton" onclick="window.location='?act=convertToOrder&quote=<%=quote%>';">
	<% Else %>
		<font color="red">&nbsp; (<b> »œÌ· »Â ”›«—‘ ‘œÂ</b>)</font>
	<% End If %>
	</CENTER>
	
	<BR>
	<TABLE class="" cellspacing="0" cellpadding="2" align="center" style="background-color:#CCCCCC; color:black; direction:RTL; width:700; border: 2 solid #555599;">
		<TR bgcolor="#555599">
			<TD align="left"><FONT COLOR="YELLOW">Õ”«»:</FONT></TD>
			<TD align="right" colspan=5 height="25px">
				<span id="customer" style="color:yellow;"><%' after any changes in this span "./Customers.asp" must be revised%>
					<span><%=RS1("AccID") & " - "& RS1("AccountTitle")%></span>.
				</span>
			</TD>
		</TR>
		
		<TR bgcolor="#555599" height=30 style="color:yellow;">
			<TD align="left">‘„«—Â «” ⁄·«„:</TD>
			<TD align="right"><%=RS1("ID")%></TD>
			<TD align="left"> «—ÌŒ:</TD>
			<TD><span dir="LTR"><%=RS1("order_date")%></span></TD>
			<TD align="left">”«⁄ :</TD>
			<TD align="right"><%=RS1("order_time")%></TD>
		</TR>
		<TR height=30>
			<TD align="left">‰«„ ‘—ﬂ :</TD>
			<TD><%=RS1("company_name")%></TD>
			<TD align="left">„Ê⁄œ «⁄ »«—:</TD>
			<TD align="right" dir=LTR><%=RS1("return_date")%></TD>
			<TD align="left">”«⁄  «⁄ »«—:</TD>
			<TD align="right" dir=LTR><%=RS1("return_time")%></TD>
		</TR>
		<TR height=30>
			<TD align="left">‰«„ „‘ —Ì:</TD>
			<TD><%=RS1("customer_name")%></TD>
			<TD align="left">‰Ê⁄ «” ⁄·«„:</TD>
			<TD><%=RS1("order_kind")%></TD>
			<TD align="left">«” ⁄·«„ êÌ—‰œÂ:</TD>
			<TD><%=RS1("salesperson")%>	</TD>
		</TR>
		<TR height=30>
			<TD align="left"> ·›‰:</TD>
			<TD><%=RS1("telephone")%></TD>
			<TD align="left">⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì·:</TD>
			<TD colspan="3"><%=RS1("order_title")%></TD>
		</TR>
	
		<TR height=30>
			<td align="left"> ⁄œ«œ:</td>
			<td><%=rs1("qtty")%></td>
			<TD align="left">”«Ì“:</TD>
			<TD><%=RS1("PaperSize")%></TD>
			<TD align="left">“„«‰  Ê·Ìœ:</TD>
			<TD>
			<%if RS1("productionDuration") = 0 then 
				response.write "<font color=red>‰«„⁄·Ê„</font>" 
			else
				response.write RS1("productionDuration") & " —Ê“"
			end if %>
			
			</TD>
		</TR>
		<TR height=30>
			<TD align="left">„—Õ·Â:</TD>
			<TD colspan="3"><%=RS1("marhale")%></TD>
			<TD align="left">ﬁÌ„  ﬂ·:</TD>
			<TD><span class="price"><%=RS1("Price")%></span></TD>
		</TR>
		<TR height=30>
			<TD align="left" valign="top" colspan="3"> Ê÷ÌÕ«  »Ì‘ —:</TD>
			<TD colspan="3">
			<%if not IsNull(RS1("Notes")) then response.write replace(RS1("Notes"),chr(13),"<br>")%>
			</TD>
		</TR>
	</TABLE>
<%
set rs = conn.Execute("select *,users.RealName from QuoteLogs inner join Users on QuoteLogs.LastUpdatedBy=users.ID where QuoteLogs.QuoteID=" & quote & " and QuoteLogs.ID in (select min(id) from QuoteLogs group by QuoteID,productionDuration)")
if not rs.eof then 
%>
	<TABLE class="" cellspacing="0" cellpadding="2" align="center" style="background-color:#CCCCCC; color:black; direction:RTL; width:700; border: 2 solid #555599;">
	<tr bgcolor="#555599" height=30 style="color:yellow;">
		<td> «—ÌŒ</td>
		<td>”«⁄ </td>
		<td>“„«‰  Ê·Ìœ</td>
		<td> Ê”ÿ</td>
	</tr>
<%
	while not rs.eof 
		response.write "<tr><td>" & rs("LastUpdatedDate") & "</td><td>" & rs("LastUpdatedTime") & "</td><td>" 
		if CInt(rs("productionDuration"))=0 then 
			response.write "<font color=red>‰«„⁄·Ê„</font>"
		else
			response.write  rs("productionDuration") & " —Ê“"
		end if
		response.write  "</td><td>" & rs("RealName") & "</td></tr>"
		rs.moveNext
	wend
%>	
	<tr>
	
	</tr>
	</table>
<%
end if
%>
	<BR>
	<%
if (not (IsNull(rs1("property")) or rs1("property")="")) then
%>
	<div>Ã“∆Ì«  «” ⁄·«„</div>

<%
	set rs=Conn.Execute("select * from OrderTraceTypes where id="&rs1("type"))
	set typeProp = server.createobject("MSXML2.DomDocument")
	set orderProp = server.createobject("MSXML2.DomDocument")
	
	orderProp.loadXML(rs1("property"))
	typeProp.loadXML(rs("property"))
	set rs=nothing
sub showKey(key)
	oldGroup="---first---"
	oldLabel="---first---"
	maxID=-1
	oldID=-1
	rowEmpty=false
	for each mykey in orderProp.SelectNodes(key)
		id=myKey.GetAttribute("id")
		if maxID<id then maxID=id
	next
	thisRow = "<div class='myRow'>"'<div class='exteraArea' id='" & Replace(key,"/","-") & "-0'>"
	for id = 0 to maxID
		For Each myKey In orderProp.SelectNodes(key & "[@id='" & id & "']")
			thisName = myKey.GetAttribute("name")
			set typeKey = typeProp.selectNodes(key & "[@name='" & thisName & "']")(0)
			thisType = typeKey.GetAttribute("type") 
			thisLabel= typeKey.GetAttribute("label")
			thisGroup= typeKey.GetAttribute("group")
			if thisType="radio" then 
				radioID = CInt(myKey.text)
				set typeKey = typeProp.selectNodes(key & "[@name='" & thisName & "']")(radioID - 1)
				thisType = typeKey.GetAttribute("type") 
				thisLabel= typeKey.GetAttribute("label")
				thisGroup= typeKey.GetAttribute("group")
			end if
			isRow =false
			if Replace(key,"/","-")="keys-service-key" then response.write "::--------::" & myKey.text
			if thisName<>"" then 
				isRow=true
				if oldID<>id then thisRow = thisRow & "<div class='exteraArea' id='" & Replace(key,"/","-") & "-" & id & "'>"
				if (oldGroup<>thisGroup and oldID=id and oldGroup <> "---first---") then thisRow = thisRow &  "</div>"
				if oldGroup<>thisGroup or oldID<>id then 
					thisRow = thisRow & "<div class='mySection'>"
					if typeKey.GetAttribute("grouplabel")<>"" then thisRow = thisRow & "<b>" & typeKey.GetAttribute("grouplabel") & "</b>"
				end if
				if oldLabel<>thisLabel and thisType<>"radio" then thisRow = thisRow &  "<label class='myLabel'>" & thisLabel & ": </label>"
				
				if left(thisType,6)="option" then set myOptions=typeKey
				myText=""
				select case thisType
					case "option"
						for each optKey in myOptions.selectNodes("option")
							if optKey.text=myKey.text then 
								myText = optKey.GetAttribute("label")
								exit for
							end if
						next
					case "option-other"
						if left(myKey.text,6)="other:" then 
							myText = mid(myKey.text,7)
						else
							for each optKey in myOptions.selectNodes("option")
								if optKey.text=myKey.text then 
									myText = optKey.GetAttribute("label")
									exit for
								end if
							next
						end if
						if myText="" then myText = myKey.text
					case "check"
						if left(myKey.text,2)="on" then myText = "<img src='/images/Checkmark-32.png' width='15px'>"
					case "radio"
						myText=thisLabel
					case else
						myText = myKey.text
				end select
				set myOptions=nothing
				thisRow = thisRow & "<span class='myProp'>" & myText & "</span>"		
			else
				if id=0 then 
					thisRow=""
					rowEmpty=true
				end if
			end if
			oldGroup=thisGroup
			oldLabel=thisLabel
			oldID=id
			if typeKey.GetAttribute("br")="yes" then thisRow = thisRow & "<br><br>"
		Next
		if isRow then thisRow = thisRow & "</div></div>"
	next
	'response.write maxID
	if not rowEmpty then thisRow = thisRow & "</div>" '"<div id='extreArea" &Replace(key,"/","-")& "'></div>"
	response.write thisRow 'prependTo 
'	response.write 
end sub
	oldTmp="---first---"
	for each tmp in orderProp.selectNodes("//key")
		if oldTmp<>tmp.parentNode.nodeName then 
			oldTmp=tmp.parentNode.nodeName
			call showKey("/keys/" & oldTmp & "/key")
		end if
	next
' 	call showKey("/keys/printing/key")
' 	call showKey("keys/binding/key")
' 	call showKey("keys/service/key")
' 	call showKey("keys/delivery/key")
	
end if
%>
<br><br>
	<table class="CustTable" cellspacing='1' align=center style="width:700; ">
		<tr>
			<td colspan="2" class="CusTableHeader"><span style="width:450;text-align:center;">Ì«œœ«‘  Â«</span><span style="width:100;text-align:left;background-color:red;"><input class="GenButton" type="button" value="‰Ê‘ ‰ Ì«œœ«‘ " onclick="window.location = '../home/message.asp?RelatedTable=quotes&RelatedID=<%=quote%>&retURL=<%=Server.URLEncode("../order/Inquiry.asp?act=show&quote="&quote)%>';"></span></td>
		</tr>
<%
	mySQL="SELECT * FROM Messages INNER JOIN Users ON Messages.MsgFrom = Users.ID WHERE (Messages.RelatedTable = 'quotes') AND (Messages.RelatedID = "& quote & ") ORDER BY Messages.ID DESC"
	Set RS = conn.execute(mySQL)
	if NOT RS.eof then

		tmpCounter=0
		Do While NOT RS.eof 
			tmpCounter=tmpCounter+1
%>
			<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>">
				<td>«“ <%=RS("RealName")%><br>
					<%=RS("MsgDate")%> <BR> <%=RS("MsgTime")%>
				</td>
				<td dir='RTL'><%=replace(RS("MsgBody"),chr(13),"<br>")%></td>
			</tr>
<%
			RS.moveNext
		Loop
	else
%>
		<tr class="CusTD3">
			<td colspan="2">ÂÌç</td>
		</tr>
<%
	end if
	RS.close
%>
	</table><BR>

<%
  end if
elseif Request.QueryString("act")="advancedSearch" then
'------  Advanced Search 
%>
<!--#include File="../include_JS_InputMasks.asp"-->
<%
	'Server.ScriptTimeout = 3600
	tmpTime=time
	tmpTime=Hour(tmpTime)&":"&Minute(tmpTime)
	if instr(tmpTime,":")<3 then tmpTime="0" & tmpTime
	if len(tmpTime)<5 then tmpTime=Left(tmpTime,3) & "0" & Right(tmpTime,1)

	if request("resultsCount")="" OR not isnumeric(request("resultsCount")) then
		resultsCount = 50
	else
		resultsCount = cint(request("resultsCount"))
	end if

%>
	<hr>
	<TABLE border="4" cellspacing="0" cellpadding="0" width="700" align="center" bordercolor="#555599">
	<FORM METHOD=POST ACTION="?act=advancedSearch" onSubmit="return checkValidation();">
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="1" dir="RTL" width="100%" bgcolor="white">
		<TR bgcolor="#EEEEEE">
			<TD><INPUT TYPE="checkbox" NAME="check_sefaresh" onclick="check_sefaresh_Click()" checked></TD>
			<TD>‘„«—Â «” ⁄·«„</TD>
			<TD><INPUT TYPE="text" NAME="az_sefaresh" dir="LTR" value="<%=request.form("az_sefaresh")%>" size="8" maxlength="6" onKeyPress="return maskNumber(this);"></TD>
			<TD> «</TD>
			<TD><INPUT TYPE="text" NAME="ta_sefaresh" dir="LTR" value="<%=request.form("ta_sefaresh")%>" size="8" maxlength="6" onKeyPress="return maskNumber(this);" ></TD>
			<td rowspan="12" style="width:1px" bgcolor="#555599"></td>
			<TD><INPUT TYPE="checkbox" NAME="check_kind" onclick="check_kind_Click()" checked></TD>
			<TD>‰Ê⁄ «” ⁄·«„</TD>
			<TD colspan="3"><SELECT NAME="order_kind_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
				<OPTION value="«›” " <%if request.form("order_kind_box")="«›” " then response.write "selected" %> >«›” </option>
				<OPTION value="œÌÃÌ «·" <%if request.form("order_kind_box")="œÌÃÌ «·" then response.write "selected" %> >œÌÃÌ «·</option>
				<OPTION value="”Ì«Â Ê ”›Ìœ" <%if request.form("order_kind_box")="”Ì«Â Ê ”›Ìœ" then response.write "selected" %> >”Ì«Â Ê ”›Ìœ</option>
				<OPTION value="ÿ—«ÕÌ" <%if request.form("order_kind_box")="ÿ—«ÕÌ" then response.write "selected" %> >ÿ—«ÕÌ</option>
				<OPTION value="’Õ«›Ì" <%if request.form("order_kind_box")="’Õ«›Ì" then response.write "selected" %> >’Õ«›Ì</option>
				<OPTION value="›Ì·„" <%if request.form("order_kind_box")="›Ì·„" then response.write "selected" %> >›Ì·„</option>
				<OPTION value="“Ì‰ﬂ" <%if request.form("order_kind_box")="“Ì‰ﬂ" then response.write "selected" %> >“Ì‰ﬂ</option>
				<OPTION value="·„Ì‰ " <%if request.form("order_kind_box")="·„Ì‰ " then response.write "selected" %> >·„Ì‰ </option>
				<OPTION value="„ ›—ﬁÂ" <%if request.form("order_kind_box")="„ ›—ﬁÂ" then response.write "selected" %> >„ ›—ﬁÂ</option>
			</SELECT></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD>
				<INPUT TYPE="checkbox" NAME="check_tarikh_sefaresh" onclick="check_tarikh_sefaresh_Click()" checked>
			</TD>
			<TD> «—ÌŒ «” ⁄·«„</TD>
			<TD>
				<INPUT TYPE="text" NAME="az_tarikh_sefaresh" dir="LTR" value="<%=request.form("az_tarikh_sefaresh")%>" size="10" onKeyPress="return maskDate(this);" onBlur="if(acceptDate(this))document.all.ta_tarikh_sefaresh.value=this.value;" maxlength="10">
			</TD>
			<TD> «</TD>
			<TD>
				<INPUT TYPE="text" NAME="ta_tarikh_sefaresh" dir="LTR" value="<%=request.form("ta_tarikh_sefaresh")%>" size="10" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10">
			</TD>
			<TD>
				<INPUT TYPE="checkbox" NAME="check_marhale" onclick="check_marhale_Click()" checked>
			</TD>
			<TD>„—Õ·Â</TD>
			<TD>
				<SELECT NAME="marhale_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
					<%
					set RS_STEP=Conn.Execute ("SELECT * FROM QuoteSteps WHERE (IsActive=1)")
					Do while not RS_STEP.eof	
					%>
						<OPTION value="<%=RS_STEP("ID")%>" <%if cint(request("marhale_box"))=cint(RS_STEP("ID")) then response.write "selected" %> ><%=RS_STEP("name")%></option>
						<%
						RS_STEP.moveNext
					loop
					RS_STEP.close
					%>
				</SELECT>
			</TD>
			<TD>
				<span id="marhale_not_check_label" style='font-weight:bold;color:red'>‰»«‘œ</span>
			</TD>
			<TD>
				<INPUT TYPE="checkbox" NAME="marhale_not_check" onclick="marhale_not_check_Click();" checked>
			</TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#EEEEEE">
			<TD>
				<INPUT TYPE="checkbox" NAME="check_tarikh_tahvil" onclick="check_tarikh_tahvil_Click()" checked>
			</TD>
			<TD> «—ÌŒ «⁄ »«—</TD>
			<TD>
				<INPUT TYPE="text" NAME="az_tarikh_tahvil" dir="LTR" value="<%=request.form("az_tarikh_tahvil")%>" size="10" onblur="acceptDate(this)" maxlength="10" onKeyPress="return maskDate(this);">
			</TD>
			<TD> «</TD>
			<TD>
				<INPUT TYPE="text" NAME="ta_tarikh_tahvil" dir="LTR" value="<%=request.form("ta_tarikh_tahvil")%>" onblur="acceptDate(this)" maxlength="10" size="10" onKeyPress="return maskDate(this);">
			</TD>
			<td colspan="5"></td>
			
		</TR>
		<TR bgcolor="#EEEEEE">
			<TD colspan="5">&nbsp;</TD>
			<TD>
				<INPUT TYPE="checkbox" NAME="check_closed" onclick="check_closed_Click()" checked>
			</TD>
			<TD colspan="4">
				<span id="check_closed_label" style='color:black;'>›ﬁÿ «” ⁄·«„ Â«Ì »«“</span>
			</TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD colspan="5">&nbsp;</TD>
			<TD>
				<INPUT TYPE="checkbox" NAME="check_salesperson" onclick="check_salesperson_Click()" checked>
			</TD>
			<TD>«” ⁄·«„ êÌ—‰œÂ</TD>
			<TD colspan="3">
				<SELECT NAME="salesperson_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
<%				set RSV=Conn.Execute ("SELECT RealName FROM Users WHERE Display=1 ORDER BY RealName") 
				Do while not RSV.eof
%>
					<option value="<%=RSV("RealName")%>" <%if RSV("RealName")=request.form("salesperson_box") then response.write " selected "%> ><%=RSV("RealName")%></option>
<%
				RSV.moveNext
				Loop
				RSV.close
%>
				</SELECT>
			</TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#EEEEEE">
			<TD>
				<INPUT TYPE="checkbox" NAME="check_company_name" onclick="check_company_name_Click()" checked>
			</TD>
			<TD>‰«„ ‘—ﬂ </TD>
			<TD colspan="3">
				<INPUT TYPE="text" NAME="company_name_box" value="<%=request.form("company_name_box")%>">
			</TD>
			<TD>
				<INPUT TYPE="checkbox" NAME="check_telephone" onclick="check_telephone_Click()" checked>
			</TD>
			<TD> ·›‰</TD>
			<TD colspan="3">
				<INPUT TYPE="text" NAME="telephone_box" value="<%=request.form("telephone_box")%>">
			</TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD>
				<INPUT TYPE="checkbox" NAME="check_customer_name" onclick="check_customer_name_Click()" checked>
			</TD>
			<TD>‰«„ „‘ —Ì</TD>
			<TD colspan="3">
				<INPUT TYPE="text" NAME="customer_name_box" value="<%=request.form("customer_name_box")%>">
			</TD>
			<TD>
				<INPUT TYPE="checkbox" NAME="check_order_title" onclick="check_order_title_Click()" checked>
			</TD>
			<TD>⁄‰Ê«‰ «” ⁄·«„</TD>
			<TD colspan="3">
				<INPUT TYPE="text" NAME="order_title_box" value="<%=request.form("order_title_box")%>">
			</TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#AAAAEE">
			<td colspan="11" style="height:30px">
			<TABLE align="center" width="50%">
			<TR>
				<TD><INPUT TYPE="submit" Name="Submit" Value=" «ÌÌœ" style="font-family:tahoma,arial; font-size:10pt;width:100px;"></TD>
				<TD align="left"><INPUT TYPE="button" Name="Cancel" Value="Å«ﬂ ﬂ‰" style="font-family:tahoma,arial; font-size:10pt;width:100px;" onclick="window.location='?act=advancedSearch';"></TD>
			</TR>
			<TR>
				<TD align="left"> ⁄œ«œ ‰ «ÌÃ:</TD>
				<TD>&nbsp;<INPUT TYPE="Text" Name="resultsCount" value="<%=resultsCount%>" maxlength="4" size="4" onKeyPress="return maskNumber(this);"></TD>
			</TR>
			</TABLE>
			</td>
		</TR>
		</TABLE>
	</TD></TR>
	</FORM>
	</TABLE>
	<hr>
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

	function check_closed_Click(){
		if (document.all.check_closed.checked) {
			document.all.check_closed_label.style.color='black'
		}
		else{
			document.all.check_closed_label.style.color='#BBBBBB'
		}
	}


	function Form_Load(){
	<%
	maybeAND = ""
	myCriteria = ""
	If request.form("check_sefaresh") = "on" Then
		if request.form("az_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & " Quotes.ID >= " & request.form("az_sefaresh")
			maybeAND=" AND "
		End If
		if request.form("ta_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & " Quotes.ID <= " & request.form("ta_sefaresh")
			maybeAND=" AND "
		End If
		If (request.form("az_sefaresh") = "") AND (request.form("ta_sefaresh") = "") then
			response.write "document.all.check_sefaresh.checked = false;" & vbCrLf
			response.write "document.all.az_sefaresh.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_sefaresh.style.visibility = 'hidden';" & vbCrLf
		End If
	Else
		response.write "document.all.check_sefaresh.checked = false;" & vbCrLf
		response.write "document.all.az_sefaresh.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.ta_sefaresh.style.visibility = 'hidden';" & vbCrLf

	End If

	If request.form("check_tarikh_sefaresh") = "on" Then
		if request.form("az_tarikh_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "order_date >= '" & request.form("az_tarikh_sefaresh") & "'"
			maybeAND=" AND "
		End If
		if request.form("ta_tarikh_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "order_date <= '" & request.form("ta_tarikh_sefaresh") & "'"
			maybeAND=" AND "
		End If 
		If (request.form("az_tarikh_sefaresh") = "") AND (request.form("ta_tarikh_sefaresh") = "") then
			response.write "document.all.check_tarikh_sefaresh.checked = false;" & vbCrLf
			response.write "document.all.az_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
		End If
	Else
		response.write "document.all.check_tarikh_sefaresh.checked = false;" & vbCrLf
		response.write "document.all.az_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.ta_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
	End If

	If request.form("check_tarikh_tahvil") = "on" Then
		if request.form("az_tarikh_tahvil") <> "" then

			myCriteria = myCriteria & maybeAND & "return_date >= '" & request.form("az_tarikh_tahvil") & "'"
			maybeAND=" AND "
		End If
		if request.form("ta_tarikh_tahvil") <> "" then
			myCriteria = myCriteria & maybeAND & "return_date <= '" & request.form("ta_tarikh_tahvil") & "'"
			maybeAND=" AND "
		End If
		If (request.form("az_tarikh_tahvil") = "") AND (request.form("ta_tarikh_tahvil") = "") then
			response.write "document.all.check_tarikh_tahvil.checked = false;" & vbCrLf
			response.write "document.all.az_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
		End If
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
	End If

	If (request.form("check_customer_name") = "on" AND request.form("customer_name_box") <> "")then
		myCriteria = myCriteria & maybeAND & "customer_name Like N'%" & request.form("customer_name_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_customer_name.checked = false;" & vbCrLf
		response.write "document.all.customer_name_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If request.form("check_kind") = "on" then
		myCriteria = myCriteria & maybeAND & "order_kind = N'" & request.form("order_kind_box") & "'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_kind.checked = false;" & vbCrLf
		response.write "document.all.order_kind_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If request("check_marhale") = "on" then
		If request("marhale_not_check") = "on" then
			myCriteria = myCriteria & maybeAND & "NOT(step = " & request("marhale_box") & ")"
		Else
			myCriteria = myCriteria & maybeAND & "step = " & request("marhale_box") 
			response.write "document.all.marhale_not_check.checked = false;" & vbCrLf
			response.write "document.all.marhale_not_check_label.style.color='#BBBBBB';"& vbCrLf
		End If
		maybeAND=" AND "
	Else
		response.write "document.all.check_marhale.checked = false;" & vbCrLf
		response.write "document.all.marhale_box.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.marhale_not_check.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.marhale_not_check_label.style.color='#BBBBBB';"& vbCrLf
	End If

	If request.form("check_salesperson") = "on" then
		myCriteria = myCriteria & maybeAND & "salesperson = N'" & request.form("salesperson_box") & "'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_salesperson.checked = false;" & vbCrLf
		response.write "document.all.salesperson_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If (request.form("check_order_title") = "on" AND request.form("order_title_box") <> "")then
		myCriteria = myCriteria & maybeAND & "order_title Like N'%" & request.form("order_title_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_order_title.checked = false;" & vbCrLf
		response.write "document.all.order_title_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If (request.form("check_telephone") = "on" AND request.form("telephone_box") <> "")then
		myCriteria = myCriteria & maybeAND & "telephone Like N'%" & request.form("telephone_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_telephone.checked = false;" & vbCrLf
		response.write "document.all.telephone_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If request("check_closed") = "on" then
		myCriteria = myCriteria & maybeAND & "Quotes.Closed=0"
	Else
		If request("Submit")=" «ÌÌœ" then
			response.write "document.all.check_closed.checked = false;" & vbCrLf
			response.write "document.all.check_closed_label.style.color='#BBBBBB';" & vbCrLf
		End If
	End If

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

			mySQL=	"SELECT Quotes.*, OrderTraceStatus.Name AS StatusName, OrderTraceStatus.Icon , Invoices.ID AS InvoiceID, Invoices.Approved, Invoices.Voided, Invoices.Issued, DRV_invoice.price "&_ 
					" FROM Quotes INNER JOIN "&_ 
					" OrderTraceStatus ON OrderTraceStatus.ID = Quotes.status LEFT OUTER JOIN "&_ 
					" Invoices INNER JOIN "&_ 
					" InvoiceQuoteRelations ON Invoices.ID = InvoiceQuoteRelations.Invoice ON Quotes.ID = InvoiceQuoteRelations.Quote "&_ 
					" left outer join (select Invoice,SUM(InvoiceLines.Price +InvoiceLines.Vat - InvoiceLines.Discount -InvoiceLines.Reverse) as price from InvoiceLines inner join Invoices on invoiceLines.invoice=invoices.id where invoices.voided=0 group by Invoice) DRV_invoice on Invoices.ID=DRV_invoice.Invoice "&_
					" WHERE ("& myCriteria & ")  "&_ 
					" ORDER BY Quotes.order_date DESC, Quotes.ID DESC "

'			mySQL=	"SELECT Quotes.*, OrderTraceStatus.Name AS StatusName, OrderTraceStatus.Icon "&_
'					" FROM Quotes JOIN OrderTraceStatus ON OrderTraceStatus.ID = Quotes.status "&_
'					" WHERE ("& myCriteria & ") ORDER BY order_date DESC, Quotes.ID DESC"
'
'response.write mySQL
			set RS1=Conn.Execute (mySQL)
			if not RS1.eof then
				tmpCounter=0
%>
			<div align="center" dir="LTR">
			<TABLE border="1" cellspacing="0" cellpadding="1" dir="RTL" borderColor="#555588">
				<TR bgcolor="#CCCCFF">
					<TD width="41"># «” ⁄·«„</TD>
					<TD width="46"> «—ÌŒ «” ⁄·«„</TD>
					<TD width="56">«⁄ »«—  «</TD>
					<TD width="116">‰«„ ‘—ﬂ </TD>
					<TD width="106">‰«„ „‘ —Ì</TD>
					<TD >⁄‰Ê«‰ ﬂ«—</TD>
					<TD width="36">‰Ê⁄</TD>
					<TD width="46">„—Õ·Â</TD>
					<TD width="36">«” ⁄·«„ êÌ—‰œÂ</TD>
					<TD width="40">›«ﬂ Ê—</TD>
					<td width="50">ﬁÌ„ </td>
				</TR>
<%				Do while not RS1.eof and tmpCounter < resultsCount
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
				Else
					tmpColor="#DDDDDD"
				End If

				if RS1("Closed") then
					tmpStyle="background-color:#FFCCCC;"
				else
					tmpStyle=""
				End If
				
				if isnull(RS1("InvoiceID")) then
					InvoiceStatus="<span style='color:red;'><b>‰œ«—œ</b></span>"
				else
					if RS1("Voided") then
						style="style='color:Red' Title='»«ÿ· ‘œÂ'"
					elseif RS1("Issued") then
						style="style='color:Red' Title='’«œ— ‘œÂ'"
					elseif RS1("Approved") then
						style="style='color:Green' Title=' «ÌÌœ ‘œÂ'"
					else
						style="style='color:#3399FF' Title=' «ÌÌœ ‰‘œÂ'"
					end if
					InvoiceStatus="<A " & style & " HREF='../AR/AccountReport.asp?act=showInvoice&invoice=" & RS1("InvoiceID")& "' Target='_blank'>" & RS1("InvoiceID") & "</A>"
				end if

%>
				<TR bgcolor="<%=tmpColor%>" title="<%=RS1("StatusName")%>">
					<TD DIR="LTR"><A HREF="?act=show&quote=<%=RS1("ID")%>" target="_blank"><%=RS1("ID")%></A></TD>
					<TD DIR="LTR"><%=RS1("order_date")%></TD>
					<TD DIR="LTR"><%=RS1("return_date") & " ("& RS1("return_time") & ")"%></TD>
					<TD ><%=RS1("company_name") & "<br> ·›‰:("& RS1("telephone")& ")"%>&nbsp;</TD>
					<TD ><%=RS1("customer_name")%>&nbsp;</TD>
					<TD ><%=RS1("order_title")%>&nbsp;</TD>
					<TD ><%=RS1("order_kind")%></TD>
					<TD style="<%=tmpStyle%>"><%=RS1("marhale")%></TD>
					<TD ><%=RS1("salesperson")%>&nbsp;</TD>
					<TD ><%=InvoiceStatus%>&nbsp;</TD>
					<td><%if isnull(RS1("price")) then response.write "----" else response.write Separate(RS1("price")) end if %></td>
				</TR>
				<TR bgcolor="#FFFFFF">
					<TD colspan="12" style="height:10px"></TD>
				</TR>
<%					RS1.moveNext
				Loop

				if tmpCounter >= resultsCount then
%>					<TR bgcolor="#CCCCCC">
						<TD align="center" colspan="10" style="padding:15px;font-size:12pt;color:red;cursor:hand;" onclick="document.all.resultsCount.focus();"><B> ⁄œ«œ ‰ «ÌÃ Ã” ÃÊ »Ì‘ «“ <%=resultsCount%> —ﬂÊ—œ «” .</B></TD>
					</TR>
<%				else
%>					<TR bgcolor="#ccccFF">
						<TD colspan="10"> ⁄œ«œ ‰ «ÌÃ Ã” ÃÊ: <%=tmpCounter%></TD>
					</TR>
<%				end if
%>
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
<%			End If
		End If
	End If
' Trace Quote End.
'-----------------------------

'-----------------------------
' Quote Input
'-----------------------------
elseif Request.QueryString("act")="quoteInpCustSearch" then
	if isnumeric(request("CustomerNameSearchBox")) then
		conn.close
		response.redirect "?act=getQuote&selectedCustomer=" & request("CustomerNameSearchBox")
	elseif request("CustomerNameSearchBox") <> "" then
		SA_TitleOrName=request("CustomerNameSearchBox")
		mySQL="SELECT * FROM Accounts WHERE (REPLACE(AccountTitle, ' ', '') LIKE REPLACE(N'%"& sqlSafe(SA_TitleOrName) & "%', ' ', '') ) ORDER BY AccountTitle"
		Set RS1 = conn.Execute(mySQL)

		if (RS1.eof) then 
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='../CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
		end if 

		SA_TitleOrName=request("CustomerNameSearchBox")
		SA_Action="return true;"
		SA_SearchAgainURL="Inquiry.asp"
		SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<br>
		<FORM METHOD=POST ACTION="?act=getType">
			<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	end if
elseif Request.QueryString("act")="getType" then 
	customerID=request("selectedCustomer")
	mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)

	if (RS1.eof) then 
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='..//CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
	end if 
%>
		<br>
		<div dir='rtl'>
			<B>ê«„ ”Ê„ : ê—› ‰ ‰Ê⁄ «” ⁄·«„</B>
		</div>
		<form method="post" action="?act=getQuote">
			<input name="selectedCustomer" type="hidden" value="<%=customerID%>">
			<SELECT NAME="OrderType" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px' tabIndex="7">
			<OPTION value="-1" style='color:red;'>«‰ Œ«» ﬂ‰Ìœ</option>
<%
			Set RS2 = conn.Execute("SELECT [User] as ID, DefaultOrderType FROM UserDefaults WHERE ([User] = "& session("ID") & ") OR (UserDefaults.[User] = 0) ORDER BY ABS(UserDefaults.[User]) DESC")
			defaultOrderType=RS2("DefaultOrderType")
			RS2.close
			Set RS2 = Nothing

			set RS_TYPE=Conn.Execute ("SELECT ID, Name FROM OrderTraceTypes WHERE (IsActive=1) ORDER BY ID")
			Do while not RS_TYPE.eof	
%>
				<OPTION value="<%=RS_TYPE("ID")%>" <%if RS_TYPE("ID")=defaultOrderType then response.write "selected"%>><%=RS_TYPE("Name")%></option>
<%
			RS_TYPE.moveNext
			loop
			RS_TYPE.close
			set RS_TYPE = nothing
%>		
			</SELECT>
			<input type="submit" value="«œ«„Â">
		</form>
<%	
elseif Request.QueryString("act")="getQuote" then
	customerID=request("selectedCustomer")
	mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)

	if (RS1.eof) then 
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='..//CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
	end if 

	AccountNo=RS1("ID")
	AccountTitle=RS1("AccountTitle")
	companyName=RS1("CompanyName")
	customerName=RS1("Dear1")& " " & RS1("FirstName1")& " " & RS1("LastName1")
	Tel=RS1("Tel1")
	'response.write request("orderType")
	set rs=Conn.Execute("select * from OrderTraceTypes where id=" & request("orderType"))
	if rs.eof then 
		conn.close
		response.redirect "?errmsg=" & server.urlEncode("‰Ê⁄ ”›«—‘ —«  ⁄ÌÌ‰ ﬂ‰Ìœ")
	end if
	orderTypeID=rs("id")
	orderTypeName=rs("name")
	set orderProp = server.createobject("MSXML2.DomDocument")
	hasProperty=false
	if rs("property")<>"" then 
		orderProp.loadXML(rs("property"))
		hasProperty=true
	end if
	rs.close
	set rs=nothing
	creationDate=shamsiToday()
	creationTime=time
	creationTime=Hour(creationTime)&":"&Minute(creationTime)
	if instr(creationTime,":")<3 then creationTime="0" & creationTime
	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)
%>
	<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
	<script type="text/javascript" src="calcOrder.js"></script>
	<script type="text/javascript">
		function checkValidation() {
			if ($('input[name="CustomerName"]').val().replace(/^\s*|\s*$/g,'')==''){
				alert("‰«„ „‘ —Ì —« Ê«—œ ﬂ‰Ìœ");
				$('input[name="CustomerName"]').focus();
				//$("input#Submit").prop("disabled",true);
				return false;
			} else if ($('input[name="SalesPerson"]').val().replace(/^\s*|\s*$/g,'')==''){
				alert("«” ⁄·«„ êÌ—‰œÂ —« Ê«—œ ﬂ‰Ìœ");
				$('input[name="SalesPerson"]').focus();
				//$("input#Submit").prop("disabled",true);
				return false;
			} else if ($('input[name="ReturnDate"]').val().replace(/^\s*|\s*$/g,'')==''){
				alert("„Ê⁄œ «⁄ »«— —« Ê«—œ ﬂ‰Ìœ");
				$('input[name="ReturnDate"]').focus();
				//$("input#Submit").prop("disabled",true);
				return false;
			} else if ($('input[name="ReturnTime"]').val().replace(/^\s*|\s*$/g,'')==''){
				alert("“„«‰ (”«⁄ ) «⁄ »«— —« Ê«—œ ﬂ‰Ìœ");
				$('input[name="ReturnTime"]').focus();
				//$("input#Submit").prop("disabled",true);
				return false;
			} else if ($('input[name="OrderTitle"]').val().replace(/^\s*|\s*$/g,'')==''){
				alert("⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì· —« Ê«—œ ﬂ‰Ìœ");
				$('input[name="OrderTitle"]').focus();
				//$("input#Submit").prop("disabled",true);
				return false;
			} else {
				$("input#Submit").prop("disabled",false);
				return true;
			} 
		}
	</script>
		

	<br>
	<div dir='rtl'><B>ê«„ çÂ«—„ : ê—› ‰ «” ⁄·«„</B>
	</div>
	<br>
<!-- ê—› ‰ «” ⁄·«„ -->
	<hr>
	<FORM METHOD=POST ACTION="?act=submitQuote" onSubmit="return checkValidation();">
		<TABLE cellspacing="0" cellpadding="2" dir="RTL" width="700" align="center" style="border: 2px solid #555599;">
			<TR bgcolor="#555599">
				<TD align="left"><FONT COLOR="YELLOW">Õ”«»:</FONT></TD>
				<TD align="right" colspan="3" height="25px">
					<FONT COLOR="YELLOW"><%=customerID & " - "& AccountTitle%></FONT>
					<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>">
				</TD>
				<td align="left"><font color="yellow">‰Ê⁄ «” ⁄·«„:</font></td>
				<td>
					<font color="red"><b><%=orderTypeName%></b></font>
					<input type="hidden" name="orderType" value="<%=orderTypeID%>">
				</td>
			</TR>
			<TR bgcolor="#555599">
				<TD align="left"><FONT COLOR="YELLOW">‘„«—Â «” ⁄·«„:</FONT></TD>
				<TD align="right">
					<!-- quote -->
					<INPUT disabled TYPE="text" NAME="quote" maxlength="6" size="8" tabIndex="1" dir="LTR" value="######">
				</TD>
				<TD align="left"><FONT COLOR="YELLOW"> «—ÌŒ:</FONT></TD>
				<TD>
					<TABLE border="0">
						<TR>
							<TD dir="LTR">
								<INPUT disabled TYPE="text" maxlength="10" size="10" value="<%=CreationDate%>">
								<INPUT TYPE="hidden" NAME="OrderDate" value="<%=CreationDate%>">
							</TD>
							<TD dir="RTL"><FONT COLOR="YELLOW"><%=weekdayname(weekday(date))%></FONT></TD>
						</TR>
					</TABLE>
				</TD>
				<TD align="left"><FONT COLOR="YELLOW">”«⁄ :</FONT></TD>
				<TD align="right">
					<INPUT disabled TYPE="text" maxlength="5" size="3" dir="LTR" value="<%=creationTime%>">
					<INPUT TYPE="hidden" NAME="OrderTime" value="<%=creationTime%>"></TD>
			</TR>
			<TR bgcolor="#CCCCCC">
				<TD align="left">‰«„ ‘—ﬂ :</TD>
				<TD align="right">
					<!-- CompanyName -->
					<INPUT TYPE="text" NAME="CompanyName" maxlength="50" size="25" tabIndex="2" value="<%=companyName%>"></TD>
				<TD align="left">„Ê⁄œ «⁄ »«—:</TD>
				<TD>
					<TABLE border="0">
						<TR>
							<TD dir="LTR"><INPUT TYPE="text" NAME="ReturnDate" onblur="acceptDate(this)" maxlength="10" size="10" tabIndex="5" value="<%=shamsiDate(dateAdd("d",10,date()))%>"></TD>
							<TD dir="RTL">(?‘‰»Â)</TD>
						</TR>
					</TABLE>
				</TD>
				<TD align="left">”«⁄  «⁄ »«—:</TD>
				<TD align="right"><INPUT TYPE="text" NAME="ReturnTime" value="<%=creationTime%>" maxlength="5" size="3" dir="LTR" tabIndex="6"  onKeyPress="return maskTime(this);" ></TD>
			</TR>
			<TR bgcolor="#CCCCCC">
				<TD align="left">‰«„ „‘ —Ì:</TD>
				<TD align="right">
					<!-- CustomerName -->
					<INPUT TYPE="text" NAME="CustomerName" maxlength="50" size="25" tabIndex="3" value="<%=customerName%>">
				</TD>
				<TD align="left">“„«‰  Ê·Ìœ:</TD>
				<TD>
					<input type="text" name="productionDuration" value="0" size="2">
					<span>—Ê“</span>
				</TD>
				<TD align="left">«” ⁄·«„ êÌ—‰œÂ:</TD>
				<TD><INPUT Type="Text" readonly NAME="SalesPerson" value="<%=CSRName%>" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px' tabIndex="888">
				</TD>
			</TR>
			<TR bgcolor="#CCCCCC">
				<TD align="left"> ·›‰:</TD>
				<TD align="right">
					<!-- Telephone -->
					<INPUT TYPE="text" NAME="Telephone" maxlength="50" size="25" tabIndex="4" value="<%=Tel%>"></TD>
				<TD align="left">⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì·:</TD>
				<TD align="right" colspan="3"><INPUT TYPE="text" NAME="OrderTitle" maxlength="255" size="50" tabIndex="9" style="width:100%"></TD>
			</TR>
			<TR bgcolor="#CCCCCC">
				<td align="left">”«Ì“:</td>
				<td>
					<input type="text" name="paperSize" tabindex="11">
				</td>
				<TD align="left" > Ê÷ÌÕ«  »Ì‘ —:</TD>
				<TD align="right"colspan="3"><TEXTAREA NAME="Notes" tabIndex="10" style="width:100%"></TEXTAREA></TD>
			</TR>
			<TR bgcolor="#CCCCCC">
				<td align="left"> Ì—«é:</td>
				<td>
					<input type="text" name="qtty" tabindex="12">
				</td>
				<td align="left">ﬁÌ„  ﬂ·:</td>
				<td colspan="3">
					<input type="text" name="totalPrice" id='totalPrice' style="background-color:#FED;border-width:0;" <%if hasProperty then response.write "readonly='readonly'"%>>
				</td>
			</TR>
			
			<tr bgcolor="#CCCCCC">
				<td colspan="6">
<%
if hasProperty then 
sub fetchKeys(key)
	oldGroup="---first---"
	oldLabel="---first---"
	thisRow="<div class='myRow'>"
	thisRow = thisRow & "<input type='hidden' value='0' id='" & Replace(key,"/","-") & "-maxID'>"
	thisRow = thisRow & "<div class='exteraArea' id='" & Replace(key,"/","-") & "-0'>"
' 	thisRow = thisRow & "<img title='Õ–› «Ì‰ Œÿ' src='/images/cancelled.gif' onclick='$(""#" & Replace(key,"/","-") & "-0"").remove();'>"
' 	thisRow = thisRow & "<img title='Õ–› «Ì‰ Œÿ' src='/images/cancelled.gif' onclick='$(this).parent().remove();'>" 
	For Each myKey In orderProp.SelectNodes(key)
	  thisType = myKey.GetAttribute("type")
	  thisName = myKey.GetAttribute("name")
	  thisLabel= myKey.GetAttribute("label")
	  thisGroup= myKey.GetAttribute("group")
	  'response.write thisName & ": " & thisLabel &"(" &thisType& ")" & "<br>"
	  if (oldGroup<>thisGroup and oldGroup <> "---first---") then thisRow = thisRow &  "</div>"
	  if oldGroup<>thisGroup then 
	  	thisRow = thisRow & "<div class='mySection' groupName='" & thisGroup & "'>"
	  	if myKey.GetAttribute("disable")="1" then 
			thisRow = thisRow &  "<input type='checkbox' value='0' name='" & thisGroup & "-disBtn' onclick='disGroup(this);"
			if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " calc_" & myKey.GetAttribute("group") & "(this);"
			thisRow = thisRow & "'>"
			disText=" disabled='disabled' "
		else
			disText=""
		end if
	  end if
	  if oldLabel<>thisLabel then thisRow = thisRow &  "<label class='myLabel'>" & thisLabel & "</label>"
		
	  select case thisType
	  	case "option"
	  		thisRow = thisRow &  "<select " & disText & " style='margin:0;padding:0;' name='" & thisName & "'"
	  		if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onchange='calc_" & myKey.GetAttribute("group") & "(this);' "
	  		thisRow = thisRow & ">"
	  		for each myOption in myKey.getElementsByTagName("option")
	  			thisRow = thisRow &  "<option value='" & myOption.text & "'"
	  			if myOption.GetAttribute("price")<>"" then 
	  				thisRow = thisRow & " price='" & myOption.GetAttribute("price") & "' "
	  			end if 
	  			thisRow = thisRow &">" & myOption.GetAttribute("label") & "</option>"
		  	next
		  	thisRow = thisRow &  "</select>"
		case "option-other"
			thisRow = thisRow &  "<select " & disText & " style='margin:0;padding:0;' name='" & thisName & "' onchange='checkOther(this);"
			if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " calc_" & myKey.GetAttribute("group") & "(this);"
			thisRow = thisRow & "'>"
	  		for each myOption in myKey.getElementsByTagName("option")
	  			thisRow = thisRow &  "<option value='" & myOption.text & "'"
	  			if myOption.GetAttribute("price")<>"" then 
	  				thisRow = thisRow & " price='" & myOption.GetAttribute("price") & "' "
	  			end if
	  			thisRow = thisRow & ">" & myOption.GetAttribute("label") & "</option>"
		  	next
		  	thisRow = thisRow &  "<option value='-1'>”«Ì—</option></select>"	
		  	thisRow = thisRow & "<input type='text' name='" &thisName & "-addValue' onblur='addOther(this);'>"	  	
		case "text"
			thisRow = thisRow &  "<input " & disText & " type='text' class='myInput' size='" & myKey.text & "' name='" & thisName & "' "
			if myKey.GetAttribute("readonly")="yes" then thisRow =thisRow & " readonly='readonly' "
			if myKey.GetAttribute("default")<>"" then thisRow = thisRow & "value='" & myKey.GetAttribute("default") & "'"
			if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onblur='calc_" & myKey.GetAttribute("group") & "(this);' "
			thisRow = thisRow & ">"
		case "textarea"
			thisRow = thisRow &  "<textarea name='" & thisName & "' style='width:600px;' cols='" & myKey.text & "'></textarea>"
		case "check"
			thisRow = thisRow & "<input type='checkbox' value='on-0' name='" & thisName & "' "
			if myKey.text="checked" then thisRow = thisRow & "checked='checked'"
			if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onclick='calc_" & myKey.GetAttribute("group") & "(this);' "
			if IsNumeric(myKey.GetAttribute("price")) then thisRow = thisRow & " price='" & myKey.GetAttribute("price") & "' "
			thisRow = thisRow & ">"
		case "radio":
				thisRow = thisRow & "<input " & disText & " type='radio' value='" & myKey.text & "' name='" & thisName & "'" 
				if myKey.GetAttribute("default")="yes" then thisRow = thisRow & " checked='checked' "
				if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onchange='calc_" & myKey.GetAttribute("group") & "(this);' "
				thisRow = thisRow & ">"
	  end select
	  if myKey.GetAttribute("force")="yes" then thisRow = thisRow &  "<span style='color:red;margin:0 0 0 2px;padding:0;'>*</span>"
	  oldGroup=thisGroup
	  oldLabel=thisLabel
	  if myKey.GetAttribute("br")="yes" then thisRow = thisRow & "<br>"
	Next
	thisRow = thisRow & "</div></div><div id='extreArea" &Replace(key,"/","-")& "'></div>"
	response.write thisRow 'prependTo 
	response.write "<div class='btn'><img title='«÷«›Â' src='/images/Plus-32.png' width='20px' onclick='cloneRow(""" & key & """);'><img title='Õ–› ¬Œ—Ì‰ Œÿ' src='/images/cancelled.gif' onclick='removeRow(""" & key & """);'></div></div>"
end sub
	oldTmp="---first---"
	for each tmp in orderProp.selectNodes("//key")
		if oldTmp<>tmp.parentNode.nodeName then 
			oldTmp=tmp.parentNode.nodeName
			call fetchKeys("/keys/" & oldTmp & "/key")
		end if
	next
' 	call fetchKeys("/keys/printing/key")
' 	call fetchKeys("keys/binding/key")
' 	call fetchKeys("keys/service/key")
' 	call fetchKeys("keys/delivery/key")
end if
%>			
				</td>
			</tr>
			<TR bgcolor="#CCCCCC">
				<TD colspan="6">
					<table align="center" width="50%" border="0">
						<tr>
							<TD><INPUT TYPE="submit" id='Submit' Name="Submit" Value=" «ÌÌœ" style="width:100px;" tabIndex="14"></TD>
							<TD><INPUT TYPE="hidden" NAME="Price" maxlength="10" size="9" dir="LTR" tabIndex="13" value="‰«„‘Œ’">&nbsp;</TD>
							<TD align="left"><INPUT TYPE="button" Name="Cancel" Value="«‰’—«›" style="width:100px;" onClick="window.location='Inquiry.asp';" tabIndex="15"></TD>
						</tr>
					</table>
				</TD>
			</TR>
		</TABLE>
	</FORM>
<%
	set orderProp=nothing
elseif Request.QueryString("act")="submitQuote" then
	//-------------------------------------------------------------------------------------------------------------------
	//-------------------------------------------------------------------------------------------------------------------
	//-------------------------------------------------------------------------------------------------------------------
	set rs=Conn.Execute("select * from OrderTraceTypes where id=" & request("orderType"))
	if rs.eof then 
		conn.close
		response.redirect "?errmsg=" & server.urlEncode("‰Ê⁄ ”›«—‘ —«  ⁄ÌÌ‰ ﬂ‰Ìœ")
	end if
	orderTypeID=rs("id")
	orderTypeName=rs("name")
	set orderProp = server.createobject("MSXML2.DomDocument")
	if rs("property")<>"" then 
		orderProp.loadXML(rs("property"))
		myXML = fetchKeyValues()
	end if
	rs.close
	set rs=nothing
	
function fetchKeyValues()
	key="---first---"
	thisRow="<?xml version=""1.0""?><keys>"
	for each tmp in orderProp.selectNodes("//key")
		if key<>tmp.parentNode.nodeName then 
			key=tmp.parentNode.nodeName
			thisRow = thisRow & "<" & key & ">"
			hasValue=0
			For Each myKey In orderProp.SelectNodes("/keys/" & key & "/key")
				thisName = myKey.GetAttribute("name")
				thisGroup= myKey.GetAttribute("group")
				id=0
'					response.write oldName& "<br>"
				if thisName<>oldName then 
					for each value in request.form(thisName)
						if value <> "" then 
							thisRow = thisRow & "<key name=""" & thisName & """ id=""" 
							select case myKey.GetAttribute("type") 
								case "check"
									thisRow = thisRow & mid(value,4)
								case else
									if request.form(thisGroup & "-disBtn")<>"" then 
										thisRow = thisRow & trim(split(request.form(thisGroup & "-disBtn"),",")(id))
									else
										thisRow = thisRow & id 
									end if
							end select
							thisRow = thisRow & """>" & value & "</key>"
							hasValue=hasValue +1
						end if
						id=id+1
					next
				end if
				oldName = thisName
			Next
			
			if hasValue>0 then 
				thisRow = thisRow & "</" & key & ">"
			else
				thisRow = Replace(thisRow,"<" & key & ">","")
			end if
		end if
	Next
	thisRow = thisRow & "</keys>"
	fetchKeyValues = thisRow 
	'response.write thisRow
	'response.end
end function

	
	
	CreationDate=shamsiToday()
	CustomerID=request.form("CustomerID") 
	if CustomerID="" OR not isNumeric(CustomerID) then
		conn.close
		response.redirect "?act=getQuote&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("Œÿ« œ— ‰«„ „‘ —Ì<br>«ÿ·«⁄«  Ê«—œ ‰‘œÖ<BR>")
	end if

	orderType=request.form("OrderType")
	if orderType="" OR not isNumeric(OrderType) then
		conn.close
		response.redirect "?act=getQuote&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ «” ⁄·«„<br>«ÿ·«⁄«  Ê«—œ ‰‘œÖ<BR>")
	end if

	set RS1=Conn.Execute ("SELECT Name FROM OrderTraceTypes WHERE (ID='"& orderType & "') AND (IsActive=1)")
	if RS1.eof then
		conn.close
		response.redirect "?act=getQuote&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ «” ⁄·«„<br>«ÿ·«⁄«  Ê«—œ ‰‘œÖ<BR>")
	else
		orderTypeName=RS1("Name")
	end if
	RS1.close

	defaultVaziat = "œ— Ã—Ì«‰" ' 1
	defaultMarhale = "À»  ‘œÂ" ' 1
	

	mySQL="SET NOCOUNT ON;INSERT INTO Quotes (CreatedDate, CreatedBy, Customer, Closed, order_date, order_time, return_date, return_time, company_name, customer_name, telephone, order_title, order_kind, Type, vazyat, marhale, salesperson, status, step, LastUpdatedDate, LastUpdatedTime, LastUpdatedBy, Notes,property,productionDuration, Qtty, price, PaperSize) VALUES ('"&_
	CreationDate & "', '"& session("ID") & "', '"& CustomerID & "', '0', N'"& sqlSafe(request.form("OrderDate")) & "', N'"& sqlSafe(request.form("OrderTime")) & "', N'"& sqlSafe(request.form("ReturnDate")) & "', N'"& sqlSafe(request.form("ReturnTime")) & "', N'"& sqlSafe(request.form("CompanyName")) & "', N'"& sqlSafe(request.form("CustomerName")) & "', N'"& sqlSafe(request.form("Telephone")) & "', N'"& sqlSafe(request.form("OrderTitle")) & "', N'"& orderTypeName & "', '"& orderType & "', N'" & defaultVaziat & "', N'" & defaultMarhale & "', N'"& sqlSafe(request.form("SalesPerson")) & "', 1, 1, N'"& CreationDate & "',  N'"& currentTime10() & "', '"& session("ID") & "', N'"& sqlSafe(request.form("Notes")) & "',N'" & myXML & "', " & sqlSafe(request.form("productionDuration")) & "," & sqlSafe(request.form("qtty")) & ",'" & sqlSafe(request.form("totalPrice")) & "', N'" & sqlSafe(request.form("PaperSize")) & "'); SELECT SCOPE_IDENTITY() AS NewQuote;SET NOCOUNT OFF"

	set RS1 = Conn.execute(mySQL)'.NextRecordSet
	QuoteID = RS1 ("NewQuote")
	RS1.close

	conn.close
	response.redirect "?act=show&quote=" & QuoteID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  «” ⁄·«„ À»  ‘œ")

' Quote Input End
'-----------------------------

'-----------------------------
' Quote Edit
'-----------------------------
elseif Request.QueryString("act")="editQuote" then

	quote = request("quote")

	if quote <> "" then
		if isNumeric(quote) then
			quote=clng(quote)
		else
			quote=""
		end if
	end if
	
	If quote="" Then
		response.write "<br><br>"
		call showAlert("‘„«—Â «” ⁄·«„ «‘ »«Â «”  .", CONST_MSG_ERROR )
		response.end
	End If 

	mySQL = "SELECT Quotes.*, Accounts.ID AS AccID, Accounts.AccountTitle FROM Quotes INNER JOIN Accounts ON Quotes.Customer = Accounts.ID WHERE (Quotes.ID ='"& quote & "')"
	set RS2=Conn.Execute (mySQL)

	if RS2.eof then 
		response.write "<br><br>"
		call showAlert("‘„«—Â «” ⁄·«„ «‘ »«Â «”  .", CONST_MSG_ERROR )
		response.end 
	else
		CustomerID=RS2("AccID")
	end if

	if RS2("salesperson")<>session("csrName") then 
		response.write "<br>"
		call showAlert("«” ⁄·«„ êÌ—‰œÂ «Ì‰ «” ⁄·«„ ‘„« ‰Ì” Ìœ..<BR>ﬁ»· «“ ÊÌ—«Ì‘ »« «” ⁄·«„ êÌ—‰œÂ Â„«Â‰ê ﬂ‰Ìœ.", CONST_MSG_ALERT ) 
	end if

	relatedApprovedInvoiceID = 0

'	mySQL="SELECT Invoices.Issued, Invoices.Approved, Invoices.ApprovedBy, Invoices.ID FROM InvoiceOrderRelations INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (InvoiceOrderRelations.[Order] = '" & quote & "') AND (Invoices.Voided = 0)"
'	set RS3=Conn.Execute (mySQL)
'	if not(RS3.eof) then 
'		FoundInvoice=RS3("ID")
'		if RS3("Issued") then 
'			Conn.Close
'			response.redirect "../order/TraceOrder.asp?act=show&order=" & quote & "&errMsg=" & Server.URLEncode("Ìﬂ ›«ﬂ Ê—  ’«œ— ‘œÂ „— »ÿ »« »« «Ì‰  «” ⁄·«„ ‘„«—Â ÅÌœ« ﬂ—œÂ «Ì„.(‘„«—Â ›«ﬂ Ê—: <A HREF='../AR/AccountReport.asp?act=showInvoice&invoice="& FoundInvoice & "' target='_blank'>" & FoundInvoice & "</a>)<BR> ·–«”  ﬂÂ «„ﬂ«‰Ì »—«Ì  €ÌÌ— œ— «Ì‰ «” ⁄·«„ ÊÃÊœ ‰œ«—œ.")
'		elseif RS3("Approved") then 
'			tmpDesc="<B> «ÌÌœ ‘œÂ</B>"
'			tmpColor="Yellow"
'			relatedApprovedInvoiceID = RS3("id")
'			relatedApprovedInvoiceBy = RS3("ApprovedBy")
'			response.write "<br>"
'			call showAlert("Ìﬂ ›«ﬂ Ê— <b> «ÌÌœ ‘œÂ</b> »« ‘„«—Â <A HREF='../AR/AccountReport.asp?act=showInvoice&invoice="& FoundInvoice &"' target='_blank'>" & FoundInvoice & "</A> »—«Ì «Ì‰ «” ⁄·«„ ÊÃÊœ œ«—œ <BR>ﬂÂ »«  €ÌÌ— «” ⁄·«„  Ê”ÿ ‘„« «“ Õ«·   «ÌÌœ Œ«—Ã ŒÊ«Âœ ‘œ." , CONST_MSG_ALERT )  
'		else
'			tmpDesc=""
'			tmpColor="#FFFFBB"
'			response.write "<br>"
'			call showAlert(" ÊÃÂ!<br>«” ⁄·«„Ì ﬂÂ ‘„« Ã” ÃÊ ﬂ—œÌœ ﬁ»·« œ— Ìﬂ ›«ﬂ Ê— <B> «ÌÌœ ‰‘œÂ</B> ÊÃÊœ œ«—œ<br><A HREF='../AR/AccountReport.asp?act=showInvoice&invoice=" & FoundInvoice & "'>‰„«Ì‘ ›«ﬂ Ê— „—»ÊÿÂ (" & FoundInvoice & ")</A>", CONST_MSG_INFORM ) 
'		end if
'	end if 
	%>
	<style>
		Table { font-size: 8pt;}
		INPUT { font-family:Tahoma,arial; font-size: 8pt;}
		SELECT { font-family:Tahoma,arial; font-size: 8pt;}
	</style>
	<font face="tahoma">

	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<FORM METHOD=POST ACTION="?act=submitEditQuote" onSubmit="return checkValidation();">
		
	<INPUT TYPE="hidden" name="relatedApprovedInvoiceBy" value="<%=relatedApprovedInvoiceBy%>">
	<INPUT TYPE="hidden" name="relatedApprovedInvoiceID" value="<%=relatedApprovedInvoiceID%>">
	<br>
	<div align="center">ÊÌ—«Ì‘ «” ⁄·«„</div>
	<br>
	<TABLE cellspacing="0" cellpadding="2" dir="RTL" width="700" align="center" style="border: 2px solid #555599;">
	<TR bgcolor="#555599">
		<TD align="left"><FONT COLOR="YELLOW">Õ”«»:</FONT></TD>
		<TD align="right" colspan=5 height="25px">
			<span id="customer" style="color:yellow;"><%' after any changes in this span "./Customers.asp" must be revised%>
				<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><span><%=customerID & " - "& RS2("AccountTitle")%></span>.
			</span>
			<INPUT class="GenButton" TYPE="button" value=" €ÌÌ—" onClick="selectCustomer();">
		</TD>
	</TR>
	<TR bgcolor="#555599">
		<TD align="left"><FONT COLOR="YELLOW">‘„«—Â «” ⁄·«„:</FONT></TD>
		<TD align="right">
			<!-- quote -->
			<INPUT TYPE="text" disabled maxlength="6" size="5" tabIndex="1" dir="LTR" value="<%=RS2("ID")%>">
			<INPUT TYPE="hidden" NAME="quote" value="<%=RS2("ID")%>">
		</TD>
		<TD align="left"><FONT COLOR="YELLOW"> «—ÌŒ:</FONT></TD>
		<TD><TABLE border="0">
			<TR>
				<TD dir="LTR">
				<INPUT disabled TYPE="text" maxlength="10" size="8"  value="<%=RS2("order_date")%>">
				<INPUT TYPE="hidden" NAME="OrderDate" value="<%=RS2("order_date")%>">
				</TD>
				<TD dir="RTL"><FONT COLOR="YELLOW"><%="ø‘‰»Â"%></FONT></TD>
			</TR>
			</TABLE></TD>
		<TD align="left"><FONT COLOR="YELLOW">”«⁄ :</FONT></TD>
		<TD align="right">
		<INPUT disabled TYPE="text" maxlength="5" size="3" dir="LTR" value="<%=RS2("order_time")%>">
		<INPUT TYPE="hidden" NAME="OrderTime" value="<%=RS2("order_time")%>"></TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD align="left">‰«„ ‘—ﬂ :</TD>
		<TD align="right">
			<!-- CompanyName -->
			<INPUT TYPE="text" NAME="CompanyName" maxlength="50" size="25" tabIndex="2"  value="<%=RS2("company_name")%>"></TD>
		<TD align="left">„Ê⁄œ «⁄ »«—:</TD>
		<TD><TABLE border="0">
			<TR>
				<TD dir="LTR"><INPUT TYPE="text" NAME="ReturnDate"  onblur="acceptDate(this)" maxlength="10" size="8" tabIndex="5" onKeyPress="return maskDate(this);" value="<%=RS2("return_date")%>"></TD>
				<TD dir="RTL">(?‘‰»Â)</TD>
			</TR>
			</TABLE></TD>
		<TD align="left">”«⁄  «⁄ »«—:</TD>
		<TD align="right"><INPUT TYPE="text" NAME="ReturnTime" maxlength="6" size="3" tabIndex="6" dir="LTR" onKeyPress="return maskTime(this);" value="<%=RS2("return_time")%>"></TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD align="left">‰«„ „‘ —Ì:</TD>
		<TD align="right">
			<!-- CustomerName -->
			<INPUT TYPE="text" NAME="CustomerName" maxlength="50" size="25" tabIndex="3" value="<%=RS2("customer_name")%>"></TD>
		<TD align="left">‰Ê⁄ «” ⁄·«„:</TD>
		<TD>
			<SELECT NAME="OrderType" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px' tabIndex="7" onchange="alert(' ÊÃÂ œ«‘ Â »«‘Ìœ ﬂÂ  €ÌÌ— ‰Ê⁄ ”›«—‘ »«⁄À «Œ ·«· œ— Ã“∆Ì«  ”›«—‘ ŒÊ«Âœ ‘œ. ¬Ì« „”Ê·Ì  ¬‰—« „ÌùÅ–Ì—Ìœø');">
	<%
			thisOrderType=RS2("Type")
			set RS_TYPE=Conn.Execute ("SELECT ID, Name FROM OrderTraceTypes WHERE (IsActive=1) ORDER BY ID")
			Do while not RS_TYPE.eof	
	%>
				<OPTION value="<%=RS_TYPE("ID")%>" <%if thisOrderType=RS_TYPE("ID") then response.write "selected" %> ><%=RS_TYPE("Name")%></option>
	<%
				RS_TYPE.moveNext
			loop
			RS_TYPE.close
			set RS_TYPE = nothing
	%>		
			</SELECT></TD>
		<TD align="left">«” ⁄·«„ êÌ—‰œÂ:</TD>
		<TD><INPUT NAME="SalesPerson" Type="TEXT"value="<%=RS2("salesperson")%>" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 100px' tabIndex="88" readonly>
		</TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD align="left"> ·›‰:</TD>
		<TD align="right">
			<!-- Telephone -->
			<INPUT TYPE="text" NAME="Telephone" maxlength="50" size="25" tabIndex="4" value="<%=RS2("telephone")%>"></TD>
		<TD align="left">⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì·:</TD>
		<TD align="right" colspan="3"><INPUT TYPE="text" NAME="OrderTitle" maxlength="255" size="50" tabIndex="9" value="<%=RS2("order_title")%>"></TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<td align="left"> Ì—«é:</td>
		<td>
			<input type="text" name="qtty" value="<%=rs2("qtty")%>" tabindex="9">
		</td>
		<td align="left">”«Ì“:</td>
		<td>
			<input type="text" value="<%=rs2("paperSize")%>" name="paperSize" tabIndex="14">
		</td>
		<td align="left">ﬁÌ„  ﬂ·:</td>
		<td >
			<input type="text" value="<%=rs2("price")%>" name="totalPrice" id='totalPrice' style="background-color:#FED;border-width:0;" <%if rs2("property")<>"" then response.write " readonly='readonly'"%>>
		</td>
	</TR>
	<TR bgcolor="#CCCCCC">
		<td> Œ„Ì‰ “„«‰  Ê·Ìœ:</td>
		<td>
			<input name="productionDuration" type="text" size="2" value="<%=RS2("productionDuration")%>">
			<span>—Ê“</span><br>
			<small>(’›— »Â „⁄‰Ì „⁄·Ê„ ‰»Êœ‰ “„«‰ «” )</small>
		</td>
		<TD align="left"> Ê÷ÌÕ«  »Ì‘ —:</TD>
		<TD align="right" colspan="3"><TEXTAREA NAME="Notes" tabIndex="10" style="width:100%"><%=RS2("Notes")%></TEXTAREA></TD>
		
	</TR>
	<tr bgcolor="#CCCCCC">
		<TD align="left">„—Õ·Â:</TD>
		<TD>
			<SELECT NAME="Marhale" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px' tabIndex="13" >
		<%
		set RS_STEP=Conn.Execute ("SELECT * FROM QuoteSteps WHERE (IsActive=1)")
		Do while not RS_STEP.eof	
		%>
			<OPTION value="<%=RS_STEP("ID")%>" <%if RS2("step")=RS_STEP("ID") then response.write "selected" %> ><%=RS_STEP("name")%></option>
			<%
			RS_STEP.moveNext
		loop
		RS_STEP.close
		set RS_STEP = nothing
		%>
			</SELECT>
		</TD>
		<td colspan="4"></td>
	</tr>
	<tr bgcolor="#CCCCCC">
		<td colspan="6">
	<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
	<script type="text/javascript" src="calcOrder.js"></script>
	<div>Ã“∆Ì«  «” ⁄·«„</div>
	<div>
		<br>»—«Ì „Õ«”»Â ŒÊœﬂ«— ﬁÌ„ ùÂ« »«Ìœ —ÊÌ Å«—«„ —Â«Ì Â— ”ÿ— »—ÊÌœ Ê ¬‰—«  €ÌÌ— œÂÌœ Ê Ì« ¬‰ ”·Ê· —«  —ﬂ ‰„«ÌÌœ.
		<br>Â„ç‰Ì‰ „Ìù Ê«‰Ìœ ﬁÌ„  ŒÊœ —« Ê«—œ ‰„«ÌÌœ «„« œ— ’Ê— Ì ﬂÂ œÊ»«—Â ÌﬂÌ «“ Å«—«„ —Â«Ì „—»ÊÿÂ —Ê  €ÌÌ— »œÌœ œÊ»«—Â ﬁÌ„  „Õ«”»Â ŒÊ«Âœ ‘œ.
		<br>„Õ«”»Â »—ŒÌ «“ ﬁÌ„ ùÂ« «“ —ÊÌ Å«—«„ —Â«Ì ”«Ì— ŒÿÊÿ ŒÊ«Âœ »Êœ. „À·«  Ì—«é œ— „«‘Ì‰°  Ì—«éÌ ŒÊ«Âœ »Êœ ﬂÂ œ— ”·›Ê‰ Ê ÌÊ ÊÌ Ê Ê—‰Ì „Õ«”»Â „Ìù‘Êœ. Ê œ— ’Ê— Ì ﬂÂ  Ì—«é „«‘Ì‰ —Ê ⁄Ê÷ ﬂ‰Ì„ »«Ìœ ”·Ê·ùÂ«Ì „—»Êÿ »Â «Ì‰ ¬Ì „ùÂ« —Ê  —ﬂ ﬂ‰Ì„  « œÊ»«—Â ﬁÌ„ ù‘Ê‰ „Õ«”»Â »‘Â
	</div>
<%
	set rs=Conn.Execute("select * from OrderTraceTypes where id="&rs2("type"))
	set typeProp = server.createobject("MSXML2.DomDocument")
	set orderProp = server.createobject("MSXML2.DomDocument")
	
	if rs2("property")<>"" then orderProp.loadXML(rs2("property"))
	if rs("property")<>"" then typeProp.loadXML(rs("property"))
	set rs=nothing

' 	for each item in typeProp.SelectNodes("//key")
' 		response.write item.parentNode.nodeName & ": " & item.xml & "<br>"
' 	next
' 	response.end
sub showKeyEdit(key)
	'key="/keys/printing/key"
	oldGroup="---first---"
	oldLabel="---first---"
	thisRow = "<div class='myRow'>"
	maxID=0
	oldID=0
	for each mykey in orderProp.SelectNodes(key)
		id=myKey.GetAttribute("id")
		if maxID<id then maxID=id
	next
	thisRow = thisRow & "<input type='hidden' value='" & maxID & "' id='" & Replace(key,"/","-") & "-maxID'>"
'	response.write "<div class='myRow'><div class='exteraArea' id='" & Replace(key,"/","-") & "-0'>"
	for id = 0 to maxID
		thisRow = thisRow & "<div class='exteraArea' id='" & Replace(key,"/","-") & "-" & id & "'>"
' 		thisRow = thisRow & "<img title='Õ–› «Ì‰ Œÿ' src='/images/cancelled.gif' onclick='$(this).parent().remove();'>"
		For Each myKey In typeProp.SelectNodes(key)
		  thisType = myKey.GetAttribute("type")
		  thisName = myKey.GetAttribute("name")
		  thisLabel= myKey.GetAttribute("label")
		  thisGroup= myKey.GetAttribute("group")
		  hasValue=false
		  	set thisValue= orderProp.SelectNodes(key & "[@id='" & id & "' and @name='" & thisName & "']")
		  	if thisValue.length>0 then hasValue=true
' 		  response.write hasValue & "<br>"
		  if (oldGroup<>thisGroup and oldID=id and oldGroup <> "---first---") then thisRow = thisRow &  "</div>"
		  if oldGroup<>thisGroup or oldID<>id then 
			thisRow = thisRow & "<div class='mySection' groupName='" & thisGroup & "'>"
			if myKey.GetAttribute("disable")="1" then 
				thisRow = thisRow & "<input type='checkbox' value='" & id & "' name='" & thisGroup & "-disBtn' onclick='disGroup(this);"
				if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " calc_" & myKey.GetAttribute("group") & "(this);"
				thisRow = thisRow & "'"
				if hasValue then 
					thisRow = thisRow & " checked='checked'"
					disText=""
				else
					disText=" disabled='disabled' "
				end if	
				thisRow = thisRow & ">"
			else
				disText=""
			end if
		  end if
		  if oldLabel<>thisLabel then thisRow = thisRow &  "<label class='myLabel'>" & thisLabel & "</label>"
			
		  select case thisType
		  	case "option"
		  		thisRow = thisRow &  "<select " & disText & " class='myInput' name='" & thisName & "'"
		  		if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onchange='calc_" & myKey.GetAttribute("group") & "(this);' "
		  		thisRow = thisRow & ">"
		  		for each myOption in myKey.getElementsByTagName("option")
		  			thisRow = thisRow & "<option value='" & myOption.text & "'"
		  			if hasValue then 
		  				if thisValue(0).text=myOption.text then thisRow = thisRow & " selected='selected' "
		  			end if
		  			if myOption.GetAttribute("price")<>"" then 
		  				thisRow = thisRow & " price='" & myOption.GetAttribute("price") & "' "
		  			end if
		  			thisRow = thisRow & ">" & myOption.GetAttribute("label") & "</option>"
			  	next
			  	thisRow = thisRow & "</select>"
			case "option-other"
				thisRow = thisRow & "<select " & disText & " class='myInput' name='" & thisName & "' onchange='checkOther(this);"
				if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " calc_" & myKey.GetAttribute("group") & "(this);"
				thisRow = thisRow & "'"
				thisRow = thisRow & ">"
		  		for each myOption in myKey.getElementsByTagName("option")
		  			thisRow = thisRow & "<option value='" & myOption.text & "'"
		  			if hasValue then 
		  				if thisValue(0).text=myOption.text then thisRow = thisRow & " selected='selected' "
		  			end if
		  			if myOption.GetAttribute("price")<>"" then 
		  				thisRow = thisRow & " price='" & myOption.GetAttribute("price") & "' "
		  			end if
		  			thisRow = thisRow & ">" & myOption.GetAttribute("label") & "</option>"
			  	next
			  	if hasValue then 
				  	if left(thisValue(0).text,6)="other:" then 
				  		thisRow = thisRow & "<option value='" & thisValue(0).text & "' selected='selected'>" & mid(thisValue(0).text,7) & "</option></select>"
				  	else
				  		thisRow = thisRow & "<option value='-1'>”«Ì—</option></select>"
				  	end if
				else
					thisRow = thisRow & "<option value='-1'>”«Ì—</option></select>"
				end if
			  	thisRow = thisRow & "<input type='text' name='" & thisName & "-addValue' onblur='addOther(this);'>"
			case "text"
				thisRow = thisRow & "<input " & disText & " type='text' class='myInput' size='" & myKey.text & "' name='" & thisName & "' "
				if myKey.GetAttribute("readonly")="yes" then thisRow =thisRow & " readonly='readonly' "
				if hasValue then 
					thisRow = thisRow & "value='" & thisValue(0).text & "'"
				else
					if myKey.GetAttribute("default")<>"" then thisRow = thisRow & "value='" & myKey.GetAttribute("default") & "'"
				end if
				if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onblur='calc_" & myKey.GetAttribute("group") & "(this);' "
				thisRow = thisRow & ">"
			case "textarea"
				thisRow = thisRow & "<textarea name='" & thisName & "' style='width:600px;' cols='" & myKey.text & "'>"
				if hasValue then thisRow = thisRow & thisValue(0).text
				thisRow = thisRow & "</textarea>"
			case "check"
				thisRow = thisRow & "<input type='checkbox' value='on-" & id & "' name='" & thisName & "' "
				if hasValue then 
					if left(thisValue(0).text,2)="on" then thisRow = thisRow & "checked='checked'"
				else
					if myKey.text="checked" then thisRow = thisRow & "checked='checked'"
				end if
				if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onclick='calc_" & myKey.GetAttribute("group") & "(this);' "
				if IsNumeric(myKey.GetAttribute("price")) then 
		  				thisRow = thisRow & " price='" & myKey.GetAttribute("price") & "' "
		  			end if
				thisRow = thisRow & ">"
			case "radio":
			'response.write hasValue
				thisRow = thisRow & "<input " & disText & " type='radio' value='" & myKey.text & "' name='" & thisName & "'" 
				if hasValue then
					if myKey.text = thisValue(0).text then thisRow = thisRow & " checked='checked' "
					
				else
					if myKey.GetAttribute("default")="yes" then thisRow = thisRow & " checked='checked' "
				end if
				if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onchange='calc_" & myKey.GetAttribute("group") & "(this);' "
				thisRow = thisRow & ">"
		  end select
		  if myKey.GetAttribute("force")="yes" then thisRow = thisRow &  "<span style='color:red;margin:0 0 0 2px;padding:0;'>*</span>"
		  oldGroup=thisGroup
		  oldLabel=thisLabel
		  oldID=id
		  if myKey.GetAttribute("br")="yes" then thisRow = thisRow & "<br>"
		Next
		thisRow = thisRow & "</div></div>"
	next
	thisRow = thisRow & "<div id='extreArea" &Replace(key,"/","-")& "'></div>"
	response.write thisRow 'prependTo 
	response.write "<div class='btn'><img title='«÷«›Â' src='/images/Plus-32.png' width='20px' onclick='cloneRow(""" & key & """);'><img title='Õ–› ¬Œ—Ì‰ Œÿ' src='/images/cancelled.gif' onclick='removeRow(""" & key & """);'></div></div>"
end sub
	
	oldTmp="---first---"
	for each tmp in typeProp.selectNodes("//key")
		if oldTmp<>tmp.parentNode.nodeName then 
			oldTmp=tmp.parentNode.nodeName
			call showKeyEdit("/keys/" & oldTmp & "/key")
		end if
	next
	
	' call showKeyEdit("/keys/printing/key")
' 	call showKeyEdit("keys/binding/key")
' 	call showKeyEdit("keys/service/key")
' 	call showKeyEdit("keys/delivery/key")
	
%>
		
		</td>
	</tr>
	<TR bgcolor="#CCCCCC">
		<TD colspan="6">

		<TABLE align="center" width="50%" border="0">
		<TR>
			<TD><INPUT TYPE="submit" Name="Submit" tabIndex="16" Value=" «ÌÌœ" onFocus="noNextField=true;" onBlur="noNextField=false;" style="width:100px;"></TD>
			<TD>&nbsp;</TD>
			<TD align="left"><INPUT TYPE="button" Name="Cancel" tabIndex="17" Value="«‰’—«›" style="width:100px;" onClick="window.location='Inquiry.asp?act=show&quote=<%=quote%>';"></TD>
		</TR>
		</TABLE>

		</TD>
	</TR>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	document.all.CompanyName.focus();

	var noNextField = false;

	function documentKeyDown() {
		var theKey = event.keyCode;
		var o = window.event.srcElement;
		if (theKey == 9) { 
			if (noNextField)			// Don't Move
				if(event.shiftKey){
					noNextField=false;
					event.keyCode=9
				}
				else
				return false; 
			else {						// send focus to next box

				if (false){
				}
				else if(o.name == 'quote'){
	// quote 
					if (o.value.length<5){
						return false;
					}
				}
				else if(o.name == 'ReturnTime'){
	// Return Time
					var soutput = o.value;
					var len = o.value.length; 
					if (len==0){
						return true;
					}
					else if (len==3){
						soutput = soutput + '00'
						o.value = soutput;
					}
					else if (len==4){
						var tempString =soutput.charAt(len-1);
						soutput = soutput.substring(0,len-1) + '0' + tempString
						o.value = soutput;
					}
					else if(len!=5)
						return false;
				}
				event.keyCode=9
				return true;
			}
		}
	}

	document.onkeydown = documentKeyDown;

	function checkValidation(){
		return true;
	}

	function selectCustomer(){
		document.all.tmpDlgArg.value="#"
		document.all.tmpDlgTxt.value="‰«„ Õ”«»Ì —« ﬂÂ „Ì ŒÊ«ÂÌœ Ã” ÃÊ ﬂ‰Ìœ Ê«—œ ﬂ‰Ìœ:"
		window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		if (document.all.tmpDlgTxt.value !="") {
			window.showModalDialog('../AR/dialog_SelectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value), document.all.tmpDlgArg, 'dialogWidth:780px; dialogHeight:500px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
			if (document.all.tmpDlgArg.value!="#"){
				Arguments=document.all.tmpDlgArg.value.split("#")
				theSpan=document.getElementById("customer");
				theSpan.getElementsByTagName("input")[0].value=Arguments[0];
				theSpan.getElementsByTagName("span")[0].innerText=Arguments[0]+" - "+Arguments[1];
			}
		}
	}

	//-->
	</SCRIPT>
	</TABLE>
	<br>
	</font>
<%
elseif Request.QueryString("act")="submitEditQuote" then

	quote = request("quote")

	if quote <> "" then
		if isNumeric(quote) then
			quote=clng(quote)
		else
			quote=""
		end if
	end if
	
	If quote="" Then
		response.write "<br><br>"
		call showAlert("‘„«—Â «” ⁄·«„ «‘ »«Â «”  .", CONST_MSG_ERROR )
		response.end
	End If 

	orderType=request.form("OrderType")
	if orderType="" OR not isNumeric(OrderType) then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ «” ⁄·«„<br>«ÿ·«⁄«  Ê«—œ ‰‘œÖ<BR>")
	else
		set RS1=Conn.Execute ("SELECT Name FROM OrderTraceTypes WHERE (ID='"& orderType & "') AND (IsActive=1)")
		if RS1.eof then
			conn.close
			response.redirect "?errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ «” ⁄·«„<br>«ÿ·«⁄«  Ê«—œ ‰‘œÖ<BR>")
		else
			orderTypeName=RS1("Name")
		end if
		RS1.close
	end if

	'set RS1=Conn.Execute ("SELECT Name FROM OrderTraceStatus WHERE (IsActive=1) and ID = " & request.form("Vazyat"))
	'statusName = RS1("name")
	set rs=Conn.Execute("select * from OrderTraceTypes where id=" & request("orderType"))
	if rs.eof then 
		conn.close
		response.redirect "?errmsg=" & server.urlEncode("‰Ê⁄ ”›«—‘ —«  ⁄ÌÌ‰ ﬂ‰Ìœ")
	end if
	orderTypeID=rs("id")
	orderTypeName=rs("name")
	set orderProp = server.createobject("MSXML2.DomDocument")
	if rs("property")<>"" then 
		orderProp.loadXML(rs("property"))
		myXML = fetchKeyValues()
	end if
	rs.close
	set rs=nothing
	
function fetchKeyValues()
	key="---first---"
	thisRow="<?xml version=""1.0""?><keys>"
	for each tmp in orderProp.selectNodes("//key")
		if key<>tmp.parentNode.nodeName then 
			key=tmp.parentNode.nodeName
			thisRow = thisRow & "<" & key & ">"
			hasValue=0
			For Each myKey In orderProp.SelectNodes("/keys/" & key & "/key")
				thisName = myKey.GetAttribute("name")
				thisGroup= myKey.GetAttribute("group")
				id=0
'					response.write oldName& "<br>"
				if thisName<>oldName then 
					for each value in request.form(thisName)
						if value <> "" then 
							thisRow = thisRow & "<key name=""" & thisName & """ id=""" 
							select case myKey.GetAttribute("type") 
								case "check"
									thisRow = thisRow & mid(value,4)
								case else
									if request.form(thisGroup & "-disBtn")<>"" then 
										thisRow = thisRow & trim(split(request.form(thisGroup & "-disBtn"),",")(id))
									else
										thisRow = thisRow & id 
									end if
							end select
							thisRow = thisRow & """>" & value & "</key>"
							hasValue=hasValue +1
						end if
						id=id+1
					next
				end if
				oldName = thisName
			Next
			
			if hasValue>0 then 
				thisRow = thisRow & "</" & key & ">"
			else
				thisRow = Replace(thisRow,"<" & key & ">","")
			end if
		end if
	Next
	thisRow = thisRow & "</keys>"
	fetchKeyValues = thisRow 
	'response.write thisRow
	'response.end
end function
	

	set RS1=Conn.Execute ("SELECT Name FROM QuoteSteps WHERE (IsActive=1) and ID = " & request.form("Marhale"))
	stepName = RS1("name")

	set RS3=Conn.Execute ("SELECT * FROM Quotes WHERE (ID = "& quote & ")")

	OrderDate =		sqlSafe(request.form("OrderDate"))
	OrderTime =		sqlSafe(request.form("OrderTime"))
	ReturnDate =	sqlSafe(request.form("ReturnDate"))
	ReturnTime =	sqlSafe(request.form("ReturnTime"))
	CompanyName	=	sqlSafe(request.form("CompanyName"))
	CustomerName =	sqlSafe(request.form("CustomerName"))
	Telephone =		sqlSafe(request.form("Telephone"))
	OrderTitle =	sqlSafe(request.form("OrderTitle"))
	'Vazyat =		sqlSafe(request.form("Vazyat"))
	Marhale =		sqlSafe(request.form("Marhale"))
	SalesPerson =	sqlSafe(request.form("SalesPerson"))
'	Qtty =			sqlSafe(request.form("Qtty"))
	Size =			sqlSafe(request.form("Size"))
'	SimplexDuplex =	sqlSafe(request.form("SimplexDuplex"))
	Price =			sqlSafe(request.form("totalPrice"))
	Notes =			sqlSafe(request.form("Notes"))
	productionDuration = sqlSafe(request.form("productionDuration"))

	mySql="UPDATE Quotes SET Customer='"& request.form("CustomerID") & "', order_date= N'"& OrderDate & "', order_time= N'"& OrderTime & "', return_date= N'"& ReturnDate & "', return_time= N'"& ReturnTime & "', company_name= N'"& CompanyName & "', customer_name= N'"& CustomerName & "', telephone= N'"& Telephone & "', order_title= N'"& OrderTitle & "', order_kind= N'"& orderTypeName & "', Type= '"& orderType & "', step= "& Marhale & ",  marhale= N'"& stepName & "', salesperson= N'"& SalesPerson & "' , LastUpdatedDate=N'"& shamsitoday() & "' , LastUpdatedTime=N'"& currentTime10() & "', LastUpdatedBy=N'"& session("ID")& "', Notes= N'"& Notes & "', property=N'" & myXML & "', paperSize= N'"& Size & "', Price= N'"& Price & "', productionDuration = " & productionDuration & "  WHERE (ID = N'"& quote & "')"	
	', qtty= N'"& Qtty & "', paperSize= N'"& Size & "', SimplexDuplex= N'"& SimplexDuplex & "', Price= N'"& Price & "'
	conn.Execute mySql
	response.write quote &" UPDATED<br>"
	response.write "<A HREF='orderEdit.asp'>Back</A>"

'	if request.form("Marhale")="10" then
'		call InformCSRorderIsReady(quote , RS3("CreatedBy"))
'	end if

'	if not request.form("relatedApprovedInvoiceID")="0" then
'		call UnApproveInvoice(request.form("relatedApprovedInvoiceID") , request.form("relatedApprovedInvoiceBy"))
'	end if

	conn.close
	response.redirect "?act=show&quote=" & quote & "&msg=" & Server.URLEncode("«ÿ·«⁄«  «” ⁄·«„ »Â —Ê“ ‘œ")
	
elseif Request.QueryString("act")="convertToOrder" then
	If  isnumeric(request("quote")) then
		quote=request("quote")
		mySQL="SELECT Quotes.Customer FROM Quotes WHERE (Quotes.ID='"& quote & "')"
		set RS1=conn.execute (mySQL)
		If RS1.EOF then
			response.write "<BR><BR><BR><BR><CENTER>‘„«—Â «” ⁄·«„ „⁄ »— ‰Ì” </CENTER>"
			conn.close
			response.end
		End If
		CustomerID = RS1("Customer")
		RS1.close
		set rs = Conn.Execute("select * from accounts where id=" & CustomerID)
		if rs.eof then 
			conn.close
			response.write "<BR><BR><BR><BR><CENTER>Œÿ«Ì ⁄ÃÌ» ‘„«—Â „‘ —Ìù „⁄ »— ‰Ì” !</CENTER>"
			response.end
		end if
		if (cdbl(rs("arBalance"))+cdbl(rs("creditLimit")) < 0) then 
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("»œÂÌ «Ì‰ Õ”«» «“ „Ì“«‰ «⁄ »«— ¬‰ »Ì‘ — ‘œÂ°<br> ·ÿ›« »« ”—Å—”  ›—Ê‘ Â„«Â‰ê ﬂ‰Ìœ.<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>‰„«Ì‘ Õ”«»</a>")
		end if
		rs.close
		set rs = nothing
		CreationDate = shamsiToday()
		OrderTime = Left(currentTime10(),5)

		' create order for the same customer
		mySQL="INSERT INTO Orders (CreatedDate, CreatedBy, Customer) VALUES ('"& CreationDate & "', '"& session("ID") & "', '"& CustomerID & "');SELECT @@Identity AS NewOrder"
		set RS1 = Conn.execute(mySQL).NextRecordSet
		OrderID = RS1("NewOrder")
		RS1.close

		' create orders_trace row and copy info from quote
		mySQL=	"INSERT INTO orders_trace (radif_sefareshat, order_date, order_time, company_name, customer_name, telephone, order_title, order_kind, Type, vazyat, marhale, salesperson, status, step, LastUpdatedDate, LastUpdatedTime, LastUpdatedBy, property) "&_
				"SELECT '" & OrderID & "', N'"& CreationDate & "', N'"& OrderTime & "', company_name, customer_name, telephone, order_title, order_kind, Type, N'œ— Ã—Ì«‰', N'œ— ’› ‘—Ê⁄', salesperson, 1, 1, N'"& CreationDate & "',  N'"& currentTime10() & "', '"& session("ID") & "',property FROM Quotes WHERE ID='" & quote & "'; "
		conn.Execute(mySQL)

		' relate invoices to the new order
		mySQL=	"INSERT INTO InvoiceOrderRelations ([Invoice], [Order]) SELECT [Invoice], '" & OrderID & "' AS [Order] FROM InvoiceQuoteRelations WHERE Quote = '" & quote & "';"
		conn.Execute(mySQL)

		' remove relation to previous quote
		mySQL=	"DELETE FROM InvoiceQuoteRelations WHERE Quote = '" & quote & "';"
		conn.Execute(mySQL)

		'close the quote
		mySQL=	"UPDATE Quotes SET Closed = 1, step = 5, marhale = ' »œÌ· »Â ”›«—‘ ‘œÂ' WHERE [ID] = '" & quote & "';"
		conn.Execute(mySQL)

		' keeping the relation
		mySQL=	"INSERT INTO QuoteOrderRelations ([QuoteId], [OrderId]) VALUES ('" & quote & "' , '" & OrderID & "');"
		conn.Execute(mySQL)

		conn.close
		response.redirect "../order/TraceOrder.asp?act=show&order=" & OrderID & "&msg=" & Server.URLEncode("«” ⁄·«„ »Â ”›«—‘ ﬂÅÌ ‘œ")
	End If

	response.write "<br><br>"
	call showAlert("‘„«—Â «” ⁄·«„ «‘ »«Â «”  .", CONST_MSG_ERROR )
	response.end
end if

Conn.Close
%>

<!--#include file="tah.asp" -->
