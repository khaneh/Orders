<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Home (0)
PageTitle= "ŒÊ«‰œ‰ ÅÌ«„Â«"
SubmenuItem=1
if not Auth(0 , 1) then NotAllowdToViewThisPage()

captionFGColor = "#C6C6D7"
captionBGColor = "#F4F4FE"
MsgBodyColor = "#F4F4FE" '"#FFFFEE"
UrgentMsgBodyColor = "#FFDDDD"
VeryUrgentMsgBodyColor = "yellow"
OrderReadyMsgBodyColor = "#33FF99" '#00CCFF
VeryUrgentMsgScrlColor = "#FFFFCC"
MsgMainColor = "#FFFFFF"

activeTabColor="#336699"
disableTabColor="#CCCCCC"

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%

panel = request("panel")
if panel="" then 
	panel=1
else
	panel=cint(panel)
end if

if request("act")="changeStatus" then
	MsgIDsCount = request.form("MsgIDs").count
	if MsgIDsCount > 0 then
		MsgIDs = "(0"
		for i=1 to MsgIDsCount
			MsgIDs = MsgIDs & ", " & clng(request.form("MsgIDs")(i))
		next 
		MsgIDs = MsgIDs & ")"
		
		Select Case request.form("MsgAct")
			Case 1:		' »«Ìê«‰Ì ‘Êœ
				MySQL="UPDATE Messages SET IsSmall=1 WHERE ID IN " & MsgIDs
			Case 2:		' Å«ﬂ ‘Êœ
				MySQL="UPDATE Messages SET IsRead=1 WHERE ID IN " & MsgIDs
			Case 3:		' ÅÌ«„ ÃœÌœ «” 
				MySQL="UPDATE Messages SET IsRead=0, IsSmall=0 WHERE ID IN " & MsgIDs
			Case else:
				response.write "<br>"
				CALL showAlert ("Œÿ«! ⁄„· «‰ Œ«» ‘œÂ «„ﬂ«‰Å–Ì— ‰Ì” .",CONST_MSG_ERROR) 
				response.end
		End Select

		Conn.Execute MySQL

		'***---------- Re-Writing Message Status:
%>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			document.getElementById("MessagesStatusPanel").innerHTML = "<%=WriteMessagesStatus%>"
		//-->
		</SCRIPT>
<%
		'------ End of Re-Writing Message Status

	end if
end if

%>
<style>
	.MsgTable {font-family:tahoma; direction: RTL; background-color:gray; width:100%; border:none;}
	.MsgTable td {vertical-align:top;border-bottom:1px solid black;}
	.MsgTable a {text-decoration:none;color:#000088}
	.MsgTable a:hover {text-decoration:underline;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CusTD3 {background-color: #DDDDDD; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; background-color:#C3DBEB; border:4 solid <%=activeTabColor%>;}

	.CusTD1 {background-color: <%=activeTabColor%>; text-align: center; }
	.CusTD1 a {color:#FFFF00; font-size:9pt;}
	.CusTD2 {background-color: <%=disableTabColor%>; text-align: center; }
	.CusTD2 a {color:#888888; font-size:9pt;}
	.MsgBodyClass {width:500px;height:60px;overflow:auto;}
	.MsgButton {border:1px solid #AAAAAA !important;vertical-align:bottom; background-color: #CCCCCC}
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
function selectAll(src){
	totalMsgIDs=document.getElementsByName("MsgIDs").length
	checked=src.checked
	for (i=0;i<totalMsgIDs;i++)
		document.getElementsByName("MsgIDs")[i].checked=checked;
}

//-->
</SCRIPT>
<table cellspacing=0 cellpadding=0 width="100%" style="border:4 solid <%=AppFgColor%>;">
<tr><td>
	<TABLE cellspacing=0 cellpadding=0 width="100%">
	<TR height='15'>
		<TD></TD>
	</TR>

	<TR class='alak' height='25'>
	<TD width=15 >&nbsp;</TD>
<%
		if panel=1 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?panel=1'>ÅÌ«„ Â«Ì ÃœÌœ</A></TD>

		<TD width=5 >&nbsp;</TD>
<%
		if panel=2 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?panel=2'>ÅÌ«„ Â«Ì »«Ìê«‰Ì ‘œÂ</A></TD>
	<TD width=400 >&nbsp;</TD>
	<TD width=* align=left>&nbsp;</TD>
	</TR>
</td></tr>
</table>
<TaBlE class="CustTable4" cellspacing="2" cellspacing="0" >
	<Tr>
	<Td valign="top" align="center">
		<FORM METHOD=POST ACTION="default.asp?act=changeStatus&panel=<%=panel%>">
		<table class="MsgTable" cellspacing='1'>
<%
		selectTop=""
'		selectTop="TOP 20"
'		session("id") = 6

		if panel=1 then 'ÅÌ«„ Â«Ì ÃœÌœ
			isSmall=0
		elseif panel=2 then 'ÅÌ«„ Â«Ì »«Ìê«‰Ì ‘œÂ
			isSmall=1
		else 
			response.end
		end if
		MySQL = "SELECT " & selectTop & " Messages.*, message_types.name as messageTypeName, message_types.id as typeID, Messages_1.MsgBody AS OrigMsgBody, Messages_1.MsgDate AS OrigMsgDate, Messages_1.MsgTime AS OrigMsgTime, Users.RealName AS Sender FROM Messages INNER JOIN Users ON Messages.MsgFrom = Users.ID LEFT OUTER JOIN Messages Messages_1 ON Messages.replyTo = Messages_1.id inner join message_types on message_types.id= messages.type WHERE (Messages.MsgTo = '" & session("id") & "') AND (Messages.IsRead = 0) AND (Messages.IsSmall = " & IsSmall & ") ORDER BY Messages.id DESC"
		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3"></td>
				<td class="CusTD3" rowspan=2><INPUT TYPE="checkbox" NAME="SelectAll" onclick="selectAll(this);"></td>
				<td class="CusTD3" colspan=4 style="text-align:right;">
					<SELECT NAME="MsgAct" style="font-family:tahoma;font-size:9pt;width:200px;">
						<option value="1">»«Ìê«‰Ì ‘Êœ</option>
						<option value="2">Å«ﬂ ‘Êœ</option>
						<option value="3">ÅÌ«„ ÃœÌœ «” </option>
					</SELECT>
					<INPUT TYPE="submit" Value=" €ÌÌ— Ê÷⁄Ì "class="genButton">
				</td>
			</tr>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3">«“</td>
				<td class="CusTD3"> «—ÌŒ</td>
				<td class="CusTD3">ÅÌ«„</td>
				<td class="CusTD3" >„—»Êÿ »Â</td>
			</tr>
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
				if RS1("Urgent")=0 then 
					tmpColor=MsgMainColor
				elseif RS1("Urgent")=2 then 
					tmpColor=VeryUrgentMsgBodyColor 
				elseif RS1("Urgent")=3 then 
					tmpColor=OrderReadyMsgBodyColor
				else 
					tmpColor=UrgentMsgBodyColor
				end if

				RelatedTable = "NaN"

				if trim(RS1("RelatedTable")) = "orders" then 
					RelatedTable="”›«—‘ "
					RelatedLink = "../order/TraceOrder.asp?act=show&order=" & RS1("RelatedID")
				elseif trim(RS1("RelatedTable")) = "accounts" then 
					RelatedTable="Õ”«» "
					RelatedLink = "../CRM/AccountInfo.asp?act=show&SelectedCustomer="& RS1("RelatedID")
				elseif trim(RS1("RelatedTable")) = "invoices" then 
					RelatedTable="›«ﬂ Ê— "
					RelatedLink = "../AR/AccountReport.asp?act=showInvoice&invoice=" & RS1("RelatedID")
				end if
%>
				<TR bgcolor="<%=tmpColor%>" onclick="javascript:void(0);">
					<TD dir="LTR" align='right'><%=tmpCounter%>&nbsp;</TD>
					<TD><INPUT TYPE="checkbox" NAME="MsgIDs" Value="<%=RS1("id")%>"></TD>
					<TD width=60 height=100%>
						<table style="width:100%;height:100%">
							<tr height=20 >
								<td colspan=2><%=RS1("Sender")%></td>
							</tr>
							<tr height=* >
								<td colspan=2 style="border-bottom:none;">&nbsp;</td>
							</tr>
							<tr height=15 >
								<td class="MsgButton" >
									<A HREF="message.asp?act=reply&id=<%=RS1("id")%>&retURL=<%=Server.URLEncode("default.asp")%>&typeID=<%=rs1("typeID")%>">Å«”Œ</a>
								</td>
								<td class="MsgButton">
									<A HREF="message.asp?act=forward&id=<%=RS1("id")%>&retURL=<%=Server.URLEncode("default.asp")%>&typeID=<%=rs1("typeID")%>">«—Ã«⁄</a>
								</td>
							</tr>
						</table>
					</TD>
					<TD dir="LTR" align='right'>
						<div><%=RS1("MsgDate")%><br><%=RS1("MsgTime")%></div>
						<div style="font-size:6pt;margin:7px 0 0 0;color:gray;"><%=RS1("messageTypeName")%></div>
					</TD>
					<TD>
						<div class="MsgBodyClass"><%=replace(RS1("MsgBody"),chr(13),"<br>")%>&nbsp;</div>
						<%if RS1("isReply") <> 0 then%>
							<div class="MsgBodyClass" style="border-top:1px solid gray;">œ— ÃÊ«» ÅÌ«„ ﬁ»·Ì ‘„« (<%=RS1("OrigMsgTime")%> - <%=replace(RS1("OrigMsgDate"),"/",".")%>) »Â ‘—Õ “Ì—:<BR> <%=replace(RS1("OrigMsgBody"),chr(13),"<br>")%>&nbsp;</div>
						<%end if%>
					</TD>
					<TD width="50"><%if RelatedTable <> "NaN" then%>
						<A HREF="<%=RelatedLink%>" target="_blank"><%=RelatedTable%> <%=RS1("RelatedID")%></A>
						<%end if%> &nbsp;
					</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
%>
			<tr>
				<td colspan="9" class="CusTableHeader" style="text-align:right;"><A HREF="?userID=<%=userID%>&showAll=on&panel=1">«œ«„Â œ«—œ ...</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="9" class="CusTD3">ÂÌç</td>
			</tr>
<%
		end if
%>
		</table>
		</FORM>
	</Td>
	</Tr>
</TaBlE>
<!--#include file="tah.asp" -->