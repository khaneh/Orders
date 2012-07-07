<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Home (0)
PageTitle= " íÛÇãåÇ "
SubmenuItem=2
if not Auth(0 , 2) then NotAllowdToViewThisPage()

sendTo = session("id")
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%

function sqlSafe (s)
  st=s
  st=replace(St,"'","`")
  st=replace(St,chr(34),"`")
  sqlsafe=st
end function
if request("act")="show" then 
	if IsNumeric(request("id")) then 
		set rs=Conn.execute ("select Messages.*,userTo.realName as toName, userFrom.realName as fromName,message_types.name as typeName from Messages inner join users as userTo on Messages.msgTo=userTo.id inner join users as userFrom on Messages.msgFrom=userFrom.id inner join message_types on Messages.type=message_types.id where Messages.id=" & request("id"))
		if rs.eof then 
			ErrorMsg	= "ÎØÇí ãÍÇá!"
			response.redirect returnURL & "errMsg=" & Server.URLEncode(ErrorMsg)
		end if
		select case Trim(rs("relatedTable"))
			case "accounts":
				response.redirect "../CRM/AccountInfo.asp?act=show&selectedCustomer="&rs("relatedID")
			case "invoices":
				response.redirect "../AR/AccountReport.asp?act=showInvoice&invoice="&rs("relatedID")
			case "orders":
				response.redirect "../order/TraceOrder.asp?act=show&order="&rs("relatedID")
			case "quotes" :
				response.redirect "../order/Inquiry.asp?act=show&quote="&rs("relatedID")
			case else 
%>
<br><br><br>

<div align="right">
	<LI>ÊÇÑíÎ ÇÑÓÇá: <span dir=ltr><%=RS("MsgDate")%></span>
	<LI>ÓÇÚÊ: <%=RS("MsgTime")%>
	<li>ÇÒ: <%=rs("fromName")%></li>
	<li>Èå: <%=rs("toName")%></li>
	<li><%=rs("typeName")%></li>
	<LI>íÇã: <%=RS("MsgBody")%>
</div>
<%					
		end select
	else
		ErrorMsg	= "ÎØÇ ÏÑ æÑæÏí."
		response.redirect returnURL & "errMsg=" & Server.URLEncode(ErrorMsg)
	end if
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ Send New Message
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="ÇÑÓÇá íÇã" then
	ON ERROR RESUME NEXT

		ErrorFound		= false
		ErrorMsg		= ""

		MsgTo			= cint(request.form("MsgTo"))
		msgTitle		= sqlSafe(request.form("msgTitle"))
		msgBody			= sqlSafe(left(request.form("msgBody"),1999))
		RelatedTable	= sqlSafe(request.form("RelatedTable"))
		relatedID		= sqlSafe(request.form("relatedID"))
		if relatedID<>"" then relatedID = clng(relatedID)
		replyTo			= sqlSafe(request.form("replyTo"))
		IsReply			= sqlSafe(request.form("IsReply"))
		urgent			= sqlSafe(request.form("urgent"))
		MsgFrom			= session("ID")
		MsgDate			= shamsiToday()
		MsgTime			= currentTime10()

		returnURL		= request.form("retURL")

		if instr(returnURL,"?") > 0 then
			returnURL = returnURL & "&"
		else
			returnURL = returnURL & "?"
		end if
		
		if IsNumeric(request.form("msgType")) then
			msgType = CInt(request.form("msgType"))
		else
			msgType=0
		end if
		
		if MsgTo <> -100 then
			set RS=Conn.Execute ("SELECT RealName FROM Users where ID="& MsgTo) 
			if RS.eof then
				ErrorFound	= true
				ErrorMsg	= "íÑäÏå íÇã æÌæÏ äÏÇÑÏ"
			else
				ReceiverName= RS("RealName")
			end if
			RS.close
		end if

		if Err.Number<>0 then
			ErrorFound	= true
			ErrorMsg	= "ÎØÇ ÏÑ æÑæÏí."
		end if
	ON ERROR GOTO 0

	if ErrorFound then
		conn.close
		response.redirect returnURL & "errMsg=" & Server.URLEncode(ErrorMsg)
	end if

	if MsgTo=-100 AND Auth(0 , 7) then 'ÇÑÓÇá íÇã Èå åãå
		msg = "íÇã ÈÑÇí "
		writeAnd = ""
		set RSV=Conn.Execute ("SELECT * FROM Users WHERE (ID <> 0) AND (Display = 1) ORDER BY RealName") 
		Do while not RSV.eof
			MsgTo=RSV("ID")
			MySQL = "INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent, type) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", '"& relatedID & "', '"& RelatedTable & "', "& urgent & ", " & msgType & ")"
			conn.Execute MySQL 
			msg = msg & writeAnd & RSV("RealName")
			writeAnd = " æ "
			RSV.moveNext
		Loop
		RSV.close
		msg = msg & " ÇÑÓÇá ÔÏ."
	else
		MySQL = "INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent, type) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", '"& relatedID & "', '"& RelatedTable & "', "& urgent & ", " & msgType & ")"
		conn.Execute MySQL 
		if MsgTo=0 then
			msg = "íÇÏÏÇÔÊ ËÈÊ ÔÏ."
		else
			msg = "íÇã ÈÑÇí " & ReceiverName &  " ÇÑÓÇá ÔÏ."
		end if
	end if
	response.redirect returnURL & "msg=" & Server.URLEncode(msg)
end if

'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
%>

<center><BR><BR><BR>
<TABLE>
<TR>
<%
replyTo = "0"
IsReply = "0"
RelatedTable = "NaN"
RelatedID = "0"
msgBody = ""
MsgTitle = ""

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------ Reply Form
'-----------------------------------------------------------------------------------------------------
if request("act") = "reply" then
	response.write "<TD valign=top> "
	replyTo = request("id")
	IsReply = "1"
	MySQL = "SELECT * FROM Messages WHERE (id = "& replyTo & ") and (MsgTo="& session("id") & ")"
	set RSM = conn.Execute (MySQL)
	if RSM.EOF then
		response.write "<BR><BR><BR><BR><CENTER>ÇÕæáÇ äíä íÇãí äÏÇÑíÏ ßå ÈÊæÇäíÏ Èå Âä ÌæÇÈ åã ÈÏåíÏ</CENTER>"
		response.end
	end if
	sendTo			= RSM("MsgFrom")
	RelatedTable	= trim(RSM("RelatedTable"))
	RelatedID		= trim(RSM("RelatedID"))
	'response.write RelatedTable
	%>
	<H3>ÇÓÎ Èå íÇã</H3>
	<TABLE style="border: solid 1pt black; width:220">
	<TR>
		<TD>
			<LI>ÊÇÑíÎ ÇÑÓÇá: <span dir=ltr><%=RSM("MsgDate")%></span>
			<LI>ÓÇÚÊ: <%=RSM("MsgTime")%>
			<LI>íÇã: <%=RSM("MsgBody")%>
		</TD>
	</TR>
	</TABLE>
	</td>
	<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------- FWD Form
'-----------------------------------------------------------------------------------------------------
elseif request("act") = "forward" then
	fwMsg = request("id")
	MySQL = "SELECT Messages.*, Users.RealName FROM Messages INNER JOIN Users ON Messages.MsgFrom = Users.ID WHERE (Messages.id = "& fwMsg & ") AND (Messages.MsgTo = "& session("id") & ")"
	set RSM = conn.Execute (MySQL)
	if RSM.EOF then
		response.write "<BR><BR><BR><BR><CENTER>ÇÕæáÇ äíä íÇãí äÏÇÑíÏ ßå ÈÊæÇäíÏ Âä  ÑÇ ÇÑÌÇÚ åã ÈÏåíÏ</CENTER>"
		response.end
	end if
	sendTo = RSM("MsgFrom")
	RelatedTable	= trim(RSM("RelatedTable"))
	RelatedID		= trim(RSM("RelatedID"))
	msgBody= "[íÇã ÇÑÌÇÚí ÇÒ "& RSM("RealName")& "] " & RSM("MsgBody")
	MsgTitle = "FWD"
	%>
	<TR>
		<TD colspan=2 align=center><H3>ÇÑÌÇÚ íÇã</H3></TD>
	</TR>
	<%
elseif  request("act") ="" then
	%>
	<TR>
		<TD colspan=2 align=center><H3>ÇÑÓÇá íÇã íÇ ÒÇÑÔ</H3></TD>
	</TR>
	<%
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ New Message Form
'-----------------------------------------------------------------------------------------------------
if request("act")<>"show" then

	if request("act")<>"reply" and request("act")<>"forward"  then 
		RelatedID=request("RelatedID")
		RelatedTable=LCase(request("RelatedTable"))
	end if
	Select Case RelatedTable
	Case "orders":
		RelatedTableName = "ÓİÇÑÔ ÔãÇÑå"
	Case "accounts":
		RelatedTableName = "ÔãÇÑå ÍÓÇÈ"
	Case "invoices":
		RelatedTableName = "İÇßÊæÑ ÔãÇÑå"
	Case "quotes":
		RelatedTableName = "ÇÓÊÚáÇã ÔãÇÑå"
	Case else:
		RelatedTableName = RelatedTable
	End Select
	if request("sendTo") <> ""  then sendTo = request("sendTo")
	'response.write sendTo
	'response.write RelatedTable
	%>
	<TD  valign=top>
	<FORM METHOD=POST ACTION="message.asp">
	<INPUT TYPE="hidden" name="replyTo" value="<%=replyTo%>">
	<INPUT TYPE="hidden" name="IsReply" value="<%=IsReply%>">
	<TABLE>
	<TR>
		<TD align=left>íÑäÏå:</TD>
		<TD align=right>
			<INPUT TYPE="hidden" NAME="retURL" value="<%=request("retURL")%>">
			<% if not (request("act") = "reply") then %>
			<select name="MsgTo" class=inputBut >
			<% set RSV=Conn.Execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
			Do while not RSV.eof
			%>
				<option value="<%=RSV("ID")%>" <%
					if cint(RSV("ID"))=cint(sendTo) then
						response.write " selected "
					end if
					%>><%=RSV("RealName")%></option>
			<%
			RSV.moveNext
			Loop
			RSV.close
			%>
	<%		if Auth(0 , 7) then
				' Has the Priviledge to SEND MESSAGE TO EVERYONE
	%>			<option disabled value="0">----------------------</option>
				<option value="-100">* åãå *</option>
	<%		end if%>
			</select> 
			<% else 
	
			if request("sendTo")<>"" then
			sendTo = request("sendTo")
			end if
	
			set RSV=Conn.Execute ("SELECT RealName FROM Users where ID = " & sendTo) 
			if RSV.EOF then
				response.redirect "message.asp"
			end if %>
			<INPUT TYPE="hidden"  NAME="MsgTo" value="<%=sendTo%>"><INPUT readonly TYPE="text" NAME="MsgTo21" value="<%=RSV("RealName")%>">
			<% end if %>
			<span dir=ltr><%=shamsiToday()%></span><BR>
		</TD>
	</TR>
	<TR>
		<TD align=left><!--ÚäæÇä--></TD>
		<TD align=right>
			<INPUT TYPE="hidden" NAME="msgTitle"  class=inputBut size=31 value="<%=MsgTitle%>">
		</TD>
	</TR>
	<TR>
		<TD align=left>íÇã</TD>
		<TD align=right>
			<TEXTAREA NAME="msgBody" ROWS="7"  class=inputBut COLS="32" maxlength=1999><%=msgBody%></TEXTAREA>
		</TD>
	</TR>
	<TR>
		<TD align=left>ãÑÈæØ ÇÓÊ Èå </TD>
		<TD align=right>
			<% if RelatedID = "" then %>
			<SELECT NAME="RelatedTable"  onchange="hideIT()" >
				<option <% if RelatedTable="NaN" then %> selected <% end if %>value="NaN">åíí</option>
				<option <% if RelatedTable="orders" then %> selected <% end if %>value="orders">ÓİÇÑÔ ÔãÇÑå </option>
				<option <% if RelatedTable="accounts" then %> selected <% end if %>value="accounts">ÔãÇÑå ÍÓÇÈ </option>
				<option <% if RelatedTable="invoices" then %> selected <% end if %>value="invoices">İÇßÊæÑ ÔãÇÑå </option>
	
				</SELECT><span name="relatedIDSpan"  id="relatedIDSpan">
				<INPUT TYPE="text" NAME="relatedID" size=9 value="<%=RelatedID%>"  onKeyPress="return maskNumber(this);" ></span>
			<% else %>
				<INPUT TYPE="hidden" NAME="RelatedTable" value="<%=RelatedTable%>"><INPUT TYPE="text" NAME="alak" value="<%=RelatedTableName%>"size=17 readonly> <INPUT TYPE="text" NAME="relatedID" size=9 value="<%=RelatedID%>"  readonly>
			<% end if %>
		</TD>
	</TR>
	<TR>
		<TD align=left>ÇæáæíÊ:</TD>
		<TD align=right>
			<span style="background-color:white"><INPUT TYPE="radio" NAME="urgent" value="0" checked>ÚÇÏí &nbsp;
			<span style="background-color:#FFDDDD"><INPUT TYPE="radio" NAME="urgent" value="1">æíå &nbsp;
			<span style="background-color:yellow"><INPUT TYPE="radio" NAME="urgent" value="2">Îíáí İæÑí &nbsp;
		</TD>
	</TR>
	<tr>
		<td align="left">äæÚ:</td>
		<td align="right">
			<select name="msgType">
			<%
			set rs= Conn.Execute("select * from message_types")
			if request("typeID")<>"" then typeID=request("typeID")
			while not rs.eof
			%>
				<option value="<%=rs("id")%>" <%if cint(typeID)=cint(rs("id")) then response.write(" selected ") %>><%=rs("name")%></option>
			<%	
				rs.moveNext
			wend
			%>
			</select>
		</td>
	</tr>
	<TR>
		<TD align=left></TD>
		<TD align=center><br><INPUT TYPE="submit" name="submit" value="ÇÑÓÇá íÇã"></TD>
	</TR>
	<TR>
		<TD align=left></TD>
		<TD align=right>
		</TD>
	</TR>
	</TABLE>
	</FORM>
<%
end if
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function hideIT()
{
if(''+document.all.RelatedTable.value!="NaN") 
	{
		document.all.relatedIDSpan.style.visibility= 'visible'
		document.all.relatedID.value = "0"
		document.all.relatedID.focus()
	}
	else
	{
		document.all.relatedIDSpan.style.visibility= 'hidden'
		document.all.relatedID.value = "0"
	}
}
hideIT();
//-->
</SCRIPT>

</TD>
</TR>
</TABLE>

<!--#include file="tah.asp" -->