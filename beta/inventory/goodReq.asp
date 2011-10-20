<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Inventory (5)
PageTitle= "ËÈÊ ÏÑÎæÇÓÊ ßÇáÇ"
SubmenuItem=10
if not Auth(5 , "A") then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
catItem1 = request("catItem")
if catItem1="" then catItem1="-1"


'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------- Delete an Inventory Request from an order
'-----------------------------------------------------------------------------------------------------
if request("di")="y" then		
	myRequestID=request("i")
	set RSX=Conn.Execute ("SELECT * FROM InventoryItemRequests WHERE id = "& myRequestID )	
	if RSX("status")="new" then
	Conn.Execute ("update InventoryItemRequests SET status = 'del' where id = "& myRequestID )	
	end if
	response.redirect "goodReq.asp?radif=" & request("r")
end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Submit an Inventory Item request
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="ËÈÊ ÏÑÎæÇÓÊ ßÇáÇ" then
	order_ID =	sqlSafe(request.form("radif"))
	item =		sqlSafe(request.form("item"))
	comment =	sqlSafe(request.form("comment"))
	qtty =		sqlSafe(request.form("qtty"))
	CreatedBy = session("id")
	if 	not item = "" then

		if  comment="" then
			comment = "-"
		end if

		if order_ID="" or item="-1" or qtty="" or (not isNumeric(qtty)) or qtty="0" then
			response.write "<br><br>"
			call showAlert("ÎØÇ! åí ßÇáÇíí ÇäÊÎÇÈ äßÑÏå ÇíÏ <br><br><A HREF='goodReq.asp'>ÈÑÔÊ</A>",CONST_MSG_ERROR)
			response.end
		else
			order_ID =	clng(order_ID)
			item =		clng(item)
			qtty =		cdbl(qtty)
		end if

		set RS4 = conn.Execute ("SELECT * FROM InventoryItems where OldItemID=" & item)
		if (RS4.eof) then
			response.write "<br><br>"
			call showAlert("ÎØÇ! åí ßÇáÇíí ÇäÊÎÇÈ äßÑÏå ÇíÏ <br><br><A HREF='goodReq.asp'>ÈÑÔÊ</A>",CONST_MSG_ERROR)
			response.end
		else
			otype = RS4("Name")
			unit = RS4("unit")
			item = RS4("id")
		end if
		RS4.close
		
		mySql="INSERT INTO InventoryItemRequests (order_ID, ItemName, ItemID, comment, ReqDate, Qtty, unit, CreatedBy) VALUES ("& order_ID & ", N'"& otype & "', "& item & ", N'"& comment & "',N'"& shamsiToday() & "', '"& Qtty & "', N'"& unit & "' , "& CreatedBy & " )"
		conn.Execute mySql
		'RS1.close
		response.redirect "goodReq.asp" 
	end if
end if
%>
<center>
	<BR><BR>
	<% call showAlert("<B>ÏŞÊ!</B><BR> ÂíÇ ÍæÇáå ãÑÈæØå ŞÈáÇÕÇÏÑ äÔÏå ÇÓÊ¿<BR>ÇÑ ÍæÇáå ÕÇÏÑ ÔÏå Èå ÈÎÔ ÎÑæÌ ÈÑæíÏ. ",CONST_MSG_ALERT)%>
	<BR><BR>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
	<TR bgcolor="black" >
		<TD align="right" colspan=2><FONT COLOR="YELLOW">ÏÑÎæÇÓÊ ßÇáÇ ÇÒ ÇäÈÇÑ:</FONT></TD>
	</TR>
		<TR bgcolor="#CCCCCC" ><td colspan=2>

		<FORM METHOD=POST ACTION="goodReq.asp"><BR>
			<INPUT TYPE="hidden" name="radif" value="-1">
			<input type="hidden" Name='tmpDlgArg' value=''>
			<input type="hidden" Name='tmpDlgTxt' value=''>
			ßÏ ßÇáÇ: &nbsp;&nbsp;&nbsp;<INPUT  dir="LTR"  TYPE="text" NAME="item" maxlength="10" size="13"   onKeyPress='return mask(this);' onBlur='check(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent">
			<br><br>
			ÊÚÏÇÏ: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="text" NAME="qtty" size=40 ><br><br>
			ÊæÖíÍÇÊ: <TEXTAREA NAME="comment" ROWS="7" COLS="32"></TEXTAREA>
			<br><center>
			<INPUT TYPE="submit" Name="Submit" Value="ËÈÊ ÏÑÎæÇÓÊ ßÇáÇ"  style="width:120px;" tabIndex="14">
			</center>
		</FORM>
		<hr>

		</FONT></TD>
	</TR>
	<%
	'Gets Request for services list from DB



	set RS3=Conn.Execute ("SELECT InventoryItemRequests.*, InventoryItemRequests.ID AS Expr2, InventoryPickuplists.Status AS Expr1 FROM InventoryPickuplists FULL OUTER JOIN InventoryPickuplistItems ON InventoryPickuplists.id = InventoryPickuplistItems.pickupListID FULL OUTER JOIN InventoryItemRequests ON InventoryPickuplistItems.RequestID = InventoryItemRequests.ID WHERE (InventoryItemRequests.Order_ID = - 1) AND (InventoryItemRequests.CreatedBy =  "& session("id") & ") AND (NOT (InventoryItemRequests.Status = 'del' or InventoryItemRequests.Status = 'pick')) AND (NOT InventoryPickuplists.Status = N'done' OR InventoryPickuplists.Status IS NULL)")
	%>
		<%
		Do while not RS3.eof
		%>
		<TR bgcolor="#CCCCCC" title="<% 
			Comment = RS3("Comment")
			if Comment<>"-" then
				response.write "ÊæÖíÍ: " & Comment
			else
				response.write "ÊæÖíÍ äÏÇÑÏ"
			end if
		%>">
			<TD align="right" valign=top><FONT COLOR="black">
			<INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RS3("id")%>" <%
			if RS3("status") = "new" then
				response.write " checked disabled "
			else 
				response.write " disabled "
			end if
			%>><%=RS3("ItemName")%> - <%=RS3("qtty")%> <%=RS3("unit")%> <small>(<%=RS3("ReqDate")%>)</small></td>
			<td align=left width=5%><%
			if RS3("status") = "new" then
			%><a href="goodReq.asp?di=y&i=<%=RS3("id")%>&r=<%=request("radif")%>"><b>ÍĞİ</b></a><%
			end if %></td>
		</tr>
		<% 
		RS3.moveNext
		Loop
		%>

	</table><BR>

	<SCRIPT LANGUAGE="JavaScript">
	<!--
	document.all.item.focus()
	var dialogActive=false;

	function mask(src){ 
		var theKey=event.keyCode;

		if (theKey==13){
			event.keyCode=9
			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="ÌÓÊÌæ ÏÑ ßÇáÇåÇí ÇäÈÇÑ"
			var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('dialog_selectInvItem.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				dialogActive=false
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					document.all.accountName.value=Arguments[1];
				}
			}
		}
	}

	function check(src){ 
		if (!dialogActive){
			badCode = false;
			if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
			objHTTP.open('GET','xml_InventoryItem.asp?id='+src.value,false)
			objHTTP.send()
			tmpStr = unescape(objHTTP.responseText)
			document.all.accountName.value=tmpStr;
			}
	}


	//-->
	</SCRIPT>
	
	<!--#include file="tah.asp" -->
