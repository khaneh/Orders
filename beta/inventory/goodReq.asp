<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Inventory (5)
PageTitle= "À»  œ—ŒÊ«”  ﬂ«·«"
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
if request.form("Submit")="À»  œ—ŒÊ«”  ﬂ«·«" then
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
			call showAlert("Œÿ«! ÂÌç ﬂ«·«ÌÌ «‰ Œ«» ‰ﬂ—œÂ «Ìœ <br><br><A HREF='goodReq.asp'>»—ê‘ </A>",CONST_MSG_ERROR)
			response.end
		else
			order_ID =	clng(order_ID)
			item =		clng(item)
			qtty =		cdbl(qtty)
		end if

		set RS4 = conn.Execute ("SELECT * FROM InventoryItems where OldItemID=" & item)
		if (RS4.eof) then
			response.write "<br><br>"
			call showAlert("Œÿ«! ÂÌç ﬂ«·«ÌÌ «‰ Œ«» ‰ﬂ—œÂ «Ìœ <br><br><A HREF='goodReq.asp'>»—ê‘ </A>",CONST_MSG_ERROR)
			response.end
		else
			otype = RS4("Name")
			unit = RS4("unit")
			item = RS4("id")
		end if
		RS4.close
		
		mySql="INSERT INTO InventoryItemRequests (orderID, ItemName, ItemID, comment, ReqDate, Qtty, unit, CreatedBy) VALUES ("& order_ID & ", N'"& otype & "', "& item & ", N'"& comment & "',getDate(), '"& Qtty & "', N'"& unit & "' , "& CreatedBy & " )"
		conn.Execute mySql
		'RS1.close
		response.redirect "goodReq.asp" 
	end if
end if
%>
<center>
	<BR><BR>
	<% call showAlert("<B>œﬁ !</B><BR> ¬Ì« ÕÊ«·Â „—»ÊÿÂ ﬁ»·«’«œ— ‰‘œÂ «” ø<BR>«ê— ÕÊ«·Â ’«œ— ‘œÂ »Â »Œ‘ Œ—ÊÃ »—ÊÌœ. ",CONST_MSG_ALERT)%>
	<BR><BR>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
	<TR bgcolor="black" >
		<TD align="right" colspan=2><FONT COLOR="YELLOW">œ—ŒÊ«”  ﬂ«·« «“ «‰»«—:</FONT></TD>
	</TR>
		<TR bgcolor="#CCCCCC" ><td colspan=2>

		<FORM METHOD=POST ACTION="goodReq.asp"><BR>
			<INPUT TYPE="hidden" name="radif" value="-1">
			<span>ﬂœ ﬂ«·«:</span>
			<INPUT  dir="LTR"  TYPE="text" NAME="item" maxlength="10" size="13">
			<INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent">
			<br><br>
			<span> ⁄œ«œ:</span>
			<INPUT TYPE="text" NAME="qtty" size=40 ><br><br>
			<span> Ê÷ÌÕ« :</span>
			<TEXTAREA NAME="comment" ROWS="7" COLS="32"></TEXTAREA>
			<br><center>
			<INPUT TYPE="submit" Name="Submit" Value="À»  œ—ŒÊ«”  ﬂ«·«"  style="width:120px;" tabIndex="14">
			</center>
		</FORM>
		<hr>

		</FONT></TD>
	</TR>
	<%
	'Gets Request for services list from DB



	set RS3=Conn.Execute ("SELECT InventoryItemRequests.*, InventoryItemRequests.ID AS Expr2, InventoryPickuplists.Status AS Expr1 FROM InventoryPickuplists FULL OUTER JOIN InventoryPickuplistItems ON InventoryPickuplists.id = InventoryPickuplistItems.pickupListID FULL OUTER JOIN InventoryItemRequests ON InventoryPickuplistItems.RequestID = InventoryItemRequests.ID WHERE (InventoryItemRequests.OrderID = - 1) AND (InventoryItemRequests.CreatedBy =  "& session("id") & ") AND (NOT (InventoryItemRequests.Status = 'del' or InventoryItemRequests.Status = 'pick')) AND (NOT InventoryPickuplists.Status = N'done' OR InventoryPickuplists.Status IS NULL)")
	%>
		<%
		Do while not RS3.eof
		%>
		<TR bgcolor="#CCCCCC" title="<% 
			Comment = RS3("Comment")
			if Comment<>"-" then
				response.write " Ê÷ÌÕ: " & Comment
			else
				response.write " Ê÷ÌÕ ‰œ«—œ"
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
			%><a href="goodReq.asp?di=y&i=<%=RS3("id")%>&r=<%=request("radif")%>"><b>Õ–›</b></a><%
			end if %></td>
		</tr>
		<% 
		RS3.moveNext
		Loop
		%>

	</table><BR>

	<SCRIPT LANGUAGE="JavaScript">
	$(document).ready(function(){
		$.ajaxSetup({
			cache: false
		});
		$("input[name=item]").focus();
		$("input[name=item]").keypress(function(event){
			if (event.keyCode == 10 || event.keyCode == 13) 
		        event.preventDefault();
		});
	    $("input[name=item]").jsonSuggest({
			url: '/service/json_getInventory.asp?act=findItem&a=1' + escape($("input[name=item]").val()),
			minCharacters: 5,
			maxResults: 20,
			width: 214,
			caseSensitive: false,
			onSelect: function(sel){
				var mySel = sel.id.split("|");
				$("input[name=accountName]").val(mySel[1]);
				$("input[name=item]").val(mySel[0]);					
			}
		});

	});
	</SCRIPT>
	
	<!--#include file="tah.asp" -->
