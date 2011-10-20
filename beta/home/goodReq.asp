<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Home (0)
PageTitle= " œ—ŒÊ«”  ﬂ«·«"
SubmenuItem=5
if not Auth(0 , 5) then NotAllowdToViewThisPage()

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
	order_ID = request.form("radif")
	item = request.form("item")
	comment = request.form("comment")
	qtty = request.form("qtty")
	CreatedBy = session("id")
	if 	not item = "" then

		if  comment="" then
			comment = "-"
		end if

		if order_ID="" or item="-1" or qtty="" or qtty="0" then
			response.write "<br><br><center>Œÿ«! ÂÌç ﬂ«·«ÌÌ «‰ Œ«» ‰ﬂ—œÂ «Ìœ"
			response.write "<br><br><A HREF='goodReq.asp'>»—ê‘ </A></center>"
			response.end
		end if

		set RS4 = conn.Execute ("SELECT * FROM InventoryItems where ID=" & item)
		if (RS4.eof) then
			otype="-unknown-"
			unit=RS4("unit")
		else
			otype=RS4("Name")
			unit=RS4("unit")
		end if
		RS4.close
		
		mySql="INSERT INTO InventoryItemRequests (order_ID, ItemName, ItemID, comment, ReqDate, Qtty, unit, CreatedBy) VALUES ("& order_ID & ", N'"& otype & "', "& item & ", N'"& comment & "',N'"& shamsiToday() & "', "& Qtty & ", N'"& unit & "' , "& CreatedBy & " )"
		conn.Execute mySql
		'RS1.close
		response.redirect "goodReq.asp" 
	end if
end if
%>
<center>
	<BR><BR>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
	<TR bgcolor="black" >
		<TD align="right" colspan=2><FONT COLOR="YELLOW">œ—ŒÊ«”  ﬂ«·« «“ «‰»«—:</FONT></TD>
	</TR>
		<TR bgcolor="#CCCCCC" ><td colspan=2>

		<FORM METHOD=POST ACTION="goodReq.asp">
			<INPUT TYPE="hidden" name="radif" value="-1">
			<SELECT NAME="catItem" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1"  onchange="document.forms[0].submit()">
			<option value="-1">œ” Â »‰œÌ ﬂ«·« —« «‰ Œ«» ﬂ‰Ìœ: </option>
			<option value="-1">----------------------------------------------</option>
<%
				set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories")
				while not (RS4.eof) %>
					<OPTION value="<%=RS4("ID")%>"<%
					if trim(catItem1) = trim(RS4("ID")) then
					response.write " selected "
					end if
					%>>* <%=RS4("Name")%> </option>
<%						RS4.MoveNext
				wend
				RS4.close
				%>
			</SELECT><br><br>
			<%
			if not catItem1="-1" then


			%>
			<SELECT NAME="item" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1">
			<option value="-1">‰Ê⁄ ﬂ«·« —« «‰ Œ«» ﬂ‰Ìœ: </option>
			<option value="-1">----------------------------------------------</option>
<%
				set RS5 = conn.Execute ("SELECT * FROM InventoryItemCategoryRelations where Cat_ID=" & catItem1)
				while not (RS5.eof) 
					set RS4 = conn.Execute ("SELECT * FROM InventoryItems where id=" & RS5("Item_ID") ) %>
					<OPTION value="<%=RS4("ID")%>">* <%=RS4("Name")%> (<%=RS4("Unit")%>)</option>
<%						RS5.MoveNext
				wend
				RS5.close
				%>
			</SELECT><br><br>
			<% end if %>
			 ⁄œ«œ: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="text" NAME="qtty" size=40 onKeyPress="return maskNumber(this);"><br><br>
			 Ê÷ÌÕ« : <TEXTAREA NAME="comment" ROWS="7" COLS="32"></TEXTAREA>
			<br><center>
			<INPUT TYPE="submit" Name="Submit" Value="À»  œ—ŒÊ«”  ﬂ«·«"  style="width:120px;" tabIndex="14"<%
			if catItem1="-1" then
				response.write " disabled "
			end if
			%>>
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
	
	<!--#include file="tah.asp" -->