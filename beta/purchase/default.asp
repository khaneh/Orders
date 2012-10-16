<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Purchase (4)
PageTitle=" œ—ŒÊ«”  Œ—Ìœ ﬂ«·«Ì «‰»«—"
SubmenuItem=1
if not Auth(4 , 1) then response.redirect "outServiceOrder.asp"
if not Auth(4 , 1) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<SCRIPT LANGUAGE="JavaScript">
<!--
function hideIT()
{
if(document.all.tavafogh.checked) 
	{
		document.all.priceTavafoghi.style.visibility= 'visible'
	}
	else
	{
		document.all.price.value= ''
		document.all.priceComment.value= ''
		document.all.priceTavafoghi.style.visibility= 'hidden'
	}
}

function hideIT2()
{
if(document.all.tavafogh2.checked) 
	{
		document.all.priceTavafoghi2.style.visibility= 'visible'
	}
	else
	{
		document.all.orderID.value= ''
		document.all.priceTavafoghi2.style.visibility= 'hidden'
	}
}
//-->
</SCRIPT>

<%
catItem1 = request("catItem")
if catItem1="" then catItem1="-1"

goodItem1 = request("goodItem")
if goodItem1="" then goodItem1="-1"


'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------- Delete an Inventory Request for Buy from an order
'-----------------------------------------------------------------------------------------------------
if request("di")="y" then		
	myRequestID=request("i")
	set RSX=Conn.Execute ("SELECT * FROM purchaseRequests WHERE id = "& myRequestID )	
	if RSX("status")="new" then
	Conn.Execute ("update purchaseRequests SET status = 'del' where id = "& myRequestID )	
	end if
	response.redirect "default.asp?radif=" & request("r")
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------ Submit an Inventory Item request For Buy
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="À»  œ—ŒÊ«”  Œ—Ìœ ﬂ«·«" then
	priceComment = request.form("priceComment")
	item = request.form("item")
	comment = request.form("comment")
	price = request.form("price")
	qtty = request.form("qtty")
	CreatedBy = session("id")
	date1 = request.form("date1")
	orderID = request.form("orderID")

	if 	not item = "" then

		if price="" then
			price = 0
		end if

		if priceComment="" then
			priceComment = "-"
		end if

		if comment="" then
			comment = "-"
		end if

		if qtty="" then
			qtty = "0"
		end if

		if orderID="" then
			orderID = "-1"
		end if

		if item="-1" then
			response.write "<br><br><center>Œÿ«! ÂÌç ﬂ«·«ÌÌ «‰ Œ«» ‰ﬂ—œÂ «Ìœ"
			response.write "<br><br><A HREF='default.asp'>»—ê‘ </A></center>"
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
		
		mySql="INSERT INTO purchaseRequests (OrderID, TypeName, TypeID, comment, ReqDate, Qtty, CreatedBy, Price, priceComment, DueDate, IsService) VALUES ( "& orderID & ", N'"& otype & "', "& item & " , N'"& comment & "',getDate(), "& Qtty & ", "& CreatedBy & " , "& Price & " , N'"& priceComment & "', dbo.udf_date_solarToDate(" & mid(date1,1,4) & "," & mid(date1,6,2) & "," & mid(date1,9,2) & "), 0 )"
		conn.Execute mySql
		'RS1.close
		response.write "<center><br><br>œ—ŒÊ«”  À»  ‘œ </center><br>"
	end if
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------ Main Inventory Item request For Buy Form
'-----------------------------------------------------------------------------------------------------
%>
<center>
<BR><BR>
<TABLE width="*">
<TR>

<TD valign=top width=50%>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
	<TR >
		<TD align="right" colspan=2><H3>œ—ŒÊ«”  Œ—Ìœ ﬂ«·«Ì «‰»«—</H3></TD>
	</TR>
		<TR bgcolor="dddddd" ><td colspan=2>

		<FORM METHOD=POST ACTION="default.asp?Submit=<%=request("Submit")%>">
			<INPUT TYPE="hidden" name="radif" value="-1">
			<SELECT NAME="catItem" style='width:350;font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1" onchange="document.forms[0].submit()">
			<option value="-1">œ” Â »‰œÌ ﬂ«·« —« «‰ Œ«» ﬂ‰Ìœ: </option>
			<option value="-1">----------------------------------------------</option>
<%
				set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories ORDER BY Replace([Name],' ','')")
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
			<SELECT NAME="item" style='width:350;font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1">
			<option value="-1">‰Ê⁄ ﬂ«·« —« «‰ Œ«» ﬂ‰Ìœ: </option>
			<option value="-1">----------------------------------------------</option>
<%
			mySQL="SELECT InventoryItems.* FROM InventoryItemCategoryRelations INNER JOIN InventoryItems ON InventoryItemCategoryRelations.Item_ID = InventoryItems.ID WHERE (InventoryItemCategoryRelations.Cat_ID = " & catItem1 & ") ORDER BY Replace([Name],' ','')" 
			set RS4 = conn.Execute (mySQL)
			while not (RS4.eof) 
%>				<OPTION value='<%=RS4("ID")%>'<%
				if trim(goodItem1) = trim(RS4("ID")) then
				response.write " selected "
				end if
				%>>* <%=RS4("OldItemID")%> - <%=RS4("Name")%> (<%=RS4("Unit")%>)</option>
<%
				RS4.MoveNext
			wend
			RS4.close
%>
			</SELECT><br><br>
			<% end if %>
			 ⁄œ«œ: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="text" NAME="qtty" size=40 onKeyPress="return maskNumber(this);" dir="LTR"><br><br>
			<INPUT TYPE="checkbox" onclick="hideIT()" name="tavafogh"> Ê«›ﬁ ﬁÌ„  Œ«’Ì ’Ê—  ê—› Â «” <BR>
			<div name="priceTavafoghi" id="priceTavafoghi" style="visibility:'hidden'">ﬁÌ„ : &nbsp;<INPUT TYPE="text" NAME="price" ID="price" size=5 onKeyPress="return maskNumber(this);"> ÿ—›: <INPUT TYPE="text" NAME="priceComment" id="priceComment" size=26 ></div>

			<INPUT TYPE="checkbox" onclick="hideIT2()" name="tavafogh2">«Ì‰ œ—ŒÊ«”  „—»Êÿ »Â ”›«—‘ Œ«’Ì «” <BR>
			<div name="priceTavafoghi2" id="priceTavafoghi2" style="visibility:'hidden'">‘„«—Â ”›«—‘: &nbsp;<INPUT TYPE="text" NAME="orderID" ID="orderID" size=10 onKeyPress="return maskNumber(this);"> </div>


			 «—ÌŒÌ ﬂÂ ﬂ«·« „Ê—œ ‰Ì«“ «” : <INPUT dir=ltr TYPE="text" NAME="date1" size=15 value="<%=shamsiToday()%>" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"><br><br>

			 Ê÷ÌÕ« : <TEXTAREA NAME="comment" ROWS="7" COLS="32"></TEXTAREA>
			<br><center>
			<INPUT class=inputBut TYPE="submit" Name="Submit" Value="À»  œ—ŒÊ«”  Œ—Ìœ ﬂ«·«" style="width:125px;" tabIndex="14"<%
			if catItem1="-1" then
				response.write " disabled "
			end if
			%>>
			</center>
		</FORM>

		</FONT></TD>
	</TR>
	<%
	'Gets Request for services list from DB
	set RS3=Conn.Execute ("SELECT * FROM purchaseRequests WHERE (status='new' and IsService=0)")
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
			%>><B><%=RS3("TypeName")%></B> &nbsp;&nbsp;<small dir=ltr>( ⁄œ«œ: <%=RS3("qtty")%> -  «—ÌŒ :<%=RS3("ReqDate")%>)</small></td>
			<td align=left width=5%><%
			if RS3("status") = "new" then
			%><a href="default.asp?di=y&i=<%=RS3("id")%>&r=<%=request("radif")%>"><b>Õ–›</b></a><%
			end if %></td>
		</tr>
		<% 
		RS3.moveNext
		Loop
		%>

	</table>
</TD>

<TD valign=top width=50%>
<FORM METHOD=POST ACTION="default.asp"><li><B> ê“«—‘ Ìﬂ <br></B>ﬂ«·«Â«ÌÌﬂÂ „ÊÃÊœÌ‘«‰ ﬂ„ — «“ Õœ«ﬁ· „ÊÃÊœÌ «” 	<BR>
<center>
<INPUT class=inputBut TYPE="submit" name="submit" value="«ÌÃ«œ ê“«—‘ Ìﬂ ">
</center>
</FORM>
<%
if trim(request("Submit"))=trim("«ÌÃ«œ ê“«—‘ Ìﬂ ") then 'or request("Submit")="" then
	%>
	<TABLE width=95% align=center>
	<TR bgcolor=#66FFFF>
		<TD>‰«„ ﬂ«·«</TD>
		<TD>„ÊÃÊœÌ</TD>
		<TD>Õœ«ﬁ·</TD>
		<TD>”›«—‘ Ã«—Ì</TD>
		<TD>Ê«Õœ</TD>
	</TR>

	<%
	'set RSX=Conn.Execute ("SELECT * FROM InventoryItems WHERE Qtty <= Minim")	

	if session("id") = 104 then ' if User is Mr Koochaki, just show items in the 1 and 5 categories
		extraCondition = " and (InventoryItemCategoryRelations.Cat_ID = 1 or InventoryItemCategoryRelations.Cat_ID = 5)"
	else
		extraCondition = ""
	end if 

	set RSX=Conn.Execute ("SELECT InventoryItems.*, InventoryItemCategoryRelations.Cat_ID FROM InventoryItems INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID WHERE (InventoryItems.Qtty < InventoryItems.Minim) " & extraCondition & " ORDER BY Replace([Name],' ','')")	
	Do while not RSX.eof
		set RSK=Conn.Execute ("SELECT sum(qtty) as sumQtty FROM PurchaseOrders WHERE IsService=0 and TypeID="& RSX("ID") &" and Status<>'OK' and Status<>'CANCEL'" )	

	%>
	<TR>
		<TD class=alak2><A HREF="default.asp?goodItem=<%=RSX("id")%>&catItem=<%=RSX("cat_ID")%>&Submit=<%=request("Submit")%>"><%=RSX("Name")%></A></TD>
		<TD><%=RSX("Qtty")%></TD>
		<TD><%=RSX("Minim")%></TD>
		<TD><%
		if not RSK.eof then
			response.write RSK("sumQtty") '& "(" & RSK("Status") &")"
		end if
		
		%></TD>
		<TD><%=RSX("Unit")%></TD>
	</TR>
	<TR>
		<TD colspan=5 bgcolor=red></TD>
	</TR>
	<%

	RSX.moveNext
	Loop
	RSX.close
	%>	
	</TABLE>
	<%

end if
%>

<FORM METHOD=POST ACTION="default.asp">
<hr>
<li><B> ê“«—‘ œÊ<br></B>ﬂ«·«Â«ÌÌ ﬂÂ «ê— Â„Â œ—ŒÊ«”  Â«Ì‘«‰ »—¬Ê—œÂ ‘Êœ „ÊÃÊœÌ‘«‰ ﬂ„ — «“ Õœ«ﬁ· „ÊÃÊœÌ „Ì ‘Êœ.<BR>
<center>
<INPUT class=inputBut TYPE="submit" name="submit" value="«ÌÃ«œ ê“«—‘ œÊ">
</center>
</FORM>

<%
if trim(request("Submit"))=trim("«ÌÃ«œ ê“«—‘ œÊ") then 'or request("Submit")="" then
	%>
	<TABLE width=95% align=center>
	<TR bgcolor=#66FFFF>
		<TD>‰«„ ﬂ«·«</TD>
		<TD>œ—ŒÊ«” </TD>
		<TD>„ÊÃÊœÌ</TD>
		<TD>Õœ«ﬁ·</TD>
		<TD>”›«—‘ Ã«—Ì</TD>
		<TD>Ê«Õœ</TD>
	</TR>

	<%
	'set RSX=Conn.Execute ("SELECT InventoryItems.id, InventoryItems.Unit, InventoryItems.Minim, InventoryItems.Qtty, InventoryItems.Name, DERIVEDTBL.sumQtty AS sumReqQtty, DERIVEDTBL_1.SQ, InventoryItemCategoryRelations.Cat_ID FROM (SELECT ItemID, SUM(Qtty) AS sumQtty FROM InventoryItemRequests WHERE (Status = N'new') GROUP BY ItemID) DERIVEDTBL INNER JOIN InventoryItems ON DERIVEDTBL.ItemID = InventoryItems.ID AND InventoryItems.Qtty - DERIVEDTBL.sumQtty < InventoryItems.Minim INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID LEFT OUTER JOIN (SELECT SUM(Qtty) AS SQ, typeID FROM PurchaseOrders WHERE (status <> N'ok' AND status <> N'cancel' AND status <> N'RETURN') GROUP BY typeID) DERIVEDTBL_1 ON InventoryItems.ID = DERIVEDTBL_1.typeID ")	

	set RSX=Conn.Execute ("SELECT InventoryItems.ID, InventoryItems.Unit, InventoryItems.Minim, InventoryItems.Qtty, InventoryItems.Name, DERIVEDTBL_1.SQ, InventoryItemCategoryRelations.Cat_ID, InventoryItems.sumReqQtty FROM (SELECT SUM(Qtty) AS SQ, typeID FROM PurchaseOrders WHERE (status <> N'ok' AND status <> N'cancel' AND status <> N'RETURN') GROUP BY typeID) DERIVEDTBL_1 RIGHT OUTER JOIN (SELECT InventoryItems.ID, InventoryItems.Unit, InventoryItems.Minim, InventoryItems.Qtty, InventoryItems.Name, ISNULL(reqsDRVTABLE.sumQtty, 0) AS sumReqQtty FROM InventoryItems LEFT OUTER JOIN (SELECT InventoryItemRequests.ItemID, ISNULL(SUM(InventoryItemRequests.Qtty), 0) AS sumQtty FROM InventoryPickuplists INNER JOIN InventoryPickuplistItems ON InventoryPickuplists.id = InventoryPickuplistItems.pickupListID RIGHT OUTER JOIN InventoryItemRequests ON InventoryPickuplistItems.RequestID = InventoryItemRequests.ID WHERE (InventoryItemRequests.Status = N'new') OR (InventoryItemRequests.Status = N'pick') AND (InventoryPickuplists.Status = N'new') GROUP BY InventoryItemRequests.ItemID) reqsDRVTABLE ON InventoryItems.ID = reqsDRVTABLE.ItemID) InventoryItems INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID ON InventoryItems.ID = DERIVEDTBL_1.typeID WHERE (InventoryItems.Minim > InventoryItems.Qtty - InventoryItems.sumReqQtty) ORDER BY Replace([Name],' ','')")

	Do while not RSX.eof

	%>
	<TR>
		<TD class=alak2><A HREF="default.asp?goodItem=<%=RSX("id")%>&catItem=<%=RSX("cat_ID")%>&Submit=<%=request("Submit")%>"><%=RSX("Name")%></A></TD>
		<TD><%=RSX("sumReqQtty")%></TD>
		<TD><%=RSX("Qtty")%></TD>
		<TD><%=RSX("Minim")%></TD>
		<TD><%
		if not isnull( RSX("SQ")) then
			response.write RSX("SQ") '& "("& RSX("status") & ")"
		end if
		%></TD>
		<TD><%=RSX("Unit")%></TD>
	</TR>
	<TR>
		<TD colspan=6 bgcolor=red></TD>
	</TR>
	<%

	RSX.moveNext
	Loop
	RSX.close
	%>	
	</TABLE>
	<%

end if
%>

</TD>

</TR>
</TABLE>
<br>
	

<!--#include file="tah.asp" -->
