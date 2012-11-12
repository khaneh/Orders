<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= " ’œÊ— ÕÊ«·Â «‰»«—"
SubmenuItem=1
if not (Auth(5 , 1) or Auth(5 , 9)) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		function boxChecked(obj)
		{
		theTR = obj.parentNode.parentNode
		rowNo = theTR.rowIndex - 2

		if(obj.checked)
			{
			document.getElementsByName("box")[rowNo].value = 1;
			}
		else
			{
			document.getElementsByName("box")[rowNo].value = 0;
			}
		}


		function conf()
		{
			var CusHavItemIDs=new Array(100);
			var CusHavItemNames=new Array(100);
			var CusHavQttys=new Array(100);
			var CusHavMax=new Array(100);
			var CusHavCount = 0;
			var CusHavFlag = true;
			var retValue = true;

			if (document.all.CustomerHaveInvItem[0]) {
				for(n=0;n<document.all.CustomerHaveInvItem.length;n++) {
					if(document.all.CustomerHaveInvItem[n].checked)
					{
						CusHavFlag = true;
						for(m=0; m<CusHavCount; m++) {
							if (CusHavItemIDs[m] == document.all.itemID[n].value) {
								CusHavQttys[m] += parseInt(document.all.qtty[n].value);
								CusHavFlag = false;
							}
						}
						if (CusHavFlag) {
								CusHavItemIDs[CusHavCount]=document.all.itemID[n].value;
								CusHavItemNames[CusHavCount]=document.all.itemName[n].value;
								CusHavQttys[CusHavCount] = parseInt(document.all.qtty[n].value);
								CusHavMax[CusHavCount] = parseInt(document.all.maxQtty[n].value);
								CusHavCount++;
							}
					}
				}
				for(m=0;m<CusHavCount;m++) 
					if (CusHavQttys[m]>CusHavMax[m]) {
					alert("Œÿ«! „ÊÃÊœÌ " + CusHavItemNames[m]+" «—”«·Ì ﬂ«›Ì ‰Ì” ")
					retValue = false
					}
			}
			else if (document.all.CustomerHaveInvItem) {
					if(document.all.CustomerHaveInvItem.checked)
					if (parseInt(document.all.qtty.value) > parseInt(document.all.maxQtty.value)) {
					alert("Œÿ«! „ÊÃÊœÌ " + document.all.itemName.value+" «—”«·Ì ﬂ«›Ì ‰Ì” ")
					retValue = false
					}
			}
		return retValue;
		}
		//-->
		</SCRIPT>

<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------- PIRAMOON Pickuplist Submit
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="«ÌÃ«œ ÕÊ«·Â ÅÌ—«„Ê‰" then
	invoice_id = trim(request.form("invoice_id"))
	if not isnumeric(invoice_id) then 
		response.write "<br><br>"
		call showAlert( "‘„«—Â ›«ﬂ Ê— Ê«—œ ‘œÂ «‘ »«Â «” ." , CONST_MSG_ERROR)
		response.end
	end if 


	'vvvvvvv ------------------------------------------ start of check for current ItemOUT
	set RSS=Conn.Execute ("SELECT InventoryLog.ID, InventoryLog.comments, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.VoidedDate, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.type, InventoryLog.ID AS Expr1, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID WHERE (InventoryLog.RelatedInvoiceID = "& invoice_id & ") AND (InventoryLog.IsInput = 0) AND (InventoryLog.Voided = 0)")

	if not RSS.EOF then
		%>
		<br><br>
		<%
		call showAlert("ﬁ»·« «ﬁ·«„ «Ì‰ ÅÌ‘ ‰ÊÌ” «“ «‰»«— Œ«—Ã ‘œÂ «‰œ", CONST_MSG_ALERT)
		response.write "<br><br>" 
		%>
		<TABLE dir=rtl align=center width=600>
			<TR bgcolor="eeeeee" >
				<TD><SMALL>ﬂœ ﬂ«·«</SMALL></TD>
				<TD width=200><SMALL>‰«„ ﬂ«·«</SMALL></TD>
				<TD><SMALL> ⁄œ«œ </SMALL></TD>
				<TD><SMALL>Ê«Õœ</SMALL></TD>
				<TD><SMALL> «—ÌŒ Œ—ÊÃ</SMALL></TD>
				<TD align=center><SMALL>‘„«—Â ÕÊ«·Â</SMALL></TD>
				<TD><SMALL> Ê”ÿ</SMALL></TD>
			</TR>
		<%
		tmpCounter=0
		do while not RSS.EOF
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

			%>
			<TR bgcolor="<%=tmpColor%>"  style="height:25pt" <% if RSS("voided") then%> disabled title="Õœ› ‘œÂ œ—  «—ÌŒ <%=RSS("VoidedDate")%>"<% end if %>>
				<TD  align=right dir=ltr><INPUT TYPE="hidden" name="id" value="<%=RSS("ID")%>"><A HREF="invReport.asp?oldItemID=<%=RSS("oldItemID")%>&logRowID=<%=RSS("ID")%>" target="_blank"><%=RSS("OldItemID")%></A></TD>
				<TD><% if RSS("voided") then%><div style="position:absolute;width:520;"><hr style="color:red;"></div><% end if %><!A HREF="default.asp?itemDetail=<%=RSS("ID")%>"><span style="font-size:10pt"><%=RSS("Name")%></A></TD>
				<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("Qtty")%></span></TD>
				<TD align=right dir=ltr><%=RSS("Unit")%></TD>
				<TD dir=ltr><%=RSS("logDate")%></span></TD>
				<TD  align=center><% 
						if RSS("type")= "2" then
						response.write "<font color=red><b>«’·«Õ „ÊÃÊœÌ</b></font>"
						elseif RSS("type")= "5" then
						response.write "<font color=orang><b>«‰ ﬁ«·</b></font>"
					   else %>
				<A HREF="default.asp?ed=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A>
					<% end if %>
					<% if RSS("owner")<> "-1" and RSS("owner")<> "-2" then
						response.write " («—”«·Ì <a href='../CRM/AccountInfo.asp?act=show&selectedCustomer="& RSS("owner") &"' target='_blank'> " & RSS("owner") & "</a>)"
					   end if %>
					<% if RSS("comments")<> "-" and RSS("comments")<> "" then
						response.write " <br><br><B> Ê÷ÌÕ:</B>  " & RSS("comments") 
					   end if %>

				</TD>
				<TD><%=RSS("RealName")%></TD>
			</TR>
				  
			<% 
			RSS.movenext
		loop
		response.write "</table><br><br>" 
		response.end

	end if
	RSS.close
	set RSS = nothing
	'^^^^^^^ ------------------------------------------ end of check for current ItemOUT

	'vvvvvvv ------------------------------------------ start of check for current Inv PickupList
	set RSS=Conn.Execute ("SELECT InventoryPickuplistItems.pickupListID FROM InventoryPickuplistItems INNER JOIN InventoryItemRequests ON InventoryPickuplistItems.RequestID = InventoryItemRequests.ID INNER JOIN InventoryPickuplists ON InventoryPickuplistItems.pickupListID = InventoryPickuplists.id LEFT OUTER JOIN Orders ON InventoryPickuplistItems.Order_ID = Orders.ID WHERE (InventoryItemRequests.RelatedInvoiceID = "& invoice_id & ") GROUP BY InventoryPickuplistItems.pickupListID, InventoryPickuplists.Status HAVING (InventoryPickuplists.Status <> N'del')"	)
	if not RSS.EOF then
		str = "»—«Ì «Ì‰ ÅÌ‘ ‰ÊÌ” ﬁ»·« ÕÊ«·Â «‰»«— ’«œ— ‘œÂ «” . <BR>»Â »Œ‘ <B>Œ—ÊÃ</B> »—ÊÌœ <BR><BR>"
		do while not RSS.EOF
			str = str & "&nbsp; <A target=_blank HREF='?show="& RSS("pickupListID") & "'>"& RSS("pickupListID") & "</A>"
			RSS.movenext
		loop
		%>
		<br><br>
		<%
		call showAlert(str, CONST_MSG_ALERT)
		response.write "<br><br>" 
		response.end
	end if
	'^^^^^^^ ------------------------------------------ end of check for current Inv PickupList

	
	set RSS=Conn.Execute ("SELECT dbo.InventoryItems.Name, dbo.InvoiceLines.AppQtty, dbo.InventoryItems.ID as itemID, dbo.InvoiceItems.RelatedInventoryItemID, dbo.InventoryItems.Unit FROM dbo.InvoiceLines INNER JOIN dbo.InvoiceItems ON dbo.InvoiceLines.Item = dbo.InvoiceItems.ID INNER JOIN dbo.InventoryItems ON dbo.InvoiceItems.RelatedInventoryItemID = dbo.InventoryItems.OldItemID WHERE (dbo.InvoiceLines.Invoice = " & invoice_id & ")")
	st = ""
	do while not RSS.EOF
		mysql = "INSERT INTO dbo.InventoryItemRequests (Order_ID, ItemName, Comment, ReqDate, Status, ItemID, unit, CreatedBy, CustomerHaveInvItem, Qtty, RelatedInvoiceID) VALUES (-1,N'" & RSS("name") & "', N'„—»Êÿ »Â ›«ﬂ Ê— ‘„«—Â "& invoice_id & "', N'" & shamsiToday() & "', 'new', " & RSS("itemID") & ", N'" & RSS("unit") & "', " & session("id") & ", 0 , " & RSS("AppQtty") & ", "& invoice_id & ")"
		Conn.Execute (mysql)
		'response.write "<br>" & mysql 
		set RSS2=Conn.Execute ("SELECT max(id) as itemReq from InventoryItemRequests")
		st = st & "&itemReq=" & RSS2("itemReq")
		RSS.movenext
	loop

	RSS.close
	set RSS = nothing
	response.redirect "?Submit=sodoor" & st
	response.end


'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------- EDIT Inventory Item Pickuplist Submit
'-----------------------------------------------------------------------------------------------------
elseif request.form("Submit")="«’·«Õ ÕÊ«·Â" then
	GiveTo = request.form("GiveTo") 
	PickupListID = request.form("PickupListID") 

	set RST=Conn.Execute ("SELECT RealName FROM Users WHERE (ID = "& GiveTo & ")" )
	GiveToRealName = RST("RealName")

	set RSY=Conn.Execute ("UPDATE InventoryPickuplists SET  GiveTo ="& GiveTo & ", LastEditTime =N'"& currentTime10() & "', LastEditDate =N'"& shamsitoday() & "', LastEditedBy ="& session("id") & "  where id= "& PickupListID )

	response.write "<BR><BR><TABLE align=center style='border: solid 2pt black'><TR><TD>"
	response.write "<li>  «—ÌŒ  «’·«Õ: " &  CreationDate & "  " & CreationTime
	response.write "<li> œ—Ì«›  ﬂ‰‰œÂ: " & GiveToRealName
	response.write "<hr>"

	for i=1 to request.form("ItemID").count
		ItemID		=	Clng(request.form("ItemID")(i))
		ID			=	Clng(request.form("ID")(i))
		ItemName	=	sqlSafe(request.form("ItemName")(i))
		unit		=	sqlSafe(request.form("unit")(i))
		qtty		=	cdbl(request.form("qtty")(i))
		box			=	sqlSafe(request.form("box")(i))

		mysql = "UPDATE InventoryPickuplistItems SET ItemName = N'"& ItemName & "', Qtty = '"& Qtty & "', CustomerHaveInvItem='" & box & "' WHERE (id = "& ID & ")"
		Conn.Execute (mysql)
		response.write "<li> " & ItemName & " (" & Qtty & " " &  unit & ")"
		if box="1" then
			response.write "<b> «—”«·Ì </b>" 
		end if
	next


	response.write "</TD></TR></TABLE><center>"

	response.write "<BR><BR>ÕÊ«·Â «’·«Õ ‘œ.</center>"
	response.redirect "default.asp?show=" & PickupListID
	response.end

end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------- Print or Show Inventory Item Pickuplist 
'-----------------------------------------------------------------------------------------------------
if request("show")<>"" then
	PickupListID = request("show")

	%>
	<BR><BR><BR>	
	<CENTER>
	<% 	ReportLogRow = PrepareReport ("InvPickupList.rpt", "InvItem_ID", PickupListID, "/beta/dialog_printManager.asp?act=Fin") %>
	<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

	<input type="button" value="«’·«Õ" Class="GenButton" onclick="window.location='default.asp?ed=<%=PickupListID%>'">

	</CENTER>

	<BR>
	<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:700; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no></iframe>
	<%

	response.end
end if




'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- EDIT Inventory Item Pickuplist
'-----------------------------------------------------------------------------------------------------
if request("ed")<>"" then
	PickupListID = request("ed")
	set RSY=Conn.Execute ("SELECT * FROM InventoryPickuplists WHERE id = "& PickupListID )
	set RSW=Conn.Execute ("SELECT InventoryItemRequests.comment,InventoryItemRequests.ReqDate, dbo.InventoryItems.OldItemID AS OldItemID, dbo.InventoryPickuplistItems.*, dbo.Users.RealName AS RealName FROM dbo.InventoryItems INNER JOIN dbo.InventoryItemRequests INNER JOIN dbo.InventoryPickuplistItems ON dbo.InventoryItemRequests.ID = dbo.InventoryPickuplistItems.RequestID ON dbo.InventoryItems.ID = dbo.InventoryPickuplistItems.ItemID INNER JOIN dbo.Users ON dbo.InventoryItemRequests.CreatedBy = dbo.Users.ID WHERE (dbo.InventoryPickuplistItems.pickupListID = " & PickupListID & ")")

	if RSY.EOF or RSW.EOF then%>
		<div align=center dir=rtl><br><br><br><br><br><br><b>Œÿ«!!<br><h3>«Ì‰ ÕÊ«·Â ÊÃÊœ ‰œ«—œ</h3></b><br></div>
		</body>
		</html>
	<%	
	else
		%><br>
		<form method=post action="default.asp" <% if RSY("Status")="done" then %> disabled <% end if %>  onsubmit="return confirm('ÊÌ—«Ì‘ ÕÊ«·Â «‰Ã«„ ‘Êœø')"><br>
		<INPUT TYPE="hidden" name=PickupListID value="<%=PickupListID%>">
		<TABLE dir=rtl align=center width=600>
		<TR bgcolor="eeeeee" >
			<TD align=center colspan=7><B>«’·«Õ ÕÊ«·Â «‰»«— </B></TD>
		</TR>
		<TR bgcolor="eeeeee" >
				<TD style="width:160;">‰Ê⁄ ﬂ«·«</TD>
				<TD style="width:70;"> «—ÌŒ œ—ŒÊ«” </TD>
				<TD style="width:70;"> ⁄œ«œ</TD>
				<TD style="width:60;">„ ﬁ«÷Ì</TD>
				<TD>‘„«—Â ”›«—‘</TD>
				<TD  style="width:160;">ﬂ«·«Ì «—”«·Ì </TD>
		</TR>
		<%
		tmpCounter=0
		totalQtty=0
		Do while not RSW.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			End if 
			%>
			<TR bgcolor="<%=tmpColor%>"  title="<% 
				response.write " ﬂœ ﬂ«·« : " & RSW("OldItemID") & " - "
				Comment = RSW("Comment")
				if Comment<>"-" then
					response.write " Ê÷ÌÕ: " & Comment
				else
					response.write " Ê÷ÌÕ ‰œ«—œ"
				end if
			%>">
				<%
				myRequestID=RSW("RequestID")
				totalQtty = totalQtty + RSW("qtty")
				%>
				<TD><INPUT TYPE="hidden" name=ID  value="<%=RSW("id")%>"><INPUT TYPE="hidden" name=itemID  value="<%=RSW("itemid")%>"><INPUT TYPE="hidden" name=unit  value="<%=RSW("unit")%>"><INPUT TYPE="text" NAME="itemName" readonly value="<%=RSW("ItemName")%>"  style="width:160; border:0"></TD>
				<TD><%=RSW("ReqDate")%></TD>
				<TD><INPUT TYPE="text" NAME="qtty" value="<%=RSW("qtty")%>"  size=7 style="width:45; border:0">  <%=RSW("unit")%></TD>
				<TD><%=RSW("RealName")%></TD>
				<TD>
				<% if RSW("order_ID")<>-1 then %>
				<a href="../order/order.asp?act=show&id=<%=RSW("order_ID")%>">
				<%=RSW("order_ID")%></a>
				<% end if %>
				</TD>
				<TD>
					<% 

					set RST=Conn.Execute ("SELECT SUM((CONVERT(tinyint, dbo.InventoryLog.IsInput) - .5) * 2 * dbo.InventoryLog.Qtty) AS sumQtty, dbo.Accounts.AccountTitle FROM dbo.Orders INNER JOIN dbo.InventoryLog ON dbo.Orders.Customer = dbo.InventoryLog.owner INNER JOIN dbo.Accounts ON dbo.Orders.Customer = dbo.Accounts.ID WHERE (dbo.InventoryLog.ItemID = " & RSW("ItemID")  & " and dbo.InventoryLog.voided=0) GROUP BY dbo.Orders.ID, dbo.Accounts.AccountTitle HAVING (dbo.Orders.ID = " & RSW("order_ID")  & ")" )
					if not RST.EOF then
						
						if clng(RST("sumQtty")) < 0 then 
							response.write "<b>Œÿ«!</b> „ÊÃÊœÌ ﬂ«·«Ì «Ì‰ „‘ —Ì „‰›Ì «” .»Â ‰Ÿ— ‘„« çÿÊ— Â„çÌ‰ çÌ“Ì „„ﬂ‰ «” ø!" 
							response.write "<INPUT TYPE='hidden' NAME='maxQtty'> <INPUT TYPE='hidden' NAME='CustomerHaveInvItem' value='-1'><INPUT TYPE='hidden' NAME='box'  value='0'>" 
						elseif clng(RST("sumQtty")) > 0 then  
							if RSW("CustomerHaveInvItem") then
								response.write "<INPUT TYPE='hidden' NAME='maxQtty' value='" & RST("sumQtty") & "'> <INPUT TYPE='checkbox' NAME='CustomerHaveInvItem' checked  value='" & RST("sumQtty") & "' onclick='boxChecked(this)'><INPUT TYPE='hidden' NAME='box'  value='1'> «“ " & RST("sumQtty") & " " & RSW("unit") & " " &  RSW("ItemName") &  "  «—”«·Ì " & RST("AccountTitle") & " ﬂ”— ‘Êœø"
							else 
								response.write "<INPUT TYPE='hidden' NAME='maxQtty'  value='" & RST("sumQtty") & "'> <INPUT TYPE='checkbox' NAME='CustomerHaveInvItem'  value='" & RST("sumQtty") & "'  onclick='boxChecked(this)'><INPUT TYPE='hidden' NAME='box'  value='0'> «“ " & RST("sumQtty") & " " & RSW("unit") & " " & RSW("ItemName") & " «—”«·Ì " & RST("AccountTitle") & " ﬂ”— ‘Êœø"
							end if 
						else
							response.write "<INPUT TYPE='hidden' NAME='maxQtty' > <INPUT TYPE='hidden' NAME='CustomerHaveInvItem'  value='-1'><INPUT TYPE='hidden' NAME='box'  value='0'>" 
						end if 
					else
						response.write "<INPUT TYPE='hidden' NAME='maxQtty' > <INPUT TYPE='hidden' NAME='CustomerHaveInvItem'  value='-1'><INPUT TYPE='hidden' NAME='box'  value='0'>" 
					end if

					%>
				</TD>
			</TR>
			<% 
			RSW.moveNext
		Loop
		CreatedBy = RSY("CreatedBy")
		%>
		<TR bgcolor="eeeeee" >
			<TD align=center colspan=7 height=10></TD>
		</TR>
		<TR >
			<TD colspan=7 align=center><BR>œ—Ì«›  ﬂ‰‰œÂ: 
				<select name="GiveTo" class=inputBut  <% if RSY("Status")="done" then %> disabled <% end if %>>
				<% set RSV=Conn.Execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
				Do while not RSV.eof
				%>
					<option value="<%=RSV("ID")%>" <%
						if RSV("ID")=RSY("GiveTo") then
							response.write " selected "
						end if
						%>><%=RSV("RealName")%></option>
				<%
				RSV.moveNext
				Loop
				RSV.close
				%>
				</select><BR><BR>
				 «—ÌŒ «ÌÃ«œ: <span dir=ltr bgcolor=red><%=RSY("CreationDate")%></span> - ”«⁄  <%=RSY("CreationTime")%>
				<BR><BR>
				Ê÷⁄Ì : <%=RSY("Status")%>
				<BR><BR>
				<% if RSY("Printed")<>"0" then%>
				 ⁄œ«œ œ›⁄«  ç«Å: <%=RSY("Printed")%>
				<% end if %>
				<BR><BR>
				<INPUT TYPE="submit"  class=inputBut Name="Submit" Value="«’·«Õ ÕÊ«·Â"  style="width:100px;" tabIndex="14"  <% if RSY("Status")="done" then %> disabled <% end if %>  onclick="return conf()">
				<!INPUT TYPE="submit"  class=inputBut Name="Submit" Value="ç«Å ÕÊ«·Â"  style="width:100px;" tabIndex="14"  <% if RSY("Status")="done" then %> disabled <% end if %>>
			</TD>
		</TR>
		</FORM>
		</body>
		</html>
		<%
	end if 
	response.end
end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Inventory Item Pickuplist Submit
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="À»  ÕÊ«·Â" then
'response.write "<br>" & replace(request.form,"&","<br>")
'esponse.end

	response.write "<BR><BR><TABLE align=center style='border: solid 2pt black'><TR><TD>"

	'===================================================
	'create an pickuplist item
	'===================================================
	GiveTo			=	request.form("GiveTo") 
	CreationDate	=	shamsiToday()
	CreationTime	=	currentTime10()
	CreatedBy		=	session("id")
	Status			=	"new"

	set RST=Conn.Execute ("SELECT RealName FROM Users WHERE (ID = "& GiveTo & ")" )
	GiveToRealName = RST("RealName")

	mysql = "INSERT INTO InventoryPickuplists (CreationTime, CreationDate, Status, CreatedBy, GiveTo) VALUES (N'"& CreationTime & "', N'"& CreationDate & "', N'new', "& CreatedBy & ", "& GiveTo & ")"
	Conn.Execute (mysql)
	'response.write mysql
	response.write "<li>  «—ÌŒ : " &  CreationDate & "  " & CreationTime
	response.write "<li> œ—Ì«›  ﬂ‰‰œÂ: " & GiveToRealName
	response.write "<hr>"
	

	'===================================================
	'create many pickuplist lines for an pickuplist item
	'===================================================
	set RSX=Conn.Execute ("SELECT * FROM InventoryPickuplists WHERE CreationTime = N'"& CreationTime & "' and CreationDate=N'" & CreationDate & "' and GiveTo=" & GiveTo )	
	pickupListID = RSX("id")
	RSX.close

	for i=1 to request.form("ItemID").count
		ItemID		=	request.form("ItemID")(i)
		ItemName	=	request.form("ItemName")(i)
		unit		=	request.form("unit")(i)
		qtty		=	request.form("qtty")(i)
		order_ID	=	request.form("Related_order_ID")(i)
		RequestID	=	request.form("RequestID")(i)
		box			=	request.form("box")(i)

		mysql = "INSERT INTO InventoryPickuplistItems (pickupListID, Order_ID, ItemName, ItemID, Qtty, unit, RequestID, CustomerHaveInvItem) VALUES ( "& pickupListID & ", "& order_ID & ", N'"& ItemName & "', "& ItemID & ", "& qtty & ", N'"& unit & "', "& RequestID & ", " &  box & ")"
		Conn.Execute (mysql)
		'response.write mysql
		response.write "<li> " & ItemName & " (" & Qtty & " " &  unit & ")"
		if box="1" then
			response.write "<b> «—”«·Ì </b>" 
		end if
	next


	'===================================================
	'change status of related requests
	'===================================================
	for i=1 to request.form("RequestID").count
		RequestID	=	request.form("RequestID")(i)

		mysql = "UPDATE InventoryItemRequests SET Status = N'pick' WHERE (ID = "& RequestID & ")"
		Conn.Execute (mysql)
		'response.write mysql
	next

	response.write "</TD></TR></TABLE><center>"
	%>
	<form method=post action="default.asp?PickupListID=<%=pickupListID%>" ><br>
	<INPUT TYPE="submit"  class=inputBut Name="Submit" Value="ç«Å ÕÊ«·Â"  style="width:100px;" tabIndex="14">
	</form>
	<%
	response.write "<BR><BR>ÕÊ«·Â ’«œ— ‘œ.</center>"
	response.redirect "default.asp?show=" & PickupListID
	response.end

end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------- Inventory Item Pickuplist Creation Form
'-----------------------------------------------------------------------------------------------------
if request("Submit")="Õ–› œ—ŒÊ«” " then
	if request("itemReq").count < 1 then%>
		<br><br>
		<%
		call showAlert("ÂÌç œ—ŒÊ«” Ì  »—«Ì Õ–› «‰ Œ«» ‰‘œÂ «” ", CONST_MSG_ERROR)
	else
		for i=1 to request("itemReq").count
			myRequestID=request("itemReq")(i)
			set RSX=Conn.Execute ("SELECT * from [dbo].[InventoryItemRequests] where id=" & myRequestID)
			itemName = RSX("Qtty") & " " & RSX("unit") & " " & RSX("ItemName")
			CSR = RSX("CreatedBy")
			set RSX=Conn.Execute ("update [dbo].[InventoryItemRequests] set status='del' where id=" & myRequestID)
			call InformRequestDenied(itemName, CSR)
			response.write "<br><br><center><li> œ—ŒÊ«”  " & itemName & " Õ–› ‘œ. </center>"
		next
	end if
	'response.end
end if
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------- Inventory Item Pickuplist Creation Form
'-----------------------------------------------------------------------------------------------------
if request("Submit")="’œÊ— ÕÊ«·Â Œ—ÊÃ" or request("Submit")="sodoor" then
	if request("itemReq").count < 1 then%>
		<div align=center dir=rtl><br><br><br><br><br><br><b>Œÿ«!!<br><h3>ÂÌç œ—ŒÊ«” Ì «‰ Œ«» ‰‘œÂ «” </h3></b><br></div>
		</body>
		</html>
	<%	
	else
		%><br>
		<form method=post action="default.asp" onsubmit="return confirm('’œÊ— ÕÊ«·Â «‰Ã«„ ‘Êœø')">
		<TABLE dir=rtl align=center width=600>
		<TR bgcolor="eeeeee" >
			<TD align=center colspan=7><B>ÕÊ«·Â «‰»«— </B></TD>
		</TR>
		<TR bgcolor="eeeeee" >
				<TD style="width:160;">‰Ê⁄ ﬂ«·«</TD>
				<TD style="width:70;"> «—ÌŒ œ—ŒÊ«” </TD>
				<TD style="width:70;"> ⁄œ«œ</TD>
				<TD style="width:60;">„ ﬁ«÷Ì</TD>
				<TD>‘„«—Â ”›«—‘</TD>
				<TD  style="width:160;">ﬂ«·«Ì «—”«·Ì </TD>
		</TR>
		<%
		tmpCounter=0
		totalQtty=0
		for i=1 to request("itemReq").count
			myRequestID=request("itemReq")(i)
			set RSX=Conn.Execute ("SELECT dbo.InventoryItems.OldItemID AS OldItemID, InventoryItems.name, dbo.InventoryItemRequests.*, dbo.Users.RealName AS RealName FROM dbo.InventoryItems INNER JOIN dbo.InventoryItemRequests ON dbo.InventoryItems.ID = dbo.InventoryItemRequests.ItemID INNER JOIN dbo.Users ON dbo.InventoryItemRequests.CreatedBy = dbo.Users.ID WHERE InventoryItemRequests.ID = "& myRequestID )	
			totalQtty = totalQtty + RSX("qtty")
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			End if 
			%>
			<TR bgcolor="<%=tmpColor%>" title="<% 
				response.write " ﬂœ ﬂ«·« : " & RSX("OldItemID") & " - "
				Comment = RSX("Comment")
				if Comment<>"-" then
					response.write " Ê÷ÌÕ: " & Comment
				else
					response.write " Ê÷ÌÕ ‰œ«—œ"
				end if
			%>">
				<TD >
					<INPUT TYPE="hidden" name=RequestID  value="<%=RSX("id")%>">
					<INPUT TYPE="hidden" name=itemID  value="<%=RSX("ItemID")%>">
					<INPUT TYPE="text" readonly NAME="itemName" value="<%=RSX("Name")%>" style="width:160; border:0">
					<br>
					<small>œ— ŒÊ«”  ‘œÂ: <%=RSX("ItemName")%></small>
				</TD>
				<TD align=center>
					<span dir=ltr><%=shamsidate(RSX("ReqDate"))%></span>
					<span dir=ltr><%=DatePart("h",RSX("ReqDate")) &":"& DatePart("n",RSX("ReqDate"))%></span>
				</TD>
				<TD>
					<INPUT TYPE="text" NAME="qtty" value="<%=RSX("qtty")%>"  size=7 style="width:45; border:0">  <%=RSX("unit")%>
					<INPUT TYPE="hidden" name=unit  value="<%=RSX("unit")%>">
				</TD>
				<TD><%=RSX("RealName")%></TD>
				<TD><INPUT TYPE="hidden" name="Related_order_ID"  value="<%=RSX("orderID")%>">
				<% if RSX("orderID")<>-1 then %>
				<a href="../order/order.asp?act=show&id=<%=RSX("orderID")%>">
				<%=RSX("orderID")%></a>
				<% end if %>
				</TD>
				<TD>
					<% 

					set RST=Conn.Execute ("SELECT SUM((CONVERT(tinyint, dbo.InventoryLog.IsInput) - .5) * 2 * dbo.InventoryLog.Qtty) AS sumQtty, dbo.Accounts.AccountTitle FROM dbo.Orders INNER JOIN dbo.InventoryLog ON dbo.Orders.Customer = dbo.InventoryLog.owner INNER JOIN dbo.Accounts ON dbo.Orders.Customer = dbo.Accounts.ID WHERE (dbo.InventoryLog.ItemID = " & RSX("ItemID")  & " and dbo.InventoryLog.voided=0) GROUP BY dbo.Orders.ID, dbo.Accounts.AccountTitle HAVING (dbo.Orders.ID = " & RSX("orderID")  & ")" )
					if not RST.EOF then
						
						if clng(RST("sumQtty")) < 0 then 
							response.write "<b>Œÿ«!</b> „ÊÃÊœÌ ﬂ«·«Ì «Ì‰ „‘ —Ì „‰›Ì «” .»Â ‰Ÿ— ‘„« çÿÊ— Â„çÌ‰ çÌ“Ì „„ﬂ‰ «” ø!" 
							response.write "<INPUT TYPE='hidden' NAME='maxQtty'> <INPUT TYPE='hidden' NAME='CustomerHaveInvItem' value='-1'><INPUT TYPE='hidden' NAME='box'  value='0'>" 
						elseif clng(RST("sumQtty")) > 0 then  
							if RSX("CustomerHaveInvItem") then
								response.write "<INPUT TYPE='hidden' NAME='maxQtty' value='" & RST("sumQtty") & "'> <INPUT TYPE='checkbox' NAME='CustomerHaveInvItem' checked  value='" & RST("sumQtty") & "' onclick='boxChecked(this)'><INPUT TYPE='hidden' NAME='box'  value='1'> «“ " & RST("sumQtty") & " " & RSX("unit") & " " &  RSX("ItemName") &  "  «—”«·Ì " & RST("AccountTitle") & " ﬂ”— ‘Êœø"
							else 
								response.write "<INPUT TYPE='hidden' NAME='maxQtty'  value='" & RST("sumQtty") & "'> <INPUT TYPE='checkbox' NAME='CustomerHaveInvItem'  value='" & RST("sumQtty") & "'  onclick='boxChecked(this)'><INPUT TYPE='hidden' NAME='box'  value='0'> «“ " & RST("sumQtty") & " " & RSX("unit") & " " & RSX("ItemName") & " «—”«·Ì " & RST("AccountTitle") & " ﬂ”— ‘Êœø"
							end if 
						else
							response.write "<INPUT TYPE='hidden' NAME='maxQtty' > <INPUT TYPE='hidden' NAME='CustomerHaveInvItem'  value='-1'><INPUT TYPE='hidden' NAME='box'  value='0'>" 
						end if 
					else
						response.write "<INPUT TYPE='hidden' NAME='maxQtty' > <INPUT TYPE='hidden' NAME='CustomerHaveInvItem'  value='-1'><INPUT TYPE='hidden' NAME='box'  value='0'>" 
					end if

					%>
				</TD>
			</TR>
			<% 
			CreatedBy = RSX("CreatedBy")
			LastRSX_typeID = RSX("ItemID")
			RSX.moveNext
		next
		%>
		<TR bgcolor="eeeeee" >
			<TD align=center colspan=7 height=10></TD>
		</TR>
		<TR >
			<TD colspan=7 align=center><BR>œ—Ì«›  ﬂ‰‰œÂ: 
				<select name="GiveTo" class=inputBut>
				<% set RSV=Conn.Execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
				Do while not RSV.eof
				%>
					<option value="<%=RSV("ID")%>" <%
						if RSV("ID")=CreatedBy then
							response.write " selected "
						end if
						%>><%=RSV("RealName")%></option>
				<%
				RSV.moveNext
				Loop
				RSV.close
				%>
				</select><BR><BR>
				 «—ÌŒ «ÌÃ«œ: <span dir=ltr bgcolor=red><%=shamsitoday()%></span> 
				<BR><BR>
				<INPUT TYPE="submit"  class=inputBut Name="Submit" Value="À»  ÕÊ«·Â"  style="width:100px;" tabIndex="14" onclick="return conf()">
			</TD>
		</TR>
		</FORM>
		</body>
		</html>
		<%
	end if 
	response.end
end if

if Request("act")="waste" then 
	if not Auth(5 , "K") then NotAllowdToViewThisPage()
 %>
<div class="inPage">
	<span>‘„«—Â ”›«—‘: </span>
	<input id="orderID" size="6" class="boot"/>
	<div id="message" class="well well-small"><ul></ul></div>
</div>
<div id="orderHeader" class="well"></div>
<div id="confirm">
	<h4></h4>
	<h5>œ·Ì· ŒÊœ —« »—«Ì «Ì‰ œ—ŒÊ«”  ‘—Õ œÂÌœ</h5>
	<input name="comment" type="text" size="20"/>
</div>
<script type="text/javascript">
$(document).ready(function(){
	$("#confirm").dialog({
		autoOpen: false,
		buttons: {" «ÌÌœ":function(){
			var thisLI = $("#message ul li input[name=add].btn-danger").closest("li");
			$.ajax({
				type:"POST",
				url:"/service/json_getInventory.asp",
				data:{
					act:"addWasteRequest",
					orderID:$("#orderID").val(),
					rowName: thisLI.find("input[name=rowName]").val(),
					rowID: thisLI.find("input[name=rowID]").val(),
					maxRowID: thisLI.find("input[name=maxRowID]").val(),
					groupName: thisLI.find("input[name=groupName]").val(),
					qtty: thisLI.find("input[name=qtty]").val(),
					reqID: thisLI.find("input[name=reqID]").val(),
					comment: $("#confirm input[name=comment]").val()
				},
				dataType:"json"
			}).done(function (data){
				$("#message").append(data.status);
				$(this).dialog("close");
			});
			//thisLI.find("input[name=add]").removeClass("btn-danger");
			//$("#message ul li input[name=add]").prop("disabled",false);
			
		}},
		title: " «ÌÌœ"
	});
	
	$("#orderID").blur(function(){
		var orderID = Number($("#orderID").val());
		$("#message ul").html("");
		if (!isNaN(orderID) && orderID!='') {
			loadXMLDoc("/service/xml_getOrderProperty.asp?act=showHead&id=" + orderID, function(orderXML){
				
				var isOrder = $(orderXML).find("status isOrder").text();
				var isClosed = $(orderXML).find("status isClosed").text();
				var isApproved = $(orderXML).find("status isApproved").text();
				var invoiceIssued = $(orderXML).find("status invoiceIssued").text();
				var invoiceApproved = $(orderXML).find("status invoiceApproved").text();
				var step = $(orderXML).find("status step").text();
				var dis="";
				if (isOrder=='0'){
					$("#message").html("«” ⁄·«„ ﬂÂ ÷«Ì⁄«  ‰œ«—Â!!!!!");
					dis = "disabled='disabled'";
				} else
				if (isClosed!='0'){
					$("#message").html("”›«—‘ »” Â ‘œÂ!");
					dis = "disabled='disabled'";
				} else if (isApproved=='0'){
					$("#message").html("”›«—‘  «ÌÌœ ‰‘œÂ");
					dis = "";
				} else if (invoiceIssued!='0') {
					$("#message").html("⁄ÃÌ»! ›«ﬂ Ê— ’«œ— ‘œÂ!");
					dis = "disabled='disabled'";
				} else if (invoiceApproved!='0'){
					$("#message").html("›«ﬂ Ê— «Ì‰ ”›«—‘  «ÌÌœ ‘œÂ° ›·–« œŒ· Ê  ’—› œ— ¬‰ „„ﬂ‰ ‰Ì” ");
					dis = "disabled='disabled'";
				} else {
					dis = "";
				}	
				
				TransformXml(orderXML, "/xsl.<%=version%>/orderShowHeader.xsl", function(result){
					$("#orderHeader").html(result);	
					$('a#customerID').click(function(e){
						window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer='+$('a#customerID').attr("myID"), 'showCustomer');
						e.preventDefault();
					});
				});
				loadXMLDoc("/service/xml_getOrderProperty.asp?act=editOrder&id=" + orderID, function(propXML){
					loadXMLDoc("/service/xml_getOrderProperty.asp?act=stock&id=" + orderID, function(stockXML){
						$(propXML).find("group[hasStock=yes]").each(function(i){
							var rowName = $(this).closest("rows").attr("name");
							var groupName = $(this).attr("name");
							var invoiceItem = $(this).attr("item");
							var rowID = $(this).closest("row").attr("id");
							var req = $(stockXML).find("req/invoiceItem:contains(" + invoiceItem + ")").find("rowID:contains(" + rowID + ")").parent().find("reqStatus:not(:contains('del'))").parent();
							if (req.size()!=0){
								$("#message ul").append("<li><span name='itemName'>" + $(req).find("ItemName").text() + 
									"</span><input name='qtty' value='" + $(req).find("Qtty").text() + 
									"' size='4'/><input name='add' " + dis + 
									" type='button' class='btn' value='À»  œ—ŒÊ«” ' onclick='addRequest(this)'/><span name='unit'>" + 
									$(req).find("unit").text() + "</span><input type='hidden' name='rowName' value='" + 
									rowName + "'/><input type='hidden' name='reqID' value='" + $(req).find("ID").text() + 
									"'/><input type='hidden' name='groupName' value='" + groupName + 
									"'/><input type='hidden' name='rowID' value='" + rowID + 
									"'/><input type='hidden' name='maxRowID' value='0'/></li>");
							}
							
						});
						$("#message ul li").each(function(i){
							var rowName = $(this).find("input[name=rowName]").val();
							var rowID = parseInt($(this).find("input[name=rowID]").val());
							var maxRowID = parseInt($(this).find("input[name=maxRowID]").val());
							if (rowID > maxRowID)
								$("#message ul li input[name=rowName]").parent().find("input[name=maxRowID]").val(rowID);
						});
					});
				});
			});
		} 
	});
});
function addRequest(e){
	var thisLI = $(e).closest("li");
	$("#confirm h4").html("¬Ì« «ÿ„Ì‰«‰ œ«—Ìœ ﬂÂ »—«Ì " + thisLI.find("span[name=itemName]").text() + " " + thisLI.find("input[name=qtty]").val() + " " + thisLI.find("span[name=unit]").text() + " »Â œ·Ì· Œ—«»Ì/÷«Ì⁄«  œ—ŒÊ«”  ê—œœø");
	thisLI.find("input[name=add]").addClass("btn-danger");
	$("#message ul li input[name=add]").prop("disabled",true);
	$("#confirm").dialog("open");
}
</script>
<%	
	Response.end
end if

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------- LIST Inventory Item Pickuplists
'-----------------------------------------------------------------------------------------------------
%>
<style>
	.changeItem{cursor: pointer;}
	td.delBtn {position: relative;}
	td.delBtn span {position: absolute;opacity: .6;cursor: pointer;top:2px;left: 20px;font-size: 8px;font-family: tahoma;}
</style>
<script type="text/javascript">
	$(document).ready(function(){
		$.ajaxSetup({
			cache: false
		});
		$("td.delBtn span").css('display','none');
		$("td.delBtn").mouseover(function(event){
			$(this).find("span").css("display","block");
		});
		$("td.delBtn").mouseout(function(event){
			$(this).find("span").css("display","none");
		});	
		$("td.delBtn span").click(function(){
			var requestID = $(this).closest("td").find("input[name=itemReq]").val();
			$("#comfirmDel input").val(requestID);
			if (parseInt($(this).closest("td").find("input[name=itemReq]").attr("invoiceitem")) > 0){
				$("#comfirmDel h3").html(" ÊÃÂ° œ— ’Ê—   «ÌÌœ „Ãœœ ”›«—‘ «Ì‰ œ—ŒÊ«”  »Â ’Ê—  ŒÊœﬂ«— „Ãœœ «ÌÃ«œ ŒÊ«Âœ ‘œ");
			} else {
				$("#comfirmDel h3").html("«“ Õ–› «Ì‰ œ—ŒÊ«”  «ÿ„Ì‰«‰ œ«—Ìœø");
			}
			$("#comfirmDel").dialog("open");
			
		});
		$('.changeItem').prop('title','ÃÂ   €ÌÌ—/ ⁄ÌÌ‰ ¬Ì „ ﬂ·Ìﬂ ﬂ‰Ìœ');
		$('.changeItem').click(function(){
			var myRow = $(this).closest("tr");
			var itemReq = myRow.find('input[name=itemReq]').val();
			var theInvoiceitem = myRow.find('input[name=itemReq]').attr('invoiceitem');
			$("#itemReq").val(itemReq);
			
			$.ajax({
				type:"POST",
				url:"/service/json_getInventory.asp",
				data:{act:"itemListFromInvoiceItem",invoiceItem:theInvoiceitem},
				dataType:"json"
			}).done(function (data){
				$("#itemID").children("option").remove();
				$.each(data,function(i,e){
					$("#itemID").append("<option value='" + e.inventoryItem + "' unit='" + e.unit + "'>" + e.name + "(" + e.unit + ")" + "</option>")
				});
			});
			$('#changeItemDlg').dialog("open");
		});
		$("#changeItemDlg").dialog({ 
			autoOpen: false,
			buttons: {" «ÌÌœ":function(){
				var invID=$("input[name=InvoiceID]").val();
				$.ajax({
					type:"POST",
					url:"/service/json_getInventory.asp",
					data:{act: "updateItemRequest" ,id: $("#itemReq").val(), unit: $("#itemID option:selected").attr("unit"), itemID:$("#itemID option:selected").val()},
					dataType:"json"
				}).done(function (data){
					if (data.status=="ok")
						$("[name=itemReq][value=" + $("#itemReq").val() + "]").prop("disabled",false);
					//location.reload();
				});
				$(this).dialog("close");
			}},
			title: "«‰ Œ«» ¬Ì „"
		});
		$("#comfirmDel").dialog({
			autoOpen: false,
			buttons: {" «ÌÌœ":function(){
				$.ajax({
					type:"POST",
					url:"/service/json_getInventory.asp",
					data:{act:"delInvRequest",id:$("#comfirmDel input").val()},
					dataType:"json"
				}).done(function(data){
					if (data.status=="ok")
						$("[name=itemReq][value=" + $("#comfirmDel input").val() + "]").closest("tr").remove();
				});
				$(this).dialog("close");
			}},
			title: "Õ–› œ—ŒÊ«” "
		});
	});
</script>
<div id='changeItemDlg'>
	<input type="hidden" id="itemReq"/>
	<select id='itemID'></select>
</div>
<div id='comfirmDel'>
	<h3></h3>
	<input type="hidden" name="reqID"/>
</div>
<BR><BR>
<br>
<FORM METHOD=POST ACTION="default.asp">
<%

sortBy=request("s")
if sortBy="5" then 
	sB="orderID"
elseif sortBy="2" then 
	sB="ReqDate"
elseif sortBy="3" then 
	sB="Qtty"
elseif sortBy="4" then 
	sB="CreatedBy"
else 
	sB="ItemName"
end if

if Auth(5 , 9) then
	set RSS=Conn.Execute ("SELECT dbo.InventoryItemRequests.*, dbo.Users.RealName AS RealName FROM dbo.InventoryItemRequests INNER JOIN dbo.Users ON dbo.InventoryItemRequests.CreatedBy = dbo.Users.ID WHERE (dbo.InventoryItemRequests.Status = 'new') AND (dbo.InventoryItemRequests.OrderID <> - 1) order by " & sB)

	if not RSS.eof then
		%>
		<TABLE dir=rtl align=center width=600>
		<TR bgcolor="eeeeee" height=20>
			<TD align=center colspan=6><INPUT TYPE="hidden" name=color1 value="#FFFFFF"><INPUT TYPE="hidden" name=color2 value="#FFFFBB"><B>œ—ŒÊ«”  Â«Ì  ﬂ«·«Ì „—»Êÿ »Â ”›«—‘ Â«</B></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD><INPUT TYPE="checkbox" NAME="" disabled><INPUT TYPE="hidden" name=color1 value="#FFFFFF"><INPUT TYPE="hidden" name=color2 value="#FFFFBB"</TD>
			<TD><A HREF="default.asp?s=1"><SMALL>‰Ê⁄  ﬂ«·«</SMALL></A></TD>
			<TD><A HREF="default.asp?s=2"><SMALL> «—ÌŒ œ—ŒÊ«” </SMALL></A></TD>
			<TD><A HREF="default.asp?s=3"><SMALL> ⁄œ«œ</SMALL></A></TD>
			<TD><A HREF="default.asp?s=4"><SMALL>œ—ŒÊ«”  ﬂ‰‰œÂ</SMALL></A></TD>
			<TD><A HREF="default.asp?s=5"><SMALL>‘„«—Â ”›«—‘</SMALL></A></TD>
		</TR>
		<%
		tmpCounter=0
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
				'tmpColor="#DDDDDD"
				'tmpColor2="#EEEEBB"
			End if 

		%>
		<TR bgcolor="<%=tmpColor%>" >
			<TD class="delBtn">
				<INPUT TYPE="hidden" name=color1 value="<%=tmpColor%>"/>
				<INPUT TYPE="hidden" name=color2 value="<%=tmpColor2%>"/>
				<INPUT TYPE="checkbox"  onclick="setPrice(this)" <%if IsNull(rss("itemID")) then response.write "disabled='disabled'"%> invoiceItem="<%=rss("invoiceItem")%>" NAME="itemReq" VALUE="<%=RSS("id")%>"/>
				<span class="label label-important">X</span>
			</TD>
			<TD <%if not IsNull(rss("invoiceItem")) then Response.write " class='changeItem'"%>><%=RSS("ItemName")%></TD>
			<TD><%=shamsidate(RSS("ReqDate"))%></small></TD>
			<TD><%=RSS("Qtty")%> <%=rss("unit")%></TD>
			<TD><%=RSS("RealName")%></TD>
			<TD><a href="../order/order.asp?act=show&id=<%=RSS("orderID")%>"><%=RSS("orderID")%></a>
			<% if RSS("CustomerHaveInvItem") then	
				response.write " <b>  («—”«·Ì) </b>" 
			end if
			%>
			</small></TD>
		</TR>
		<!--TR bgcolor="<%=tmpColor%>" >
			<TD></TD>
			<TD colspan=5> Ê÷ÌÕ: <%=RSS("comment")%></small></TD>
		</TR-->
			  
		<% 
		RSS.moveNext
		Loop
		%>
		<TR >
			<TD align=center colspan=6 height=50><INPUT TYPE="hidden" name=color1 value=""><INPUT TYPE="hidden" name=color2 value="">
				<center><INPUT TYPE="submit" Name="Submit" Value="Õ–› œ—ŒÊ«” "  style="width:150px;" tabIndex="14" class="btn" onclick="return confirm('¬Ì« „ÿ„∆‰« „Ì ŒÊ«ÂÌœ «Ì‰ œ—ŒÊ«”  —« Õ–› ﬂ‰Ìœø')"> <INPUT TYPE="submit" Name="Submit" Value="’œÊ— ÕÊ«·Â Œ—ÊÃ"  style="width:150px;" class="btn" tabIndex="14"> 
				</center>	<BR>
			</TD>
		</TR>
		</TABLE>
		<%
	end if 
end if 

if Auth(5 , 1) then
	set RSS=Conn.Execute ("SELECT dbo.InventoryItemRequests.*, dbo.Users.RealName AS RealName FROM dbo.InventoryItemRequests INNER JOIN dbo.Users ON dbo.InventoryItemRequests.CreatedBy = dbo.Users.ID WHERE (dbo.InventoryItemRequests.Status = 'new') AND (dbo.InventoryItemRequests.OrderID = - 1) order by " & sB)

	if not RSS.eof then
		%>
		<TABLE dir=rtl align=center width=600>
		<TR bgcolor="eeeeee" >
			<TD align=center colspan=6 height=20><INPUT TYPE="hidden" name=color1 value="#FFFFFF"><INPUT TYPE="hidden" name=color2 value="#FFFFBB"><B>œ—ŒÊ«”  Â«Ì  ﬂ«·«Ì „ ›—ﬁÂ</B></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD><INPUT TYPE="hidden" name=color1 value="#FFFFFF"><INPUT TYPE="hidden" name=color2 value="#FFFFBB"><INPUT TYPE="checkbox" NAME="" disabled></TD>
			<TD><A HREF="default.asp?s=1"><SMALL>‰Ê⁄  ﬂ«·«</SMALL></A></TD>
			<TD><A HREF="default.asp?s=2"><SMALL> «—ÌŒ œ—ŒÊ«” </SMALL></A></TD>
			<TD><A HREF="default.asp?s=3"><SMALL> ⁄œ«œ</SMALL></A></TD>
			<TD><A HREF="default.asp?s=4"><SMALL>œ—ŒÊ«”  ﬂ‰‰œÂ</SMALL></A></TD>
		</TR>
		<%
		tmpCounter = tmpCounter + 1
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
				'tmpColor="#DDDDDD"
				'tmpColor2="#EEEEBB"
			End if 

		%>
		<TR bgcolor="<%=tmpColor%>" >
			<TD><INPUT TYPE="hidden" name=color1 value="<%=tmpColor%>"><INPUT TYPE="hidden" name=color2 value="<%=tmpColor2%>"><INPUT TYPE="checkbox"  onclick="setPrice(this)"  NAME="itemReq" VALUE="<%=RSS("id")%>"></TD>
			<TD><%=RSS("ItemName")%></TD>
			<TD><%=shamsidate(RSS("ReqDate"))%></small></TD>
			<TD><%=RSS("Qtty")%> <%=rss("unit")%></TD>
			<TD><%=RSS("RealName")%></TD>
		</TR>
		<!--TR bgcolor="<%=tmpColor%>" >
			<TD></TD>
			<TD colspan=5> Ê÷ÌÕ: <%=RSS("comment")%></small></TD>
		</TR-->
			  
		<% 
		RSS.moveNext
		Loop
		%>
		</TABLE><br>
		<center>
			<INPUT TYPE="submit" Name="Submit" Value="Õ–› œ—ŒÊ«” "  style="width:150px;" tabIndex="14" class="btn" onclick="return confirm('¬Ì« „ÿ„∆‰« „Ì ŒÊ«ÂÌœ «Ì‰ œ—ŒÊ«”  —« Õ–› ﬂ‰Ìœø')"/> 
			<INPUT TYPE="submit" Name="Submit" Value="’œÊ— ÕÊ«·Â Œ—ÊÃ"  style="width:150px;" tabIndex="14" class="btn"/>
			<a href="?act=waste" class="btn" style="width:150px;">’œÊ— ÕÊ«·Â »—ê‘ Ì</a>
		</center>
		</form>
	<% end if %>
	<BR><BR>
	<TABLE align=center width=50%>
	<TR>
		<TD align=center style="border:solid 1pt black">
			<BR>
			<FORM METHOD=POST ACTION="default.asp">
			<B>’œÊ— ÕÊ«·Â Ì ÅÌ—«„Ê‰ <BR></B><BR>
			‘„«—Â ÅÌ‘‰Â«œ ﬁÌ„  : <INPUT TYPE="text" NAME="invoice_id" dir=ltr ><BR><BR> <INPUT TYPE="submit"  name="submit"  value="«ÌÃ«œ ÕÊ«·Â ÅÌ—«„Ê‰">
			<BR>
			</FORM>
		</TD>
	</TR>
	</TABLE><BR>
	<%
end if 
%>

<SCRIPT LANGUAGE="JavaScript">

function setPrice(obj)
{
theTR = obj.parentNode.parentNode
ii = theTR.rowIndex -1

if(obj.checked)
	{
	//alert(ii)
	theTR.setAttribute("bgColor",document.getElementsByName('color2')[ii].value)
	}
else
	{
	theTR.setAttribute("bgColor",document.getElementsByName('color1')[ii].value)
	}
}

</SCRIPT>


<!--#include file="tah.asp" -->