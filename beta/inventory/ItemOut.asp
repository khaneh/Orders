<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= " Œ—ÊÃ ﬂ«·«"
SubmenuItem=2
if not Auth(5 , 2) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->


<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------- Item Out
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")=" «ÌÌœ" then
	PickID		=	request.form("PickID")
	comments	=	request.form("comments")
	exitDate	=	request.form("exitDate")

	if Auth(5 , "C") AND "" & exitDate <> "" then ' À»  Ê—Êœ/Œ—ÊÃ œ—  «—ÌŒ œ·ŒÊ«Â
		logDate=exitDate
	else
		logDate=shamsiToday()
	end if
	
	'===================================================
	' check status of pick list to be 'New' 
	' and change it to 'done'
	'===================================================
	set RSS=Conn.Execute ("SELECT * FROM InventoryPickuplists WHERE (id = "& PickID & " and status='new') " )
	if RSS.eof then 
		response.write "<BR><BR><BR><BR><center>Œÿ«! <BR><BR>«Ì‰ ÕÊ«·Â ﬁ»·« «“ «‰»«— Œ«—Ã ‘œÂ «” . </center>"
		'response.write pickID
		response.end
	end if
	mySQL = "update InventoryPickuplists set status='done' where (id = "& PickID & ")"
'	response.write mySQL
	Conn.Execute (mySQL)


	'===================================================
	' insert items in InventoryLogOut table and
	' change the Qtty of item in InventoryItems table
	'===================================================


	set RSI=Conn.Execute ("SELECT dbo.InventoryPickuplistItems.*, ISNULL(dbo.Orders.Customer, - 1) AS owner, dbo.InventoryItemRequests.RelatedInvoiceID FROM dbo.InventoryPickuplistItems INNER JOIN dbo.InventoryItemRequests ON dbo.InventoryPickuplistItems.RequestID = dbo.InventoryItemRequests.ID LEFT OUTER JOIN dbo.Orders ON dbo.InventoryPickuplistItems.Order_ID = dbo.Orders.ID WHERE (dbo.InventoryPickuplistItems.pickupListID = "& PickID & ")")
	if RSI.eof then 
		response.write "<BR><BR><BR><BR><center>Œÿ«! <BR><BR>ÕÊ«·Â Œ«·Ì «” </center>"
		response.end
	end if

	response.write "<BR><BR><BR><BR><center>«ﬁ·«„ “Ì— «“ «‰»«— Œ«—Ã ‘œ:"
	Do while not RSI.eof
		ItemID		=	RSI("ItemID")
		ItemName	=	RSI("ItemName")
		unit		=	RSI("unit")
		qtty		=	RSI("qtty")
		order_ID	=	RSI("order_ID")
		RequestID	=	RSI("RequestID")
		RelatedInvoiceID	=	RSI("RelatedInvoiceID")
		if trim(RelatedInvoiceID) = "" or isnull(RelatedInvoiceID) then
			RelatedInvoiceID = 0
		end if

		if RSI("CustomerHaveInvItem") then
			owner	=	RSI("owner")
		else
			owner	= "-1"
		end if
		mySQL = "SET NOCOUNT ON;INSERT INTO InventoryLog (ItemID, RelatedID, Qtty, logDate, CreatedBy, owner, IsInput, comments, type, RelatedInvoiceID,price,gl_update) VALUES ("& ItemID & ","& PickID& ","& Qtty& ",N'"& logDate& "', "& session("ID") & ", " & owner & " , 0, N'" & comments & "', 1, "& RelatedInvoiceID & ", null,0);select @@identity as newID;SET NOCOUNT OFF;"
		'response.write mySQL 
		'response.end
		set rs = Conn.Execute(mySQL)
		newOut = rs.Fields("newID").value
		rs.close
		set rs= nothing
		if RSI("CustomerHaveInvItem") then
			set RSW=Conn.Execute ("SELECT SUM((CONVERT(tinyint, dbo.InventoryLog.IsInput) - .5) * 2 * dbo.InventoryLog.Qtty) AS sumQtty, dbo.Accounts.AccountTitle FROM dbo.Orders INNER JOIN dbo.InventoryLog ON dbo.Orders.Customer = dbo.InventoryLog.owner INNER JOIN dbo.Accounts ON dbo.Orders.Customer = dbo.Accounts.ID WHERE (dbo.InventoryLog.ItemID = " & ItemID & " and dbo.InventoryLog.voided=0) GROUP BY dbo.Orders.ID, dbo.Accounts.AccountTitle HAVING (dbo.Orders.ID = " & order_ID  & ")")
			newItemQtty = RSW("sumQtty")
			RSW.close
			set RSW = nothing
		else
			set RSW=Conn.Execute ("SELECT * FROM InventoryItems WHERE (id = "& ItemID & ")" )
			newItemQtty = RSW("qtty")
			RSW.close
			set RSW = nothing
		end if
		response.write "<li> " & ItemName & " " & qtty & " " & unit & " (»«ﬁÌ„«‰œÂ: " & newItemQtty & " " & unit & ")"
		if clng(newItemQtty)<0 then 
			response.write "error!"
			response.end
		end if
'		Conn.Execute ("update InventoryItems set qtty="& newItemQtty & " where (id = "& ItemID & ")")
' --------------- TEmPORARY DISABLE
'		Conn.Execute("execute dbo.outFIFO")
		set rs=conn.Execute("select * from InventoryFIFORelations where outID=" & newOut)
		while not rs.eof
			response.write("<li> Œ—ÊÃ «“ Ê—Êœ <a href='invReport.asp?oldItemID=-1&itemID=" & itemID & "'>" & rs("inID") & "</a> »Â  ⁄œ«œ " & rs("qtty") & " «‰Ã«„ ‘œ.<br>")
			rs.moveNext
		wend
		RSI.moveNext
	loop 
	RSI.close
	response.write "</center>"
	
%>
	<BR>	
	<CENTER>
		<% 	ReportLogRow = PrepareReport ("InvPickupList.rpt", "InvItem_ID", PickID, "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">
	</CENTER>

	<BR>
	<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:700; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no></iframe>
<%
response.end
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------- Inventory Item Pickup list Submit
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="Õ–› ÕÊ«·Â" then
	

	PickID =  request.form("PickID")
	if PickID = "" then%>
		<br><br>
		<%
		call showAlert("ÂÌç ÕÊ«·Â «Ì »—«Ì Õ–› «‰ Œ«» ‰‘œÂ «” ", CONST_MSG_ERROR)
	else
			set RSX=Conn.Execute ("update InventoryPickuplists set status='del', LastEditedBy=" & session("id") & ", LastEditDate=N'" & shamsiToday() & "', LastEditTime=N'" & currentTime10() & "' where id=" & PickID)
			response.write "<br><br>"
			call showAlert(" ÕÊ«·Â ‘„«—Â " & PickID & " Õ–› ‘œ. ", CONST_MSG_INFORM)
				
			set RSX=Conn.Execute ("SELECT dbo.InventoryPickuplistItems.RequestID FROM         dbo.InventoryPickuplists INNER JOIN dbo.InventoryPickuplistItems ON dbo.InventoryPickuplists.id = dbo.InventoryPickuplistItems.pickupListID where InventoryPickuplists.id=" & PickID)


			Do while not RSX.eof
				Conn.Execute ("UPDATE dbo.InventoryItemRequests SET status='new' WHERE (ID = "& RSX("RequestID") & ")")
				RSX.moveNext
			loop
			RSX.close

			response.write "<br><br>"
			call showAlert(" œ—ŒÊ«”  Â«Ì „—»ÊÿÂ »Â ’› ’œÊ— ÕÊ«·Â »—ê‘ ‰œ. ", CONST_MSG_INFORM)
			
	end if

end if
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------- Inventory Item Pickup list Submit
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="«‰ Œ«» »—«Ì Œ—ÊÃ" then
	 PickID =  request.form("PickID")
	 if PickID ="" then
		response.redirect "ItemOut.asp"
	end if

	set RSS=Conn.Execute ("SELECT * FROM InventoryPickuplists WHERE (id = "& PickID & ") " )
	if RSS.eof then 
		response.redirect "ItemOut.asp"
	end if

	set RSI=Conn.Execute ("SELECT dbo.InventoryPickuplistItems.*, ISNULL(dbo.Orders.Customer, - 1) AS owner FROM dbo.InventoryPickuplistItems LEFT OUTER JOIN dbo.Orders ON dbo.InventoryPickuplistItems.Order_ID = dbo.Orders.ID WHERE (dbo.InventoryPickuplistItems.pickupListID = "& PickID & ") " )
	if RSI.eof then 
		response.redirect "ItemOut.asp"
	end if

	set RST=Conn.Execute ("SELECT RealName FROM Users WHERE (ID = "& RSS("GiveTo") & ")" )
	GiveToRealName = RST("RealName")

	response.write "<center><br><br><br><BR><BR><BR>‘„« œ— Õ«· À»  Œ—ÊÃ ﬂ«·«Ì „—»Êÿ »Â ÕÊ«·Â “Ì— «“ «‰»«— Â” Ìœ: "
	response.write "<BR><br><br><BR><TABLE align=center style='border: solid 2pt black; '><TR><TD style='font-size:12pt'>"
	response.write "<li>‘„«—Â ÕÊ«·Â: " & PickID
	response.write "<li>  «—ÌŒ ’œÊ—: " &  RSS("CreationDate") & "  (" & RSS("CreationTime") & ")"
	response.write "<li> œ—Ì«›  ﬂ‰‰œÂ: " & GiveToRealName
	response.write "<hr>"
	Do while not RSI.eof
		ItemID		=	RSI("ItemID")
		ItemName	=	RSI("ItemName")
		unit		=	RSI("unit")
		qtty		=	RSI("qtty")
		order_ID	=	RSI("order_ID")
		RequestID	=	RSI("RequestID")
		if RSI("CustomerHaveInvItem") then
			owner	=	RSI("owner")
		else
			owner	= "-1"
		end if

		set RSW=Conn.Execute ("SELECT * FROM InventoryItems WHERE (id = "& ItemID & ")" )
		OldItemID = RSW("OldItemID")
		if owner = "-1" then
			itemQtty = RSW("qtty")
		else
			itemQtty = RSW("cusQtty")
		end if

		if itemQtty < Qtty then 
			response.write "<span><li  style='background-color: #FF9933'> "& OldItemID & " - " & ItemName & " (œ—ŒÊ«” Ì: " & Qtty & " " &  unit & " - „ÊÃÊœÌ: "& itemQtty & " "& unit & ")</span>"
			notAvailable = "yes"
		else
			response.write "<li> "& OldItemID & " - " & ItemName & " (" & Qtty & " " &  unit & ")"
			if not owner = "-1" then
				response.write "<b> «—”«·Ì " & owner & " </b>" 
			end if
		end if
		RSI.moveNext
	loop
	RSI.close
	 
response.write "</TD></TR></TABLE><br><br><FORM METHOD=POST ACTION='ItemOut.asp'><INPUT TYPE='hidden' name='PickID' value='"& PickID & "'>"
if notAvailable <> "yes" then
	
	if Auth(5 , "C") then ' À»  Ê—Êœ/Œ—ÊÃ œ—  «—ÌŒ œ·ŒÊ«Â
%>
	 «—ÌŒ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT dir=ltr TYPE="text" NAME="exitDate" style="width:200px;" value="<%=shamsiToday()%>" onblur="acceptDate(this)"><br><br>
<%
	End if
%>
	 Ê÷ÌÕ«  <TEXTAREA NAME='comments' style="width:200px;" ></TEXTAREA><br><br>
	<INPUT TYPE='submit' name='submit' class=inputBut value=' «ÌÌœ'> &nbsp;&nbsp;
	<INPUT TYPE='submit' name='submit' class=inputBut value='«‰’—«›'>
<%
else
	response.write "»Â œ·Ì· „ÊÃÊœ ‰»Êœ‰ »—ŒÌ «“ «ﬁ·«„ «Ì‰ ÕÊ«·Â ﬁ«»· ﬁ»Ê· ‰Ì” . "
	response.write "<A HREF='default.asp?ed=" & RSS("id") & "'>(«’·«Õ ÕÊ«·Â)</A>"
end if
	response.write "<BR><BR></center></FORM>"
response.end

end if
if Request("act")="" then 
set RSS=Conn.Execute ("SELECT InventoryPickuplists.*, Users.RealName AS RealName FROM InventoryPickuplists INNER JOIN Users ON InventoryPickuplists.GiveTo = Users.ID WHERE (InventoryPickuplists.Status = 'new') ORDER BY InventoryPickuplists.ID" )
 %>
<BR>
<FORM METHOD=POST ACTION="">
<TABLE dir=rtl align=center width=600 cellspacing=0>
<TR bgcolor="eeeeee" >
	<TD align=center colspan=6><B>ÕÊ«·Â Â«Ì Œ—ÊÃ ﬂ«·«</B></TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD style="border-bottom: solid 1pt black" ><INPUT TYPE="radio" NAME="" disabled></TD>
	<TD style="border-bottom: solid 1pt black" ><!A HREF="default.asp?s=1"><SMALL>‘„«—Â </SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black" ><!A HREF="default.asp?s=2"><SMALL> «—ÌŒ ÕÊ«·Â </SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black" ><!A HREF="default.asp?s=3"><SMALL>œ—Ì«›  ﬂ‰‰œÂ</SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black" ><!A HREF="default.asp?s=4"><SMALL>«ﬁ·«„</SMALL></A></TD>
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

set RSF=Conn.Execute ("SELECT * FROM InventoryPickuplistItems WHERE (pickupListID = "& RSS("ID") & ")" )
%>
<TR style="cursor:hand" onclick="this.getElementsByTagName('input')[0].click();this.getElementsByTagName('input')[0].focus();">
	<TD  style="border-bottom: solid 1pt black; border-left: solid 1pt black; border-right: solid 1pt black" ><INPUT TYPE="radio" NAME="PickID" onclick="setColor(this)" VALUE="<%=RSS("id")%>"></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black" ><A HREF="default.asp?show=<%=RSS("id")%>"><%=RSS("id")%></A></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black" dir=l><span dir=ltr><%=RSS("CreationDate")%></span><!-- (”«⁄  <%=RSS("CreationTime")%>)--></small></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black" ><%=RSS("RealName")%></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black" >
	<%
	
	
	response.write "<TABLE width=100% >"

	Do while not RSF.eof
		response.write "<TR><TD>" & RSF("ItemName") & " (" & RSF("qtty") & " " & RSF("unit") & ") "
		if RSF("CustomerHaveInvItem") then
			response.write "<b style='color:red;background-color:white'> «—”«·Ì </b>" 
		end if
		response.write "</TD><TD align=left dir=ltr>" 
		if RSF("Order_ID")<>-1 then
			response.write RSF("Order_ID") & "</TD></TR>"
		end if
	RSF.moveNext
	Loop
	response.write "</TABLE>"
	%>

	</TD>
</TR>
<% 
RSS.moveNext
Loop
%>
</TABLE><br>
<center>
	<INPUT TYPE="submit" Name="Submit" Value="Õ–› ÕÊ«·Â" onclick="return confirm('¬Ì« „ÿ„∆‰« „Ì ŒÊ«ÂÌœ «Ì‰ œ—ŒÊ«”  —« Õ–› ﬂ‰Ìœø')" class="btn" style="width:150px;" tabIndex="14"> 
	<INPUT TYPE="submit" Name="Submit" Value="«‰ Œ«» »—«Ì Œ—ÊÃ" class="btn" style="width:150px;" tabIndex="14">
</center>
</form>
<%
		
	end if
	
%>
<SCRIPT LANGUAGE="JavaScript">
<!--

function setColor(obj)
{
	for(i=0; i<document.all.PickID.length; i++)
		{
		theTR = document.all.PickID[i].parentNode.parentNode
		theTR.setAttribute("bgColor","<%=AppFgColor%>")
		}
	theTR = obj.parentNode.parentNode
	theTR.setAttribute("bgColor","#FFFFFF")
}
//-->
</SCRIPT>

<!--#include file="tah.asp" -->
