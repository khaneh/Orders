<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= " ê“«—‘ Ê—Êœ Ê Œ—ÊÃ »Â «‰»«—"
SubmenuItem=5
if not Auth(5 , 5) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
fromDate = request("fromDate") 
toDate = request("toDate") 
InRep = request("InRep") 
outRep = request("outRep") 

if inRep="" then
	InRepSt = " "
else
	InRepSt = "checked"
end if

if outRep="" then
	outRepSt = " "
else
	outRepSt = "checked"
end if

if toDate="" or fromDate="" then
	toDate = shamsiToday()
	fromDate = shamsiDate(Date())
	InRepSt = "checked"
	outRepSt = "checked"
	flag = "first"
end if


'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------- Delete Log Line
'-----------------------------------------------------------------------------------------------------
if request("act") = "del" and isNumeric(request("rowID")) then
	rowID = clng(request("rowID"))
	desc = request("desc") ' Description : Deletion Reason
	if "" & desc = "" then
		desc = "-"
	else
		desc = sqlSafe(desc)
	end if

	set RSS=Conn.Execute ("SELECT * from InventoryLog where (ID = " & rowID & ") and voided = 0")
	if RSS.EOF then
		response.write "<br><br>" 
		call showAlert ("Œÿ«! ç‰Ì‰ ”ÿ—Ì „ÊÃÊœ ‰Ì”  Ì« ﬁ»·« Õ–› ‘œÂ «” ." , CONST_MSG_ERROR)
	else
		if rss("gl_update")="1" then 
			response.write "<br><br>" 
			call showAlert ("Œÿ«! »—«Ì «Ì‰ ”ÿ— ﬁ»·« ”‰œ Õ”«»œ«—Ì ’«œ— ‘œÂ° ›·–« ›⁄·« ﬁ«œ— »Â Õ–› ¬‰ ‰Ì” Ìœ!." , CONST_MSG_ERROR)
		else
			type1 = RSS("type")
			RelatedID = RSS("RelatedID")
			isInput = RSS("isInput")
	
			if type1 = 1 and not isInput then
				Conn.Execute ("UPDATE InventoryPickuplists SET Status = 'new' WHERE (id = " & RelatedID & ")")
				response.write "<br><br>" 
				call showAlert ("ÕÊ«·Â Œ—Ìœ „—»ÊÿÂ »Â ’› ¬„«œÂ Œ—ÊÃ »«“ê‘ ." , CONST_MSG_INFORM)
			end if
	
			if (not isInput) and (RelatedID > 0) then
				set RSB=Conn.Execute ("SELECT id from InventoryLog where (RelatedID = " & RelatedID & " and isInput=0) and voided = 0 ")
				Do while not RSB.eof
					Conn.Execute ("UPDATE InventoryLog SET Voided =1, VoidedBy =" & session("id") & ", VoidedDate =N'" & shamsiToday() & "', comments = LEFT(comments + '<br>[œ·Ì· Õ–›: " & desc & "]', 250) WHERE (ID = " & RSB("id") & ")")
					RSB.movenext
				loop
			else
				Conn.Execute ("UPDATE InventoryLog SET Voided =1, VoidedBy =" & session("id") & ", VoidedDate =N'" & shamsiToday() & "', comments = LEFT(comments + '<br>[œ·Ì· Õ–›: " & desc & "]', 250) WHERE (ID = " & rowID & ") AND (voided = 0)")
			end if 
	
			response.write "<br><br>" 
			call showAlert ("Õ–› ”ÿ— «‰Ã«„ ‘œ." , CONST_MSG_INFORM)
			Conn.Execute("update InventoryLog set price=null, pricedQtty=null where isInput=0 and owner=-1 and voided=0 and itemID=" & rss("itemID"))
			Conn.Execute("update InventoryLog set pricedQtty = qtty where isInput=1 and owner=-1 and voided=0 and itemID=" & rss("itemID"))
			Conn.Execute("delete from InventoryFIFORelations where inID in (select id from InventoryLog where isInput=1 and owner=-1 and voided=0 and itemID=" & rss("itemID")) &")"
			Conn.Execute("delete from InventoryFIFORelations where outID in (select id from InventoryLog where isInput=0 and owner=-1 and voided=0 and itemID=" & rss("itemID")) &")"
			'Conn.Execute("execute dbo.outFIFO")
		end if
	end if

end if
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- Inventory Transaction Log Form
'-----------------------------------------------------------------------------------------------------
%>
<BR><BR>
<FORM METHOD=POST ACTION="report.asp">
<TABLE border=0 align=center>
<TR>
	<TD colspan=5 align=center><H3>ê“«—‘ ⁄„·Ì«  «‰»«—</H3></TD>
</TR>
<TR>
	<TD align=left>«“  «—ÌŒ</TD>
	<TD align=right><INPUT TYPE="text" NAME="fromDate" value="<%=fromDate%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
	<TD align=left width=20></TD>
	<TD align=left> «  «—ÌŒ</TD>
	<TD align=right><INPUT TYPE="text" NAME="toDate" value="<%=toDate%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
</TR>

<TR height=10>
	<TD colspan=5 align=center></TD>
</TR>

<TR>
	<TD align=left></TD>
	<TD align=right><INPUT TYPE="checkbox" NAME="outRep" <%=outRepSt%>> Œ—ÊÃ ﬂ«·«</TD>
	<TD align=left width=20><INPUT TYPE="checkbox" NAME="InRep" <%=inRepSt%>></TD>
	<TD align=left>Ê—Êœ ﬂ«·«</TD>
	<TD align=left><INPUT TYPE="submit" NAME="submit" class=inputBut value="„‘«ÂœÂ"></TD>
</TR>

</TABLE>
</FORM>
<%
if flag = "first" then
	response.end
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------- Transaction Log
'-----------------------------------------------------------------------------------------------------

if request("submit")="„‘«ÂœÂ"then

	%>
	<TABLE dir=rtl align=center width=95%>
	<%

	if outRep = "on" then
'		mySQL="SELECT InventoryLog.ID, InventoryLog.comments, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.VoidedDate, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.type, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName, Users_1.RealName AS GiveTo FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID INNER JOIN InventoryPickuplists ON InventoryLog.RelatedID = InventoryPickuplists.id INNER JOIN Users Users_1 ON InventoryPickuplists.GiveTo = Users_1.ID WHERE (InventoryLog.logDate >= N'"& fromDate & "' and InventoryLog.logDate <= N'"& toDate & "' and IsInput=0) ORDER BY InventoryLog.ID DESC"

		mySQL="SELECT InventoryLog.ID, InventoryLog.comments, InventoryLog.Voided, InventoryLog.VoidedDate, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.type, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName, Users_1.RealName AS GiveTo, Users_2.RealName AS VoidedBy, InventoryFIFORelations.inID,InventoryFIFORelations.qtty as inQtty FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID left outer JOIN InventoryPickuplists ON InventoryLog.RelatedID = InventoryPickuplists.id and InventoryLog.RelatedID>0 left outer JOIN Users Users_1 ON InventoryPickuplists.GiveTo = Users_1.ID LEFT OUTER JOIN Users Users_2 ON InventoryLog.VoidedBy = Users_2.ID left outer join InventoryFIFORelations on inventoryLog.id = InventoryFIFORelations.outID WHERE (InventoryLog.logDate >= N'"& fromDate & "') AND (InventoryLog.logDate <= N'"& toDate & "') AND (InventoryLog.IsInput = 0) ORDER BY InventoryLog.ID DESC"
'response.write mySQL
		set RSS=Conn.Execute (mySQL)	
		%>
		<TR bgcolor="eeeeee" >
			<TD colspan=9><H4>Œ—ÊÃ ﬂ«·«</H4></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD>Õ–›</TD>
			<TD>ﬂœﬂ«·«</TD>
			<TD width=200>‰«„ ﬂ«·«</TD>
			<TD> ⁄œ«œ </TD>
			<TD>Ê«Õœ</TD>
			<TD> «—ÌŒ Œ—ÊÃ</TD>
			<TD>êÌ—‰œÂ</TD>
			<TD align=center>‘„«—Â ÕÊ«·Â</TD>
			<TD> Ê”ÿ</TD>
		</TR>
		<%
		tmpCounter=0
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

		%>
		<TR bgcolor="<%=tmpColor%>" style="height:25pt" <% if RSS("voided") then%> disabled title="Õœ› ‘œÂ œ—  «—ÌŒ <%=RSS("VoidedDate")%>  Ê”ÿ <%=RSS("VoidedBy")%>"<% end if %>>
			<TD align=center dir=ltr><small disabled><%=RSS("id")%></small><% if not RSS("voided") then%><A HREF="javascript:confirmDelete(<%=RSS("ID")%>);">X</A> <% end if %></TD>
			<TD align=right dir=ltr><INPUT TYPE="hidden" name="id" value="<%=RSS("ID")%>"><A HREF="invReport.asp?oldItemID=<%=RSS("oldItemID")%>&logRowID=<%=RSS("ID")%>" target="_blank"><%=RSS("OldItemID")%></A></TD>
			<TD><% if RSS("voided") then%><div style="position:absolute;width:520;"><hr style="color:red;"></div><% end if %><!A HREF="default.asp?itemDetail=<%=RSS("ID")%>"><span style="font-size:10pt"><%=RSS("Name")%></A></TD>
			<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("Qtty")%></span></TD>
			<TD align=right dir=ltr><%=RSS("Unit")%></TD>
			<TD dir=ltr><%=RSS("logDate")%></span></TD>
			<TD><%=RSS("GiveTo")%></TD>
			<TD align=center><% 
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
					response.write " <br><br><B> Ê÷ÌÕ:</B> " & RSS("comments") 
				 end if 
				 if not isNull(rss("inID")) then 
				 	response.write("<br>ﬁÌ„  ê–«—Ì ‘œÂ°" & rss("inQtty") & " «“ Ê—Êœ <b>" & rss("inID") & "</b>")
				 else
				 	response.write("<br><b>ﬁÌ„  ê–«—Ì ‰‘œÂ!</b>")
				 end if
				 %>

			</TD>
			<TD><%=RSS("RealName")%></TD>
		</TR>
			 
		<% 
		RSS.moveNext
		Loop
	end if

	if inRep = "on" then

		mySQL="SELECT InventoryLog.RelatedInvoiceID, InventoryLog.ID, InventoryLog.comments, InventoryLog.Voided, InventoryLog.VoidedDate, InventoryLog.CreatedBy, InventoryLog.owner, InventoryLog.logDate, InventoryLog.type, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryItems.OldItemID, InventoryItems.Name, InventoryItems.Unit, Users.RealName, Users_1.RealName AS VoidedBy, InventoryFIFORelations.outID, InventoryFIFORelations.qtty as outQtty FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID LEFT OUTER JOIN Users Users_1 ON InventoryLog.VoidedBy = Users_1.ID left outer join InventoryFIFORelations on inventoryLog.id = InventoryFIFORelations.inID WHERE (InventoryLog.logDate >= N'"& fromDate & "') AND (InventoryLog.logDate <= N'"& toDate & "') AND (InventoryLog.IsInput = 1) ORDER BY InventoryLog.ID DESC"

		set RSS=Conn.Execute (mySQL)	

		%>
		<TR height=50>
			<TD colspan=9></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD colspan=9><H4>Ê—Êœ ﬂ«·«</H4></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD> Õ–› </TD>
			<TD>ﬂœﬂ«·«</TD>
			<TD width=200>‰«„ ﬂ«·«</TD>
			<TD> ⁄œ«œ </TD>
			<TD>Ê«Õœ</TD>
			<TD> «—ÌŒ Ê—Êœ</TD>
			<TD colspan=2 align=center> Ê÷ÌÕ« </TD>
			<TD> Ê”ÿ</TD>
		</TR>
		<%
		tmpCounter=0
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

		%>
		<TR bgcolor="<%=tmpColor%>" style="height:25pt" <% if RSS("voided") then%> disabled title="Õœ› ‘œÂ œ—  «—ÌŒ <%=RSS("VoidedDate")%>  Ê”ÿ <%=RSS("VoidedBy")%>"<% end if %>>
			<TD align=center dir=ltr><% if not RSS("voided") then%><A HREF="javascript:confirmDelete(<%=RSS("ID")%>);">X</A> <small disabled><%=RSS("id")%></small><% end if %></TD>
			<TD align=right dir=ltr><INPUT TYPE="hidden" name="id" value="<%=RSS("ID")%>"><A HREF="invReport.asp?oldItemID=<%=RSS("oldItemID")%>&logRowID=<%=RSS("ID")%>" target="_blank"><%=RSS("OldItemID")%></A</TD>
			<TD><% if RSS("voided") then%><div style="position:absolute;width:520;"><hr style="color:red;"></div><% end if %><!A HREF="default.asp?itemDetail=<%=RSS("ID")%>"><span style="font-size:10pt"><%=RSS("Name")%></A></TD>
			<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("Qtty")%></span></TD>
			<TD align=right dir=ltr><%=RSS("Unit")%></TD>
			<TD dir=ltr><%=RSS("logDate")%></span></TD>
			<TD colspan=2 align=center><% if RSS("type")= "1" and RSS("RelatedID")<> "-1" then
					%>
						”›«—‘ Œ—Ìœ 
				 <A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A>
				 <%
				 elseif RSS("type")= "2" then
					response.write "<font color=red><b>«’·«Õ „ÊÃÊœÌ</b></font>"
				 elseif RSS("type")= "3" then
					response.write "<font color=green><b>„—ÃÊ⁄Ì</b></font> (‘„«—Â ›«ﬂ Ê— :<font color=green> "& RSS("RelatedInvoiceID") & "</font>)"
					elseif RSS("type")= "4" then
						response.write "<font color=blue><b> ⁄—Ì› ﬂ«·«Ì ÃœÌœ </b></font>"
					elseif RSS("type")= "5" then
						response.write "<font color=orang><b>«‰ ﬁ«·</b></font>"
					elseif RSS("type")= "6" then
						response.write "<font color=#6699CC><b>Ê—Êœ «“  Ê·Ìœ</b></font>"
					elseif RSS("type")= "7" then
						response.write "<font color=#FF9966><b>Ê—Êœ «“ «‰»«— ‘Â—Ì«—</b></font>"
					elseif RSS("type")="9" then 
						response.write "<font color=#AAAA00><b>Ê—Êœ ÷«Ì⁄« </b></font>"
				 else 
					response.write " "
				 end if	%>
				 
				<% if RSS("owner")<> "-1" and RSS("owner")<> "-2" then
					response.write " («—”«·Ì <a href='../CRM/AccountInfo.asp?act=show&selectedCustomer="& RSS("owner") &"' target='_blank'> " & RSS("owner") & "</a>)"
				 end if %>
				<% if RSS("comments")<> "-" and RSS("comments")<> "" then
					response.write " <br><br><B> Ê÷ÌÕ:</B> " & RSS("comments") 
				 end if 
				 if not isNull(rss("outID")) then 
				 	response.write("<br>ﬁÌ„  ê–«—Ì ‘œÂ°" & rss("outQtty") & " «“ Œ—Êç<b> " & rss("outID") & "</b>")
				 else
				 	response.write("<br><b>ﬁÌ„  ê–«—Ì ‰‘œÂ!</b>")
				 end if
				 %>
				 </TD>
			<TD dir='LTR' align='right'><%=RSS("RealName")%></TD>
		</TR>
			 
<% 
		RSS.moveNext
		Loop
	end if
%>
	</TABLE><br>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function confirmDelete(rowID){
		if (confirm('¬Ì« „ÿ„∆‰ Â” Ìœø')){

			var desc='';
			var tmpDlgTxt= new Object();
			tmpDlgTxt.value = ' Ê÷ÌÕÌ œ— «Ì‰ —«»ÿÂ: ';

			dialogActive=true
			window.showModalDialog('../dialog_GenInput.asp',tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false
			desc=tmpDlgTxt.value
			window.location = "?act=del&rowID=" + rowID + "&desc=" + escape(desc);
		}
	}
	//-->
	</SCRIPT>
<%

end if
%>

<!--#include file="tah.asp" -->