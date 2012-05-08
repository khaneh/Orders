<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= "ÒÇÑÔ ãæÌæÏí ÇäÈÇÑ"
SubmenuItem=8
if not Auth(5 , 8) then NotAllowdToViewThisPage()
ResultDate = sqlsafe(Request("ResultDate"))

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------- Report page Hearder
'-----------------------------------------------------------------------------------------------------
oldItemID = request("oldItemID") & ""
itemDetail = request("itemDetail") & ""
if itemDetail="" and oldItemID ="" then
	catItem1 = request("catItem")
	%>
	<BR><BR>
	<FORM METHOD=POST ACTION="invReport.asp">
	<center>
	<div style="width:500px;border:1px solid #ffffff;"  align=center>
	<br>
	<TABLE dir=rtl align=center >
	<TR ><!bgcolor="#DDDDDD" >
		<TD align=left >ÏÓÊå ÈäÏí ßÇáÇ :</td>
		<td>
			<SELECT NAME="catItem" size="1" rem_onchange="document.forms[0].submit()" class=inputBut >
			<option value="-1">ÏÓÊå ÈäÏí ßÇáÇ </option>
			<option value="-1">------------------</option>
			<%
				set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories")
				while not (RS4.eof) %>
					<OPTION value="<%=RS4("ID")%>"<%
					if trim(catItem1) = trim(RS4("ID")) then
					response.write " selected "
					end if
					%>>* <%=RS4("Name")%> </option>
					<%	
					RS4.MoveNext
				wend
				RS4.close
				%>
			<option value="-1">------------------</option>
			<option value="-2" <%
			if catItem1 = "-2" then
			response.write " selected "
			end if
			%>>ßÇáÇåÇí ÇÑÓÇáí ÏíÑÇä</option>
			</SELECT>
		</TD>
	</TR>
	<Tr>
		<TD align=left >ÊÇÑíÎ :</td>
		<Td >
			<INPUT TYPE="text" NAME="ResultDate" Value="<%=ResultDate%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);">
		</Td>
	</Tr>
	<Tr>
		<Td colspan=2 align=left >
			<INPUT TYPE="submit" Value=" äãÇíÔ " >
		</Td>
	</Tr>
	</FORM>
	</table>
	<br>
	</div>
	</center>
	<br>
	<%
		if request("catItem") = "" then
			%>
			<FORM METHOD=POST ACTION=""><BR><center>
				<INPUT TYPE="hidden" name="radif" value="-1">
				<input type="hidden" Name='tmpDlgArg' value=''>
				<input type="hidden" Name='tmpDlgTxt' value=''>
				íÇ ßÏ ßÇáÇ ÑÇ æÇÑÏ ßäíÏ:   &nbsp;&nbsp;&nbsp;<INPUT  dir="LTR"  TYPE="text" NAME="oldItemID" maxlength="10" size="13"   onKeyPress='return mask(this);' onBlur='check(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent">
				&nbsp;&nbsp;<INPUT TYPE="submit" Name="Submit" Value="äãÇíÔ"  style="width:80px;" tabIndex="14">
				</center>
			</FORM>
			<SCRIPT LANGUAGE="JavaScript">
			<!--
			document.all.oldItemID.focus()
			var dialogActive=false;

			function mask(src){ 
				var theKey=event.keyCode;

				if (theKey==13){
					event.keyCode=9
					dialogActive=true
					document.all.tmpDlgArg.value="#"
					document.all.tmpDlgTxt.value="ÌÓÊÌæ ÏÑ ßÇáÇåÇí ÇäÈÇÑ"
					var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
					if (document.all.tmpDlgTxt.value != "") {
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

			<%
		end if

end if


'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Update Item Properties
'-----------------------------------------------------------------------------------------------------

if request("catItem") ="-2" then
	set RSS=Conn.Execute ("SELECT SUM(DERIVEDTBL.sumQtty) AS sumQttys, DERIVEDTBL.AccountID, Accounts.AccountTitle, InventoryItems.Name, InventoryItems.OldItemID, InventoryItems.Unit FROM (SELECT SUM((CONVERT(tinyint, dbo.InventoryLog.IsInput) - .5) * 2 * dbo.InventoryLog.Qtty) AS sumQtty, dbo.InventoryLog.owner AS AccountID, InventoryLog.itemid AS invid FROM dbo.InventoryLog where dbo.InventoryLog.voided=0 GROUP BY dbo.InventoryLog.owner, InventoryLog.itemid) DERIVEDTBL INNER JOIN Accounts ON DERIVEDTBL.AccountID = Accounts.ID INNER JOIN InventoryItems ON DERIVEDTBL.invid = InventoryItems.ID GROUP BY DERIVEDTBL.AccountID, Accounts.AccountTitle, InventoryItems.Name, InventoryItems.OldItemID, InventoryItems.Unit having SUM(DERIVEDTBL.sumQtty) <> 0 ")	
	%><BR>
	<CENTER><B>ÇÑÓÇáí ÏíÑÇä</B></CENTER><BR>
	<TABLE dir=rtl align=center width=90%>
	<TR bgcolor="eeeeee">
		<TD><SMALL>äÇã ßÇáÇ</SMALL></TD>
		<TD align=center><SMALL>ßÏ ßÇáÇ</SMALL></TD>
		<TD align=center><SMALL>ÔãÇÑå ÍÓÇÈ</SMALL></TD>
		<TD><SMALL>ÚäæÇä ÍÓÇÈ</SMALL></TD>
		<TD align=center><SMALL>ÊÚÏÇÏ</SMALL></TD>
		<TD align=center><SMALL>æÇÍÏ</SMALL></TD>
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
	<TR bgcolor="<%=tmpColor%>" height=25>
		<TD><%=RSS("name")%></TD>
		<TD align=center dir=ltr><%=RSS("OldItemID")%></TD>
		<TD align=center dir=ltr><%=RSS("AccountID")%></TD>
		<TD><%=RSS("AccountTitle")%></TD>
		<TD align=center dir=ltr><%=RSS("sumQttys")%></TD>
		<TD align=center><%=RSS("Unit")%></TD>
	</TR>
		  
	<% 
	RSS.moveNext
	Loop
	%>
	</TABLE><br>
	</FORM>
	<%
end if
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------- List Items of a category
'-----------------------------------------------------------------------------------------------------

if request("catItem") >0 then

if (len(ResultDate)>0) Then
	sqlstr = "SELECT ISNULL((SELECT Top 1 sumQtty FROM InventoryLog WHERE (Logdate<='"& ResultDate &"') AND dbo.InventoryLog.ItemId = dbo.InventoryItems.Id ORDER BY Logdate DESC, Id DESC), 0) As Qtty, ISNULL((SELECT Top 1 SumCusQtty FROM InventoryLog WHERE (Logdate<='"& ResultDate &"') AND dbo.InventoryLog.ItemId = dbo.InventoryItems.Id ORDER BY Logdate DESC, Id DESC), 0) As CusQtty, dbo.InventoryItems.ID, dbo.InventoryItems.OldItemID, dbo.InventoryItems.owner, dbo.InventoryItems.Minim, dbo.InventoryItems.Name, dbo.InventoryItems.Unit, dbo.InventoryItems.costingMethod, dbo.InventoryItemCategories.Name AS Expr1 FROM dbo.InventoryItems INNER JOIN dbo.InventoryItemCategoryRelations ON dbo.InventoryItems.ID = dbo.InventoryItemCategoryRelations.Item_ID INNER JOIN dbo.InventoryItemCategories ON dbo.InventoryItemCategoryRelations.Cat_ID = dbo.InventoryItemCategories.ID WHERE (InventoryItemCategories.ID ="& catItem1 & ") and InventoryItems.enabled=1 ORDER BY OldItemID"
else
	sqlstr = "SELECT InventoryItems.ID, InventoryItems.OldItemID, InventoryItems.owner, InventoryItems.Minim, InventoryItems.Name, InventoryItems.Qtty, InventoryItems.CusQtty, InventoryItems.Unit, InventoryItems.costingMethod, InventoryItemCategories.Name AS Expr1 FROM InventoryItems INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID INNER JOIN InventoryItemCategories ON InventoryItemCategoryRelations.Cat_ID = InventoryItemCategories.ID WHERE (InventoryItemCategories.ID ="& catItem1 & ") and InventoryItems.enabled=1 ORDER BY OldItemID"
end if
set RSS=Conn.Execute (sqlstr)	
%>
<TABLE dir=rtl align=center width=600>
<TR bgcolor="eeeeee">
	<TD align=center><SMALL>ßÏ ßÇáÇ</SMALL></A></TD>
	<TD><SMALL>äÇã ßÇáÇ</SMALL></A></TD>
	<TD align=center><SMALL>ÊÚÏÇÏ</SMALL></A></TD>
	<TD align=center><SMALL>ÊÚÏÇÏ ÇÑÓÇáí</SMALL></A></TD>
	<TD align=center><SMALL>ÍÏÇŞá ãæÌæÏí</SMALL></A></TD>
	<TD><SMALL>æÇÍÏ</SMALL></A></TD>
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
<TR bgcolor="<%=tmpColor%>" height=25>
	<TD align=center dir=ltr><%=RSS("OldItemID")%></TD>
	<TD><A HREF="invReport.asp?itemDetail=<%=RSS("ID")%><% if Len(ResultDate)>0 then %>&dateTo=<%=ResultDate%><% end if %>"><%=RSS("Name")%></A></TD>
	<TD align=center dir=ltr><%=RSS("Qtty")%></TD>
	<TD align=center dir=ltr><a style="cursor:hand" onclick="window.open('cusItemDetails.asp?id=<%=RSS("ID")%>&name=<%=RSS("Name")%>','CusItemDetails','width=400,height=300')"><%=RSS("CusQtty")%></a></TD>
	<TD align=center dir=ltr><%=RSS("Minim")%></TD>
	<TD><%=RSS("Unit")%></TD>
</TR>
	  
<% 
RSS.moveNext
Loop
%>
</TABLE><br>
<%
end if
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Details of an item
'-----------------------------------------------------------------------------------------------------
if isNumeric(oldItemID) then
	if CInt(oldItemID)=-1 then 
		itemDetail = request("itemID")
		set rs = Conn.Execute("select * from InventoryItems where id=" & itemDetail)
		oldItemID = rs("oldItemID")
		rs.close
	end if

'response.write oldItemID
	set RS3 = conn.Execute ("SELECT * from InventoryItems WHERE (InventoryItems.OldItemID = "&  Clng(oldItemID) & ")")
	if not RS3.EOF then
		itemDetail = RS3("id")
	end if
end if 
if isNumeric(itemDetail) then
	sqlstr = "SELECT InventoryItemCategories.id as catID, InventoryItems.ID, InventoryItems.outByOrder, InventoryItems.OldItemID, InventoryItems.owner, InventoryItems.Name, InventoryItems.Minim, InventoryItems.Qtty, InventoryItems.CusQtty, InventoryItems.Unit, InventoryItems.costingMethod FROM InventoryItems INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID INNER JOIN InventoryItemCategories ON InventoryItemCategoryRelations.Cat_ID = InventoryItemCategories.ID WHERE (InventoryItems.ID = "&  Clng(itemDetail) & ")"
	set RS3 = conn.Execute (sqlstr)
	if RS3.EOF then
		response.write "<center><br><br>ÎØÇ!<br><br>äíä íÒí ÏÑ ÇäÈÇÑ äÏÇÑíã</center>"
		response.end
	end if
	%>
	<BR>

	<TABLE border=0 align=center>
	<TR>
		<TD colspan=2 align=center><H3>ßÇÑÊ ÇäÈÇÑ</H3></TD>
	</TR>
	<TR>
		<TD align=left>ßÏ ßÇáÇ</TD>
		<TD align=right><INPUT TYPE="text" NAME="OldItemID" value="<%=RS3("OldItemID")%>" readonly></TD>
	</TR>
	<TR>
		<TD align=left>äæÚ ßÇáÇ</TD>
		<TD align=right>
			<%
				set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories where id=" & RS3("catID"))
				if not (RS4.eof) then %>
					<INPUT TYPE="text" NAME="ItemName" value="<%=RS4("Name")%>" size=64 readonly> 
					<%	
				end if
				RS4.close
				%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD align=left>äÇã ßÇáÇ</TD>
		<TD align=right><INPUT TYPE="text" NAME="ItemName" value="<%=RS3("Name")%>" size=64 readonly></TD>
	</TR>
	</table>
	<TABLE border=0 align=center>
	<TR>
		<TD>
			<TABLE> 
			<TR>
				<TD align=left>ÊÚÏÇÏ İÚáí</TD>
				<TD align=right><INPUT TYPE="text" NAME="qtty"  value="<%=(clng(RS3("Qtty")*100))/100%>" size=7 readonly></TD>
			</TR>
			<TR>
				<TD align=left>ÊÚÏÇÏ ÇÑÓÇáí</TD>
				<TD align=right><INPUT TYPE="text" NAME="cusQtty"  value="<%=(clng(RS3("cusQtty")*100))/100%>" size=7 readonly> <a style="cursor:hand" onclick="window.open('cusItemDetails.asp?id=<%=RS3("ID")%>&name=<%=RS3("Name")%>','CusItemDetails','width=400,height=300')">...</a> </TD>
			</TR>
			<TR>
				<TD align=left></TD>
				<TD align=right height=20>
				<BR>
				</TD>
			</TR>
			</TABLE> 
		</TD>
		<TD>
			<TABLE> 
			<TR>
				<TD align=left>ÍÏÇŞá ãæÌæÏí</TD>
				<TD align=right><INPUT TYPE="text" NAME="Minim" value="<%=RS3("Minim")%>" size=7 readonly></TD>
			</TR>
			<TR>
				<TD align=left>æÇÍÏ ÇäÏÇÒå íÑí</TD>
				<TD align=right><INPUT TYPE="text" NAME="Unit" value="<%=RS3("Unit")%>" size=7 readonly></TD>
			</TR>
			<TR>
				<TD align=left></TD>
				<TD align=right height=20>
				<BR>
				</TD>
			</TR>
			<!--TR>
				<TD align=left>&nbsp;&nbsp;&nbsp;&nbsp;äÍæå ÇÑÒíÇÈí</TD>
				<TD align=right><INPUT TYPE="radio" disabled NAME="cost" <%
				if RS3("costingMethod") = "AVRG" then
					response.write " checked"
				end if
				%> value="AVRG">ãÊæÓØ <br>
				<INPUT TYPE="radio" NAME="cost" disabled <%
				if RS3("costingMethod")  = "LIFO" then
					response.write " checked"
				end if
				%> value="LIFO">LIFO <br>
				<INPUT TYPE="radio" NAME="cost" disabled <%
				if RS3("costingMethod") = "FIFO" then
					response.write " checked"
				end if
				%> value="FIFO">FIFO</TD>
			</TR-->
			</TABLE> 
		</TD>
	</TR>
	</TABLE>
	<CENTER>
	<INPUT TYPE="checkbox" NAME="outByOrder" <% if RS3("outByOrder") then%> disabled checked <% end if %>>ÎÑæÌ ÈÑ ÇÓÇÓ ÓİÇÑÔ<BR><BR><BR>

	<INPUT TYPE="submit" value="ÈÑÔÊ Èå áíÓÊ ßÇáÇåÇ" onclick="window.location='?catItem=<%=RS3("catID")%>'">
	<INPUT TYPE="button" value="Ç Çíä ÕİÍå" onclick="window.open('myreport.asp?dateto=<%=request("dateto")%>&datefrom=<%=request("datefrom")%>&itemDetail=<%=itemDetail%>','_blank');"></CENTER>
	<%
	dateFrom = request("dateFrom")
	dateTo = request("dateTo")

	if dateTo="" then dateTo=shamsiToday()
	if dateFrom="" then  dateFrom = left(shamsiToday(),4) & "/01/01"
	'if dateFrom="" then dateFrom=session("FiscalYear")&"/01/01"
	Set HRS = Conn.Execute("SELECT ISNULL(dbo.CalcSumQttyInv("&  itemDetail & ",'"&dateTo&"',0),0) As OurSumQtty,ISNULL(dbo.CalcSumQttyInv("&  itemDetail & ",'"&dateTo&"',1),0) As TheirSumQtty")
	if Not HRS.EOF Then
		mysumQtty = cdbl(HRS("OurSumQtty"))
		ursumQtty = cdbl(HRS("TheirSumQtty"))
	End if 
	HRs.Close

	sqlstr = "SELECT InventoryLog.type, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.comments, InventoryLog.VoidedDate, InventoryLog.IsInput, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.ID, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName, InventoryLog.price FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID WHERE (InventoryItems.ID = "&  itemDetail & ") and InventoryLog.logDate>= N'"&dateFrom&"' and InventoryLog.logDate<= N'"&dateTo&"' ORDER BY InventoryLog.Logdate DESC, InventoryLog.ID DESC"
	set RSS=Conn.Execute (sqlstr)	

	%><BR><BR>
	<table align=center width=95%>
	<TR bgcolor="eeeeee" >
		<FORM METHOD=POST ACTION="">
		<TD colspan=10>
		<INPUT TYPE="hidden" NAME="OldItemID" value="<%=RS3("OldItemID")%>"><H4>
		ÓÇÈŞå æÑæÏ æ ÎÑæÌ ßÇáÇ  (ÇÒ ÊÇÑíÎ <INPUT TYPE="text" NAME="dateFrom" style="border:0pt" size=10 value="<%=dateFrom%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" > ÊÇ ÊÇÑíÎ  <INPUT TYPE="text" NAME="dateTo" style="border:0pt" size=10 value="<%=dateTo%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" >) <INPUT TYPE="submit" value="äãÇíÔ">
		</H4></TD>
		</FORM>
	</TR>
	<TR bgcolor="eeeeee" >
		<td align="center"><small>ÑÏíİ</small></td>
		<TD align="center"><SMALL>æÑæÏ</SMALL></A></TD>
		<TD align="center"><SMALL>ÎÑæÌ</SMALL></A></TD>
		<td align="center"><small>ÈåÇ</small></td>
		<TD align="center"><SMALL>ÊÑÇÒ ßá</SMALL></A></TD>
		<TD align="center"><SMALL>ÊÑÇÒ ßá<br>ÇÑÓÇáí</SMALL></A></TD>
		<TD align="center"><SMALL>æÇÍÏ</SMALL></A></TD>
		<TD align=center><SMALL>ÊÇÑíÎ </SMALL></A></TD>
		<TD align="center"><SMALL>ÊæÖíÍÇÊ</SMALL></A></TD>
		<TD align="center"><SMALL>ÊæÓØ</SMALL></A></TD>
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
		if trim(request("logRowID"))=trim(RSS("id")) then
			tmpColor = "yellow"
		end if
	%>
	<TR bgcolor="<%=tmpColor%>"  style="height:25pt" <% if RSS("voided") then%> disabled title="ÍÏİ ÔÏå ÏÑ ÊÇÑíÎ <%=RSS("VoidedDate")%>"<% end if %>>
		<td align=right dir=ltr><span style="font-size:10pt"><%=RSS("id")%></span></td>
		<TD align=right dir=ltr><span style="font-size:10pt"><% if RSS("IsInput") then %><%=RSS("Qtty")%><% end if %></span></TD>
		<TD align=right dir=ltr style="position:relative;"><% if RSS("voided") then%><div style="right:0px;position:absolute;width:520;"><hr style="color:red;"></div><% end if %><span style="font-size:10pt"><% if not RSS("IsInput") then %><%=RSS("Qtty")%><% end if %></span></TD>
		<td align=right dir=ltr title="İí: <%if not isNull(rss("price")) then response.write cdbl(rss("price"))/cdbl(rss("qtty"))%>"><span style="font-size:10pt"><%=rss("price")%></span></td>
		<TD align=right dir=ltr><span style="font-size:10pt"><%=mySumQtty%></span></TD>
		<TD align=right dir=ltr><span style="font-size:10pt"><%=urSumQtty%></span></TD>
		<TD align=right dir=ltr ><%=RSS("Unit")%></TD>
		<TD dir=ltr align=center><%=RSS("logDate")%></span></TD>
		<TD><% if RSS("type")= "2" then
					response.write "<font color=red><b>ÇÕáÇÍ ãæÌæÏí</b></font>"
				elseif RSS("type")= "3" then
					response.write "<font color=green><b>ãÑÌæÚí</b></font>"
				elseif RSS("type")= "4" then
					response.write "<font color=blue><b>ÊÚÑíİ ßÇáÇí ÌÏíÏ </b></font>"
				elseif RSS("type")= "5" then
					response.write "<font color=orang><b>ÇäÊŞÇá</b></font>"
				elseif RSS("type")= "6" then
					response.write "<font color=#6699CC><b>æÑæÏ ÇÒ ÊæáíÏ</b></font>"
				elseif RSS("type")= "7" then
					response.write "<font color=#FF9966><b>æÑæÏ ÇÒ ÇäÈÇÑ ÔåÑíÇÑ</b></font>"
				elseif RSS("type")="9" then
					response.write "<font color=#AAAA00><b>æÑæÏ ÖÇíÚÇÊ</b></font>"
				elseif RSS("RelatedID")= "-1" then %> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			   <% else 
					if RSS("IsInput") then 
						%> ÔãÇÑå ÓİÇÑÔ ÎÑíÏ: <A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
					else
						%>ÔãÇÑå ÍæÇáå ÎÑæÌ: <A HREF="default.asp?ed=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
					end if
				end if %>
			<% if RSS("owner")<> "-1" and RSS("owner")<> "-2" then
				response.write " (ÇÑÓÇáí <a href='../CRM/AccountInfo.asp?act=show&selectedCustomer="& RSS("owner") &"' target='_blank'> " & RSS("owner") & "</a>)"
			   end if %>
			<% if trim(RSS("comments"))<> "-" and RSS("comments")<> "" then
				response.write " <br><br><B>ÊæÖíÍ:</B>  " & RSS("comments") 
			   end if 
			   set rs=Conn.Execute("select logDate,realName,price from InventoryPriceLog inner join users on InventoryPriceLog.userID = users.id where logID=" & rss("id") & " group by logDate,realName,price having count(price)>1")
			   if not rs.eof then response.write("<br><b>ÊÛííÑÇÊ ŞíãÊ:</b><br>")
			   while not rs.eof
			   	response.write("<small>ÏÑ ÊÇÑíÎ " & rs("logDate") & "¡ ÊæÓØ " & rs("realName") & " " & rs("price") & " ÑíÇá" &"</small>")
			   	rs.moveNext
			   wend
			   rs.close
			   set rs=nothing
			   if CDbl(rss("qtty"))>0 and not rss("voided") then 
				   if rss("isInput") then 
				   	mySQL = "select * from InventoryFIFORelations where inID=" & rss("id")
				   	msg1="áíÓÊ ÎÑæÌåÇ ÇÒ Çíä æÑæÏ:"
				   	msg2="åäæÒ ÇÒ Çíä æÑæÏ ÎÑæÌ ÇäÌÇã äÔÏå"
				   else
				   	mySQL = "select * from InventoryFIFORelations where outID=" & rss("id")
				 	msg1="áíÓÊ æÑæÏåÇ ÈÑÇí Çíä ÎÑæÌ:"
				 	msg2="Çíä ÎÑæÌ ŞíãÊ ĞÇÑí äÔÏå! áØİÇ æÑæÏåÇí ŞÈáí ÑÇ ŞíãÊ ĞÇÑí ßäíÏ"
				   end if
				   set rs = Conn.Execute(mySQL)
				   if rs.eof then 
				   	response.write ("<br><b>" & msg2&"</b><br>")
				   else
				   	response.write ("<br><b>" & msg1&"</b><br>")
				   end if
				   while not rs.eof
				   	if rss("isInput") then 
				   		link = rs("outID")
				   	else
				   		link = rs("inID")
				   	end if
				   	response.write ("<small>" & rs("qtty") & " ÚÏÏ ÇÒ " & link & "</small><br>")
				   	rs.moveNext
				   wend
				   rs.close
				   set rs=nothing
			   end if
			   %>
		</TD>
		<TD><%=RSS("RealName")%></TD>
		<%
		if Not RSS("Voided") Then
			if RSS("IsInput") Then
				if Rss("owner")=-1 then
					mySumQtty = mySumQtty - cdbl(RSS("Qtty"))
				else
					urSumQtty = urSumQtty - cdbl(RSS("Qtty"))
				end if 
			else
				if Rss("owner")=-1 then
					mySumQtty = mySumQtty + cdbl(RSS("Qtty"))
				else
					urSumQtty = urSumQtty + cdbl(RSS("Qtty"))
				end if 
			end if 
		end if

		%>
	</TR>
	  
	<% 
		if not RSS("voided") then

			if RSS("IsInput") then
				if Rss("owner")=-1 then
					inputs = inputs + cdbl(RSS("Qtty"))
				else
					urinputs = urinputs + cdbl(RSS("Qtty"))
				end if 
			else
				if Rss("owner")=-1 then
					outputs = outputs + cdbl(RSS("Qtty"))
				else
					uroutputs = uroutputs + cdbl(RSS("Qtty"))
				end if 
			end if
		end if
	RSS.moveNext
	Loop
	%>
	<TR bgcolor="666666" height=25 style="color:white;">
		<td></td>
		<TD align=right dir=ltr><%=inputs+urinputs%></A></TD>
		<TD align=right dir=ltr><%=outputs+uroutputs%></A></TD>
		<TD colspan=7>ÑÏÔ æÑæÏ æ ÎÑæÌ  ÇÒ ÊÇÑíÎ <span dir=ltr><%=dateFrom%></span> ÊÇ ÊÇÑíÎ <span dir=ltr><%=dateTo%></span></A></TD>
	</TR>
	<TR bgcolor="cccccc" height=25 style="color:black;">
		<td></td>
		<TD align=right dir=ltr><%=inputs%></A></TD>
		<TD align=right dir=ltr><%=outputs%></A></TD>
		<TD colspan=7>ÑÏÔ æÑæÏ æ ÎÑæÌ  ÇÒ ÊÇÑíÎ <span dir=ltr><%=dateFrom%></span> ÊÇ ÊÇÑíÎ <span dir=ltr><%=dateTo%></span> ÎæÏãÇä</A></TD>
	</TR>
	<TR bgcolor="cccccc" height=25 style="color:black;">
		<td></td>
		<TD align=right dir=ltr><%=urinputs%></A></TD>
		<TD align=right dir=ltr><%=uroutputs%></A></TD>
		<TD colspan=7>ÑÏÔ æÑæÏ æ ÎÑæÌ  ÇÒ ÊÇÑíÎ <span dir=ltr><%=dateFrom%></span> ÊÇ ÊÇÑíÎ <span dir=ltr><%=dateTo%> </span> ÏíÑÇä</A></TD>
	</TR>
	<% 
		sqlstr = "SELECT ROUND(ISNULL((SELECT Sum(Qtty) From InventoryLog Where (Voided=0) And (ItemID = "&  itemDetail & ") And Logdate<'"&dateFrom&"' And isInput=1),0),2) As InputQtty, ROUND(ISNULL((SELECT Sum(Qtty) From InventoryLog Where (Voided=0) And (ItemID = "&  itemDetail & ") And Logdate<'"&dateFrom&"' And isInput=0 ),0), 2) As OutputQtty"
		set DRS1 = conn.Execute(sqlstr)
		if (not DRS1.EOF) then
		%>
	<Tr><td colspan=11 Height=20 ></td></tr>
	<TR bgcolor="#FFFFFF"  style="height:25pt;" >
		<td></td>
		<TD align=right dir=ltr ><span style="font-size:10pt"><%=DRS1("InputQtty")%></span></TD>
		<TD align=right dir=ltr ><span style="font-size:10pt"><%=DRS1("OutputQtty")%></span></TD>
		<TD align=right dir=ltr ><span style="font-size:10pt"><%=cdbl(DRS1("InputQtty")) - cdbl(DRS1("OutputQtty")) %></span></TD>
		<TD align=right Colspan=6  >
		ãæÌæÏ æ ÑÏÔ ÊÇ ŞÈá ÇÒ ÊÇÑíÎ <span dir=ltr><%=dateFrom%></span>
		</TD>
	</TR>
		<%
		End if
		DRS1.Close
		%>
	</table><BR><BR>
<%
end if
%>
<!--#include file="tah.asp" -->
