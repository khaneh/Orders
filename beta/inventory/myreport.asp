<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<TITLE> گزارش انبار</TITLE>
<style>
	body { font-family: tahoma; font-size: 8pt;}
	body A { Text-Decoration : none ;}
	Input { font-family: tahoma; font-size: 9pt;}
	td { font-family: tahoma; font-size: 8pt;}
	.tt { font-family: tahoma; font-size: 10pt; color:yellow;}
	.tt2 { font-family: tahoma; font-size: 8pt; color:yellow;}
	.inputBut { font-family: tahoma; font-size: 8pt; richness: 10}
	.t7pt { font-size: 8pt;}
	.t8pt { font-size: 10pt;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
</HEAD>
<BODY topmargin=0 leftmargin=0 align=center dir=rtl >

	<%
	dateFrom = request("dateFrom")
	dateTo = request("dateTo")
	itemDetail = request("itemDetail") & ""
	sqlstr = "SELECT InventoryItemCategories.id as catID, InventoryItems.ID, InventoryItems.outByOrder, InventoryItems.OldItemID, InventoryItems.owner, InventoryItems.Name, InventoryItems.Minim, InventoryItems.Qtty, InventoryItems.CusQtty, InventoryItems.Unit, InventoryItems.costingMethod FROM InventoryItems INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID INNER JOIN InventoryItemCategories ON InventoryItemCategoryRelations.Cat_ID = InventoryItemCategories.ID WHERE (InventoryItems.ID = "&  Clng(itemDetail) & ")"
	set RS3 = conn.Execute (sqlstr)
	if RS3.EOF then
		response.write "<center><br><br>خطا!<br><br>چنين چيزي در انبار نداريم</center>"
		response.end
	end if
	%>
	<BR>

	<center><H3>كارت انبار </H3></center>
	<TABLE border=0 align=center style="border:2px dotted #000000;" >
	<TR>
		<TD align=left><B>كد كالا :</b></TD>
		<TD align=right><%=RS3("OldItemID")%></TD>
	</TR>
	<TR>
		<TD align=left width=100 ><B>نوع كالا :</B></TD>
		<TD align=right>
			<%
				set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories where id=" & RS3("catID"))
				if not (RS4.eof) then %>
					<%=RS4("Name")%>
					<%	
				end if
				RS4.close
			%>
		</TD>
	</TR>
	<TR>
		<TD align=left width=100 ><B>نام كالا :</B></TD>
		<TD align=right width=400 ><%=RS3("Name")%></TD>
	</TR>
	</table>
	<br>
	<TABLE border=0 align=center style="border:2px dotted #000000;" >
	<TR>
		<TD>
			<TABLE> 
			<TR>
				<TD align=left width=100 ><B>تعداد فعلي : </B></TD>
				<TD align=right width=100 ><%=RS3("Qtty")%></TD>
			</TR>
			<TR>
				<TD align=left><B>تعداد ارسالي :</B></TD>
				<TD align=right><%=RS3("cusQtty")%></TD>
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
				<TD align=left width=100 ><B>حداقل موجودي :</B></TD>
				<TD align=right width=100 ><%=RS3("Minim")%></TD>
			</TR>
			<TR>
				<TD align=left><B>واحد اندازه گيري : </B></TD>
				<TD align=right><%=RS3("Unit")%></TD>
			</TR>
			<TR>
				<TD align=left></TD>
				<TD align=right height=20>
				<BR>
				</TD>
			</TR>
			</TABLE> 
		</TD>
	</TR>
	</TABLE>
	<CENTER>
	<%
	if dateTo="" then dateTo=shamsiToday()
	if dateFrom="" then  dateFrom = left(shamsiToday(),4) & "/01/01"
	'if dateFrom="" then dateFrom=session("FiscalYear")&"/01/01"
	Set HRS = Conn.Execute("SELECT ISNULL(dbo.CalcSumQttyInv("&  itemDetail & ",'"&dateTo&"',0),0) As OurSumQtty,ISNULL(dbo.CalcSumQttyInv("&  itemDetail & ",'"&dateTo&"',1),0) As TheirSumQtty")
	if Not HRS.EOF Then
		mysumQtty = clng(HRS("OurSumQtty"))
		ursumQtty = clng(HRS("TheirSumQtty"))
	End if 
	HRs.Close

	sqlstr = "SELECT InventoryLog.type, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.comments, InventoryLog.VoidedDate, InventoryLog.IsInput, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.ID, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID WHERE (InventoryItems.ID = "&  itemDetail & ") and InventoryLog.logDate>= N'"&dateFrom&"' and InventoryLog.logDate<= N'"&dateTo&"' ORDER BY InventoryLog.ID DESC"
	set RSS=Conn.Execute (sqlstr)	

	%><BR><BR>
	<table align=center width=90% border=1 bordercolor=black cellspacing=1 cellpadding=1>
	<TR >
		<FORM METHOD=POST ACTION="">
		<TD colspan=9>
		<H4>
		سابقه ورود و خروج كالا  (از تاريخ <%=dateFrom%> تا تاريخ  <%=dateTo%>) 
		</H4></TD>
		</FORM>
	</TR>
	<TR >
		<TD><SMALL>ورود</SMALL></A></TD>
		<TD><SMALL>خروج</SMALL></A></TD>
		<TD><SMALL>تراز كل</SMALL></A></TD>
		<TD><SMALL>تراز كل ارسالي</SMALL></A></TD>
		<TD><SMALL>واحد</SMALL></A></TD>
		<TD align=center><SMALL>تاريخ </SMALL></A></TD>
		<TD><SMALL>توضيحات</SMALL></A></TD>
		<TD><SMALL>توسط</SMALL></A></TD>
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
	<TR style="height:25pt" <% if RSS("voided") then%> disabled title="حدف شده در تاريخ <%=RSS("VoidedDate")%>"<% end if %>>
		<TD align=right dir=ltr><span style="font-size:10pt"><% if RSS("IsInput") then %><%=RSS("Qtty")%><% end if %></span></TD>
		<TD align=right dir=ltr style="position:relative;"><% if RSS("voided") then%><div style="right:0px;position:absolute;width:520;"><hr style="color:red;"></div><% end if %><span style="font-size:10pt"><% if not RSS("IsInput") then %><%=RSS("Qtty")%><% end if %></span></TD>
		<TD align=right dir=ltr><span style="font-size:10pt"><%=mySumQtty%></span></TD>
		<TD align=right dir=ltr><span style="font-size:10pt"><%=urSumQtty%></span></TD>
		<TD align=right dir=ltr ><%=RSS("Unit")%></TD>
		<TD dir=ltr align=center><%=RSS("logDate")%></span></TD>
		<TD><% if RSS("type")= "2" then
					response.write "<font color=red><b>اصلاح موجودي</b></font>"
				elseif RSS("type")= "3" then
					response.write "<font color=green><b>مرجوعي</b></font>"
				elseif RSS("type")= "4" then
					response.write "<font color=blue><b>تعريف كالاي جديد </b></font>"
				elseif RSS("type")= "5" then
					response.write "<font color=orang><b>انتقال</b></font>"
				elseif RSS("type")= "6" then
					response.write "<font color=#6699CC><b>ورود از توليد</b></font>"
				elseif RSS("type")= "7" then
					response.write "<font color=#FF9966><b>ورود از انبار شهريار</b></font>"
				elseif RSS("RelatedID")= "-1" then %> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			   <% else 
					if RSS("IsInput") then 
						%> شماره سفارش خريد: <A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
					else
						%>شماره حواله خروج: <A HREF="default.asp?ed=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
					end if
				end if %>
			<% if RSS("owner")<> "-1" and RSS("owner")<> "-2" then
				response.write " (ارسالي <a href='../CRM/AccountInfo.asp?act=show&selectedCustomer="& RSS("owner") &"' target='_blank'> " & RSS("owner") & "</a>)"
			   end if %>
			<% if trim(RSS("comments"))<> "-" and RSS("comments")<> "" then
				response.write " <br><br><B>توضيح:</B>  " & RSS("comments") 
			   end if %>

		</TD>
		<TD><%=RSS("RealName")%></TD>
	</TR>
	  
	<% 
		if Not RSS("Voided") Then
			if RSS("IsInput") Then
				if Rss("owner")=-1 then
					mySumQtty = mySumQtty - clng(RSS("Qtty"))
				else
					urSumQtty = urSumQtty - clng(RSS("Qtty"))
				end if 
			else
				if Rss("owner")=-1 then
					mySumQtty = mySumQtty + clng(RSS("Qtty"))
				else
					urSumQtty = urSumQtty + clng(RSS("Qtty"))
				end if 
			end if 
		end if
		if not RSS("voided") then

			if RSS("IsInput") then
				if Rss("owner")=-1 then
					inputs = inputs + clng(RSS("Qtty"))
				else
					urinputs = urinputs + clng(RSS("Qtty"))
				end if 
			else
				if Rss("owner")=-1 then
					outputs = outputs + clng(RSS("Qtty"))
				else
					uroutputs = uroutputs + clng(RSS("Qtty"))
				end if 
			end if
		end if

	RSS.moveNext
	Loop
	%>
	<TR height=25 style="color:black;">
		<TD align=right dir=ltr><%=inputs+urinputs%></A></TD>
		<TD align=right dir=ltr><%=outputs+uroutputs%></A></TD>
		<TD colspan=6>گردش ورود و خروج  از تاريخ <span dir=ltr><%=dateFrom%></span> تا تاريخ <span dir=ltr><%=dateTo%></span></A></TD>
	</TR>
	<TR  height=25 style="color:black;">
		<TD align=right dir=ltr><%=inputs%></A></TD>
		<TD align=right dir=ltr><%=outputs%></A></TD>
		<TD colspan=6>گردش ورود و خروج  از تاريخ <span dir=ltr><%=dateFrom%></span> تا تاريخ <span dir=ltr><%=dateTo%></span> خودمان</A></TD>
	</TR>
	<TR height=25 style="color:black;">
		<TD align=right dir=ltr><%=urinputs%></A></TD>
		<TD align=right dir=ltr><%=uroutputs%></A></TD>
		<TD colspan=6>گردش ورود و خروج  از تاريخ <span dir=ltr><%=dateFrom%></span> تا تاريخ <span dir=ltr><%=dateTo%> </span> ديگران</A></TD>
	</TR>
	<% 
		sqlstr = "SELECT ISNULL((SELECT Sum(Qtty) From InventoryLog Where (Voided=0) And (ItemID = "&  itemDetail & ") And Logdate<'"&dateFrom&"' And isInput=1),0) As InputQtty,ISNULL((SELECT Sum(Qtty) From InventoryLog Where (Voided=0) And (ItemID = "&  itemDetail & ") And Logdate<'"&dateFrom&"' And isInput=0 ),0) As OutputQtty"
		set DRS1 = conn.Execute(sqlstr)
		if (not DRS1.EOF) then
		%>
	<Tr><td colspan=7 Height=20 ></td></tr>
	<TR style="height:25pt;" >
		<TD align=right dir=ltr ><span style="font-size:10pt"><%=DRS1("InputQtty")%></span></TD>
		<TD align=right dir=ltr ><span style="font-size:10pt"><%=DRS1("OutputQtty")%></span></TD>
		<TD align=right dir=ltr ><span style="font-size:10pt"><%=cdbl(DRS1("InputQtty")) - cdbl(DRS1("OutputQtty")) %></span></TD>
		<TD align=right Colspan=5  >
		موجود و گردش تا قبل از تاريخ <span dir=ltr><%=dateFrom%></span>
		</TD>
	</TR>
		<%
		End if
		DRS1.Close
		%>
	</table><BR><BR>

</BODY>
</HTML>
