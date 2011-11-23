<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Response.Buffer = False
'Purchase (4)
PageTitle=" ÅÌêÌ—Ì Œ—Ìœ"
SubmenuItem=4
if not Auth(4 , 4) then NotAllowdToViewThisPage()

'OutService Page Trace
'By Alix - Last changed: 81/01/13
'By Alix - Last changed: 83/12/06 - Enable changing related order (job) of this purchase order
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<SCRIPT LANGUAGE="JavaScript">
<!--
function edit_order_id(reqID)
{
	tmp = document.getElementById("orderID_span").innerText
	document.getElementById("orderID_span").innerHTML="<INPUT id='orderID_edit_field' style='border:0pt; width:40pt' Type='text' value='"+tmp+"'> <input type='hidden' id='ireqID' value='"+reqID+"'>"
	document.getElementById("orderID_edit_field").select()
	document.getElementById("dokme_sabt").innerHTML=" &nbsp;<A HREF='javascript:xmlsend()'>[À» ]</A>"
}

function xmlsend()
{
	ordID = document.getElementById("orderID_edit_field").value
	reqID = document.getElementById("ireqID").value
	url = "/beta/purchase/XMLchangeRelatedOrder.asp?ordID="+ordID + "&reqID="+ reqID
	if (window.XMLHttpRequest) {
		var xmldoc=new XMLHttpRequest();
	} else if (window.ActiveXObject) {
		var xmldoc = new ActiveXObject("Microsoft.XMLHTTP");
	}
	//objHTTP.open('GET',url,false)
	//objHTTP.send()
	//tmpStr = unescape(objHTTP.responseText)

	document.getElementsByTagName('body')[0].style.cursor = 'wait';

	url = url + "&<%=currentTime10()%>";
	
	xmldoc.open('GET',url,false);
	xmldoc.setRequestHeader("Cache-Control","no-cache");
	try {xmldoc.send(valval);} catch (e) {xmldoc.send();}
	{
		//alert(xmldoc.status);
		if (xmldoc.readyState==4)
		{ 
		 document.getElementsByTagName('body')[0].style.cursor = 'auto';
		 returnobj = xmldoc.responseText;
		}
		if (returnobj=="ok") 
		{
		document.getElementById("orderID_span").innerHTML="<a href='../shopfloor/manageOrder.asp?radif="+ordID+"'>"+ordID+"</a>"
		document.getElementById("dokme_sabt").innerHTML=""	
		alert(returnobj)
		
		} else {alert("Ê—ÊœÌ ‰«œ—”  «” ")};
		
	};

}
//-->
</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Remove Relation to Inventory Log
'---------------------------------------------------------------------------------------84-10-18-Alix-
if request("removeRelation")="yes" then	
	invLogID	=	request("invLogID")
	ordID		=	request("ordID")
	newComment = "(Õ–› —«»ÿÂ »« ”›«—‘ Œ—Ìœ " & ordID & " œ—  «—ÌŒ " & shamsitoday() & "  Ê”ÿ " & session("CSRName") & ")"
	Conn.Execute ("update InventoryLog set relatedID=-1, comments = comments + N'" & newComment & "' where id=" &invLogID & "")	
	response.write "<br>"
	alertMsg = "—«»ÿÂ ”›«—‘ Œ—Ìœ ‘„«—Â " & ordID & " »« Ê—Êœ »Â «‰»«— ‘„«—Â " & invLogID & " œ—  «—ÌŒ " & shamsitoday() & " Õ–› ‘œ"
	'call showAlert( alertMsg , CONST_MSG_INFORM)
	
	response.redirect "outServiceTrace.asp?od=" & ordID & "&msg=" & alertMsg


end if
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------ Show an Out Service Order's detail
'-----------------------------------------------------------------------------------------------------
if request("od")<>"" then		
	ordID=request("od")
	set RSOD=Conn.Execute ("SELECT * FROM purchaseOrders WHERE id = "& ordID )	
	if not RSOD.eof then
		VendorID=RSOD("Vendor_ID")
		comment=RSOD("comment")
		otypeName=RSOD("TypeName")
		TypeID=RSOD("TypeID")
		Price=RSOD("Price")
		status=RSOD("status")
		printed=RSOD("printed")
		qtty=RSOD("qtty")
		HasVoucher=RSOD("HasVoucher")
		%><br>
		<center><h2>Ã“ÌÌ«  ”›«—‘ Œ—Ìœ </h2>
		‘„«—Â ”›«—‘: <%=ordID%>
		</center>
		<TABLE dir=rtl align=center width=600>
		<TR bgcolor="#CCCCCC">
			<TD width=50%>
				<li>‰Ê⁄ ﬂ«—: <%=otypeName%><br>
				<li>ﬁÌ„ : <%=Price%><br>
				<li> ⁄œ«œ: <%=qtty%><br>
				<li> Ê÷ÌÕ« : <%=comment%><br>
				<li>Ê÷⁄Ì : <%=status%><br>
					<%
					'===================================================
					' show voucher, if exists
					'===================================================
					set RSV=Conn.Execute ("SELECT Vouchers.verified, Vouchers.id, Vouchers.paid FROM Vouchers INNER JOIN VoucherLines ON Vouchers.id = VoucherLines.Voucher_ID WHERE (VoucherLines.RelatedPurchaseOrderID = "& ordID & ") AND voided=0")

					if not RSV.eof then
						response.write "<hR>"
						'linkto = "verify"
						'if not RSV("verified")=0 and RSV("paid")=0 then
						'	linkto = "payment"
						'end if 
						
						response.write "»—«Ì «Ì‰ ”›«—‘ Ìﬂ "
						response.write "<A target='_blank' HREF='../AP/AccountReport.asp?act=showVoucher&voucher="& RSV("ID") & "'> ›«ﬂ Ê— »Â ‘„«—Â " & RSV("ID") & " </A>"
						response.write "’«œ— ‘œÂ ﬂÂ  «ÌÌœ"
						if RSV("verified")=0 then
							response.write " ‰‘œÂ"
							isPaid = 0
						else
							response.write " ‘œÂ"
							isPaid = 1
						end if
						response.write " Ê Å—œ«Œ "
						if RSV("paid")=0 then
							response.write " ‰‘œÂ"
							isPaid = 0
						else
							response.write " ‘œÂ"
							isPaid = 1
						end if
						response.write " «” . <br>"
						if Not HasVoucher then
							response.write "<BR>Ê·Ì Ìﬂ «‘ﬂ«·Ì Â„ œ— œ«œÂ Â« Â” !"
							response.write "<BR><BR>‘„«—Â «Ì‰ ”›«—‘ —« »Â „œÌ— ”Ì” „ «ÿ·«⁄ œÂÌœ"
						end if
					else
						if HasVoucher then
							response.write "<BR><HR>‰« Â„«Â‰êÌ œ— œ«œÂ Â«! «‰ Ÿ«— „Ì —›  «Ì‰ ”›«—‘ Ìﬂ ›«ﬂ Ê— œ«‘ Â »«‘œ. Ê·Ì ‰œ«—œ."
							response.write "<BR><BR>‘„«—Â «Ì‰ ”›«—‘ —« »Â „œÌ— ”Ì” „ «ÿ·«⁄ œÂÌœ"
						end if
					end if
					%>
			</TD>
			<TD width=50%>
				<%
				set RSOD=Conn.Execute ("SELECT * FROM Accounts WHERE ID = "& VendorID )	
				if not RSOD.eof then
					%>
					<li><%=RSOD("accounttitle")%> (<%=VendorID%>)<br>
					<li>‰«„ ›—Ê‘‰œÂ: <%=RSOD("firstName1")%> <%=RSOD("lastName1")%><br>
					<li>‰«„ ‘—ﬂ : <%=RSOD("companyName")%><br>
					<li>‰‘«‰Ì: <%=RSOD("city1")%> - <%=RSOD("Address1")%> -<BR> Tel: <%=RSOD("tel1")%> - Fax: <%=RSOD("fax1")%> - Email: <%=RSOD("email1")%><br><br>

					<% 	ReportLogRow = PrepareReport ("purchaseOrder.rpt", "Pord_ID", ordID, "/beta/dialog_printManager.asp?act=Fin") %>
					<INPUT TYPE="button" value=" ç«Å »—ê ”›«—‘ Œœ„«  " style="height:25px; border:2 solid <%=SelectedMenuColor%>; width:100%; cursor:hand; background-Color:'white'; " onMouseOver="this.style.borderColor='white';" onMouseOut="this.style.borderColor='<%=SelectedMenuColor%>';" onclick="printThisReport(this,<%=ReportLogRow%>);">
					<%
				else
					response.write "<li> Œÿ«: ›—Ê‘‰œÂ „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ <br>" 
				end if 
				%>
			</TD>
		</TR>
		<TR>
			<TD>
				<%
				set RSOD=Conn.Execute ("SELECT * FROM purchaseOrderStatus WHERE Ord_ID = "& ordID )	
				'response.write "SELECT * FROM purchaseOrderStatus WHERE Ord_ID = "& ordID 
				Do while not RSOD.eof
					%>
					- <span dir=ltr><%=RSOD("StatusDate")%></span> <span dir=ltr>(<%=RSOD("StatusTime")%>)</span> : <%=RSOD("StatusDetail")%><br>
					<%
				RSOD.moveNext
				Loop
				%>
			</TD>
			<TD>
				<FORM METHOD=POST ACTION="?">
					<INPUT TYPE="hidden" name="ordID" value="<%=ordID%>">
					À»  Ê÷⁄Ì  ÃœÌœ:<br>
					<INPUT TYPE="radio" NAME="st" value="1" <% if status="OK" and isPaid=1 then %> disabled <% else %> checked <% end if %>>
					<SELECT NAME="stdet1" onfocus="document.all.st[0].checked='true'" <% if status="OK" and isPaid=1 then %> disabled <% end if %>>
					<option <% if status="OUT" then%>selected<% end if %> value="11">”›«—‘ »—«Ì ›—Ê‘‰œÂ «—”«· ‘œ</option>
					<option <% if status="RETURN" then%>selected<% end if %> value="12">”›«—‘ »Â ‘—ﬂ  »«“ê‘ </option>
					<option <% if status="CANCEL" then%>selected<% end if %> value="13">”›«—‘ ·€Ê ‘œ</option>
					<option <% if status="OK" then%>selected<% end if %> value="20">”›«—‘  «ÌÌœ ‘œ</option>
					<option <% if status="Unknown" then%>selected<% end if %> value="14">‰«„⁄·Ê„</option>
					</SELECT><br>
					<INPUT TYPE="radio" NAME="st" value="2"><INPUT TYPE="text" NAME="stdet12" onfocus="document.all.st[1].checked='true'"><br>
					 «—ÌŒ<INPUT TYPE="text" NAME="stDate" value="<%=shamsiToday()%>" dir=ltr onblur="acceptDate(this);" onKeyPress="return maskDate(this);" ><br>
					<center>
					<INPUT TYPE="submit" Name="Submit" Value="À»  Ê÷⁄Ì " style="width:100px;" tabIndex="14">
					</center>
				</FORM>
			</TD>
		</TR>
		<TR >
			<TD colspan=2>
				<!------------------------------------------------------------>
				<%
				set RSS=Conn.Execute ("SELECT InventoryLog.comments, InventoryLog.CreatedBy, InventoryLog.owner, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.ID, InventoryItems.OldItemID, InventoryItems.Name, InventoryItems.Unit, Users.RealName FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID WHERE InventoryLog.IsInput=1 and InventoryLog.Voided=0 and (InventoryLog.RelatedID = "& ordID & ")")	
				if not RSS.EOF then
				%>
				<TABLE dir=rtl align=center width=600>
					<TR bgcolor="eeeeee" >
						<TD colspan=7><H4>Ê—Êœ ﬂ«·« »Â «‰»«— „—»Êÿ »Â «Ì‰ ”›«—‘ Œ—Ìœ</H4></TD>
					</TR>
					<TR bgcolor="eeeeee" >
						<TD><!A HREF="default.asp?s=1"><SMALL>Õ–› —«»ÿÂ</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=1"><SMALL>ﬂœ ﬂ«·«</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=2"><SMALL>‰«„ ﬂ«·«</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=3"><SMALL> ⁄œ«œ Ê—Êœ</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=5"><SMALL> «—ÌŒ Ê—Êœ</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=6"><SMALL>‘„«—Â ”›«—‘ Œ—Ìœ</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=7"><SMALL> Ê”ÿ</SMALL></A></TD>
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
					<TR bgcolor="<%=tmpColor%>" title="<%=RSS("comments")%>">
						<TD align=center><a onclick="return confirm('¬Ì« Ê«ﬁ⁄« „Ì ŒÊ«ÂÌœ «Ì‰ Ê—Êœ »Â «‰»«— —« «“ «Ì‰ ”›«—‘ Œ—Ìœ »—œ«—Ìœø')" href="?removeRelation=yes&invLogID=<%=RSS("ID")%>&ordID=<%=ordID%>">X</a></TD>
						<TD><INPUT TYPE="hidden" name="id" value="<%=RSS("ID")%>"><%=RSS("OldItemID")%></TD>
						<TD><!A HREF="default.asp?itemDetail=<%=RSS("ID")%>"><%=RSS("Name")%></A></TD>
						<TD><%=RSS("Qtty")%>&nbsp;<%=RSS("Unit")%></TD>
						<TD><span dir=ltr><%=RSS("logDate")%></span></TD>
						<TD><% if RSS("RelatedID")= "-1" then
								response.write "‰œ«—œ"
							 else
								response.write RSS("RelatedID")
							 end if	%></TD>
						<TD><%=RSS("RealName")%></TD>
					</TR>
						 
					<% 
					RSS.moveNext
					Loop
				%>
				</TABLE><br>
				<% end if %>
				<!------------------------------------------------------------>
				<TABLE width=100% bgcolor="#CCCCCC">
				<TR><TD COLSPAN=5><CENTER><B>—”Ìœ œ—ŒÊ«”  Â«Ì „—»Êÿ »Â «Ì‰ ”›«—‘</B></CENTER><hr></TD></TR>
				<tr>
				<TD>‰Ê⁄ ⁄„·Ì« </TD>
				<TD> «—ÌŒ œ—ŒÊ«” </TD>
				<TD>ﬁÌ„  „Ê—œ ‰Ÿ—</TD>
				<TD>”›«—‘ „—»ÊÿÂ</TD>
				<TD width=8%>œ—’œ </TD>
				</tr>
				<TR><TD COLSPAN=5><hr></TD></TR>
				<%
				set RSOD=Conn.Execute ("SELECT * FROM purchaseRequestOrderRelations WHERE Ord_ID = "& OrdID )	
				Do while not RSOD.eof
					set RSX=Conn.Execute ("SELECT * FROM purchaseRequests WHERE id = "& RSOD("req_ID") )	
					%>
					<tr>
					<TD><%=RSX("typeName")%></TD>
					<TD><span dir=ltr><%=RSX("ReqDate")%></span></TD>
					<TD><%=RSX("price")%></TD>
					<TD><% if RSX("order_ID")=-1 then%>‰œ«—œ<% else %><span id="orderID_span"><a href="../shopfloor/manageOrder.asp?radif=<%=RSX("order_ID")%>"><%=RSX("order_ID")%></a></span><% if Auth(4 , 6) then %><span id="dokme_sabt"> &nbsp;<A HREF="javascript:edit_order_id(<%=RSX("id")%>)">[ €ÌÌ—]</A></span><% end if %><% end if %></TD>
					<TD width=8%><%=RSOD("PercentOfAll")%></TD>
					</tr>
					<%	
					RSOD.moveNext 
				loop %>
			</TD>
		</TR>
		</TABLE>
	<%
	end if
	'response.redirect "?radif=" & request("r")
	RSOD.close
	response.end

end if

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------- change status of an OutService Order
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="À»  Ê÷⁄Ì " then
	ordID = request.form("ordID") 
	st = request.form("st")
	stDate = request.form("stDate")
	hasUnpaidVoucher = "no"

	if st=1 then
		stid=request.form("stdet1")
		if stid="11" then 
			stdetl="”›«—‘ »—«Ì ›—Ê‘‰œÂ «—”«· ‘œ"
			orderStatus = "OUT"
		elseif stid="12" then 
			stdetl="”›«—‘ »Â ‘—ﬂ  »«“ê‘ "
			orderStatus = "RETURN"
		elseif stid="13" then 
			stdetl="”›«—‘ ·€Ê ‘œ"
			orderStatus = "CANCEL"
		elseif stid="20" then 
			stdetl="”›«—‘  «ÌÌœ ‘œ"
			orderStatus = "OK"
		elseif stid="14" then 
			stdetl="”›«—‘ »Â Õ«·  ‰«„⁄·Ê„ œ—¬„œ"
			orderStatus = "Unknown"
		end if
	else
		stdetl=request.form("stdet12")
		stid=3
		orderStatus = "NOT CHANGED"
	end if

	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusTime, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", '"& currentTime10() & "', '"& stDate & "', "& stid & ", N'"& stdetl & "' )")

	if NOT orderStatus = "NOT CHANGED" then 
		Conn.Execute ("update purchaseOrders SET status = '"& orderStatus & "' where id = "& OrdID )	
	end if 

	if orderStatus = "OK" then
		'===================================================
		' check to see if exists
		'===================================================
		set RSV=Conn.Execute ("SELECT Vouchers.* FROM Vouchers INNER JOIN VoucherLines ON Vouchers.id = VoucherLines.Voucher_ID WHERE (VoucherLines.RelatedPurchaseOrderID = "& ordID & ")")

		if not RSV.eof then
			if RSV("paid")=0 then
				VoucherID = RSV("ID")
				VendorID = RSV("VendorID")
				TotalPrice = RSV("TotalPrice")
				hasUnpaidVoucher = "yes"

				'---------------------------------------------------------------------------------------------------
				'------ Next line has been Commented by Alix - 82-02-18 
				'------ Routine changed: Account will be updated when Voucher verified. (only in AP/Verify.asp page)
				'---------------------------------------------------------------------------------------------------

				'Conn.Execute ("UPDATE Accounts SET APBalance=APBalance+"& TotalPrice & " WHERE (ID = "& VendorID & ")")	
			else
				response.write "<BR><BR>"
				response.write "Œÿ«! «Ì‰ ”›«—‘ ﬁ»·« Å—œ«Œ  ‘œÂ «” . œ— Õ«·Ì ﬂÂ Â„Ê“  «ÌÌœ ‰‘œÂ »ÊœÂ"
				response.write "<BR><BR>"
				response.write "ﬁ÷ÌÂ „‘ﬂÊﬂ «” !"
			end if
		end if
		'===================================================
	end if 

	response.redirect "?od=" & ordID

end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------ Main
'-----------------------------------------------------------------------------------------------------
'Gets all Request for services list from DB

set RST=Conn.Execute ("SELECT count(*) as CNT FROM purchaseOrders WHERE (Status = 'NEW')")
C1 = RST("CNT")
set RST=Conn.Execute ("SELECT count(*) as CNT FROM purchaseOrders WHERE (Status = 'OUT')")
C2 = RST("CNT")
set RST=Conn.Execute ("SELECT count(*) as CNT FROM purchaseOrders WHERE (Status = 'RETURN')")
C3 = RST("CNT")

eventDate	=	sqlSafe(request("eventDate"))

%>
<br><br>
<CENTER>
<FORM METHOD=POST ACTION="">
‘„«—Â ”›«—‘ Œ—Ìœ: <INPUT TYPE="text" NAME="od" size=6><INPUT TYPE="submit" value="‰„«Ì‘ Ã“ÌÌ« ">
</FORM>
</CENTER>
<br><TABLE dir=rtl align=center width=600 class=t8pt>
<TR bgcolor="eeeeee">
	<TD align=center bgcolor=ffffff><B>ÅÌêÌ—Ì ”›«—‘ Â«Ì Œ—Ìœ </B></TD>
	<TD align=center><A HREF="?lstOrd=NEW"><IMG SRC="../images/folder<% if C1=0 then %>0<% else%>1<% end if %>.gif" BORDER=0><br>ÃœÌœ </A>(<%=C1%>)</TD>
	<TD align=center><A HREF="?lstOrd=OUT"><IMG SRC="../images/folder<% if C2=0 then %>0<% else%>1<% end if %>.gif" BORDER=0><br>Œ«—Ã «“ ‘—ﬂ  </A>(<%=C2%>)</TD>
	<TD align=center><A HREF="?lstOrd=RETURN"><IMG SRC="../images/folder<% if C3=0 then %>0<% else%>1<% end if %>.gif" BORDER=0><br>»—ê‘ Â »Â ‘—ﬂ  </A>(<%=C3%>)</TD>
	<TD align=center><A HREF="?lstOrd=CANCEL"><IMG SRC="../images/folder2.gif" BORDER=0 alt="¬—‘ÌÊ"><br>·€Ê ‘œÂ </A></TD>
	<TD align=center><A HREF="?lstOrd=OK"><IMG SRC="../images/folder2.gif" BORDER=0 alt="¬—‘ÌÊ"><br> «ÌÌœ ‘œÂ </A></TD>
	<TD align=center><A HREF="?lstOrd=Unknown"><IMG SRC="../images/folder2.gif" BORDER=0 alt="¬—‘ÌÊ"><br>‰«„⁄·Ê„ </A></TD>
</TR>
<TR bgcolor="eeeeee">
	<FORM METHOD=POST ACTION="?act=showEvents">
	<TD align=center bgcolor=ffffff><B>ÅÌêÌ—Ì œ— Ìﬂ  «—ÌŒ </B></TD>
	<TD align=right valign=bottom colspan=6 >
		<INPUT dir="LTR" TYPE="text" NAME="eventDate" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=eventDate%>">
		<INPUT TYPE="submit" Value="‰„«Ì‘" class="genButton">
	</TD>
	</FORM>
</TR>
</TABLE><br>
<%
sortBy=request("s")
if sortBy="2" then 
	sB="PurchaseOrders.typeName"
elseif sortBy="3" then 
	sB="PurchaseOrders.OrdDate"
elseif sortBy="4" then 
	sB="PurchaseOrders.price"
elseif sortBy="5" then 
	sB="PurchaseOrders.Qtty"
else 
	sB="PurchaseOrders.id DESC"
end if


if request("act") = "showEvents" then			


	'mySQL="SELECT * FROM PurchaseOrderStatus INNER JOIN PurchaseOrders ON PurchaseOrderStatus.Ord_ID = PurchaseOrders.ID WHERE (PurchaseOrderStatus.StatusDate = '"& eventDate & "') ORDER BY PurchaseOrderStatus.Ord_ID DESC, PurchaseOrderStatus.StatusTime DESC"
	mySQL="SELECT * FROM PurchaseOrderStatus INNER JOIN PurchaseOrders ON PurchaseOrderStatus.Ord_ID = PurchaseOrders.ID WHERE (PurchaseOrderStatus.StatusDate = '"& eventDate & "') ORDER BY PurchaseOrderStatus.Ord_ID DESC, PurchaseOrderStatus.ID DESC"

	set RSS=Conn.Execute (mySQL)
	%><br>
	<TABLE dir=rtl align=center width=600>
	<TR bgcolor="eeeeee" >
		<TD align=center colspan=6><B>Ê÷⁄Ì  Â«Ì ”›«—‘Â«Ì Œ—Ìœ</B></TD>
	</TR>
	<TR bgcolor="eeeeee" height=20>
		<TD><SMALL>‘„«—Â</SMALL></TD>
		<TD width=200><SMALL>‰Ê⁄ ”›«—‘</SMALL></TD>
		<TD><SMALL> «—ÌŒ ”›«—‘</SMALL></TD>
		<TD width=200><SMALL>‘—Õ</SMALL></TD>
		<TD><SMALL>ﬁÌ„ </SMALL></TD>
		<TD><SMALL> ⁄œ«œ</SMALL></TD>
	</TR>
	<%
	tmpCounter=0
	Do while not RSS.eof
		if not tmpid = RSS("id") then
			tmpCounter = tmpCounter + 1
		end if 
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 
		tmpid = RSS("id")
	%>
	<TR bgcolor="<%=tmpColor%>" height=20 title="<% 
		Comment = RSS("Comment")
		if Comment<>"-" then
			response.write " Ê÷ÌÕ: " & Comment
		else
			response.write " Ê÷ÌÕ ‰œ«—œ"
		end if
	%>">
		<TD><%=RSS("id")%></TD>
		<TD><A HREF="?od=<%=RSS("id")%>" target="_blank"><%=RSS("TypeName")%></A></TD>
		<TD><span dir=ltr><%=RSS("OrdDate")%></span></TD>
		<TD><%=RSS("StatusDetail")%></TD>
		<TD><%=RSS("price")%></TD>
		<TD><%=RSS("qtty")%></TD>
	 
	<% 
	RSS.moveNext
	Loop
	%>
	</TABLE><br>

	</body></html>
	<% 
	response.end
end if 

if request("lstOrd") = "" then			
response.end
end if 
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ List OutService Orders 
'---------------------------------------------------------------------- (NEW, OUT, RETURN, CANCEL, OK)
'-----------------------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------------------
' OK or cancel Orders
'-----------------------------------------------------------------------------------------------------


'======================== old query (changed 82-10-15 by ALix) =======================================
'SELECT PurchaseOrders.*, PurchaseRequests.Order_ID AS Order_ID, Invoices.Approved AS Approved, Invoices.Issued AS Issued FROM PurchaseRequests FULL OUTER JOIN Invoices FULL OUTER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice ON PurchaseRequests.Order_ID = InvoiceOrderRelations.[Order] FULL OUTER JOIN PurchaseOrders LEFT OUTER JOIN PurchaseRequestOrderRelations ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID ON PurchaseRequests.ID = PurchaseRequestOrderRelations.Req_ID where ISNULL(Invoices.voided,0)=0 and (PurchaseOrders.Status =

if request("lstOrd") <> "" then			
	lstOrd = request("lstOrd")
	set RSS=Conn.Execute ("SELECT top 100 dbo.PurchaseOrders.*, dbo.PurchaseRequests.Order_ID AS Order_ID, ISNULL(DERIVEDTBL.Approved2, 0) AS Approved, ISNULL(DERIVEDTBL.Issued2, 0) AS Issued, DERIVEDTBL.id AS invoice_ID FROM dbo.PurchaseRequests FULL OUTER JOIN (SELECT Invoices.id, isnull(Invoices.Approved, 0) AS Approved2, isnull(Invoices.Issued, 0) AS Issued2 FROM dbo.Invoices WHERE isnull(Invoices.voided, 0) = 0 and (isnull(Invoices.Approved, 0) = 1 OR isnull(Invoices.Issued, 0) = 1)) DERIVEDTBL INNER JOIN dbo.InvoiceOrderRelations ON DERIVEDTBL.id = dbo.InvoiceOrderRelations.Invoice ON dbo.PurchaseRequests.Order_ID = dbo.InvoiceOrderRelations.[Order] FULL OUTER JOIN dbo.PurchaseOrders LEFT OUTER JOIN dbo.PurchaseRequestOrderRelations ON dbo.PurchaseOrders.ID = dbo.PurchaseRequestOrderRelations.Ord_ID ON dbo.PurchaseRequests.ID = dbo.PurchaseRequestOrderRelations.Req_ID WHERE (dbo.PurchaseOrders.Status = '"& lstOrd & "') order by " & sB)
	%><br>
	<TABLE dir=rtl align=center width=600>
	<TR bgcolor="eeeeee" >
		<TD align=center colspan=6><B>”›«—‘Â«Ì Œ—Ìœ (<%=lstOrd%>)</B></TD>
	</TR>
	<TR bgcolor="eeeeee" height=20>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=1"><SMALL>‘„«—Â</SMALL></A></TD>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=2"><SMALL>‰Ê⁄ ”›«—‘</SMALL></A></TD>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=3"><SMALL> «—ÌŒ ”›«—‘</SMALL></A></TD>
		<TD align=center><A HREF="?lstOrd=<%=lstOrd%>&s=3"><SMALL>‘„«—Â ”›«—‘</SMALL></A></TD>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=4"><SMALL>ﬁÌ„ </SMALL></A></TD>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=5"><SMALL> ⁄œ«œ</SMALL></A></TD>
	</TR>
	<%
	tmpCounter=0
	Do while not RSS.eof
		if not tmpid = RSS("id") then
			tmpCounter = tmpCounter + 1
		end if 
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 
		tmpid = RSS("id")
	%>
	<TR bgcolor="<%=tmpColor%>" height=20 title="<% 
		Comment = RSS("Comment")
		if Comment<>"-" then
			response.write " Ê÷ÌÕ: " & Comment
		else
			response.write " Ê÷ÌÕ ‰œ«—œ"
		end if
	%>">
		<TD><%=RSS("id")%></TD>
		<TD><A HREF="?od=<%=RSS("id")%>"><%=RSS("TypeName")%></A></TD>
		<TD><span dir=ltr><%=RSS("OrdDate")%></span></TD>
		<TD align=center><%
		if RSS("order_ID") = "-1" then
			response.write "-"
		else
			response.write "<A target='_blank' HREF='../shopfloor/manageOrder.asp?radif=" & RSS("order_ID") & "'>" & RSS("order_ID") & "</a>"

			if RSS("Approved") then
				response.write "<br>" & "›«ﬂ Ê—  «ÌÌœ ‘œÂ"
			end if 

			if RSS("Issued") then
				response.write "<br>" & " ÕÊÌ· „‘ —Ì ‘œÂ"
			end if 

		end if
		%></TD>
		<TD><%=RSS("price")%></TD>
		<TD><%=RSS("qtty")%></TD>
	 
	<% 
	RSS.moveNext
	Loop
	%>
	</TABLE><br>

	</body></html>
	<% 
	response.end
end if 
Conn.Close
%>
<!--#include file="tah.asp" -->
