<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Response.Buffer = False
'Purchase (4)
PageTitle=" ����� ����"
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
	document.getElementById("dokme_sabt").innerHTML=" &nbsp;<A HREF='javascript:xmlsend()'>[���]</A>"
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
		
		} else {alert("����� ������ ���")};
		
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
	newComment = "(��� ����� �� ����� ���� " & ordID & " �� ����� " & shamsitoday() & " ���� " & session("CSRName") & ")"
	Conn.Execute ("update InventoryLog set relatedID=-1, comments = comments + N'" & newComment & "' where id=" &invLogID & "")	
	response.write "<br>"
	alertMsg = "����� ����� ���� ����� " & ordID & " �� ���� �� ����� ����� " & invLogID & " �� ����� " & shamsitoday() & " ��� ��"
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
		<center><h2>������ ����� ���� </h2>
		����� �����: <%=ordID%>
		</center>
		<TABLE dir=rtl align=center width=600>
		<TR bgcolor="#CCCCCC">
			<TD width=50%>
				<li>��� ���: <%=otypeName%><br>
				<li>����: <%=Price%><br>
				<li>�����: <%=qtty%><br>
				<li>�������: <%=comment%><br>
				<li>�����: <%=status%><br>
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
						
						response.write "���� ��� ����� �� "
						response.write "<A target='_blank' HREF='../AP/AccountReport.asp?act=showVoucher&voucher="& RSV("ID") & "'> ������ �� ����� " & RSV("ID") & " </A>"
						response.write "���� ��� �� �����"
						if RSV("verified")=0 then
							response.write " ����"
							isPaid = 0
						else
							response.write " ���"
							isPaid = 1
						end if
						response.write " � ������"
						if RSV("paid")=0 then
							response.write " ����"
							isPaid = 0
						else
							response.write " ���"
							isPaid = 1
						end if
						response.write " ���. <br>"
						if Not HasVoucher then
							response.write "<BR>��� �� ������ �� �� ���� �� ���!"
							response.write "<BR><BR>����� ��� ����� �� �� ���� ����� ����� ����"
						end if
					else
						if HasVoucher then
							response.write "<BR><HR>�� ������ �� ���� ��! ������ �� ��� ��� ����� �� ������ ����� ����. ��� �����."
							response.write "<BR><BR>����� ��� ����� �� �� ���� ����� ����� ����"
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
					<li>��� �������: <%=RSOD("firstName1")%> <%=RSOD("lastName1")%><br>
					<li>��� ����: <%=RSOD("companyName")%><br>
					<li>�����: <%=RSOD("city1")%> - <%=RSOD("Address1")%> -<BR> Tel: <%=RSOD("tel1")%> - Fax: <%=RSOD("fax1")%> - Email: <%=RSOD("email1")%><br><br>

					<% 	ReportLogRow = PrepareReport ("purchaseOrder.rpt", "Pord_ID", ordID, "/beta/dialog_printManager.asp?act=Fin") %>
					<INPUT TYPE="button" value=" �ǁ �ѐ ����� ����� " style="height:25px; border:2 solid <%=SelectedMenuColor%>; width:100%; cursor:hand; background-Color:'white'; " onMouseOver="this.style.borderColor='white';" onMouseOut="this.style.borderColor='<%=SelectedMenuColor%>';" onclick="printThisReport(this,<%=ReportLogRow%>);">
					<%
				else
					response.write "<li> ���: ������� ���� ��� ���� ��� <br>" 
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
					��� ����� ����:<br>
					<INPUT TYPE="radio" NAME="st" value="1" <% if status="OK" and isPaid=1 then %> disabled <% else %> checked <% end if %>>
					<SELECT NAME="stdet1" onfocus="document.all.st[0].checked='true'" <% if status="OK" and isPaid=1 then %> disabled <% end if %>>
					<option <% if status="OUT" then%>selected<% end if %> value="11">����� ���� ������� ����� ��</option>
					<option <% if status="RETURN" then%>selected<% end if %> value="12">����� �� ���� ��Ґ��</option>
					<option <% if status="CANCEL" then%>selected<% end if %> value="13">����� ��� ��</option>
					<option <% if status="OK" then%>selected<% end if %> value="20">����� ����� ��</option>
					<option <% if status="Unknown" then%>selected<% end if %> value="14">�������</option>
					</SELECT><br>
					<INPUT TYPE="radio" NAME="st" value="2"><INPUT TYPE="text" NAME="stdet12" onfocus="document.all.st[1].checked='true'"><br>
					�����<INPUT TYPE="text" NAME="stDate" value="<%=shamsiToday()%>" dir=ltr onblur="acceptDate(this);" onKeyPress="return maskDate(this);" ><br>
					<center>
					<INPUT TYPE="submit" Name="Submit" Value="��� �����" style="width:100px;" tabIndex="14">
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
						<TD colspan=7><H4>���� ���� �� ����� ����� �� ��� ����� ����</H4></TD>
					</TR>
					<TR bgcolor="eeeeee" >
						<TD><!A HREF="default.asp?s=1"><SMALL>��� �����</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=1"><SMALL>�� ����</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=2"><SMALL>��� ����</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=3"><SMALL>����� ����</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=5"><SMALL>����� ����</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=6"><SMALL>����� ����� ����</SMALL></A></TD>
						<TD><!A HREF="default.asp?s=7"><SMALL>����</SMALL></A></TD>
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
						<TD align=center><a onclick="return confirm('��� ����� �� ������ ��� ���� �� ����� �� �� ��� ����� ���� ������Ͽ')" href="?removeRelation=yes&invLogID=<%=RSS("ID")%>&ordID=<%=ordID%>">X</a></TD>
						<TD><INPUT TYPE="hidden" name="id" value="<%=RSS("ID")%>"><%=RSS("OldItemID")%></TD>
						<TD><!A HREF="default.asp?itemDetail=<%=RSS("ID")%>"><%=RSS("Name")%></A></TD>
						<TD><%=RSS("Qtty")%>&nbsp;<%=RSS("Unit")%></TD>
						<TD><span dir=ltr><%=RSS("logDate")%></span></TD>
						<TD><% if RSS("RelatedID")= "-1" then
								response.write "�����"
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
				<TR><TD COLSPAN=5><CENTER><B>���� ������� ��� ����� �� ��� �����</B></CENTER><hr></TD></TR>
				<tr>
				<TD>��� ������</TD>
				<TD>����� �������</TD>
				<TD>���� ���� ���</TD>
				<TD>����� ������</TD>
				<TD width=8%>���� </TD>
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
					<TD><% if RSX("order_ID")=-1 then%>�����<% else %><span id="orderID_span"><a href="../shopfloor/manageOrder.asp?radif=<%=RSX("order_ID")%>"><%=RSX("order_ID")%></a></span><% if Auth(4 , 6) then %><span id="dokme_sabt"> &nbsp;<A HREF="javascript:edit_order_id(<%=RSX("id")%>)">[�����]</A></span><% end if %><% end if %></TD>
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
if request.form("Submit")="��� �����" then
	ordID = request.form("ordID") 
	st = request.form("st")
	stDate = request.form("stDate")
	hasUnpaidVoucher = "no"

	if st=1 then
		stid=request.form("stdet1")
		if stid="11" then 
			stdetl="����� ���� ������� ����� ��"
			orderStatus = "OUT"
		elseif stid="12" then 
			stdetl="����� �� ���� ��Ґ��"
			orderStatus = "RETURN"
		elseif stid="13" then 
			stdetl="����� ��� ��"
			orderStatus = "CANCEL"
		elseif stid="20" then 
			stdetl="����� ����� ��"
			orderStatus = "OK"
		elseif stid="14" then 
			stdetl="����� �� ���� ������� �����"
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
				response.write "���! ��� ����� ���� ������ ��� ���. �� ���� �� ���� ����� ���� ����"
				response.write "<BR><BR>"
				response.write "���� ����� ���!"
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
����� ����� ����: <INPUT TYPE="text" NAME="od" size=6><INPUT TYPE="submit" value="����� ������">
</FORM>
</CENTER>
<br><TABLE dir=rtl align=center width=600 class=t8pt>
<TR bgcolor="eeeeee">
	<TD align=center bgcolor=ffffff><B>����� ����� ��� ���� </B></TD>
	<TD align=center><A HREF="?lstOrd=NEW"><IMG SRC="../images/folder<% if C1=0 then %>0<% else%>1<% end if %>.gif" BORDER=0><br>���� </A>(<%=C1%>)</TD>
	<TD align=center><A HREF="?lstOrd=OUT"><IMG SRC="../images/folder<% if C2=0 then %>0<% else%>1<% end if %>.gif" BORDER=0><br>���� �� ���� </A>(<%=C2%>)</TD>
	<TD align=center><A HREF="?lstOrd=RETURN"><IMG SRC="../images/folder<% if C3=0 then %>0<% else%>1<% end if %>.gif" BORDER=0><br>�ѐ��� �� ���� </A>(<%=C3%>)</TD>
	<TD align=center><A HREF="?lstOrd=CANCEL"><IMG SRC="../images/folder2.gif" BORDER=0 alt="�����"><br>��� ��� </A></TD>
	<TD align=center><A HREF="?lstOrd=OK"><IMG SRC="../images/folder2.gif" BORDER=0 alt="�����"><br>����� ��� </A></TD>
	<TD align=center><A HREF="?lstOrd=Unknown"><IMG SRC="../images/folder2.gif" BORDER=0 alt="�����"><br>������� </A></TD>
</TR>
<TR bgcolor="eeeeee">
	<FORM METHOD=POST ACTION="?act=showEvents">
	<TD align=center bgcolor=ffffff><B>����� �� �� ����� </B></TD>
	<TD align=right valign=bottom colspan=6 >
		<INPUT dir="LTR" TYPE="text" NAME="eventDate" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=eventDate%>">
		<INPUT TYPE="submit" Value="�����" class="genButton">
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
		<TD align=center colspan=6><B>����� ��� �������� ����</B></TD>
	</TR>
	<TR bgcolor="eeeeee" height=20>
		<TD><SMALL>�����</SMALL></TD>
		<TD width=200><SMALL>��� �����</SMALL></TD>
		<TD><SMALL>����� �����</SMALL></TD>
		<TD width=200><SMALL>���</SMALL></TD>
		<TD><SMALL>����</SMALL></TD>
		<TD><SMALL>�����</SMALL></TD>
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
			response.write "�����: " & Comment
		else
			response.write "����� �����"
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
		<TD align=center colspan=6><B>�������� ���� (<%=lstOrd%>)</B></TD>
	</TR>
	<TR bgcolor="eeeeee" height=20>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=1"><SMALL>�����</SMALL></A></TD>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=2"><SMALL>��� �����</SMALL></A></TD>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=3"><SMALL>����� �����</SMALL></A></TD>
		<TD align=center><A HREF="?lstOrd=<%=lstOrd%>&s=3"><SMALL>����� �����</SMALL></A></TD>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=4"><SMALL>����</SMALL></A></TD>
		<TD><A HREF="?lstOrd=<%=lstOrd%>&s=5"><SMALL>�����</SMALL></A></TD>
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
			response.write "�����: " & Comment
		else
			response.write "����� �����"
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
				response.write "<br>" & "������ ����� ���"
			end if 

			if RSS("Issued") then
				response.write "<br>" & "����� ����� ���"
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
