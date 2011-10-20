<%'
  'This Include file needs the variables below:
  '
  '		Conn	
  '		cusID	(Customer ID)
  '
%>
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;height:30;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
</STYLE>
	  <Tr>
		<Td colspan="2" valign="center" align="center">
			<BR>
			<% if Auth(7 , 1) then %><input class="GenButton" type="button" value="Ê—Êœ ›«ﬂ Ê—" onclick="window.open('../AP/voucherInput.asp?act=show&selectedCustomer=<%=cusID%>');" <% if AccountIsDisabled then %> disabled <% end if %>> <% end if %> 
			<% if Auth(7 , 4) then %><input class="GenButton" type="button" value="Ê—Êœ «⁄·«„ÌÂ" onclick="window.open('../AP/MemoInput.asp?act=getMemo&selectedCustomer=<%=cusID%>');" <% if AccountIsDisabled then %> disabled <% end if %>> <% end if %> 
			<% if Auth(7 , 6) then %><input class="GenButton" type="button" value="œÊŒ ‰" onclick="window.open('../AP/ItemsRelation.asp?sys=AR&act=relate&selectedCustomer=<%=CusID%>');"> <% end if %> 
			<% if Auth(7 , 5) then %><input class="GenButton" type="button" value="ê“«—‘ Õ”«»" onclick="window.open('../AP/AccountReport.asp?sys=AR&act=show&selectedCustomer=<%=CusID%>');"> <% end if %> 
			<% if Auth(7 , 3) then %><input class="GenButton" type="button" value="ê“«—‘ «·›" onclick="window.open('../AP/IsA-VouchersReport.asp?selectedCustomer=<%=CusID%>');"> <% end if %> 

			<BR><BR>
		</Td>
	  </Tr>
	  <Tr>
		<Td colspan="2" valign="top" align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="5" class="CusTableHeader">”›«—‘ Â«Ì Œ—Ìœ œ— Ã—Ì«‰ (ﬂ‰”· ‰‘œÂ Ê  „«„ ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT * FROM PurchaseOrders WHERE (Vendor_ID = '"& cusID & "') AND (Status <> N'OK' AND Status <> N'CANCEL') ORDER BY OrdDate DESC, ID"
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
%>
					<tr>
						<td colspan="5" class="CusTD3">ÂÌç</td>
					</tr>
<%
				else
%>					<tr>
						<td class="CusTD3">#</td>
						<td class="CusTD3">‘„«—Â</td>
						<td class="CusTD3"> «—ÌŒ</td>
						<td class="CusTD3">‰Ê⁄</td>
						<td class="CusTD3">„—Õ·Â</td>
					</tr>
<%					tmpCounter=0
					Do while not RS1.eof 
						tmpCounter = tmpCounter + 1
						if tmpCounter mod 2 = 1 then
							tmpColor="#FFFFFF"
							tmpColor2="#FFFFBB"
						Else
							tmpColor="#DDDDDD"
							tmpColor2="#EEEEBB"
						End if 
						'alert(this.getElementByTagName('td').items(0).data);
%>
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../purchase/outServiceTrace.asp?od=<%=RS1("ID")%>');">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD dir="LTR" align='right'><%=RS1("OrdDate")%>&nbsp;</TD>
							<TD><%=RS1("TypeName")%>&nbsp;</TD>
							<TD><%=RS1("Status")%>&nbsp;</TD>
						</TR>
<%						RS1.moveNext
					Loop
				end if
				%>
			</table>
		</Td>
	  </Tr>
	  <Tr>
		<Td colspan="2" valign="top" align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="7" class="CusTableHeader" style="background-color:#CCCC99;">›«ﬂ Ê— Â«Ì Œ—Ìœ œ— Ã—Ì«‰ ( «ÌÌœ ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Vouchers.id, Vouchers.CreationDate, Users.RealName AS Creator, Vouchers.TotalPrice FROM Vouchers INNER JOIN Users ON Vouchers.CreatedBy = Users.ID WHERE (Vouchers.Voided = 0) AND (Vouchers.VendorID = '"& cusID & "') AND (Vouchers.Verified = 0) ORDER BY Vouchers.CreationDate DESC, Vouchers.id"
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
%>
					<tr>
						<td colspan="7" class="CusTD3">ÂÌç</td>
					</tr>
<%
				else
%>					<tr>
						<td class="CusTD3">#</td>
						<td class="CusTD3"># ›«ﬂ Ê— Œ—Ìœ</td>
						<td class="CusTD3">«ÌÃ«œ ﬂ‰‰œÂ</td>
						<td class="CusTD3"> «—ÌŒ «ÌÃ«œ</td>
						<td class="CusTD3">„—»Êÿ »Â ”›«—‘ Œ—Ìœ</td>
						<td class="CusTD3">„»·€</td>
					</tr>
<%					tmpCounter=0
					Do while not RS1.eof 
						RelatedPurchaseOrders = ""
						tmpsep=""

						mySQL="SELECT RelatedPurchaseOrderID FROM VoucherLines WHERE (Voucher_ID = '"& RS1("ID") & "')"

						Set RS2 = conn.execute(mySQL)
						Do While not RS2.eof 
							RelatedPurchaseOrders = RelatedPurchaseOrders & tmpsep & RS2("RelatedPurchaseOrderID")
							tmpsep=", "
							RS2.Movenext
						Loop
						RS2.close
						Set Rs2=Nothing
						tmpCounter = tmpCounter + 1
						if tmpCounter mod 2 = 1 then
							tmpColor="#FFFFFF"
							tmpColor2="#FFFFBB"
						Else
							tmpColor="#DDDDDD"
							tmpColor2="#EEEEBB"
						End if 
						'alert(this.getElementByTagName('td').items(0).data);
%>
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AP/AccountReport.asp?act=showVoucher&voucher=<%=RS1("ID")%>');">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD><%=RS1("ûCreator")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("CreationDate")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RelatedPurchaseOrders%>&nbsp;</TD>
							<TD><%=Separate(RS1("TotalPrice"))%>&nbsp;</TD>
						</TR>
<%						RS1.moveNext
					Loop
				end if
				%>
			</table>
		</Td>
	  </Tr>
	  <Tr>
		<Td colspan="2" valign="top" align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="10" class="CusTableHeader" style="background-color:#CCAA99;">›«ﬂ Ê— Â«Ì Œ—Ìœ ’«œ— ‘œÂ ( ”ÊÌÂ ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Vouchers.id, Vouchers.CreationDate, Users.RealName AS Creator, Vouchers.TotalPrice, Users_1.RealName AS Approver, APItems.CreatedDate AS ApprovedDate, APItems.RemainedAmount, APItems.EffectiveDate FROM Vouchers INNER JOIN Users ON Vouchers.CreatedBy = Users.ID INNER JOIN APItems ON Vouchers.id = APItems.Link INNER JOIN Users Users_1 ON APItems.CreatedBy = Users_1.ID WHERE (Vouchers.VendorID = '"& cusID & "') AND (Vouchers.Voided = 0) AND (Vouchers.Verified = 1) AND (APItems.RemainedAmount > 0) AND (APItems.Type = 6) ORDER BY Vouchers.CreationDate DESC, Vouchers.id"

'response.write mySQL
'response.end
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
%>
					<tr>
						<td colspan="10" class="CusTD3">ÂÌç</td>
					</tr>
<%
				else
%>					<tr>
						<td class="CusTD3">#</td>
						<td class="CusTD3">‘„«—Â</td>
						<td class="CusTD3">«ÌÃ«œ</td>
						<td class="CusTD3"> «—ÌŒ «ÌÃ«œ</td>
						<td class="CusTD3"> «ÌÌœ</td>
						<td class="CusTD3"> «—ÌŒ  «ÌÌœ</td>
						<td class="CusTD3"> «—ÌŒ „ÊÀ—</td>
						<td class="CusTD3">„»·€</td>
						<td class="CusTD3">„«‰œÂ</td>
					</tr>
<%					tmpCounter=0
					Do while not RS1.eof 
						tmpCounter = tmpCounter + 1
						if tmpCounter mod 2 = 1 then
							tmpColor="#FFFFFF"
							tmpColor2="#FFFFBB"
						Else
							tmpColor="#DDDDDD"
							tmpColor2="#EEEEBB"
						End if 
						'alert(this.getElementByTagName('td').items(0).data);
%>
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AP/AccountReport.asp?act=showVoucher&voucher=<%=RS1("ID")%>');">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD><%=RS1("ûCreator")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("CreationDate")%>&nbsp;</TD>
							<TD><%=RS1("Approver")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("ApprovedDate")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("EffectiveDate")%>&nbsp;</TD>
							<TD><%=Separate(RS1("TotalPrice"))%>&nbsp;</TD>
							<TD><%=Separate(RS1("RemainedAmount"))%>&nbsp;</TD>
						</TR>
<%						RS1.moveNext
					Loop
				end if
				%>
			</table>
		</Td>
	  </Tr>