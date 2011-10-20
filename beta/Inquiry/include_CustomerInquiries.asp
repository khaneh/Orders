<%'
  'This Include file needs the variables below:
  '
  '		Conn	
  '		cusID	(Customer ID)
  '		AccountTitle
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
	  <TaBlE style="width:100%; background-color:#C3DBEB;" cellspacing="2" cellspacing="2">
	  <Tr>
			<Td align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr><td class="CusTableHeader" style="background-color: #CCCCCC; height:50;">
					<B><A target="_blank" HREF="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=cusID%>">„—»Êÿ »Â <%=AccountTitle%> (<%=CusID%>)</A></B></td>
				</tr>
			</table>
			</Td>
	  </Tr>
	  <Tr>
		<Td colspan="2" valign="top" align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="7" class="CusTableHeader" style="background-color:#CCCC99;">›«ﬂ Ê— Â«Ì œ— Ã—Ì«‰ ( «ÌÌœ ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Inquiries.ID, Inquiries.CreatedDate, Users.RealName AS Creator, Inquiries.TotalReceivable FROM Inquiries INNER JOIN Users ON Inquiries.CreatedBy = Users.ID WHERE (Inquiries.Customer='"& cusID & "') AND (Inquiries.Voided = 0) AND (Inquiries.Approved = 0) ORDER BY Inquiries.CreatedDate DESC, Inquiries.ID"
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
						<td class="CusTD3">‘„«—Â ›«ﬂ Ê—</td>
						<td class="CusTD3">«ÌÃ«œ ﬂ‰‰œÂ</td>
						<td class="CusTD3"> «—ÌŒ</td>
						<td class="CusTD3">„—»Êÿ »Â ”›«—‘ Â«Ì</td>
						<td class="CusTD3">„»·€</td>
					</tr>
<%					tmpCounter=0
					Do while not RS1.eof 
						RelatedOrders = ""
						tmpsep=""
						mySQL="SELECT [Order] FROM InquiryOrderRelations WHERE Inquiry='"& RS1("ID") & "'"
						Set RS2 = conn.execute(mySQL)
						Do While not RS2.eof 
							RelatedOrders = RelatedOrders & tmpsep & RS2("Order")
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
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.location='../AR/AccountReport.asp?act=Inquiry&inquiry=<%=RS1("ID")%>';">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD><%=RS1("ûCreator")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RelatedOrders%>&nbsp;</TD>
							<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
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
					<td colspan="7" class="CusTableHeader" style="background-color:#33BB99;">›«ﬂ Ê— Â«Ì  «ÌÌœ ‘œÂ (’«œ— ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Inquiries.ID, Inquiries.CreatedDate, Users.RealName AS Creator, Inquiries.ApprovedDate, Inquiries.TotalReceivable, Users_1.RealName AS Approver FROM Inquiries INNER JOIN Users ON Inquiries.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Inquiries.ApprovedBy = Users_1.ID WHERE (Inquiries.Customer='"& cusID & "') AND (Inquiries.Voided = 0) AND (Inquiries.Issued = 0) ORDER BY Inquiries.CreatedDate DESC, Inquiries.ID"
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
						<td class="CusTD3">‘„«—Â ›«ﬂ Ê—</td>
						<td class="CusTD3">«ÌÃ«œ ﬂ‰‰œÂ</td>
						<td class="CusTD3"> «—ÌŒ</td>
						<td class="CusTD3"> «ÌÌœ ﬂ‰‰œÂ</td>
						<td class="CusTD3"> «—ÌŒ  «ÌÌœ</td>
						<td class="CusTD3">„»·€</td>
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
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.location='../AR/AccountReport.asp?act=showInquiry&inquiry=<%=RS1("ID")%>';">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD><%=RS1("ûCreator")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
							<TD><%=RS1("Approver")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("ApprovedDate")%>&nbsp;</TD>
							<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
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
					<td colspan="10" class="CusTableHeader" style="background-color:#CCAA99;">›«ﬂ Ê— Â«Ì ’«œ— ‘œÂ ( ”ÊÌÂ ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Inquiries.ID, Inquiries.CreatedDate, Users.RealName AS Creator, Inquiries.ApprovedDate, Inquiries.TotalReceivable, Users_1.RealName AS Approver,  Users_2.RealName AS Issuer, Inquiries.IssuedDate, ARItems.RemainedAmount FROM Inquiries INNER JOIN Users ON Inquiries.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Inquiries.ApprovedBy = Users_1.ID INNER JOIN Users Users_2 ON Inquiries.IssuedBy = Users_2.ID INNER JOIN ARItems ON Inquiries.ID = ARItems.Link WHERE (Inquiries.Customer = '"& cusID & "') AND (Inquiries.Voided = 0) AND (Inquiries.Issued = 1) AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) ORDER BY Inquiries.CreatedDate DESC, Inquiries.ID"
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
						<td class="CusTD3">’œÊ—</td>
						<td class="CusTD3"> «—ÌŒ ’œÊ—</td>
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
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.location='../AR/AccountReport.asp?act=showInquiry&inquiry=<%=RS1("ID")%>';">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD><%=RS1("ûCreator")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
							<TD><%=RS1("Approver")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("ApprovedDate")%>&nbsp;</TD>
							<TD><%=RS1("Issuer")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("IssuedDate")%>&nbsp;</TD>
							<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
							<TD><%=Separate(RS1("RemainedAmount"))%>&nbsp;</TD>
						</TR>
<%						RS1.moveNext
					Loop
				end if
				%>
			</table>
		</Td>
	  </Tr>
 	</TaBlE>
