	  <Tr><Td colspan="2" height="10px">
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
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
	.CusTD5 {background-color: yellow; direction: LTR; text-align: center; font-size:9pt;}
	.from {color: blue;}
	.to {color: red;}
</STYLE>


	  &nbsp;</Td></Tr>
	  <Tr>
		<Td colspan="2" valign="top" align="center">
				<% if Auth(1 , 3) then %><input class="GenButton" type="button" value="ÊÌ—«Ì‘" onclick="window.location='../CRM/AccountEdit.asp?act=editaccount&selectedCustomer=<%=cusID%>';"><% end if %>
				<input class="GenButton" type="button" value="‰Ê‘ ‰ Ì«œœ«‘ "onclick="window.location = '../home/message.asp?sendTo=<%=RS1("CSR")%>&RelatedTable=accounts&RelatedID=<%=cusID%>&retURL=<%=Server.URLEncode("../CRM/AccountInfo.asp?act=show&selectedCustomer="&RS1("ID"))%>';">
				<BR><BR>
		</Td>
	  </Tr>

	  <Tr><Td valign="top" align="left">
		<table class="CustTable3" cellspacing='1' cellspacing='1'>
<!-- 		<tr>
			<td class="CusTD1" width="80">⁄‰Ê«‰ Õ”«» : </td>
			<td class="CusTD2"><span dir="RTL"><%=RS1("AccountTitle")%></span></td>
		</tr> -->
		<tr>
			<td class="CusTD1">‘„«—Â Õ”«» :</td>
			<td class="CusTD2"><%=RS1("ID")%></td>
		</tr>
<%		if not RS1("IsPersonal") then %>
			<tr>
				<td class="CusTD1">‰«„ ‘—ﬂ  : </td>
					<td class="CusTD2"><%=RS1("CompanyName")%></td>
			</tr>
<%		end if%>
		<tr>
			<td class="CusTD1"> «—ÌŒ «ÌÃ«œ : </td>
			<td class="CusTD2"><%=RS1("CreatedDate")%></td>
		</tr>
		<tr>
			<td class="CusTD1">„”ÊÊ· ÅÌêÌ—Ì : </td>
			<td class="CusTD2"><%=RS1("CSRName")%></td>
		</tr>
		<tr>
			<td class="CusTD1">„«‰œÂ ›—Ê‘ : </td>
			<td class="CusTD2"><%=Separate(RS1("ARBalance"))%></td>
		</tr>
		<tr>
			<td class="CusTD1">„«‰œÂ Œ—Ìœ: </td>
			<td class="CusTD2"><%=Separate(RS1("APBalance"))%></td>
		</tr>
		<%if Auth("B" , 0) then %>
		<tr>
			<td class="CusTD1">„«‰œÂ ”«Ì—:</td>
			<td class="CusTD2"><%=Separate(RS1("AOBalance"))%></td>
		</tr>
		<%end if %>
		<tr>
			<td class="CusTD1">”ﬁ› «⁄ »«— : </td>
			<td class="CusTD2"><%=separate(RS1("CreditLimit"))%></td>
		</tr>
		<tr>
			<td class="CusTD1"><%If CBool(RS1("IsPersonal")) Then response.write "‘„«—Â „·Ì:" Else response.write "‘„«—Â À» :" End if%></td>
			<td class="CusTD2"><%=RS1("NorRCode")%></td>
		</tr>
		<tr>
			<td class="CusTD1">ﬂœ «ﬁ ’«œÌ: </td>
			<td class="CusTD2"><%=RS1("EconomicalCode")%></td>
		</tr>
		<%
		dim typeColor(4)
		typeColor(1)="blue"
		typeColor(2)="fuchsia"
		typeColor(3)="red"
		typeColor(4)="lime"
		set rs2= conn.Execute("select * from accountGroupRelations inner join AccountGroups on accountGroupRelations.AccountGroup=AccountGroups.id where account="&RS1("ID"))
		if not rs2.eof then 

		%>
		<tr>
			<td class="CusTD1"><%If CBool(rs2("isPartner")) Then response.write "Â„ﬂ«—" Else response.write " Ã«—Ì" End if%></td>
			<td class="CusTD2"><%=rs2("Name")%></td>			
		</tr>
		<%
		end if
		rs2.close
		set rs2=nothing
		%>
		</table>
	  </Td>
	  <Td valign="top" align="center">
		<table class="CustTable3" cellspacing='1'>
		<tr>
			<td class="CusTD1" valign="top">—«»ÿ «’·Ì : 
			
			<BR><BR><BR><BR>
			Å”  <INPUT TYPE="checkbox" NAME="postable1" disabled <% if RS1("postable1") then response.write "checked" %> >
			</td>
			<td style="background-color: <%=typeColor(cint(rs1("type")))%>;">
				<TABLE class="CustContactTable">
					<TR>
						<TD><%=RS1("Dear1")%> - <%=RS1("FirstName1")%> - <%=RS1("LastName1")%></TD>
						<TD width="140" style="border-right:1 solid black;"> ·›‰ : <%=echoTels(RS1("Tel1"))%></TD>
					</TR>
					<TR>
						<TD rowspan="2" valign="top">¬œ—” : <%=RS1("City1")%> - <%=RS1("Address1")%></TD>
						<TD style="border-right:1 solid black;">›«ﬂ” : <%=echoTels(RS1("Fax1"))%></TD>
					</TR>
					<TR>
						<TD style="border-right:1 solid black;">«Ì„Ì· : 
							<span><%=RS1("Email1")%></span>
							<%if RS1("email1")<>"" then %>
							<span title="»« ﬂ·Ìﬂ —ÊÌ «Ì‰ œﬂ„Â »Â ’Ê—  ŒÊœﬂ«— Ìﬂ «Ì„Ì· »Â „‘ —Ì «—”«· ŒÊ«Âœ ‘œ ﬂÂ Õ«ÊÌ ‘„«—Â „‘ —Ì ÊÌ ŒÊ«Âœ »Êœ">
								<form method="post" action="http://my.pdhco.com/sendMail.php">
									<input type="hidden" name="id" value="<%=RS1("ID")%>">
									<input type="hidden" name="Email" value="<%=RS1("Email1")%>">
									<input type="hidden" name="AccountTitle" value="<%=RS1("AccountTitle")%>">
									<input type="hidden" name="Dear" value="<%=RS1("Dear1")%>">
									<input type="hidden" name="FirstName" value="<%=RS1("FirstName1")%>">
									<input type="hidden" name="LastName" value="<%=RS1("LastName1")%>">
									<input type="submit" name="jameSend" value="‘„«—Â „‘ —Ì «Ì„Ì· ‘Êœ">
								</form>
							</span>
							<%end if%>
						</TD>
					</TR>
					<TR>
						<TD>ﬂœÅ” Ì: <%=RS1("PostCode1")%></TD>
						<TD style="border-right:1 solid black;">„Ê»«Ì·: <span  Dir="LTR"><%=echoTels(RS1("Mobile1"))%></span></TD>
					</TR>
				</TABLE>
			</td>
		</tr>
		<tr>
			<td class="CusTD1" valign="top">—«»ÿ œÊ„ : 
			
			<BR><BR><BR><BR>
			Å”  <INPUT TYPE="checkbox" NAME="postable2" disabled <% if RS1("postable2") then response.write "checked" %> >
			
			</td>
			<td style="background-color: <%=typeColor(cint(rs1("type")))%>;">
				<TABLE class="CustContactTable">
					<TR>
						<TD><%=RS1("Dear2")%> - <%=RS1("FirstName2")%> - <%=RS1("LastName2")%></TD>
						<TD width="140"> ·›‰ : <%=echoTels(RS1("Tel2"))%></TD>
					</TR>
					<TR>
						<TD rowspan="2" valign="top">¬œ—” : <%=RS1("City2")%> - <%=RS1("Address2")%></TD>
						<TD>›«ﬂ” : <%=echoTels(RS1("Fax2"))%></TD>
					</TR>
					<TR>
						<TD>«Ì„Ì· : 
							<span><%=RS1("Email2")%></span>
							<%if RS1("email2")<>"" then %>
							<span title="»« ﬂ·Ìﬂ —ÊÌ «Ì‰ œﬂ„Â »Â ’Ê—  ŒÊœﬂ«— Ìﬂ «Ì„Ì· »Â „‘ —Ì «—”«· ŒÊ«Âœ ‘œ ﬂÂ Õ«ÊÌ ‘„«—Â „‘ —Ì ÊÌ ŒÊ«Âœ »Êœ">
								<form method="post" action="http://my.pdhco.com/sendMail.php">
									<input type="hidden" name="id" value="<%=RS1("ID")%>">
									<input type="hidden" name="Email" value="<%=RS1("Email2")%>">
									<input type="hidden" name="AccountTitle" value="<%=RS1("AccountTitle")%>">
									<input type="hidden" name="Dear" value="<%=RS1("Dear2")%>">
									<input type="hidden" name="FirstName" value="<%=RS1("FirstName2")%>">
									<input type="hidden" name="LastName" value="<%=RS1("LastName2")%>">
									<input type="submit" name="jameSend" value="‘„«—Â „‘ —Ì «Ì„Ì· ‘Êœ">
								</form>
							</span>
							<%end if%>
						</TD>
					</TR>
					<TR>
						<TD>ﬂœÅ” Ì: <%=RS1("PostCode2")%></TD>
						<TD>„Ê»«Ì·: <span  Dir="LTR"><%=RS1("Mobile2")%></span></TD>
					</TR>
				</TABLE>
			</td>
		</tr>
		</table>
	  </Td></Tr>
	  <Tr>
	  	<Td colspan="2"><BR><BR>
			<div id="orderMessages"></div>
			<input id="customerID" type="hidden" value="<%=RS1("ID")%>"/>
			<script type="text/javascript">
				$(document).ready(function(){
					TransformXmlURL("/service/xml_getMessage.asp?act=related&table=accounts&id=" + $("#customerID").val(),"/xsl/showRelatedMessage.xsl", function(result){
						$("#orderMessages").html(result);
						$("td.msgBody").each(function(i){
							$(this).html($(this).html().replace(/\n/gi,"<br/>"));
							});
							$("input[name=addNewMessage]").click(function(){
								document.location="../home/message.asp?RelatedTable=accounts&RelatedID=" + $("#customerID").val() + "&retURL=" + escape("../CRM/AccountInfo.asp?act=show&selectedCustomer=" + $("#customerID").val());
						});
					});
				});
			</script>	  	
<BR><BR>
			<%
			mySQL="SELECT AccountsEditLog.EditOn, AccountsEditLog.AutoKey, Users.RealName FROM AccountsEditLog INNER JOIN Users ON AccountsEditLog.EditBy = Users.ID where AccountsEditLog.id=" & cusID & " ORDER BY AccountsEditLog.AutoKey"
			Set RS1 = conn.execute(mySQL)
			if not RS1.eof then
			%>
			<table class="CustTable" cellspacing='1'>
				<tr>
					<td colspan="2" class="CusTableHeader"> «—ÌŒçÂ ÊÌ—«Ì‘  «ÿ·«⁄«  «Ì‰ Õ”«»</td>
				</tr>
				<%
				tmpCounter=0
				while not RS1.eof 
					tmpCounter=tmpCounter+1
				%>
					<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>">
						<td dir=ltr><%=RS1("EditOn")%></td>
						<td dir=rtl>
						<span style="cursor:hand" onclick="window.open('AccountsEditLogViewer.asp?AutoKey=<%=RS1("AutoKey")%>','AccLog','menubar=no, width=500')"> Ê”ÿ <%=RS1("RealName")%></span>
						</td>
					</tr>
				<%
					RS1.moveNext
				wend
				%>
			</table><BR><BR>
			<% end if %>
	  </Td></Tr>
