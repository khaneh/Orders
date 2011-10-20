<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Purchase (4)
PageTitle=" ”›«—‘ Œ—Ìœ"
SubmenuItem=3
if not Auth(4 , 3) then NotAllowdToViewThisPage()

'OutService Page Order
'By Alix - Last changed: 81/01/13
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<style>
	.outServTable1 {font-family: Tahoma;font-size: 8pt; border:1 solid navy;}
	.outServTable1 th {height:25px;border-bottom:1 solid navy;Background-Color:ddddee}
	.outServTable1 td {height:25px;}
	.outServTR1 {Background-Color:#CCCCDD;}
</style>
<%
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- OutService Order Creation Form
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="«ÌÃ«œ ”›«—‘ Œ—Ìœ" then
	if request.form("outReq").count < 1 then%>
		<div align=center dir=rtl><br><br><br><br><br><br><b>Œÿ«!!<br><h3>ÂÌç œ—ŒÊ«” Ì «‰ Œ«» ‰‘œÂ «” </h3></b><br></div>
		</body>
		</html>
	<%	
	else
		%><br>
		<!--#include File="../include_JS_InputMasks.asp"-->
		<form method=post action="outServiceOrder.asp">
		<TABLE dir=rtl align=center width=600>
		<TR bgcolor="eeeeee" >
			<TD align=center colspan=7><B>«ÌÃ«œ ”›«—‘ Œ—Ìœ </B></TD>
		</TR>
		<TR bgcolor="eeeeee" >
				<TD>‰Ê⁄ œ—ŒÊ«” </TD>
				<TD> «—ÌŒ œ—ŒÊ«” </TD>
				<TD> «—ÌŒ ‰Ì«“</TD>
				<TD>„»·€ œ—ŒÊ«” </TD>
				<TD>‘„«—Â ”›«—‘</TD>
				<TD> Ê”ÿ</TD>
				<TD width=8%>œ—’œ</TD>
		</TR>
		<%
		tmpCounter=0
		pp=100
		totalPrice=0
		allComments = ""
		TypeID_IsService = ""
		for i=1 to request.form("outReq").count
			myRequestID = clng(request.form("outReq")(i))
			percent = cint(pp / (request.form("outReq").count - tmpCounter))
			pp = pp - percent
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
			set RSX=Conn.Execute ("SELECT * FROM purchaseRequests WHERE id = "& myRequestID )	
			if TypeID_IsService="" then
				TypeID_IsService = RSX("TypeID") & "_" & RSX("IsService")
			end if
			if TypeID_IsService <>  RSX("TypeID") & "_" & RSX("IsService") then
				conn.close
				response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br> œ— Ìﬂ ”›«—‘ Œ—Ìœ Â„Â œ—ŒÊ«”  Â« »«Ìœ Ìﬂ”«‰ »«‘‰œ.")
			end if
			totalPrice = totalPrice + RSX("price")
			totalQtty = totalQtty + RSX("qtty")
			Comment = RSX("Comment")
			if Comment<>"-" then
				Comment = " Ê÷ÌÕ: " & Comment
				allComments = allComments & Comment & chr(13) &  chr(13) 
			else
				Comment = " Ê÷ÌÕ ‰œ«—œ"
			end if
			set RST=Conn.Execute ("SELECT RealName FROM Users WHERE (ID = "& RSX("CreatedBy") & ")" )

			%>
			<TR bgcolor="<%=tmpColor%>" title="<%=Comment%>">
				<TD><%=RSX("typeName")%></TD>
				<TD><span dir=ltr><%=RSX("ReqDate")%></span></TD>
				<TD><span dir=ltr><%=RSX("DueDate")%></span></TD>
				<TD><%=RSX("price")%>
				<%
				if RSX("price") <> 0 then
					response.write "<br><small>" & RSX("priceComment") & "</small>"
				end if
				%>
				</TD>
				<TD>
				<% if clng(RSX("IsService"))=0 or clng(RSX("Qtty"))<>0 then %>
				( ⁄œ«œ: <%=RSX("qtty")%>)
				<% end if %>
				<% if clng(RSX("order_ID"))<>-1 then %>
				<%=RSX("order_ID")%>
				<% end if %>
				</TD>
				<TD><%=RST("RealName")%></TD>
				<TD width=8%><input type="hidden" name="reqID" value="<%=RSX("id")%>">
				<input type="text" size=2 name="percent" value="<%=percent%>">%
				</TD>
			</TR>
			<% 
			RSX_typeID = clng(RSX("typeID"))
			RSX_typeName = RSX("typeName")
			DueDate = RSX("DueDate")
			IsService =  RSX("IsService")
			RSX.moveNext
			if DueDate="" then DueDate =shamsiToday()

		next
		%>
		<TR bgcolor="eeeeee" >
			<TD align=center colspan=7 height=10></TD>
		</TR>
		<TR >
		<input type="hidden" Name='tmpDlgArg' value=''>
		<input type="hidden" Name='tmpDlgTxt' value=''>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		function selectVendor(){
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="‰«„ Õ”«»Ì —« ﬂÂ „Ì ŒÊ«ÂÌœ Ã” ÃÊ ﬂ‰Ìœ Ê«—œ ﬂ‰Ìœ:"
			window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			if (document.all.tmpDlgTxt.value !="") {
				window.showModalDialog('../AP/dialog_SelectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value), document.all.tmpDlgArg, 'dialogWidth:780px; dialogHeight:500px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					theSpan=document.getElementById("AccountSpan");
					theSpan.getElementsByTagName("input")[0].value=Arguments[0];
					theSpan.getElementsByTagName("span")[0].innerText=Arguments[1]+"["+Arguments[0]+"]";
				}
			}
		}

		//-->
		</SCRIPT>
			<TD colspan=1 align=left>›—Ê‘‰œÂ</TD>
			<TD colspan=6 align=right>
				<INPUT TYPE="hidden" name="IsService" value=<%=IsService%>>
				<span id="AccountSpan"><%' after any changes in this span "../AR/Customers.asp" must be revised%>
				<INPUT TYPE="hidden" NAME="vendor" value=""><b><span>................................</span></b>
				</span>
				<INPUT class="GenButton" TYPE="button" value="«‰ Œ«»" onClick="selectVendor();">
			</TD>
		</TR>
		<TR >
			<%
			if (IsService) then
				if RSX_typeID <> 0 then
					%>
					<TD colspan=1 align=left>‰Ê⁄ Œœ„ </TD>
					<TD colspan=6 align=right>
						<SELECT NAME="type"  class=inputBut  size="1" >
						<%
						set RS4 = conn.Execute ("SELECT * FROM OutServices")
						while not (RS4.eof) %>
							<OPTION value="<%=RS4("ID")%>" <%
							if RS4("ID")=RSX_typeID then
								response.write " selected "
							end if
							%>><%=RS4("Name")%></option>
							<%
							RS4.MoveNext
						wend
						RS4.close
						%>
						</SELECT>
					</TD>
				<% else %>
					<TD colspan=1 align=left>‰Ê⁄ Œœ„  Ì« ﬂ«·«</TD>
					<TD colspan=6 align=right>
						<INPUT TYPE="hidden" name=type value=0>
						<INPUT TYPE="text"  class=inputBut NAME="typeName"  size=51  value="<%=RSX_typeName%>">
					</td>
				<% end if %>
			<% else %>
				<TD colspan=1 align=left>‰Ê⁄ ﬂ«·«</TD>
				<TD colspan=6 align=right>
					<SELECT NAME="type"  class=inputBut  size="1" >
					<%
					set RS4 = conn.Execute ("SELECT * FROM InventoryItems")
					while not (RS4.eof) %>
						<OPTION value="<%=RS4("ID")%>" <%
						if RS4("ID")=RSX_typeID then
							response.write " selected "
						end if
						%>><%=RS4("Name")%> (<%=RS4("Unit")%>)</option>
						<%
						RS4.MoveNext
					wend
					RS4.close
					%>
					</SELECT>
				</TD>
			<% end if %>
		</TR>
		<TR >
			<TD colspan=1 align=left> ⁄œ«œ</TD>
			<TD colspan=6 align=right> 
				<INPUT TYPE="text" NAME="qtty"  class=inputBut size=51 value="<%=totalQtty%>" onKeyPress="return maskNumber(this);">
			</TD>
		</TR>
		<TR >
			<TD colspan=1 align=left>ﬁÌ„  „Ê—œ «‰ Ÿ«—</TD>
			<TD colspan=6 align=right> 
				<INPUT TYPE="text" NAME="price"  class=inputBut size=51 value="<%=totalPrice%>" onKeyPress="return maskNumber(this);">
			</TD>
		</TR>
		<TR >
			<TD colspan=1 align=left>œ— «Ì‰  «—ÌŒ »«Ìœ «—”«· ‘Êœ</TD>
			<TD colspan=6 align=right>
				<INPUT dir=ltr TYPE="text" NAME="date1"  class=inputBut size=51 value="<%=shamsiToday()%>" onKeyPress="return maskDate(this);"  onblur="acceptDate(this)">
			</TD>
		</TR>
		<TR >
			<TD colspan=1 align=left>œ— «Ì‰  «—ÌŒ »«Ìœ œ—Ì«›  ‘Êœ</TD>
			<TD colspan=6 align=right>
				<INPUT dir=ltr TYPE="text" NAME="date2"  class=inputBut size=51 value="<%=DueDate%>" onKeyPress="return maskDate(this);" onblur="acceptDate(this)">
			</TD>
		</TR>
		<TR >
			<TD colspan=1 align=left> Ê÷ÌÕ« </TD>
			<TD colspan=6 align=right>
				<TEXTAREA NAME="comment"  class=inputBut ROWS="7" COLS="32" ><%=allComments%></TEXTAREA>
			</TD>
		</TR>
		<TR >
			<TD colspan=7 align=center>
				<INPUT TYPE="submit"  class=inputBut Name="Submit" Value="À»  ”›«—‘"  style="width:100px;" tabIndex="14">
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
'-------------------------------------------------------------------------- Submit an OutService Order
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="À»  ”›«—‘" then
	vendorID =		request.form("vendor") 
	otypeID =		clng(request.form("type"))
	otype =			request.form("TypeName") 
	TotalPrice =	request.form("price") 
	comment =		request.form("comment") 
	ordDate =		shamsiToday()
	date1 =			request.form("date1") 
	date2 =			request.form("date2") 
	IsService =		request.form("IsService")
	qtty =			request.form("qtty") 

	a = qtty
	if isNumeric(a) and a<>"" then
		qttyLetter = ConvertIT(a)
	else
		response.write "<br>" 
		call showAlert ("Œÿ«!  ⁄œ«œ Ê«—œ ‰‘œÂ «” .", CONST_MSG_ERROR) 
		response.end
	end if

	if TotalPrice="" then TotalPrice=0
	if comment = "" then
		comment = "-"
	end if

	if vendorID="" then
		response.write "<br>" 
		call showAlert ("Œÿ«! ›—Ê‘‰œÂ «‰ Œ«» ‰‘œÂ «” .", CONST_MSG_ERROR) 
		response.end
	end if

	TypeID_IsService = otypeID
	if IsService then
		TypeID_IsService = TypeID_IsService & "_True"
		IsServiceFlag = 1
		if otypeID <> 0 then
			set RS4 = conn.Execute ("SELECT * FROM OutServices where ID=" & otypeID)
			if (RS4.eof) then
				otype="-unknown-"
			else
				otype=RS4("Name")
			end if
			RS4.close
		end if
	else
		TypeID_IsService = TypeID_IsService & "_False"
		IsServiceFlag = 0
		set RS4 = conn.Execute ("SELECT * FROM InventoryItems where ID=" & otypeID)
		if (RS4.eof) then
			otype="-unknown-"
		else
			otype=RS4("Name")
		end if
		RS4.close
	end if

	for i=1 to request.form("ReqID").count
		RequestID=	clng(request.form("ReqID")(i))
		set RSX=Conn.Execute ("SELECT * FROM purchaseRequests WHERE id = "& RequestID)	
		if TypeID_IsService <>  RSX("TypeID") & "_" & RSX("IsService") then
			conn.close
			response.write "<BR><BR>"
			call showAlert ("Œÿ«!<br>‰Ê⁄ œ—ŒÊ«”  Â«Ì Œ—Ìœ »« ‰Ê⁄ ”›«—‘ Œ—Ìœ „€«Ì—  œ«—œ.", CONST_MSG_ERROR) 
			response.end
		end if
		RSX.close
	next

	mySql="INSERT INTO purchaseOrders (Vendor_ID, typeName, typeID, comment, OrdDate, price, IsService, qtty, qttyLetter) VALUES ("& clng(VendorID) & ", N'"& sqlSafe(otype) & "', "& otypeID & ", N'"& sqlSafe(comment) & "',N'"& sqlSafe(ordDate) & "', "& sqlSafe(TotalPrice) & " , "& IsServiceFlag & ", "& sqlSafe(qtty) & ", N'" & sqlSafe(qttyLetter) & "');SELECT @@Identity AS NewPO"
	set RS4 = Conn.execute(mySQL).NextRecordSet
	OrdID = RS4 ("NewPO")
	RS4.close

	for i=1 to request.form("ReqID").count
		RequestID=request.form("ReqID")(i)
		Percent=request.form("Percent")(i)
		conn.Execute ("INSERT INTO purchaseRequestOrderRelations (Req_ID, Ord_ID, PercentOfAll) VALUES ("& RequestID & ", "& ordID & ", "& Percent & " )")
		Conn.Execute ("UPDATE purchaseRequests SET status = 'ord' WHERE ID = "& RequestID )	
		response.write RequestID & "- per: "
		response.write Percent & " % <br>"
	next

	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusTime, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", '"& currentTime10() & "', N'"& ordDate & "', "& 0 & ", N' «—ÌŒ «ÌÃ«œ' )")
'	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusTime, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", N'"& Date1 & "', "& 1 & ", N'œ— «Ì‰  «—ÌŒ »«Ìœ «—”«· ‘Êœ' )")
'	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusTime, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", N'"& Date2 & "', "& 2 & ", N'œ— «Ì‰  «—ÌŒ »«Ìœ œ—Ì«›  ‘Êœ' )")
'	
'	Changed By Kid 821224 Because There was An Error : Number of VALUES does not match the number of Fields.
'
	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", N'"& Date1 & "', "& 1 & ", N'œ— «Ì‰  «—ÌŒ »«Ìœ «—”«· ‘Êœ' )")
	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", N'"& Date2 & "', "& 2 & ", N'œ— «Ì‰  «—ÌŒ »«Ìœ œ—Ì«›  ‘Êœ' )")
	

	response.redirect "outServiceTrace.asp?od=" & ordID

end if


'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- Submit an Inventory Item Order
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="À»  ”›«—‘ ﬂ«·«" then
	vendorID =		request.form("vendor") 
	otypeID =		clng(request.form("type"))
	qtty = request.form("qtty") 
	comment = request.form("comment") 
	ordDate = shamsiToday()
	date1 = request.form("date1") 
	date2 = request.form("date2") 

	if qtty="" then qtty=0


	set RS4 = conn.Execute ("SELECT * FROM InventoryItems where ID=" & otypeID)
	if (RS4.eof) then
		otype="-unknown-"
	else
		otype=RS4("Name")
	end if
	RS4.close

	mySql="INSERT INTO purchaseOrders (Vendor_ID, typeName, typeID, comment, OrdDate, price) VALUES ("& VendorID & ", N'"& otype & "', "& otypeID & ", N'"& comment & "',N'"& ordDate & "', "& TotalPrice & " )"	
	conn.Execute mySql

	set RS4 = conn.Execute ("SELECT max(id) as nid FROM purchaseOrders where Vendor_ID=" & VendorID & " and ordDate='" & ordDate & "' and typeID=" & otypeID )
	if (RS4.eof) then
		response.write "BAD ERROR!@<br> Contact system admin"
		response.end
	else
		OrdID=RS4("nid")
	end if
	RS4.close

	for i=1 to request.form("ReqID").count
		RequestID=request.form("ReqID")(i)
		Percent=request.form("Percent")(i)
		conn.Execute ("INSERT INTO purchaseRequestOrderRelations (Req_ID, Ord_ID, PercentOfAll) VALUES ("& RequestID & ", "& ordID & ", "& Percent & " )")
		Conn.Execute ("update purchaseRequests SET status = 'ord' where id = "& RequestID )	
		response.write RequestID & "- per: "
		response.write Percent & " % <br>"
	next

	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", '"& ordDate & "', "& 0 & ", N' «—ÌŒ «ÌÃ«œ' )")
	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", '"& Date1 & "', "& 1 & ", N'œ— «Ì‰  «—ÌŒ »«Ìœ «—”«· ‘Êœ' )")
	conn.Execute ("INSERT INTO purchaseOrderStatus (Ord_ID, statusDate, statusCode, StatusDetail) VALUES ("& OrdID & ", '"& Date2 & "', "& 2 & ", N'œ— «Ì‰  «—ÌŒ »«Ìœ œ—Ì«›  ‘Êœ' )")
	

	response.redirect "outServiceTrace.asp?od=" & ordID

end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------ Main
'-----------------------------------------------------------------------------------------------------

sortBy=request("s")
if sortBy="4" then 
	sB="order_ID"
elseif sortBy="2" then 
	sB="ReqDate"
elseif sortBy="3" then 
	sB="price"
else 
	sB="typeName"
end if

%>
<br>
<SCRIPT LANGUAGE="JavaScript">
<!--
function checkValidation(){
	result = true;
	try{
		var slctn = document.getElementsByName("outReq");
		var TypeID_IsService = "";
		for(i=0;i<slctn.length;i++){
			if(slctn[i].checked){
				if (TypeID_IsService == ""){
					TypeID_IsService = slctn[i].parentNode.getElementsByTagName("INPUT")[0].value;
				}
				if (TypeID_IsService != slctn[i].parentNode.getElementsByTagName("INPUT")[0].value){
					tmpBG=slctn[i].parentNode.parentNode.bgColor;
					slctn[i].parentNode.parentNode.bgColor="red";
					slctn[i].focus();
					alert('Œÿ«!\n œ— Ìﬂ ”›«—‘ Œ—Ìœ Â„Â œ—ŒÊ«”  Â« »«Ìœ Ìﬂ”«‰ »«‘‰œ.');
					slctn[i].parentNode.parentNode.bgColor=tmpBG;
					result = false;
					break;
				}
			}
		}
	}
	catch(e){
		alert("Error: \n"+e.message)
		result = false;
	}
	return result;
}
//-->
</SCRIPT>
<FORM METHOD="POST" ACTION="OutServiceOrder.asp" OnSubmit="return checkValidation();">
<%
	set RSS=Conn.Execute ("SELECT * FROM purchaseRequests WHERE (Status = 'new' AND TypeID <> 0 AND IsService=1) ORDER BY " & sB)
	if not RSS.EOF then
%>
		<TABLE dir=rtl align=center width=600 class="outServTable1">
		<TR class="outServTR1">
			<TH align=center colspan=6>œ—ŒÊ«”  Â«Ì Œ—Ìœ Œœ„«  ÃœÌœ</TH>
		</TR>
		<TR class="outServTR1">
			<TD><INPUT TYPE="checkbox" NAME="" disabled></TD>
			<TD><A HREF="outServiceOrder.asp?s=1"><SMALL>‰Ê⁄ ⁄„·Ì« </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=2"><SMALL> «—ÌŒ œ—ŒÊ«” </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=3"><SMALL>ﬁÌ„ </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=3"><SMALL> Ê”ÿ</SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=4"><SMALL>‘„«—Â ”›«—‘</SMALL></A></TD>
		</TR>
<%

		tmpCounter=0
		Do while not RSS.eof
			set RST=Conn.Execute ("SELECT RealName FROM Users WHERE (ID = "& RSS("CreatedBy") & ")" )
			tmpCounter = tmpCounter + 1
			Comment = RSS("Comment")
			if Comment<>"-" then
				Comment = " Ê÷ÌÕ: " & Comment
			else
				Comment = " Ê÷ÌÕ ‰œ«—œ"
			end if
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
%>
			<TR bgcolor="<%=tmpColor%>" title="<%=Comment%>">
				<TD><INPUT TYPE="Hidden" Name="TypeID_IsService" Value="<%=RSS("TypeID")&"_"&RSS("IsService")%>"><INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RSS("id")%>"></TD>
				<TD><%=RSS("typeName")%></TD>
				<TD><span dir=ltr><%=RSS("ReqDate")%></span></TD>
				<TD><%=RSS("price")%></TD>
				<TD><%=RST("RealName")%></TD>
				<TD><a href="../shopfloor/manageOrder.asp?radif=<%=RSS("order_ID")%>"><%=RSS("order_ID")%></a></small></TD>
			</TR>
<% 
		RSS.moveNext
		Loop
%>
		</TABLE>
<%
	end if
	RSS.close
		  

	set RSS=Conn.Execute ("SELECT * FROM purchaseRequests WHERE (Status = 'new' AND IsService=0) ORDER BY " & sB)
	if not RSS.EOF then
%>
		<BR>
		<TABLE dir=rtl align=center width=600 class="outServTable1">
		<TR class="outServTR1">
			<TH align=center colspan=6>œ—ŒÊ«”  Â«Ì Œ—Ìœ ﬂ«·« Â«Ì «‰»«—</TH>
		</TR>
		<TR class="outServTR1">
			<TD><INPUT TYPE="checkbox" NAME="" disabled></TD>
			<TD><A HREF="outServiceOrder.asp?s=1"><SMALL>‰Ê⁄ ⁄„·Ì« </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=2"><SMALL> «—ÌŒ œ—ŒÊ«” </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=3"><SMALL>ﬁÌ„ </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=3"><SMALL> Ê”ÿ</SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=4"><SMALL>‘„«—Â ”›«—‘</SMALL></A></TD>
		</TR>
<%

		tmpCounter=0
		Do while not RSS.eof
			set RST=Conn.Execute ("SELECT RealName FROM Users WHERE (ID = "& RSS("CreatedBy") & ")" )
			tmpCounter = tmpCounter + 1
			Comment = RSS("Comment")
			if Comment<>"-" then
				Comment = " Ê÷ÌÕ: " & Comment
			else
				Comment = " Ê÷ÌÕ ‰œ«—œ"
			end if
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

			orderIDpos = clng(RSS("Order_ID"))
			if orderIDpos<>-1 then 
				orderLink="<small><a href='../shopfloor/manageOrder.asp?radif=" & orderIDpos & "'>" & orderIDpos & "</a></small>"
			else
				orderLink="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			end if
%>
			<TR bgcolor="<%=tmpColor%>" title="<%=Comment%>">
				<TD><INPUT TYPE="Hidden" Name="TypeID_IsService" Value="<%=RSS("TypeID")&"_"&RSS("IsService")%>"><INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RSS("id")%>"></TD>
				<TD><%=RSS("typeName")%></TD>
				<TD><span dir=ltr><%=RSS("ReqDate")%></span></TD>
				<TD><%=RSS("price")%></TD>
				<TD><%=RST("RealName")%></TD>
				<TD><%=orderLink%> ( ⁄œ«œ: <%=RSS("qtty")%>) </TD>
			</TR>
			  
<% 
		RSS.moveNext
		Loop
%>
		</TABLE>
<%
	end if
	RSS.close

	set RSS=Conn.Execute ("SELECT * FROM purchaseRequests WHERE (Status = 'new' AND TypeID=0) ORDER BY " & sB)

	if not RSS.EOF then
	%>
		<BR>
		<TABLE dir=rtl align=center width=600 class="outServTable1">
		<TR class="outServTR1">
			<TH align=center colspan=6>œ—ŒÊ«”  Â«Ì Œ—Ìœ ”«Ì— ﬂ«·«Â« Ê Œœ„« </TH>
		</TR>
		<TR class="outServTR1">
			<TD><INPUT TYPE="checkbox" NAME="" disabled></TD>
			<TD><A HREF="outServiceOrder.asp?s=1"><SMALL>‰Ê⁄ ⁄„·Ì« </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=2"><SMALL> «—ÌŒ œ—ŒÊ«” </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=3"><SMALL>ﬁÌ„ </SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=3"><SMALL> Ê”ÿ</SMALL></A></TD>
			<TD><A HREF="outServiceOrder.asp?s=4"><SMALL>‘„«—Â ”›«—‘</SMALL></A></TD>
		</TR>
<%
		tmpCounter=0
		Do while not RSS.eof
			set RST=Conn.Execute ("SELECT RealName FROM Users WHERE (ID = "& RSS("CreatedBy") & ")" )
			Comment = RSS("Comment")
			if Comment<>"-" then
				Comment = " Ê÷ÌÕ: " & Comment
			else
				Comment = " Ê÷ÌÕ ‰œ«—œ"
			end if
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
			orderIDpos = clng(RSS("Order_ID"))
			if orderIDpos<>-1 then 
				orderLink="<small><a href='../shopfloor/manageOrder.asp?radif=" & orderIDpos & "'>" & orderIDpos & "</a></small>"
			else
				orderLink="&nbsp;"
			end if
		
			if (RSS("qtty")="" or RSS("qtty")="0") then
				if orderIDpos = -1 then
					orderQtty = "-"
				else
					orderQtty = ""
				end if
			else
				orderQtty ="( ⁄œ«œ:"& RSS("qtty") & ")"
			end if 
%>
			<TR bgcolor="<%=tmpColor%>" title="<%=Comment%>">
				<TD><INPUT TYPE="Hidden" Name="TypeID_IsService" Value="<%=RSS("TypeID")&"_"&RSS("IsService")%>"><INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RSS("id")%>"></TD>
				<TD><%=RSS("typeName")%></TD>
				<TD><span dir=ltr><%=RSS("ReqDate")%></span></TD>
				<TD><%=RSS("price")%></TD>
				<TD><%=RST("RealName")%></TD>
				<TD><%=orderLink%> <%=orderQtty%> </TD>
			</TR>
<% 
		RSS.moveNext
		Loop
%>
		</TABLE>
<%
	end if
	RSS.close
%>
<BR>
<center><INPUT TYPE="submit" Name="Submit" Value="«ÌÃ«œ ”›«—‘ Œ—Ìœ"  style="width:150px;" tabIndex="14"></center>
<br>
</FORM>
<%
Conn.Close
%>

<!--#include file="tah.asp" -->