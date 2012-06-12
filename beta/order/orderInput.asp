<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="Ê—Êœ ”›«—‘"
SubmenuItem=1
if not Auth(2 , 1) then NotAllowdToViewThisPage()

'Copy from old trace/orderTraceInput.asp
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style >
	.mySection{border: 1px #F90 dashed;margin: 15px 10px 0 15px;}
	.myRow{border: 2px #F05 dashed;margin: 10px 0 10px 0;padding: 0 3px 5px 0;}
	.exteraArea{border: 1px #33F dotted;margin: 5px 0 0 5px;padding: 0 3px 5px 0;}
	.myLabel {margin: 0 3px 0 0;white-space: nowrap;}
	.myProp {font-weight: bold;color: #40F; margin: 0 3px 0 3px;}
	div.btn label{background-color:yellow;color: blue;padding: 3px 30px 3px 30px;cursor: pointer;}
	div.btn{margin: -5px 250px 0px 5px;}
	div.btn img{margin: 0px 20px -5px 0;cursor: pointer;}
</style>
<%
if request("act")="submitsearch" then
	if isnumeric(request("CustomerNameSearchBox")) then
		conn.close
		response.redirect "?act=getorder&selectedCustomer=" & request("CustomerNameSearchBox")
	elseif request("CustomerNameSearchBox") <> "" then
		SA_TitleOrName=request("CustomerNameSearchBox")
		mySQL="SELECT * FROM Accounts WHERE (REPLACE(AccountTitle, ' ', '') LIKE REPLACE(N'%"& sqlSafe(SA_TitleOrName) & "%', ' ', '') ) ORDER BY AccountTitle"
		Set RS1 = conn.Execute(mySQL)

		if (RS1.eof) then 
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='../CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
		end if 

		SA_TitleOrName=request("CustomerNameSearchBox")
		SA_Action="return true;"
		SA_SearchAgainURL="OrderInput.asp"
		SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<br>
		<FORM METHOD=POST ACTION="?act=getType">
			<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	end if
elseif Request.QueryString("act")="getType" then 
	customerID=request("selectedCustomer")
	mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)

	if (RS1.eof) then 
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='../CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
	end if 
	'if (cdbl(rs1("arBalance"))+cdbl(rs1("apBalance"))+cdbl(rs1("aoBalance"))+cdbl(rs1("creditLimit")) < 0) then 
	if (cdbl(rs1("arBalance"))+cdbl(rs1("creditLimit")) < 0) then 
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("»œÂÌ «Ì‰ Õ”«» «“ „Ì“«‰ «⁄ »«— ¬‰ »Ì‘ — ‘œÂ°<br> ·ÿ›« »« ”—Å—”  ›—Ê‘ Â„«Â‰ê ﬂ‰Ìœ.<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>‰„«Ì‘ Õ”«»</a>")
	end if
%>
		<br>
		<div dir='rtl'>
			<B>ê«„ ”Ê„ : ê—› ‰ ‰Ê⁄ ”›«—‘</B>
		</div>
		<form method="post" action="?act=getorder">
			<input name="selectedCustomer" type="hidden" value="<%=customerID%>">
			<SELECT NAME="OrderType" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px' tabIndex="7">
			<OPTION value="-1" style='color:red;'>«‰ Œ«» ﬂ‰Ìœ</option>
<%
			Set RS2 = conn.Execute("SELECT [User] as ID, DefaultOrderType FROM UserDefaults WHERE ([User] = "& session("ID") & ") OR (UserDefaults.[User] = 0) ORDER BY ABS(UserDefaults.[User]) DESC")
			defaultOrderType=RS2("DefaultOrderType")
			RS2.close
			Set RS2 = Nothing

			set RS_TYPE=Conn.Execute ("SELECT ID, Name FROM OrderTraceTypes WHERE (IsActive=1) ORDER BY ID")
			Do while not RS_TYPE.eof	
%>
				<OPTION value="<%=RS_TYPE("ID")%>" <%if RS_TYPE("ID")=defaultOrderType then response.write "selected"%>><%=RS_TYPE("Name")%></option>
<%
			RS_TYPE.moveNext
			loop
			RS_TYPE.close
			set RS_TYPE = nothing
%>		
			</SELECT>
			<input type="submit" value="«œ«„Â">
		</form>
<%	

elseif request("act")="getorder" then
	customerID=request("selectedCustomer")
	mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)

	if (RS1.eof) then 
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='..//CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
	end if 

	AccountNo=RS1("ID")
	AccountTitle=RS1("AccountTitle")
	companyName=RS1("CompanyName")
	customerName=RS1("Dear1")& " " & RS1("FirstName1")& " " & RS1("LastName1")
	Tel=RS1("Tel1")
	set rs=Conn.Execute("select * from OrderTraceTypes where id=" & request("orderType"))
	if rs.eof then 
		conn.close
		response.redirect "?errmsg=" & server.urlEncode("‰Ê⁄ ”›«—‘ —«  ⁄ÌÌ‰ ﬂ‰Ìœ")
	end if
	orderTypeID=rs("id")
	orderTypeName=rs("name")
	set orderProp = server.createobject("MSXML2.DomDocument")
	hasProperty=false
	if rs("property")<>"" then 
		orderProp.loadXML(rs("property"))
		hasProperty=true
	end if
	rs.close
	set rs=nothing
	creationDate=shamsiToday()
	creationTime=time
	creationTime=Hour(creationTime)&":"&Minute(creationTime)
	if instr(creationTime,":")<3 then creationTime="0" & creationTime
	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)
%>
	<br>
	<div dir='rtl'><B>ê«„ çÂ«—„ : ê—› ‰ ”›«—‘</B>
	</div>
	<br>
<!-- ê—› ‰ ”›«—‘ -->
	<hr>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" width="700" align="center">
	<FORM METHOD=POST ACTION="?act=submitorder" onSubmit="return checkValidation();">
	<TR bgcolor="black">
		<TD align="left"><FONT COLOR="YELLOW">Õ”«»:</FONT></TD>
		<TD align="right" colspan=3 height="25px">
			<FONT COLOR="YELLOW"><%=customerID & " - "& AccountTitle%></FONT>
			<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>">
		</TD>
		<td align="left"><font color="yellow">‰Ê⁄ ”›«—‘:</font></td>
		<td>
			<font color="red"><b><%=orderTypeName%></b></font>
			<input type="hidden" name="orderType" value="<%=orderTypeID%>">
		</td>
	</TR>
	<TR bgcolor="black">
		<TD align="left"><FONT COLOR="YELLOW">‘„«—Â ”›«—‘:</FONT></TD>
		<TD align="right">
			<!-- Radif -->
			<INPUT disabled TYPE="text" NAME="Radif" maxlength="6" size="8" tabIndex="1" dir="LTR" value="######">
		</TD>
		<TD align="left"><FONT COLOR="YELLOW"> «—ÌŒ:</FONT></TD>
		<TD><TABLE border="0">
			<TR>
				<TD dir="LTR">
					<INPUT disabled TYPE="text" maxlength="10" size="10" value="<%=CreationDate%>">
					<INPUT TYPE="hidden" NAME="OrderDate" value="<%=CreationDate%>">
				</TD>
				<TD dir="RTL"><FONT COLOR="YELLOW"><%=weekdayname(weekday(date))%></FONT></TD>
			</TR>
			</TABLE></TD>
		<TD align="left"><FONT COLOR="YELLOW">”«⁄ :</FONT></TD>
		<TD align="right">
			<INPUT disabled TYPE="text" maxlength="5" size="3" dir="LTR" value="<%=creationTime%>">
			<INPUT TYPE="hidden" NAME="OrderTime" value="<%=creationTime%>"></TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD align="left">‰«„ ‘—ﬂ :</TD>
		<TD align="right">
			<!-- CompanyName -->
			<INPUT TYPE="text" NAME="CompanyName" maxlength="50" size="25" tabIndex="2" value="<%=companyName%>"></TD>
		<TD title=" ÊÃÂ ›—„«ÌÌœ ﬂÂ œ— ’Ê—  Å— ‘œ‰ «Ì‰  «—ÌŒ° €Ì— ﬁ«»· ÊÌ—«Ì‘ ŒÊ«Âœ »Êœ" align="left"> «—ÌŒ  ÕÊÌ· ﬁ—«—œ«œ:</TD>
		<TD>
			<TABLE border="0">
				<TR>
					<TD dir="LTR">
						<INPUT TYPE="text" NAME="ReturnDate" onblur="acceptDate(this)" maxlength="10" size="10" >
					</TD>
					<TD dir="RTL">(?‘‰»Â)</TD>
				</TR>
			</TABLE>
		</TD>
		<TD align="left">”«⁄   ÕÊÌ·:</TD>
		<TD align="right">
			<INPUT TYPE="text" NAME="ReturnTime" maxlength="5" size="3" dir="LTR" onKeyPress="return maskTime(this);" >
		</TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD align="left">‰«„ „‘ —Ì:</TD>
		<TD align="right">
			<!-- CustomerName -->
			<INPUT TYPE="text" NAME="CustomerName" maxlength="50" size="25" tabIndex="3" value="<%=customerName%>">
		</TD>
		<TD align="left"> ·›‰:</TD>
		<TD align="right">
			<!-- Telephone -->
			<INPUT TYPE="text" NAME="Telephone" maxlength="50" size="25" tabIndex="4" value="<%=Tel%>">
		</TD>
		<TD title="Â— ê«Â ﬂÂ ·«“„ »«‘Â «Ì‰ ›Ì·œ ﬁ«»· ÊÌ—«Ì‘ ŒÊ«Âœ »Êœ" align="left"> «—ÌŒ  ÕÊÌ· ⁄„·Ì:</TD>
		<TD>
			<INPUT TYPE="text" NAME="actualReturn_date" onblur="acceptDate(this)" maxlength="10" size="10" tabIndex="5">
		</td>
	</TR>
	<TR bgcolor="#CCCCCC">
		<td align="left">”«Ì“:</td>
		<td>
			<input type="text" name="paperSize" tabindex="6">
		</td>
		<td align="left"> Ì—«é:</td>
		<td>
			<input type="text" name="qtty" tabindex="7">
		</td>
		<td align="left">ﬁÌ„  ﬂ·:</td>
		<td colspan="3">
			<input type="text" name="totalPrice" id='totalPrice' style="background-color:#FED;border-width:0;" <%if hasProperty then response.write " readonly='readonly' "%>  tabindex="8">
		</td>
	</TR>
	<TR bgcolor="#CCCCCC">
		
		<TD align="left">⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì·:</TD>
		<TD align="right" colspan="3">
			<INPUT TYPE="text" NAME="OrderTitle" maxlength="255" size="50" tabIndex="9">
		</TD>
		<TD align="left">”›«—‘ êÌ—‰œÂ:</TD>
		<TD>
			<INPUT Type="Text" readonly NAME="SalesPerson" value="<%=CSRName%>" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px' tabIndex="888">
		</TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD colspan="6" height="30px">&nbsp;</TD>
	</TR>
	<tr bgcolor="#CCCCCC">
	<td colspan="6">
<%
if hasProperty then 
sub fetchKeys(key)
	oldGroup="---first---"
	oldLabel="---first---"
	thisRow="<div class='myRow'>"
	thisRow = thisRow & "<input type='hidden' value='0' id='" & Replace(key,"/","-") & "-maxID'>"
	thisRow = thisRow & "<div class='exteraArea' id='" & Replace(key,"/","-") & "-0'>"
' 	thisRow = thisRow & "<img title='Õ–› «Ì‰ Œÿ' src='/images/cancelled.gif' onclick='$(""#" & Replace(key,"/","-") & "-0"").remove();'>"
' 	thisRow = thisRow & "<img title='Õ–› «Ì‰ Œÿ' src='/images/cancelled.gif' onclick='$(this).parent().remove();'>" 
	For Each myKey In orderProp.SelectNodes(key)
	  thisType = myKey.GetAttribute("type")
	  thisName = myKey.GetAttribute("name")
	  thisLabel= myKey.GetAttribute("label")
	  thisGroup= myKey.GetAttribute("group")
	  'response.write thisName & ": " & thisLabel &"(" &thisType& ")" & "<br>"
	  if (oldGroup<>thisGroup and oldGroup <> "---first---") then thisRow = thisRow &  "</div>"
	  if oldGroup<>thisGroup then 
	  	thisRow = thisRow & "<div class='mySection' groupName='" & thisGroup & "'>"
	  	if myKey.GetAttribute("disable")="1" then 
			thisRow = thisRow &  "<input type='checkbox' value='0' name='" & thisGroup & "-disBtn' onclick='disGroup(this);"
			if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " calc_" & myKey.GetAttribute("group") & "(this);"
			thisRow = thisRow & "'>"
			disText=" disabled='disabled' "
		else
			disText=""
		end if
	  end if
	  if oldLabel<>thisLabel then thisRow = thisRow &  "<label class='myLabel'>" & thisLabel & "</label>"
		
	  select case thisType
	  	case "option"
	  		thisRow = thisRow &  "<select " & disText & " style='margin:0;padding:0;' name='" & thisName & "'"
	  		if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onchange='calc_" & myKey.GetAttribute("group") & "(this);' "
	  		thisRow = thisRow & ">"
	  		for each myOption in myKey.getElementsByTagName("option")
	  			thisRow = thisRow &  "<option value='" & myOption.text & "'"
	  			if myOption.GetAttribute("price")<>"" then 
	  				thisRow = thisRow & " price='" & myOption.GetAttribute("price") & "' "
	  			end if 
	  			thisRow = thisRow &">" & myOption.GetAttribute("label") & "</option>"
		  	next
		  	thisRow = thisRow &  "</select>"
		case "option-other"
			thisRow = thisRow &  "<select " & disText & " style='margin:0;padding:0;' name='" & thisName & "' onchange='checkOther(this);"
			if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " calc_" & myKey.GetAttribute("group") & "(this);"
			thisRow = thisRow & "'>"
	  		for each myOption in myKey.getElementsByTagName("option")
	  			thisRow = thisRow &  "<option value='" & myOption.text & "'"
	  			if myOption.GetAttribute("price")<>"" then 
	  				thisRow = thisRow & " price='" & myOption.GetAttribute("price") & "' "
	  			end if
	  			thisRow = thisRow & ">" & myOption.GetAttribute("label") & "</option>"
		  	next
		  	thisRow = thisRow &  "<option value='-1'>”«Ì—</option></select>"	
		  	thisRow = thisRow & "<input type='text' name='" &thisName & "-addValue' onblur='addOther(this);'>"	  	
		case "text"
			thisRow = thisRow &  "<input " & disText & " type='text' class='myInput' size='" & myKey.text & "' name='" & thisName & "' "
			if myKey.GetAttribute("readonly")="yes" then thisRow =thisRow & " readonly='readonly' "
			if myKey.GetAttribute("default")<>"" then thisRow = thisRow & "value='" & myKey.GetAttribute("default") & "'"
			if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onblur='calc_" & myKey.GetAttribute("group") & "(this);' "
			thisRow = thisRow & ">"
		case "text area"
			thisRow = thisRow &  "<textarea name='" & thisName & "' style='width:600px;' cols='" & myKey.text & "'></textarea>"
		case "check"
			thisRow = thisRow & "<input type='checkbox' value='on-0' name='" & thisName & "' "
			if myKey.text="checked" then thisRow = thisRow & "checked='checked'"
			if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onclick='calc_" & myKey.GetAttribute("group") & "(this);' "
			if IsNumeric(myKey.GetAttribute("price")) then thisRow = thisRow & " price='" & myKey.GetAttribute("price") & "' "
			thisRow = thisRow & ">"
		case "radio":
				thisRow = thisRow & "<input " & disText & " type='radio' value='" & myKey.text & "' name='" & thisName & "'" 
				if myKey.GetAttribute("default")="yes" then thisRow = thisRow & " checked='checked' "
				if myKey.GetAttribute("blur")="yes" then thisRow = thisRow & " onchange='calc_" & myKey.GetAttribute("group") & "(this);' "
				thisRow = thisRow & ">"
	  end select
	  if myKey.GetAttribute("force")="yes" then thisRow = thisRow &  "<span style='color:red;margin:0 0 0 2px;padding:0;'>*</span>"
	  oldGroup=thisGroup
	  oldLabel=thisLabel
	  if myKey.GetAttribute("br")="yes" then thisRow = thisRow & "<br>"
	Next
	thisRow = thisRow & "</div></div><div id='extreArea" &Replace(key,"/","-")& "'></div>"
	response.write thisRow 'prependTo 
	response.write "<div class='btn'><img title='«÷«›Â' src='/images/Plus-32.png' width='20px' onclick='cloneRow(""" & key & """);'><img title='Õ–› ¬Œ—Ì‰ Œÿ' src='/images/cancelled.gif' onclick='removeRow(""" & key & """);'></div></div>"
end sub

	oldTmp="---first---"
	for each tmp in orderProp.selectNodes("//key")
		if oldTmp<>tmp.parentNode.nodeName then 
			oldTmp=tmp.parentNode.nodeName
			call fetchKeys("/keys/" & oldTmp & "/key")
		end if
	next
' 	call fetchKeys("/keys/printing/key")
' 	call fetchKeys("keys/binding/key")
' 	call fetchKeys("keys/service/key")
' 	call fetchKeys("keys/delivery/key")
end if
%>			

	</td>
	</tr>
	<TR bgcolor="#CCCCCC">
		<TD colspan="6">
		<TABLE align="center" width="50%" border="0">
		<TR>
			<TD><INPUT TYPE="submit" Name="Submit" Value=" «ÌÌœ" style="width:100px;" tabIndex="14"></TD>
			<TD><INPUT TYPE="hidden" NAME="Price" maxlength="10" size="9" dir="LTR" tabIndex="13" value="‰«„‘Œ’">&nbsp;</TD>
			<TD align="left"><INPUT TYPE="button" Name="Cancel" Value="«‰’—«›" style="width:100px;" onClick="window.location='OrderInput.asp';" tabIndex="15"></TD>
		</TR>
		</TABLE>
		</TD>
	</TR>
	</FORM>
	</TABLE>
	<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
	<script type="text/javascript" src="calcOrder.js"></script>
	<SCRIPT LANGUAGE="JavaScript">
		
	<!--
		function checkValidation(){
		//TRIM : str = str.replace(/^\s*|\s*$/g,""); 
		if(document.all.CustomerName.value.replace(/^\s*|\s*$/g,'') == ''){
			alert("‰«„ „‘ —Ì —« Ê«—œ ﬂ‰Ìœ")
			document.all.CustomerName.focus();
			return false;
		}
		else if(document.all.SalesPerson.value.replace(/^\s*|\s*$/g,"") == ''){
			alert("”›«—‘ êÌ—‰œÂ —« Ê«—œ ﬂ‰Ìœ")
			document.all.SalesPerson.focus();
			return false;
		}
		/*else if(document.all.ReturnDate.value.replace(/^\s*|\s*$/g,'') == ''){
			alert("„Ê⁄œ  ÕÊÌ· —« Ê«—œ ﬂ‰Ìœ")
			document.all.ReturnDate.focus();
			return false;
		}
		else if(document.all.ReturnTime.value.replace(/^\s*|\s*$/g,'') == ''){
			alert("“„«‰ (”«⁄ )  ÕÊÌ· —« Ê«—œ ﬂ‰Ìœ")
			document.all.ReturnTime.focus();
			return false;
		}*/
		else if(document.all.OrderTitle.value.replace(/^\s*|\s*$/g,'') == ''){
			alert("⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì· —« Ê«—œ ﬂ‰Ìœ")
			document.all.OrderTitle.focus();
			return false;
		}
		else if(document.all.paperSize.value.replace(/^\s*|\s*$/g,'') == ''){
			alert("”«Ì“ —« Ê«—œ ﬂ‰Ìœ")
			document.all.paperSize.focus();
			return false;
		}
		else if(document.all.qtty.value.replace(/^\s*|\s*$/g,'') == ''){
			alert(" Ì—«é —« Ê«—œ ﬂ‰Ìœ")
			document.all.qtty.focus();
			return false;
		}
		else{
			document.all.Submit.disabled=true;
			return true;
		}
	}

		document.all.CompanyName.focus();
	//-->
	</SCRIPT>
<%elseif request("act")="submitorder" then
	CreationDate=shamsiToday()
	CustomerID=request.form("CustomerID") 
	if CustomerID="" OR not isNumeric(CustomerID) then
		conn.close
		response.redirect "?act=getorder&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("Œÿ« œ— ‰«„ „‘ —Ì<br>«ÿ·«⁄«  Ê«—œ ‰‘œ...<BR>")
	end if

'	TraceID=request.form("radif")
'
'	if TraceID="" OR not isNumeric(TraceID) then
'		conn.close
'		response.redirect "?act=getorder&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("Œÿ« œ— ‘„«—Â ”›«—‘<br>«ÿ·«⁄«  Ê«—œ ‰‘œ...<BR>")
'	end if
'
'	set RS1=Conn.Execute ("SELECT * FROM Orders WHERE (ID='"& TraceID & "')")
'	if not RS1.eof then
'		conn.close
'		response.redirect "?act=getorder&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("‘„«—Â ”›«—‘  ﬂ—«—Ì «” <br>«ÿ·«⁄«  Ê«—œ ‰‘œ...<BR>")
'	end if
'	RS1.close

	orderType=request.form("OrderType")
	if orderType="" OR not isNumeric(OrderType) then
		conn.close
		response.redirect "?act=getorder&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ ”›«—‘<br>«ÿ·«⁄«  Ê«—œ ‰‘œ...<BR>")
	end if
	set rs=Conn.Execute("select * from OrderTraceTypes where id=" & orderType)
	if rs.eof then 
		conn.close
		response.redirect "?errmsg=" & server.urlEncode("‰Ê⁄ ”›«—‘ —«  ⁄ÌÌ‰ ﬂ‰Ìœ")
	end if
	set RS1=Conn.Execute ("SELECT Name FROM OrderTraceTypes WHERE (ID='"& orderType & "') AND (IsActive=1)")
	if RS1.eof then
		conn.close
		response.redirect "?act=getorder&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ ”›«—‘<br>«ÿ·«⁄«  Ê«—œ ‰‘œ...<BR>")
	else
		orderTypeName=RS1("Name")
	end if
	RS1.close
	'----------------------------------------------------------------------------
	'----------------------------------------------------------------------------
	'----------------------------------------------------------------------------
	orderTypeID=rs("id")
	orderTypeName=rs("name")
	set orderProp = server.createobject("MSXML2.DomDocument")
	if rs("property")<>"" then 
		orderProp.loadXML(rs("property"))
		myXML = fetchKeyValues()
	end if
	rs.close
	set rs=nothing
	
function fetchKeyValues()
	key="---first---"
	thisRow="<?xml version=""1.0""?><keys>"
	for each tmp in orderProp.selectNodes("//key")
		if key<>tmp.parentNode.nodeName then 
			key=tmp.parentNode.nodeName
			thisRow = thisRow & "<" & key & ">"
			hasValue=0
			For Each myKey In orderProp.SelectNodes("/keys/" & key & "/key")
				thisName = myKey.GetAttribute("name")
				thisGroup= myKey.GetAttribute("group")
				id=0
'					response.write oldName& "<br>"
				if thisName<>oldName then 
					for each value in request.form(thisName)
						if value <> "" then 
							thisRow = thisRow & "<key name=""" & thisName & """ id=""" 
							select case myKey.GetAttribute("type") 
								case "check"
									thisRow = thisRow & mid(value,4)
								case else
									if request.form(thisGroup & "-disBtn")<>"" then 
										thisRow = thisRow & trim(split(request.form(thisGroup & "-disBtn"),",")(id))
									else
										thisRow = thisRow & id 
									end if
							end select
							thisRow = thisRow & """>" & value & "</key>"
							hasValue=hasValue +1
						end if
						id=id+1
					next
				end if
				oldName = thisName
			Next
			
			if hasValue>0 then 
				thisRow = thisRow & "</" & key & ">"
			else
				thisRow = Replace(thisRow,"<" & key & ">","")
			end if
		end if
	Next
	thisRow = thisRow & "</keys>"
	fetchKeyValues = thisRow 
	'response.write thisRow
	'response.end
end function

	mySQL="INSERT INTO Orders (CreatedDate, CreatedBy, Customer) VALUES ('"& CreationDate & "', '"& session("ID") & "', '"& CustomerID & "');SELECT @@Identity AS NewOrder"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	OrderID = RS1 ("NewOrder")
	RS1.close
	'--------------------------------SAM-----------------------------------------
	'dim fs,f
	'set fs = Server.CreateObject("Scripting.FileSystemObject")
	'set f = fs.CreateFolder("d:\OrdersStorage\" & cstr(OrderID))
	'set f = nothing
	'set fs = nothing
	'----------------------------------------------------------------------------
	qtty = 0
	if IsNumeric(request.form("qtty")) then 
		qtty = CInt(request.form("qtty"))
	end if
	actualReturn_date = "null"
	ReturnDate = "null"
	ReturnTime = "null"
	if request.form("actualReturn_date")<>"" then actualReturn_date = "'" & request.form("actualReturn_date") & "'"
	if request.form("actualReturn_date")<>"" then ReturnDate = "'" & request.form("ReturnDate") & "'"
	if request.form("actualReturn_date")<>"" then ReturnTime = "'" & request.form("ReturnTime") & "'"
		
	mySQL="INSERT INTO orders_trace (radif_sefareshat, order_date, order_time, return_date, return_time, company_name, customer_name, telephone, order_title, order_kind, Type, vazyat, marhale, salesperson, status, step, LastUpdatedDate, LastUpdatedTime, LastUpdatedBy,property,qtty,paperSize,price, actualReturn_date) VALUES ('"&_
	OrderID & "', N'"& sqlSafe(request.form("OrderDate")) & "', "& ReturnTime & ", "& ReturnDate & ", N'"& sqlSafe(request.form("ReturnTime")) & "', N'"& sqlSafe(request.form("CompanyName")) & "', N'"& sqlSafe(request.form("CustomerName")) & "', N'"& sqlSafe(request.form("Telephone")) & "', N'"& sqlSafe(request.form("OrderTitle")) & "', N'"& orderTypeName & "', '"& orderType & "', N'œ— Ã—Ì«‰', N'œ— ’› ‘—Ê⁄', N'"& sqlSafe(request.form("SalesPerson")) & "', 1, 1, N'"& CreationDate & "',  N'"& currentTime10() & "', "& session("ID") & " ,N'" & myXML& "'," & qtty & ",'" & request.form("paperSize") & "','" & request.form("totalPrice") & "'," & actualReturn_date & ")"	
	response.write mySQL
	conn.Execute(mySQL)

	conn.close
	response.redirect "../order/TraceOrder.asp?act=show&order=" & OrderID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  ”›«—‘ À»  ‘œ")
else%>
<!-- Ã” ÃÊ »—«Ì ‰«„ Õ”«» -->
	<br>
	<FORM METHOD=POST ACTION="?act=submitsearch" onsubmit="if (document.all.CustomerNameSearchBox.value=='') return false;">
	<div dir='rtl'><B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«»</B>
		<INPUT TYPE="text" NAME="CustomerNameSearchBox">&nbsp;
		<INPUT TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.CustomerNameSearchBox.focus();
	//-->
	</SCRIPT>
<%
end if
conn.Close
%>
</font>
<!--#include file="tah.asp" -->