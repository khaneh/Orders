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
		<FORM METHOD=POST ACTION="?act=getorder">
			<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	end if
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
	
	creationDate=shamsiToday()
	creationTime=time
	creationTime=Hour(creationTime)&":"&Minute(creationTime)
	if instr(creationTime,":")<3 then creationTime="0" & creationTime
	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)
%>
	<br>
	<div dir='rtl'><B>ê«„ ”Ê„ : ê—› ‰ ”›«—‘</B>
	</div>
	<br>
<!-- ê—› ‰ ”›«—‘ -->
	<hr>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" width="700" align="center">
	<FORM METHOD=POST ACTION="?act=submitorder" onSubmit="return checkValidation();">
	<TR bgcolor="black">
		<TD align="left"><FONT COLOR="YELLOW">Õ”«»:</FONT></TD>
		<TD align="right" colspan=5 height="25px">
			<FONT COLOR="YELLOW"><%=customerID & " - "& AccountTitle%></FONT>
			<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>">
		</TD>
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
		<TD align="left">„Ê⁄œ  ÕÊÌ·:</TD>
		<TD><TABLE border="0">
			<TR>
				<TD dir="LTR"><INPUT TYPE="text" NAME="ReturnDate" onblur="acceptDate(this)" maxlength="10" size="10" tabIndex="5"></TD>
				<TD dir="RTL">(?‘‰»Â)</TD>
			</TR>
			</TABLE></TD>
		<TD align="left">”«⁄   ÕÊÌ·:</TD>
		<TD align="right"><INPUT TYPE="text" NAME="ReturnTime" maxlength="5" size="3" dir="LTR" tabIndex="6"  onKeyPress="return maskTime(this);" ></TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD align="left">‰«„ „‘ —Ì:</TD>
		<TD align="right">
			<!-- CustomerName -->
			<INPUT TYPE="text" NAME="CustomerName" maxlength="50" size="25" tabIndex="3" value="<%=customerName%>"></TD>
		<TD align="left">‰Ê⁄ ”›«—‘:</TD>
		<TD>
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
		</SELECT></TD>
		<TD align="left">”›«—‘ êÌ—‰œÂ:</TD>
		<TD><INPUT Type="Text" readonly NAME="SalesPerson" value="<%=CSRName%>" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px' tabIndex="888">
		</TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD align="left"> ·›‰:</TD>
		<TD align="right">
			<!-- Telephone -->
			<INPUT TYPE="text" NAME="Telephone" maxlength="50" size="25" tabIndex="4" value="<%=Tel%>"></TD>
		<TD align="left">⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì·:</TD>
		<TD align="right" colspan="4"><INPUT TYPE="text" NAME="OrderTitle" maxlength="255" size="50" tabIndex="9"></TD>
	</TR>
	<TR bgcolor="#CCCCCC">
		<TD colspan="6" height="30px">&nbsp;</TD>
	</TR>
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
		else if(document.all.ReturnDate.value.replace(/^\s*|\s*$/g,'') == ''){
			alert("„Ê⁄œ  ÕÊÌ· —« Ê«—œ ﬂ‰Ìœ")
			document.all.ReturnDate.focus();
			return false;
		}
		else if(document.all.ReturnTime.value.replace(/^\s*|\s*$/g,'') == ''){
			alert("“„«‰ (”«⁄ )  ÕÊÌ· —« Ê«—œ ﬂ‰Ìœ")
			document.all.ReturnTime.focus();
			return false;
		}
		else if(document.all.OrderType.value == -1){
			alert("‰Ê⁄ ”›«—‘ —« Ê«—œ ﬂ‰Ìœ")
			document.all.OrderType.focus();
			return false;
		}
		else if(document.all.OrderTitle.value.replace(/^\s*|\s*$/g,'') == ''){
			alert("⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì· —« Ê«—œ ﬂ‰Ìœ")
			document.all.OrderTitle.focus();
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

	set RS1=Conn.Execute ("SELECT Name FROM OrderTraceTypes WHERE (ID='"& orderType & "') AND (IsActive=1)")
	if RS1.eof then
		conn.close
		response.redirect "?act=getorder&selectedCustomer=" & CustomerID & "&errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ ”›«—‘<br>«ÿ·«⁄«  Ê«—œ ‰‘œ...<BR>")
	else
		orderTypeName=RS1("Name")
	end if
	RS1.close

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
	mySQL="INSERT INTO orders_trace (radif_sefareshat, order_date, order_time, return_date, return_time, company_name, customer_name, telephone, order_title, order_kind, Type, vazyat, marhale, salesperson, status, step, LastUpdatedDate, LastUpdatedTime, LastUpdatedBy) VALUES ('"&_
	OrderID & "', N'"& sqlSafe(request.form("OrderDate")) & "', N'"& sqlSafe(request.form("OrderTime")) & "', N'"& sqlSafe(request.form("ReturnDate")) & "', N'"& sqlSafe(request.form("ReturnTime")) & "', N'"& sqlSafe(request.form("CompanyName")) & "', N'"& sqlSafe(request.form("CustomerName")) & "', N'"& sqlSafe(request.form("Telephone")) & "', N'"& sqlSafe(request.form("OrderTitle")) & "', N'"& orderTypeName & "', '"& orderType & "', N'œ— Ã—Ì«‰', N'œ— ’› ‘—Ê⁄', N'"& sqlSafe(request.form("SalesPerson")) & "', 1, 1, N'"& CreationDate & "',  N'"& currentTime10() & "', "& session("ID") & " )"	
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