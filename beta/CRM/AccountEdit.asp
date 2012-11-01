<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'CRM (1)
PageTitle= "«ÿ·«⁄«  Õ”«»"
if request("act")="getAccount" or  request("act")="submitcustomer" then
	SubmenuItem=2
	if not Auth(1 , 2) then NotAllowdToViewThisPage()
else
	SubmenuItem=3
	if not Auth(1 , 3) then NotAllowdToViewThisPage()
end if

editFlag = 0
AccType = 2 

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<style>
	.CustGenTable  { font-family:tahoma; font-size: 9pt;}
	.CustGenInput { font-family:tahoma; font-size: 9pt;}
</STYLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;

var tempKeyBuffer;
function myKeyDownHandler(){
	tempKeyBuffer=window.event.keyCode;
}
function myKeyPressHandler(){
//	alert (tempKeyBuffer)
	if (tempKeyBuffer>=65 && tempKeyBuffer<=90){
		window.event.keyCode=tempKeyBuffer+32;
	}
	else if(tempKeyBuffer==186){
		window.event.keyCode=59;
	}
	else if(tempKeyBuffer==188){
		window.event.keyCode=44;
	}
	else if(tempKeyBuffer==190){
		window.event.keyCode=46;
	}
	else if(tempKeyBuffer==191){
		window.event.keyCode=47;
	}
	else if(tempKeyBuffer==192){
		window.event.keyCode=96;
	}
	else if(tempKeyBuffer>=219 && tempKeyBuffer<=221){
		window.event.keyCode=tempKeyBuffer-128;
	}
	else if(tempKeyBuffer==222){
		window.event.keyCode=39;
	}
}
//-->
</SCRIPT>
<font face="tahoma">
<!-- <div dir='rtl'><B>Ê—Êœ ”›«—‘ </B>
</div> -->
<%
if request("act")="submitsearch" then
	if request("CustomerNameSearchBox") <> "" then
		SA_TitleOrName=request("CustomerNameSearchBox")
		mySQL="SELECT * FROM Accounts WHERE (REPLACE(AccountTitle, ' ', '') LIKE REPLACE(N'%"& SA_TitleOrName & "%', ' ', '') ) ORDER BY AccountTitle"
		Set RS1 = conn.Execute(mySQL)

		if (RS1.eof) then 
					response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='../CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")

		' Not Found	%>
		<div dir='rtl'><B>ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ</B> &nbsp; <A HREF="AccountEdit.asp" style='font-size:7pt;'>Ã” ÃÊÌ „Ãœœ</A>
		</div><br>
	<%	else 
			SA_TitleOrName=request("CustomerNameSearchBox")
			SA_Action="return true;"
			SA_SearchAgainURL="AccountEdit.asp"
			SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<FORM METHOD=POST ACTION="AccountEdit.asp?act=editaccount">
			<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
		end if
	end if
elseif request("act")="editaccount" AND (request("selectedCustomer") <> ""  or request("NextOf") <> "" or request("PrevOf") <> "") then 

	'--------------------------------- START
	NextOf=request("NextOf")
	PrevOf=request("PrevOf")
	CusID=request("selectedCustomer")
	if NextOf <> "" AND isNumeric(NextOf) then
		NextOf=clng(NextOf)
		mySQL="SELECT TOP 1 AccountTitle FROM Accounts WHERE (ID = '"& NextOf & "')"
		set RS1=Conn.execute(mySQL)
		theTitle = RS1("AccountTitle")

		'mySQL="SELECT TOP 1 ID FROM Accounts WHERE (AccountTitle > N'"& theTitle & "') OR ((ID > '"& NextOf & "') AND (AccountTitle = N'"& theTitle & "')) Order BY AccountTitle, ID"

		'-------- changed by Alix (83-4-9) : be khaste mohaghegh tartibe hesaab haa az Alefbaa be Shomare Hesab tagheer kard
		mySQL="SELECT TOP 1 ID FROM Accounts WHERE (ID > '"& NextOf & "') Order BY ID"


		set RS1=Conn.execute(mySQL)
		if not Rs1.eof then
			CusID = RS1("ID")
			LinkToNext="<a href=""?act=show&NextOf="& CusID & """>»⁄œÌ &gt;</a>"
		else
			CusID = NextOf
			LinkToNext="»⁄œÌ ‰œ«—œ"
		end if
		LinkToPrev="<a href=""?act=show&PrevOf="& CusID & """>&lt; ﬁ»·Ì</a>"
	elseif PrevOf <> "" AND isNumeric(PrevOf) then
		PrevOf = clng(PrevOf)
		mySQL="SELECT TOP 1 AccountTitle FROM Accounts WHERE (ID = '"& PrevOf & "')"
		set RS1=Conn.execute(mySQL)
		theTitle = RS1("AccountTitle")

		'mySQL="SELECT TOP 1 ID FROM Accounts WHERE (AccountTitle < N'"& theTitle & "') OR ((ID < '"& PrevOf & "') AND (AccountTitle = N'"& theTitle & "')) Order BY AccountTitle DESC, ID DESC"

		mySQL="SELECT TOP 1 ID FROM Accounts WHERE (ID < '"& PrevOf & "') Order BY ID DESC"
		set RS1=Conn.execute(mySQL)
		if not Rs1.eof then
			CusID = RS1("ID")
			LinkToPrev="<a href=""?act=show&PrevOf="& CusID & """>&lt; ﬁ»·Ì</a>"
		else
			CusID = PrevOf 
			LinkToPrev="ﬁ»·Ì ‰œ«—œ"
		end if
		LinkToNext="<a href=""?act=show&NextOf="& CusID & """>»⁄œÌ &gt;</a>"
	elseif CusID <> "" AND isNumeric(CusID) then
		CusID=clng(CusID)
		LinkToNext="<a href=""?act=show&NextOf="& CusID & """>»⁄œÌ &gt;</a>"
		LinkToPrev="<a href=""?act=show&PrevOf="& CusID & """>&lt; ﬁ»·Ì</a>"
	else
		Conn.close
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â Õ”«» „⁄ »— ‰Ì” ")
	end if
	'--------------------------------- END
	mySQL="SELECT * FROM Accounts WHERE (ID='"& CusID & "')"
	CustomerID=CusID
	Set RS1 = conn.Execute(mySQL)
	editFlag = 1
	ID				=	RS1("id")
	AccountTitle	=	RS1("AccountTitle")
	CreditLimit		=	RS1("CreditLimit")
	IsPersonal		=	RS1("IsPersonal")
	IsADefault		=	RS1("IsADefault")
	EconomicalCode	=	RS1("EconomicalCode")
	CSR				=	RS1("CSR")
	CompanyName		=	RS1("CompanyName")
	AccType			=	RS1("Type")
	Status			=	RS1("Status")
	Postable1		=	RS1("Postable1")
	Dear1			=	RS1("Dear1")
	FirstName1		=	RS1("FirstName1")
	LastName1		=	RS1("LastName1")
	City1			=	RS1("City1")
	Address1		=	RS1("Address1")
	Tel1			=	RS1("Tel1")
	Fax1			=	RS1("Fax1")
	Email1			=	RS1("Email1")
	Mobile1			=	RS1("Mobile1")
	PostCode1		=	RS1("PostCode1")
	Postable2		=	RS1("Postable2")
	Dear2			=	RS1("Dear2")
	FirstName2		=	RS1("FirstName2")
	LastName2		=	RS1("LastName2")
	City2			=	RS1("City2")
	Address2		=	RS1("Address2")
	Tel2			=	RS1("Tel2")
	Fax2			=	RS1("Fax2")
	Email2			=	RS1("Email2")
	Mobile2			=	RS1("Mobile2")
	PostCode2		=	RS1("PostCode2")
	NorRCode		=	RS1("NorRCode")
	employee		=	RS1("employee")
	JobTitle1		=	RS1("JobTitle1")
	JobTitle2		=	RS1("JobTitle2")
	website			=	RS1("website")
end if
if request("act")="getAccount" or editFlag = 1 then
if CSR = "" then CSR = session("ID")
%>
<!-- Ê—Êœ «ÿ·«⁄«  „‘ —Ì  -->
		<FORM METHOD=POST ACTION="AccountEdit.asp?act=<% if editFlag = 0 then%>submitcustomer<%else%>submitEdit<%end if%>" onsubmit="if (document.all.AccountTitle.value=='') return false;">
		<br><br><INPUT TYPE="hidden" name=ID value=<%=ID%>>
		<table class="CustGenTable" Border="0" align='center' Width="600" Cellspacing="1" Cellpadding="5" Dir="RTL">
			<tr bgcolor='#C3C300'>
				<td align='center' colspan="5"><b>
				<% if editFlag = 0 then
					response.write "Ê—Êœ «ÿ·«⁄«  Õ”«»  "& CusID
				   else
					response.write "ÊÌ—«Ì‘ «ÿ·«⁄«  Õ”«»  "& CusID
				   end if %>
				</b></td>
			</tr>
			<tr>
				<td align='left' <% if not Auth(1,7) and editFlag=1 then response.write "title='‘„« „Ã«“ »Â ÊÌ—«Ì‘ «Ì‰ ¬Ì „ ‰Ì” Ìœ!'"%>><span <% if Auth(1,7) or editFlag=0 then response.write "onclick='setAccountTitle();'" %>>⁄‰Ê«‰ Õ”«» :</span></td>
				<td colspan='3' <% if not Auth(1,7) and editFlag=1 then response.write "title='‘„« „Ã«“ »Â ÊÌ—«Ì‘ «Ì‰ ¬Ì „ ‰Ì” Ìœ!'"%>>
					<INPUT class="CustGenInput" TYPE="text" NAME="AccountTitle" VALUE="<%=AccountTitle%>" size="30" MaxLength="200" onfocus="//if (this.value=='') this.value=document.all.Dear.value+' '+document.all.Name.value;" <%if Not Auth(1,7) and editFlag=1 then response.write "readonly"%>> 
					<INPUT class="CustGenInput" TYPE="checkbox" NAME="IsPersonal" <% if IsPersonal then %> checked <% end if %> onclick='showCompanyName();'><span>Õ”«» ‘Œ’Ì «” </span>
				</td>
				<td width='130px' align='left'>„”Ê·: 
					<%if Auth(1 , 4) then %>
					<select name="CSR" class="CustGenInput" style="width:90">
						<option value=""> «‰ Œ«» ﬂ‰Ìœ </option>
					<% set RSV=Conn.Execute ("SELECT * FROM Users ORDER BY RealName") ' WHERE Display=1 
					Do while not RSV.eof
					%>
						<option value="<%=RSV("ID")%>" <%
							if RSV("ID")=CSR then
								response.write " selected "
							end if
							%>><%=RSV("RealName")%></option>
					<%
					RSV.moveNext
					Loop
					RSV.close
					%>

					</select> 
					<% else %>
<%
						set RSV=Conn.Execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
						Do while not RSV.eof
							if RSV("ID")=CSR then
								a = RSV("RealName")
							end if
						RSV.moveNext
						Loop
						RSV.close					
%> 						
						<INPUT TYPE="text" NAME="" value="<%=a%>" size=9 readonly>
						<INPUT TYPE="hidden" NAME="CSR" value="<%=CSR%>" size=9 readonly>
					<% end if %>
					
				</td >
			</tr>
			<tr>
				<td align='left'>‰«„ ‘—ﬂ  : </td>
				<td colspan='1'><INPUT class="CustGenInput" TYPE="text"  <% if IsPersonal then %> style="visibility:hidden;" <% end if %>NAME="CompanyName" value="<%=CompanyName%>" size="30" onfocus="//if (this.value=='') this.value=document.all.Dear.value+' '+document.all.Name.value;">
				</td>
				<td align='right' colspan=2>&nbsp;</td>
				<td align='left'>‰Ê⁄:
					<select name="type" class="CustGenInput">
					<%
					set rs=Conn.Execute("select * from accountTypes")
					while not rs.eof
						%>
						<option <% if AccType = rs("id") then response.write "selected" %> value=<%=rs("id")%>><%=rs("name")%></option>	
						<%
						rs.moveNext
					wend
					rs.close
%> 
					</select>
				</td>
			</tr>
			<tr>
				<td colspan='2' align=left>ﬂœ «ﬁ ’«œÌ : </td>
				<td align='right' colspan=2>
					<INPUT class="CustGenInput" NAME="EconomicalCode" value="<%=EconomicalCode%>" TYPE="text" size="15" maxlength="20" style="direction:LTR;"> 
				</td>
				<td>
				    <INPUT class="CustGenInput" TYPE="checkbox" NAME="IsADefault" <% if IsADefault or editFlag<>1 then response.write "checked" %>>ÅÌ‘ ›—÷ «·›</td>
			</tr>
			<tr>
				<td colspan='2' align='left'>”ﬁ› «⁄ »«—: </td>
				<td align='right' colspan=2>
					<INPUT dir=LTR class="CustGenInput" TYPE="text" name=CreditLimit value="<%=separate(CreditLimit)%>" size=12 onblur="this.value=val2txt(txt2val(this.value))" maxlength=15 <%if not Auth(1 , 4) then response.write "readonly" %>> —Ì«·
				</td>
				<td align='left'>Ê÷⁄Ì :
<%				if editFlag=0 then Status=1
				if Auth(1,5) then%>	
					<select name="Status" class="CustGenInput">
					<option <% if Status = 1 then response.write "selected" %> value=1>›⁄«·
					<option <% if Status = 2 then response.write "selected" %> value=2>„ Êﬁ›
					<option <% if Status = 3 then response.write "selected" %> value=3>„”œÊœ
					</select>
				<%else
					select case Status
						Case 1: StatusText = "›⁄«·"
						Case 2: StatusText = "„ Êﬁ›"
						Case 3: StatusText = "„”œÊœ"
					end select
%>					<INPUT class="CustGenInput" TYPE="text" name="StatusText" value="<%=StatusText%>" size=12 readonly>
					<INPUT TYPE="Hidden" name="Status" value="<%=Status%>">
				<%end if%>
				</td>
			</tr>
			<tr>
				<td colspan='2' align=left>
					<input name="lblNorRCode" type="text" readonly value="<%If IsPersonal Then %> ‘„«—Â „·Ì: <%Else%> ‘„«—Â À» : <%End if%>" size=10 style='background-color:transparent;border:0 transparent hidden;text-align:left;'>
				</td>
				<td align=right colspan='3'>
					<input class="CustGenInput" name="NorRCode" value="<%=NorRCode%>" type="text" size="15" maxlength="20">
				</td>
			</tr>
			<tr>
				<td colspan="2" align="left"> ⁄œ«œ ‰›—« :</td>
				<td align="right" colspan="3">
					<select name="employee" id="employee" <% if IsPersonal then response.write " style='visibility:hidden;' " %> >
						<option value="0" <%if IsNull(employee) or employee=0 or IsPersonal then response.write " selected='selected' "%>>«‰ Œ«» ‰‘œÂ</option>
						<option value="1" <%if employee=1 then response.write " selected='selected' "%>>0  « 5 ‰›—</option>
						<option value="2" <%if employee=2 then response.write " selected='selected' "%>>6  « 20 ‰›—</option>
						<option value="3" <%if employee=3 then response.write " selected='selected' "%>>21  « 100 ‰›—</option>
						<option value="4" <%if employee=4 then response.write " selected='selected' "%>>101 »Â »«·«</option>
					</select>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="left">Ê» ”«Ì :</td>
				<td align=right colspan=3>
					<input class="CustGenInput" name="website" dir="ltr" value="<%=website%>" type="text" size="30" maxlength="100">
				</td>
			</tr>
			<tr bgcolor='#C3C300'>
				<td align='center' colspan='4'><b> «ÿ·«⁄«  —«»ÿ «’·Ì </b></td>
				<td align='center'>
				<INPUT TYPE="checkbox" NAME="postable1" <% if postable1 then response.write "checked" %> >
				ﬁ«»· Å”  ﬂ—œ‰ </td>
			</tr>
			<tr>
				<td align='center'><BR><SELECT class="CustGenInput" NAME="Dear1" style="width:70"><option value="">«‰ Œ«» ﬂ‰Ìœ</option><option <% if Dear1="¬ﬁ«Ì" then%> selected <% end if %>value="¬ﬁ«Ì">¬ﬁ«Ì</option><option <% if Dear1="Œ«‰„" then%> selected <% end if %> value="Œ«‰„">Œ«‰„</option></SELECT></td>
				<td colspan=2 align='center'><BR>‰«„ <INPUT class="CustGenInput" NAME="FirstName1" value="<%=FirstName1%>" TYPE="text" size="13">
				›«„Ì· <INPUT class="CustGenInput" NAME="LastName1" value="<%=LastName1%>" TYPE="text" size="13"></td>
				<td align='center'>‘Â— <BR><INPUT class="CustGenInput" TYPE="text" NAME="City1" size="15" value="<%=City1%>"></td>
				<td align='center'>ﬂœÅ” Ì <BR><INPUT class="CustGenInput" TYPE="text" NAME="PostCode1" size="15" value="<%=PostCode1%>"></td>
			</tr>
			<tr>
				<td align='center' rowspan='3' colspan='3'> ¬œ—” <BR><br><TEXTAREA class="CustGenInput" NAME="Address1" ROWS="6" COLS="60"><%=Address1%></TEXTAREA></td>
				<td valign='top' align='center'> ·›‰<BR><INPUT class="CustGenInput" Dir="LTR" NAME="Tel1" value="<%=Tel1%>" TYPE="text" size="15"></td>
				<td valign='top' align='center'>›«ﬂ”<BR><INPUT class="CustGenInput" Dir="LTR" NAME="Fax1" value="<%=Fax1%>" TYPE="text" size="15"></td>
			</tr>
			<tr>
				<td align='center'>„Ê»«Ì· <BR><INPUT class="CustGenInput" TYPE="text"  Dir="LTR"  NAME="Mobile1" size="15" value="<%=Mobile1%>"></td>
				<td valign='top' align='center'>«Ì„Ì·<BR><INPUT class="CustGenInput" Dir="LTR" NAME="Email1"  value="<%=Email1%>" TYPE="text" size="15" onkeyDown="return myKeyDownHandler();" onKeyPress="return myKeyPressHandler();"></td>
			</tr>
			<tr>
				<td colspan="2" align="center">”„  <br><input class="CustGenInput" type="text" dir="rtl" name="JobTitle1" value="<%=JobTitle1%>" size="40"></td>
			</tr>
			<tr bgcolor='#C3C300'>
				<td align='center' colspan='4'><b> «ÿ·«⁄«  —«»ÿ œÊ„ </b></td>
				<td align='center'>
				<INPUT TYPE="checkbox" NAME="postable2" <% if postable2 then response.write "checked" %> >
				ﬁ«»· Å”  ﬂ—œ‰ </td>
			</tr>
			<tr>
				<td align='center'><BR><SELECT class="CustGenInput" NAME="Dear2" style="width:70"><option value="">«‰ Œ«» ﬂ‰Ìœ</option><option <% if Dear2="¬ﬁ«Ì" then%> selected <% end if %>value="¬ﬁ«Ì">¬ﬁ«Ì</option><option <% if Dear2="Œ«‰„" then%> selected <% end if %> value="Œ«‰„">Œ«‰„</option></SELECT></td>
				<td colspan=2 align='center'><BR>‰«„ <INPUT class="CustGenInput" NAME="FirstName2" value="<%=FirstName2%>" TYPE="text" size="13">
				›«„Ì· <INPUT class="CustGenInput" NAME="LastName2" value="<%=LastName2%>" TYPE="text" size="13"></td>
				<td align='center'>‘Â— <BR><INPUT class="CustGenInput" TYPE="text" NAME="City2" size="15" value="<%=City2%>"></td>
				<td align='center'>ﬂœÅ” Ì <BR><INPUT class="CustGenInput" TYPE="text" NAME="PostCode2" size="15" value="<%=PostCode2%>"></td>
			</tr>
			<tr>
				<td align='center' rowspan='3' colspan='3'> ¬œ—” <BR><TEXTAREA class="CustGenInput" NAME="Address2" ROWS="6" COLS="60"><%=Address2%></TEXTAREA></td>
				<td valign='top' align='center'> ·›‰<BR><INPUT class="CustGenInput" Dir="LTR" NAME="Tel2" value="<%=Tel2%>" TYPE="text" size="15"></td>
				<td valign='top' align='center'>›«ﬂ”<BR><INPUT class="CustGenInput" Dir="LTR" NAME="Fax2" value="<%=Fax2%>" TYPE="text" size="15"></td>
			</tr>
			<tr>
				<td align='center'>„Ê»«Ì· <BR><INPUT class="CustGenInput" TYPE="text"  Dir="LTR"  NAME="Mobile2" size="15" value="<%=Mobile2%>"></td>
				<td valign='top' align='center'>«Ì„Ì·<BR><INPUT class="CustGenInput" Dir="LTR" NAME="Email2"  value="<%=Email2%>" TYPE="text" size="15" onkeyDown="return myKeyDownHandler();" onKeyPress="return myKeyPressHandler();"></td>
			</tr>
			<tr>
				<td colspan="2" align="center">”„  <br><input class="CustGenInput" type="text" dir="rtl" name="JobTitle2" value="<%=JobTitle2%>" size="40"></td>
			</tr>
			<%' ----------------------------S A M    E D I T  ----------------------------------------
			%>
			<tr bgcolor='#C3C300'>
				<td align='center' colspan='5'><b>œ” Â »‰œÌ</b></td>
			</tr>
			<tr>
				<td colspan=5>
					<table>
			<%
			accountGroup=0
			IF CusID <>"" then 
				mySQL="select * from accountGroupRelations where account="&CusID
				set RSV=conn.Execute(mySQL)
				if not RSV.eof then accountGroup=RSV("AccountGroup")
				RSV.close
				set RSV=nothing
			end if
			
			col=0
			oldState="True"
			set RSV=Conn.Execute ("SELECT * FROM accountGroups ORDER BY isPartner DESC, Name") 
			Do while not RSV.eof
				if (RSV("isPartner")<>oldState) then 
					if col <> 0 then 
						col = 0
						response.write "</tr>"
					end if
				end if
				if (col = 0) then response.write "<tr>"
				col=col+1
				%>
				<td <%
				if RSV("isPartner")="True" then response.write "bgcolor='#C33300'"%> >
					<INPUT type='radio' name='accountGroup' 
						<% if RSV("ID")=accountGroup then response.write(" CHECKED ")%> 
						value='<%=RSV("ID")%>'>
					<%=RSV("name")%>
				</td>
				<%
				if col=4 then 
					response.write "</tr>"
					col=0
				end if
				oldState = RSV("isPartner")
			RSV.moveNext
			Loop
			RSV.close
			set RSV=nothing
			%>
					</table>
				</td>
			</tr>
			<tr bgcolor='#C3C300'>
				<td align='center' colspan='5'><b>”Ê«·« </b></td>
			</tr>
			<tr>
				<td colspan="5">
					<table>
						<%
						if CLng(CusID)>0 then 
							set rs=Conn.Execute("select * from accountGroupRelations where account=" & CusID)
							if rs.eof then 
							%>
							<tr>
								<td><b>‘—„‰œÂ! ”Ê«·«  ›ﬁÿ œ— ’Ê— Ì ‰„«Ì‘ œ«œÂ ŒÊ«Âœ ‘œ ﬂÂ œ” Â »‰œÌ «Ì —« «‰ Œ«» ﬂ—œÂ »«‘Ìœ</b></td>
							</tr>
							<%						
							else
								group=rs("accountGroup")
								rs.close
								set rs=Conn.Execute("select * from account_questions where [group]=" & group)
								if rs.eof then
								%>
								<tr>
									<td><b>»—«Ì «Ì‰ œ” Â ÂÌç ”Ê«·Ì ‰œ«—„</b></td>
								</tr>
								<%
								else
									while not rs.eof
										%>
								<tr>
									<td><%=rs("name")%><input name="questionID" type="hidden" value="<%=rs("id")%>"></td>
									<td>
										<%
										'response.write ("select * from account_answers where account_id=" & CusID & " and question=" & rs("id"))
										set rss=Conn.Execute("select * from account_answers where account_id=" & CusID & " and question=" & rs("id"))
										if CInt(rs("type"))=0 then
										%>
										<input name="answer" type="text" <%if not rss.eof then response.write (" value='" & rss("answer") & "'")%>>
										<%
										elseif CInt(rs("type"))=1 then
										%>
										<select name="answer">
										<%
											choice=Split(rs("choice"),",")
											for i=0 to UBound(choice)
											%>
											<option value="<%=trim(choice(i))%>" <%if not rss.eof then if rss("answer")=choice(i) then response.write(" selected ")%>><%=choice(i)%></option>
											<%
											next
										%>
										</select>
										<%
										end if
										rss.close
										%>
									</td>
								</tr>
										<%
										rs.moveNext
									wend
									rs.close
								end if
							end if
						end if
						%>
					</table>
				</td>
			</tr>
			<%' ----------------------------S A M    E D I T  ----------------------------------------
			%>
			<tr bgcolor='#C3C300'>
				<td align='center' colspan="5"><INPUT class="CustGenInput" TYPE="submit"  NAME="submit" value="À» "> <% if request("act")="editaccount" then %><INPUT class="CustGenInput" TYPE="submit" NAME="submit" value="À»  Ê »⁄œÌ">
				<% end if %>
				</td>
			</tr>
			<tr>
				<td colspan="5">&nbsp;</td>
			</tr>
			<tr bgcolor='#C3C300'>
				<td align='center' colspan="5"><b>«ÿ·«⁄«  ﬁœÌ„Ì</b></td>
			</tr>
			<tr dir="LTR">
				<td align='center' colspan="5">
			<% OLD_ACID = ID%>
			<!--#include File="include_OldCusData.asp"-->
				</td>
			</tr>
		</table>
		</FORM>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			//document.all.Name.focus();
		//-->
		</SCRIPT>
<%
elseif request("act")="submitEdit" then 
	EditDate		= shamsiToday()
	EditBy			= session("ID")
	ID				= sqlSafe(Request.form("id"))

	AccountTitle	= left(sqlSafe(Request.form("AccountTitle")) , 200) ' Field width is 200
	IsPersonal		= sqlSafe(Request.form("IsPersonal"))
	IsADefault		= sqlSafe(Request.form("IsADefault"))
	EconomicalCode	= left(sqlSafe(Request.form("EconomicalCode")) , 20)
	CompanyName		= left(sqlSafe(Request.form("CompanyName")) , 50)
	NorRCode		= Left(sqlSafe(Request.Form("NorRCode")) , 20)

	Postable1		= sqlSafe(Request.form("Postable1"))
	Dear1			= left(sqlSafe(Request.form("Dear1")) , 20)
	FirstName1		= left(sqlSafe(Request.form("FirstName1")) , 50)
	LastName1		= left(sqlSafe(Request.form("LastName1")) , 50)
	JobTitle1		= left(sqlSafe(Request.form("JobTitle1")) , 50)
	Tel1			= left(sqlSafe(Request.form("Tel1")) , 50)
	Mobile1			= left(sqlSafe(Request.form("Mobile1")) , 50)
	Fax1			= left(sqlSafe(Request.form("Fax1")) , 50)
	Email1			= left(sqlSafe(Request.form("Email1")) , 50)
	City1			= left(sqlSafe(Request.form("City1")) , 50)
	Address1		= left(sqlSafe(Request.form("Address1")) , 255)
	PostCode1		= left(sqlSafe(Request.form("PostCode1")) , 50)

	Postable2		= sqlSafe(Request.form("Postable2"))
	Dear2			= left(sqlSafe(Request.form("Dear2")) , 20)
	FirstName2		= left(sqlSafe(Request.form("FirstName2")) , 50)
	LastName2		= left(sqlSafe(Request.form("LastName2")) , 50)
	JobTitle2		= left(sqlSafe(Request.form("JobTitle2")) , 50)
	Tel2			= left(sqlSafe(Request.form("Tel2")) , 50)
	Mobile2			= left(sqlSafe(Request.form("Mobile2")) , 50)
	Fax2			= left(sqlSafe(Request.form("Fax2")) , 50)
	Email2			= left(sqlSafe(Request.form("Email2")) , 50)
	City2			= left(sqlSafe(Request.form("City2")) , 50)
	Address2		= left(sqlSafe(Request.form("Address2")) , 255)
	PostCode2		= left(sqlSafe(Request.form("PostCode2")) , 50)

	CreditLimit		= cdbl(text2value(Request.form("CreditLimit")))
	accountGroup	= cdbl(text2value(request.form("accountGroup")))
	employee		= cint(text2value(request.form("employee")))
	website			= left(sqlSafe(Request.form("website")) , 100)
	
	if not Auth(1 , 4) then ' Doesn't have the permission to set CSR / credit limit
		CSR			= ""
	else
		CSR			= cint(sqlSafe(Request.form("CSR")))

		mySQL="SELECT RealName FROM Users WHERE (ID = "& CSR & ")"
		Set RS1=Conn.Execute(mySQL)
		if RS1.eof then
			Conn.close
			response.redirect "?msg=" & Server.URLEncode("ç‰Ì‰ ﬂ«—»—Ì ÅÌœ« ‰‘œ.")
		else
			NewCSRName=		RS1("RealName")
		end if
		RS1.close

		mySQL="SELECT Accounts.CSR, Users.RealName, Accounts.CreditLimit FROM Accounts INNER JOIN Users ON Accounts.CSR = Users.ID WHERE (Accounts.ID = "& ID & ")"
		Set RS1=Conn.Execute(mySQL)
		if RS1.eof then
			Conn.close
			response.redirect "?msg=" & Server.URLEncode("ç‰Ì‰ ›«ﬂ Ê—Ì ÅÌœ« ‰‘œ.")
		else
			OldCreditLimit=	cdbl(RS1("CreditLimit"))
			OldCSRID=		cint(RS1("CSR"))
			OldCSRName=		RS1("RealName")
		end if
		RS1.close
		Set RS1=Nothing

		if OldCSRID <> CSR OR OldCreditLimit <> CreditLimit then

			mySQL = "SELECT [User] as ID, OnAccountCSRChangeSendMessageTo FROM UserDefaults WHERE (([User] = "& session("ID") & ") OR (UserDefaults.[User] = 0)) AND (OnAccountCSRChangeSendMessageTo IS NOT NULL) ORDER BY ABS(UserDefaults.[User]) DESC"
			Set RS2 = Conn.Execute (mySQL)

			MessageTo=RS2("OnAccountCSRChangeSendMessageTo")
			RS2.close
			Set RS2 = Nothing

			msg = "" '" Ê”ÿ " & session("CSRName") & " "
			tmpAnd = ""
			if OldCSRID <> CSR then
				msg= "„”ƒÊ· ÅÌêÌ—Ì «“ '" & OldCSRName & "' »Â '" & NewCSRName & "' "
				tmpAnd = "Ê "
			end if

			if OldCreditLimit <> CreditLimit then
				msg = msg & tmpAnd &  "”ﬁ› «⁄ »«— «“  '" & Separate(OldCreditLimit) & "' »Â '" & Separate(CreditLimit) & "' "
			end if
			msg = msg & " €ÌÌ— Ì«› ."

			MsgTo			=	MessageTo
			msgTitle		=	"Account changed"
			msgBody			=	sqlSafe(msg)
			RelatedTable	=	"accounts"
			relatedID		=	ID
			replyTo			=	0
			IsReply			=	0
			urgent			=	1
			MsgFrom			=	session("ID")
			MsgDate			=	shamsiToday()
			MsgTime			=	currentTime10()
			mySQL="INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")"
			Conn.Execute (mySQL)

		end if

	end if

	AccType			= sqlSafe(Request.form("Type"))
	Status			= sqlSafe(Request.form("Status"))
	'ACCID			= sqlSafe(Request.form("ACCID"))
	' ----------------------------S A M    E D I T  ----------------------------------------
	mySQL="select * from AccountGroupRelations where account="&id 
	set rs1= Conn.Execute(mySQL)
	if rs1.eof then 
		oldAccountGroup=0
	else 
		oldAccountGroup=rs1("AccountGroup")
	end if
	rs1.close
	set rs1=nothing
	' ----------------------------S A M    E D I T  ----------------------------------------
	if IsPersonal="on" then
		IsPersonal=1
		CompanyName=""
	else 
		IsPersonal=0
	end if

	if IsADefault="on" then
		IsADefault=1
	else 
		IsADefault=0
	end if

	if Postable1="on" then
		Postable1=1
	else 
		Postable1=0
	end if

	if Postable2="on" then
		Postable2=1
	else 
		Postable2=0
	end if

'	Added By kid 820727 
'	-------------------------------
'	Log The Edition Before Updating

	conn.Execute("INSERT INTO AccountsEditLog SELECT '"& EditDate & "' AS EditedOn, '"& EditBy & "' AS EditedBy, * FROM Accounts WHERE (ID = "& ID & ")")

'	End of Log
'   -------------------------------
	if CSR="" then
		mySQL="UPDATE Accounts SET LastEditOn='"& EditDate & "', LastEditBy='"& EditBy& "', Type ="&AccType & ", Status ="&Status & ", AccountTitle =N'"& AccountTitle& "', Postable1 = "& Postable1 & ", Postable2 = "& Postable2 & ", IsADefault = "& IsADefault & ", EconomicalCode =N'"& EconomicalCode& "', IsPersonal ="& IsPersonal& ", CompanyName =N'"& CompanyName& "', Dear1 =N'"& Dear1& "', FirstName1 =N'"& FirstName1& "', LastName1 =N'"& LastName1& "', JobTitle1 =N'"& JobTitle1& "', Tel1 =N'"& Tel1& "', Fax1 =N'"& Fax1& "', EMail1 =N'"& EMail1& "', Mobile1 =N'"& Mobile1& "', PostCode1 =N'"& PostCode1& "', City1 =N'"& City1& "', Address1 =N'"& Address1& "', Dear2 =N'"& Dear2& "', FirstName2 =N'"& FirstName2& "', LastName2 =N'"& LastName2& "', JobTitle2 =N'"& JobTitle2 & "', Tel2 =N'"& Tel2& "', Fax2 =N'"& Fax2& "', EMail2 =N'"& EMail2& "', Mobile2 =N'"& Mobile2& "', PostCode2 =N'"& PostCode2& "', City2 =N'"& City2& "', Address2 =N'"& Address2& "', NorRCode = N'" & NorRCode & "' , employee=" & employee & ",website=N'" & website & "' WHERE (ID = "& ID & ")"
	else
		mySQL="UPDATE Accounts SET CSR ="& CSR & ", LastEditOn='"& EditDate & "', LastEditBy='"& EditBy& "', Type ="&AccType & ", Status ="&Status & ", AccountTitle =N'"& AccountTitle& "', Postable1 = "& Postable1 & ", Postable2 = "& Postable2 & ", IsADefault = "& IsADefault & ", EconomicalCode =N'"& EconomicalCode& "', CreditLimit="& CreditLimit & ", IsPersonal ="& IsPersonal& ", CompanyName =N'"& CompanyName& "', Dear1 =N'"& Dear1& "', FirstName1 =N'"& FirstName1& "', LastName1 =N'"& LastName1& "', JobTitle1 =N'"& JobTitle1& "', Tel1 =N'"& Tel1& "', Fax1 =N'"& Fax1& "', EMail1 =N'"& EMail1& "', Mobile1 =N'"& Mobile1& "', PostCode1 =N'"& PostCode1& "', City1 =N'"& City1& "', Address1 =N'"& Address1& "', Dear2 =N'"& Dear2& "', FirstName2 =N'"& FirstName2& "', LastName2 =N'"& LastName2& "', JobTitle2 =N'"& JobTitle2 & "', Tel2 =N'"& Tel2& "', Fax2 =N'"& Fax2& "', EMail2 =N'"& EMail2& "', Mobile2 =N'"& Mobile2& "', PostCode2 =N'"& PostCode2& "', City2 =N'"& City2& "', Address2 =N'"& Address2& "', NorRCode = N'" & NorRCode & "',employee=" & employee & ",website=N'" & website & "' WHERE (ID = "& ID & ")"
	end If
	'response.write mySQL
	'response.end
	conn.Execute(mySQL)
	' ----------------------------S A M    E D I T  ----------------------------------------
'	response.write oldAccountGroup&"<br>"&accountGroup&"<br>"&id&"<br>"
	if oldAccountGroup<>accountGroup then
		conn.Execute("DELETE FROM AccountGroupRelations where account="&id&" AND AccountGroup="&oldAccountGroup)
		conn.Execute("INSERT INTO AccountGroupRelations (Account,AccountGroup) VALUES ("&id&","&accountGroup&")") 
	end if
	questionID=split(request.form("questionID"),",")
	'response.write (request.form("questionID"))
	answers=split(request("answer"), ",")
	set rs=Conn.Execute("select account_answers.* from account_answers inner join account_questions on account_questions.id=account_answers.question where account_questions.[group]=" & accountGroup & " and account_answers.account_id=" & id)
	while not rs.eof
		answer=""
		for i=0 to UBound(questionID)
			if cint(rs("id"))=cint(questionID(i)) then answer=trim(answers(i))
			'response.write rs("id") &","&questionID(i)&","& answers(i)&"<br>"
			'response.write answer
		next 
		'response.end
		Conn.Execute("update account_answers set answer=N'" & answer & "' where id=" & rs("id"))
		'response.write("update account_answers set answer=N'" & answer & "' where id=" & rs("id") &"<br>")
		rs.moveNext
	wend
	
	'response.end
	rs.close
	set rs=Conn.Execute("select * from account_questions where [group]=" & accountGroup & " and id not in (select question from account_answers where account_id=" & id & ")")
	while not rs.eof
		answer=""
		'response.w
		for i=0 to UBound(questionID)
			if cint(rs("id"))=cint(questionID(i)) then answer=trim(answers(i))
		next 
		Conn.Execute("insert into account_answers (question,answer,account_id) values (" & rs("id") & ",N'" & answer & "'," & id & ")")
		rs.moveNext
	wend
	rs.close
	' ----------------------------S A M    E D I T  ----------------------------------------
	if request("submit")="À»  Ê »⁄œÌ" then 
		response.redirect "AccountEdit.asp?act=editaccount&NextOf=" & ID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  Õ”«»  ﬁ»·Ì »Â —Ê“ ‘œ.")
	else
		response.redirect "AccountInfo.asp?act=show&selectedCustomer=" & ID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  Õ”«» »Â —Ê“ ‘œ.")
	end if

elseif request("act")="submitcustomer" then 
	CreatedDate		=	shamsiToday()
	CreatedBy		=	session("ID")

	AccountTitle	= left(sqlSafe(Request.form("AccountTitle")) , 200) ' Field width is 200
	CreditLimit		= text2value(Request.form("CreditLimit"))
	IsPersonal		= sqlSafe(Request.form("IsPersonal"))
	IsADefault		= sqlSafe(Request.form("IsADefault"))
	EconomicalCode	= left(sqlSafe(Request.form("EconomicalCode")) , 20)
	CompanyName		= left(sqlSafe(Request.form("CompanyName")) , 50)
	NorRCode		= Left(sqlSafe(Request.Form("NorRCode")) , 20)

	Postable1		= sqlSafe(Request.form("Postable1"))
	Dear1			= left(sqlSafe(Request.form("Dear1")) , 20)
	FirstName1		= left(sqlSafe(Request.form("FirstName1")) , 50)
	LastName1		= left(sqlSafe(Request.form("LastName1")) , 50)
	JobTitle1		= left(sqlSafe(Request.form("JobTitle1")) , 50)
	Tel1			= left(sqlSafe(Request.form("Tel1")) , 50)
	Mobile1			= left(sqlSafe(Request.form("Mobile1")) , 50)
	Fax1			= left(sqlSafe(Request.form("Fax1")) , 50)
	Email1			= left(sqlSafe(Request.form("Email1")) , 50)
	City1			= left(sqlSafe(Request.form("City1")) , 50)
	Address1		= left(sqlSafe(Request.form("Address1")) , 255)
	PostCode1		= left(sqlSafe(Request.form("PostCode1")) , 50)

	Postable2		= sqlSafe(Request.form("Postable2"))
	Dear2			= left(sqlSafe(Request.form("Dear2")) , 20)
	FirstName2		= left(sqlSafe(Request.form("FirstName2")) , 50)
	LastName2		= left(sqlSafe(Request.form("LastName2")) , 50)
	JobTitle2		= left(sqlSafe(Request.form("JobTitle2")) , 50)
	Tel2			= left(sqlSafe(Request.form("Tel2")) , 50)
	Mobile2			= left(sqlSafe(Request.form("Mobile2")) , 50)
	Fax2			= left(sqlSafe(Request.form("Fax2")) , 50)
	Email2			= left(sqlSafe(Request.form("Email2")) , 50)
	City2			= left(sqlSafe(Request.form("City2")) , 50)
	Address2		= left(sqlSafe(Request.form("Address2")) , 255)
	PostCode2		= left(sqlSafe(Request.form("PostCode2")) , 50)

	CreditLimit		= cdbl(text2value(Request.form("CreditLimit")))
	accountGroup	= cdbl(text2value(request.form("accountGroup")))
	employee		= cint(text2value(request.form("employee")))
	website			= left(sqlSafe(Request.form("website")) , 100)
	
	if not Auth(1 , 4) then ' Doesn't have the permission to set CSR / credit limit
		CreditLimit = 0
		CSR			= 0
	else
		CSR			= cint(sqlSafe(Request.form("CSR")))
	end if

	AccType			= sqlSafe(Request.form("Type"))
	Status			= sqlSafe(Request.form("Status"))
	'ACCID			= Request.form("ACCID")

	if IsPersonal="on" then
		IsPersonal=1
	else 
		IsPersonal=0
	end if

	if IsADefault="on" then
		IsADefault=1
	else 
		IsADefault=0
	end if

	if Postable1="on" then
		Postable1=1
	else 
		Postable1=0
	end if

	if Postable2="on" then
		Postable2=1
	else 
		Postable2=0
	end if


	mySQL="SELECT MAX(ID) AS MaxID FROM Accounts WHERE ([ID] <900000)"	
	Set RS1=Conn.Execute(mySQL)
	NewAccID=RS1("MAXID")+1
	RS1.close
	Set RS1=Nothing

	mySQL="INSERT INTO Accounts (ID, CreatedDate, CreatedBy, LastEditOn, LastEditBy, CSR, IsADefault, EconomicalCode, Type, AccountTitle, IsPersonal, CompanyName, Postable1, Dear1, FirstName1, LastName1, JobTitle1, Tel1, Fax1, EMail1, Mobile1, PostCode1, City1, Address1, Postable2, Dear2, FirstName2, LastName2, JobTitle2, Tel2, Fax2, EMail2, Mobile2, PostCode2, City2, Address2, NorRCode, employee, website) VALUES ("&_
	NewAccID & ", N'"& CreatedDate & "', "& CreatedBy & ", N'"& CreatedDate & "', "& CreatedBy & ", "& CSR & ", "& IsADefault & ", N'"& EconomicalCode & "', "& AccType & ", N'"& AccountTitle & "',"& IsPersonal & ", N'"& CompanyName & "', " & Postable1 & ", N'"& Dear1 & "', N'"& FirstName1 & "', N'"& LastName1 & "', N'"& JobTitle1 & "', N'"& Tel1 & "', N'"& Fax1 & "', N'"& Email1 & "', N'"& Mobile1 & "', N'"& PostCode1 & "', N'"& City1 & "', N'"& Address1 & "', " & Postable2 & ",N'"& Dear2 & "', N'"& FirstName2 & "', N'"& LastName2 & "',  N'"& JobTitle2 & "', N'"& Tel2 & "', N'"& Fax2 & "', N'"& Email2 & "', N'"& Mobile2 & "', N'"& PostCode2 & "', N'"& City2 & "', N'"& Address2 & "', N'" & NorRCode & "', " & employee & ",N'" & website & "')"

	conn.Execute(mySQL)
'	Added By kid 820727 
'	-------------------------------
'	Log The Insertion
	conn.Execute("INSERT INTO AccountsEditLog SELECT '"& CreatedDate & "' AS EditedOn, '"& CreatedBy & "' AS EditedBy,* FROM Accounts WHERE (ID = "& NewAccID & ")")
'	End of Log
'   -------------------------------

	response.redirect "AccountInfo.asp?act=show&selectedCustomer=" & NewAccID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  Õ”«» À»  ‘œ.")
%>
<!-- «ÿ·«⁄«  „‘ —Ì À»  ‘œ -->
	<div dir='rtl' align=center><B>«ÿ·«⁄«  Õ”«» À»  ‘œ...</B><br><br>
	</div>
<%
	response.end
	RS1.close

else if request("act")="" then %>
<!-- Ã” ÃÊ »—«Ì ‰«„ Õ”«» -->
	<BR><BR>
	<FORM METHOD=POST ACTION="AccountEdit.asp?act=submitsearch" onsubmit="if (document.all.CustomerNameSearchBox.value=='') return false;">
	<div dir='rtl'>&nbsp;&nbsp;<B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«»</B>
		<INPUT TYPE="text" NAME="CustomerNameSearchBox">&nbsp;
		<INPUT TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>
	<BR>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.CustomerNameSearchBox.focus();
	//-->
	</SCRIPT>
<%
end if
end if
conn.Close
%>
</font>
<br>
<script language="JavaScript">
<!--
function setAccountTitle(){
	with (document.all){
		if (IsPersonal.checked){
			AccountTitle.value= Dear1.value + " " +FirstName1.value + " " +LastName1.value + " "
		}
		else{
			if(Dear1.value+FirstName1.value+LastName1.value != ''){
				AccountTitle.value=CompanyName.value + " (" + Dear1.value + " " +FirstName1.value + " " +LastName1.value + ")"
			}
			else{
				AccountTitle.value=CompanyName.value
			}
		}
	}
}

function showCompanyName(){
	if (document.all.IsPersonal.checked){
		document.all.CompanyName.style.visibility="hidden";
		document.all.lblNorRCode.value="‘„«—Â „·Ì:";
		document.all.employee.style.visibility="hidden";
	}
	else{
		document.all.CompanyName.style.visibility="visible"
		document.all.CompanyName.focus();
		document.all.lblNorRCode.value="‘„«—Â À» :";
		document.all.employee.style.visibility="visible";
	}
}
//-->
</script>
<!--#include file="tah.asp" -->
