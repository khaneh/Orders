<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="ê“«—‘ ›—Ê‘"
SubmenuItem=11
if not Auth("C" , 0) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {border:1pt solid white;vertical-align:top;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTable2 th {font-size:9pt; padding:5px; background-color:#0080C0;height:25px;}
	.RepTable2 td {height:25px;}
	.RepTable2 input {font-family:tahoma; font-size:9pt; border:1 solid black;}
	.RepTable2 select {font-family:tahoma; font-size:9pt; border:1 solid black;}
</STYLE>
<BR>
<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Search Form
'-----------------------------------------------------------------------------------------------------
if request("act")="" then

StartOfTheYear = left(shamsiToday(),4) & "/01/01"
%>

<TABLE class="RepTable" align=center cellspacing=10 >
	<TR height=110>
		<TD style="border:1pt solid white;width:130pt" align=center>
		<% if Auth("C" , 1) then %>
			<FORM METHOD=POST ACTION="?act=dailySale">
			<table class="RepTable2" id="AInvoices">
			<tr>
				<th colspan="2">ê“«—‘ ›—Ê‘ —Ê“«‰Â</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<% 
			if Auth("C" , 2) then 
			%>
				<tr>
					<td align=left>«“&nbsp; «—ÌŒ</td>
					<td><INPUT TYPE="text" NAME="input_date_start" value="<%=shamsiToday()%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
				</tr>
				<tr>
					<td align=left> «&nbsp; «—ÌŒ</td>
					<td><INPUT TYPE="text" NAME="input_date_end"  value="<%=shamsiToday()%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
				</tr>
			<%
			else
			%>
				<tr>
					<td colspan="2" align=left>
					<SELECT NAME="input_date">
						<option value="<%=shamsiToday()%>">			«„—Ê“		<%=replace(shamsitoday(),"/" ,".")%> </option>
						<option value="<%=shamsidate(date()-1)%>">	œÌ—Ê“		<%=replace(shamsidate(date()-1),"/" ,".")%> </option>
						<option value="<%=shamsidate(date()-2)%>">	2 —Ê“ ﬁ»·	<%=replace(shamsidate(date()-2),"/" ,".")%> </option>
						<option value="<%=shamsidate(date()-3)%>">	3 —Ê“ ﬁ»·	<%=replace(shamsidate(date()-3),"/" ,".")%> </option>
						<option value="<%=shamsidate(date()-4)%>">	4 —Ê“ ﬁ»·	<%=replace(shamsidate(date()-4),"/" ,".")%> </option>
					</SELECT>
					</td>
				</tr>
			<%
			end if
			%>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			</table>
				<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ç«Å ">&nbsp;
				<br><br>
				<input type=hidden name='fullyApplied' value='on'>
				<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ " onclick="document.forms[0].action='rep_dailySale.asp';">&nbsp;
			</FORM>
		<% end if %>
		&nbsp;
		</TD>
		<TD style="border:1pt solid white;width:130pt" align=center>
		<% if Auth("C" , 3) then %>
			<FORM METHOD=POST ACTION="Rep_InvoiceItems.asp?act=show">
			<table class="RepTable2" id="InvoiceItems">
			<tr>
				<th colspan="2">ê“«—‘ ¬Ì „ Â«Ì ›—Ê‘</td>
			</tr>
			<tr>
				<td align=left>«“&nbsp; «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="FromDate" Value="<%=StartOfTheYear%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left> «&nbsp; «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="ToDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left>ê—ÊÂ</td>
				<td align=left>
					<SELECT NAME="Category" style="width:120px;">
						<option value="">--- «‰ Œ«» ﬂ‰Ìœ ---</option>
						<option value="0" <%if Category=0 then response.write "selected " : GroupName="[ Â„Â ] "%>>Â„Â</option>
<%
					mySQL = "SELECT * FROM InvoiceItemCategories WHERE ID > 0"
					Set RS1 = Conn.Execute(mySQL)
					Do Until RS1.eof
%>				
						<option value="<%=RS1("ID")%>" <%if RS1("ID")=Category then response.write "selected "%>><%=RS1("Name")%></option>
<%
						RS1.moveNext
					Loop
%>
					</SELECT>
				</td>
			</tr>
			<tr>
				<td align=left> ⁄œ«œ</td>
				<td><INPUT TYPE="text" NAME="ResultsInPage" style="width:75px;direction:LTR;text-align:right;" maxlength=5 value="50"></td>
			</tr>
			<tr>
				<td align=left>›«ﬂ Ê—Â«</td>
				<td align=left>
					<SELECT NAME="isA" style="width:120px;">
						<option value="0" <%if isA=0 then response.write "selected " %>>Â„Â </option>
						<option value="1" <%if isA=1 then response.write "selected " %>>»«  Ìﬂ «·›</option>
						<option value="2" <%if isA=2 then response.write "selected " %>>»œÊ‰  Ìﬂ «·›</option>
					</SELECT>
				</td>
			</tr>
			</table>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ ">&nbsp;
			</FORM>
		<% end if %>
		&nbsp;
		</TD>
		<TD align=center>
		<% if Auth("C" , 4) then %>
			<FORM METHOD=POST ACTION="Rep_AInvoices.asp?act=show">
			<table class="RepTable2" id="AInvoices">
			<tr>
				<th colspan="2">ê“«—‘ ›«ﬂ Ê— Â« («·›)</td>
			</tr>
			<tr>
				<td align=left>«“  «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="FromDate" Value="<%=StartOfTheYear%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left> «  «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="ToDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left> ⁄œ«œ</td>
				<td><INPUT TYPE="text" NAME="ResultsInPage" style="width:75px;direction:LTR;text-align:right;" maxlength=5 value="50"></td>
			</tr>
			</table>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ ">&nbsp;
			<br><br>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="Button" value=" ‰„«Ì‘ —Ê“«‰Â " onclick="document.forms[2].action='Rep_AInvoices.asp?act=showByDay';document.forms[2].submit();">&nbsp;
			</FORM>
		<% end if %>
		&nbsp;
		</TD>
	</TR>
	<TR height=110>
		<TD style="border:1pt solid white;width:130pt" align=center>
		<% if Auth("C" , 5) then %>
			<FORM METHOD=POST ACTION="Rep_AReceipt.asp?act=show">
			<table class="RepTable2" id="AInvoices">
			<tr>
				<th colspan="2">ê“«—‘ —”Ìœ Â«</td>
			</tr>
			<tr>
				<td align=left>«“  «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="FromDate" Value="<%=StartOfTheYear%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left> «  «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="ToDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left> ⁄œ«œ</td>
				<td><INPUT TYPE="text" NAME="ResultsInPage" style="width:75px;direction:LTR;text-align:right;" maxlength=5 value="50"></td>
			</tr>
			<tr>
				<td align=left>›«ﬂ Ê—Â«Ì</td>
				<td align=left>
					<SELECT NAME="isA" style="width:90px;">
						<option value="0" >Â„Â </option>
						<option value="1" selected >»«  Ìﬂ «·›</option>
						<option value="2" >»œÊ‰  Ìﬂ «·›</option>
					</SELECT>
				</td>
			</tr>
			</table>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ ">&nbsp;
			</FORM>		
		<% end if %>
		&nbsp;
		</TD>
		<TD style="border:1pt solid white;width:130pt" align=center>
		<% if Auth("C" , 7) then %>
			<FORM METHOD=POST ACTION="Rep_COGS.asp?act=show">
			<table class="RepTable2" id="AInvoices">
			<tr>
				<th colspan="2">ê“«—‘ ﬁÌ„   „«„ ‘œÂ</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td align=left>«“  «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="FromDate" Value="<%=StartOfTheYear%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left> «  «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="ToDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			</table>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ ">&nbsp;
			</FORM>		
		<% end if %>
		 &nbsp;
		</TD>
		
<%
		'sam add this!	
%>
		<TD align=center>
		<% if Auth("C" , 4) then %>
			<FORM METHOD=POST ACTION="Rep_BInvoices.asp?act=show">
			<table class="RepTable2" id="AInvoices">
			<tr>
				<th colspan="2">ê“«—‘ ›«ﬂ Ê—Â« (»)</td>
			</tr>
			<tr>
				<td align=left>«“  «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="FromDate" Value="<%=StartOfTheYear%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left> «  «—ÌŒ</td>
				<td><INPUT TYPE="text" NAME="ToDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			</tr>
			<tr>
				<td align=left> ⁄œ«œ</td>
				<td><INPUT TYPE="text" NAME="ResultsInPage" style="width:75px;direction:LTR;text-align:right;" maxlength=5 value="50"></td>
			</tr>
			</table>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ ">&nbsp;
			<br><br>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="Button" value=" ‰„«Ì‘ —Ê“«‰Â " onclick="document.forms[2].action='Rep_BInvoices.asp?act=showByDay';document.forms[2].submit();">&nbsp;
			</FORM>
		<% end if %>
		&nbsp;
		</TD>

	</TR>
	<!---------------------------------------------------- SAM -->
	<TR height=110>
		<TD style="border:1pt solid white;width:130pt" align=center>
		<% if Auth("C" , 8) then %>
			<FORM METHOD=POST ACTION="Rep_4col.asp?act=show">
			<table class="RepTable2" id="col4">
			<tr>
				<th colspan="2">ê“«—‘  —«“ çÂ«— ” Ê‰Ì</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
				<tr>
					<td align=left>«“&nbsp; «—ÌŒ</td>
					<td><INPUT TYPE="text" NAME="FromDate" value="<%=shamsiToday()%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
				</tr>
				<tr>
					<td align=left> «&nbsp; «—ÌŒ</td>
					<td><INPUT TYPE="text" NAME="ToDate"  value="<%=shamsiToday()%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
				</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			</table>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ ">&nbsp;
			</FORM>
		<% end if %>
		&nbsp;
		</TD>
		<!---------------------------------------------------------------------------------->
		<TD style="border:1pt solid white;width:130pt" align=center>
		<% if Auth("C" , 9) then %>
			<table class="RepTable2" id="col8">
			<tr>
				<td colspan="2"><a href="Rep_CustomerNoCredit.asp?act=show">ê“«—‘ „‘ —Ì«‰Ì ﬂÂ «⁄ »«— ¬‰Â« ’›— ‰Ì” </a></td>
			</tr>
			<tr>
				<td colspan="2"><a href="arReport.asp">Õ”«»ùÂ«Ì »œÂﬂ«—</a></td>
			</tr>
			</table>
		<% end if %>
		&nbsp;
		</TD>
		<!------------------------------------------------------------>
		<TD align=center>
		&nbsp;
		<% if Auth("C", "A") then %>
			<form method="post" action="rep_reversSales.asp?act=show">
				<table class="RepTable2">
					<tr>
						<th colspan="2">ê“«—‘ »—ê‘  «“ ›—Ê‘</th>
					</tr>
					<tr>
						<td>«“  «—ÌŒ</td>
						<td><input type="text" name="fromDate" value="<%=shamsiToday()%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
					</tr>
					<tr>
						<td> «  «—ÌŒ</td>
						<td><input type="text" name="toDate" value="<%=shamsiToday()%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
					</tr>
				</table>
				<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ ">&nbsp;
			</form>
		<%end if%>
		</TD>
		<!------------------------------------------------------------>
	</TR>
</TABLE>
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------ Daily Sale
'-----------------------------------------------------------------------------------------------------
elseif request("act")="dailySale" then
	input_date=			sqlSafe(request("input_date"))
	input_date_start=	sqlSafe(request("input_date_start"))
	input_date_end=		sqlSafe(request("input_date_end"))

	if Auth("C" , 1) then 
		if Auth("C" , 2) then 
			'FrameSrc="/Reports/dailySale.aspx?input_date_start=" & input_date_start & "&input_date_end=" & input_date_end 
			RepParamValues = input_date_start & "" & input_date_end
		else
			'FrameSrc="/Reports/dailySale.aspx?input_date_start=" & input_date & "&input_date_end=" & input_date 
			RepParamValues = input_date & "" & input_date
		end if
		%>

		<BR><BR>
		<CENTER>
		<% 	ReportLogRow = PrepareReport ("dailySale1.rpt", "input_date_startinput_date_end", RepParamValues, "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

		</CENTER>
		<BR>
		<BR><iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no ></iframe>

		<%
	else
		NotAllowdToViewThisPage()
	end if
'---------------------------------------------------------------
end if
%>

<!--#include file="tah.asp" -->
