<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' RFM
PageTitle="œ‘»Ê—œ „‘ —Ì«‰"
SubmenuItem=7
if not Auth(1 , 8) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.myTable {border-color: black;border-width: medium;border-style: solid;}
	.myTable tr:nth-child(1) {border-color: gray;border-width: thin;border-style: solid; background-color: lime ;text-align: center;border-bottom-width: thick;border-bottom-color: black;border-bottom-style: solid;}
	.myTable tr:nth-child(2n+2) { background-color:#eee; }
	.myTable tr:nth-child(2n+3) { background-color:#C3DBEB; }
</STYLE>
<BR>
<%
dim color(4)
dim text(4)
color(0)="red"
color(1)="white"
color(2)="yellow"
color(3)="green"
text(0)="ﬂ„"
text(1)="„ Ê”ÿ"
text(2)="“Ì«œ"
text(3)="ŒÌ·Ì “Ì«œ"
if request("act")="" then 
%>
<table class="myTable" cellpadding="8px" cellspacing="0">
	<tr>
		<td>Recency</td>
		<td>Frequency</td>
		<td>Value</td>
		<td></td>
	</tr>
<%

	mySQL="select count(rfv_value.customer) as c,rfv_value.Value,rfv_recency.Recency,rfv_frequency.Frequency from rfv_frequency inner join rfv_recency on rfv_frequency.customer=rfv_recency.customer inner join rfv_value on rfv_recency.customer=rfv_value.customer group by rfv_value.Value,rfv_recency.Recency,rfv_frequency.Frequency order by Recency,Frequency,Value"
				'response.write mySQL
				'response.end
	set rs=Conn.Execute(mySQL)
	while not rs.eof
%>
	<tr>
		<td style="background-color:<%=color(rs("Recency"))%>"><%=text(rs("Recency"))%></td>
		<td style="background-color:<%=color(rs("Frequency"))%>"><%=text(rs("Frequency"))%></td>
		<td style="background-color:<%=color(rs("Value"))%>"><%=text(rs("Value"))%></td>
		<td><a href="rfmModel.asp?act=show&value=<%=rs("Value")%>&recency=<%=rs("Recency")%>&frequency=<%=rs("Frequency")%>"><%=Separate(rs("c"))%></a></td>
	</tr>
<%
		rs.moveNext
	wend
%>
</table>
<%
elseif request("act")="show" then 
%>
<table class="myTable" cellpadding="5px" cellspacing="0">
	<tr>
		<td width="30px">¬Œ—Ì‰ ”›«—‘ (—Ê“)</td>
		<td width="30px">«Ê·Ì‰ ”›«—‘  « «„—Ê“ („«Â)</td>
		<td width="30px"> ⁄œ«œ ”›«—‘</td>
		<td width="30px">„Ì«‰êÌ‰ ”›«—‘ œ— „«Â</td>
		<td width="50px">Ã⁄⁄ ﬂ· ”›«—‘</td>
		<td>„‘ —Ì</td>
		<td>Ã„⁄ „«‰œÂ</td>
	</tr>
<%

	mySQL="select rfv_frequency.customer,rfv_frequency.[count],rfv_frequency.[Month],rfv_value.Amount,rfv_recency.[day],accounts.accountTitle, accounts.arBalance+accounts.aoBalance+accounts.apBalance as Remain from rfv_frequency inner join rfv_recency on rfv_frequency.customer=rfv_recency.customer inner join rfv_value on rfv_recency.customer=rfv_value.customer inner join Accounts on rfv_value.Customer = Accounts.id where Recency=" & request("Recency") & " and Value=" & request("Value") & " and Frequency=" & request("Frequency") & " order by [day]"
	'response.write mySQL
	'response.end
	set rs=Conn.Execute(mySQL)
	i=0
	while not rs.eof
		i=i+1
%>
	<tr>
		<td><%=rs("Day")%></td>
		<td><%=rs("Month")%></td>
		<td><%=rs("Count")%></td>
		<td><%if rs("Month")=0 then response.write "N/A" else response.write Round(CInt(rs("Count"))/Cint(rs("Month")),2) end if%></td>
		<td><%=Separate(rs("Amount"))%></td>
		<td><a href="AccountInfo.asp?act=show&selectedCustomer=<%=rs("Customer")%>"><%=rs("accountTitle")%></a></td>
		<td><%=Separate(rs("Remain"))%></td>
	</tr>
<%
		rs.moveNext
	wend
%>
	<tr>
		<td>Ã„⁄ ﬂ·:</td>
		<td colspan="6"><%=i%></td>
	</tr>
</table>
<%
end if
%>

<!--#include file="tah.asp" -->
