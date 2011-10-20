<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle=" ”  „—ﬂ“ Â“Ì‰Â"
SubmenuItem=7
if not Auth(3 , 6) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->

<%
if request("act")="" then 
	response.write "<span>œ—«ÌÊ „Ê—œ ‰Ÿ— —« «‰ Œ«» ò‰Ìœ</span>"
	set rs=Conn.Execute("SELECT cost_centers.id as costID, cost_centers.name as costName,cost_drivers.* from cost_centers inner join cost_drivers on cost_centers.id=cost_drivers.cost_center_id")
	oldCostID=-1
	%>
<form method=post action='?act=setDriver'>
	<select name='driver'>
	<%
	while not rs.eof
		if oldCostID<>rs("costID") then response.write("<optgroup label='"&rs("costName")&"'>")
	%>
		<option value='<%=rs("id")%>'><%=rs("name")%></option>
	<%
	
		oldCostID=rs("costID")
		rs.moveNext
		if not rs.eof then 
			if oldCostID<>rs("costID") then response.write("</optgroup>")
		end if
	wend
	rs.close
	%>

	</select>
	<input type=submit value='Ê—Êœ „ﬁœ«—'> 
</form>
<%	
elseif request("act")="setDriver" then 
	
	set rs=Conn.Execute("select * from cost_drivers where id="&request.form("driver"))
	%>
<form method=post action='?act=addCost'>
	<span><%=rs("name")%></span>
	<input name='diverValue' type=text>
	<%
	if rs("has_order")="True" then response.write("<input name='order' type=text>")
	%>
	<span><%=rs("unitSize")%></span>
</form>
	<%
end if
%>
<!--#include file="tah.asp" -->
