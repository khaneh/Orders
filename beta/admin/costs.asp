<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="„—ﬂ“ Â“Ì‰ÂùÂ«"
SubmenuItem=11
%>
<!--#include file="top.asp" -->
<%
function echoUnitSize(unitID)
	result="<select name='unitSize'>"
	set rss = Conn.Execute("select * from cost_unitSizes")
	while not rss.eof
		result=result&"<option value='"&rss("id")&"'"
		if cint(rss("id"))=unitID then result=result&" selected='selected'"
		result=result&">"&rss("name")&"</option>"
		rss.moveNext
	wend
	result=result&"</select>"
	echoUnitSize=result
	rss.close
end function
if request("step")="" or request("step") = "costCenters" then 
	if request("act")="" then
%>
		<form method=post action='?act=add&step=costCenters'>
			<span>„—ﬂ“ Â“Ì‰Â:</span>
			<input type=text name='costCenterName'>
			<input type=submit value='«÷«›Â ‘Êœ'>
		</form>
		<%
		set rs=Conn.Execute("SELECT cost_centers.id,cost_centers.name, count(cost_drivers.id) as driverCount FROM cost_centers left outer join cost_drivers on cost_centers.id=cost_drivers.cost_center_id group by cost_centers.id,cost_centers.name")
		while not rs.eof
			response.write("<a title='ÃÂ  ‰„«Ì‘ „—ﬂ“ Â“Ì‰Â ﬂ·Ìﬂ ﬂ‰Ìœ' href='?act=show&step=driver&id="&rs("id")&"'>"&rs("Name")&"(‘«„· "&rs("driverCount")&" œ—«ÌÊ—)</a><br>")
			rs.moveNext
		wend
		rs.close
	elseif request("act")="add" then  
		if request.form("costCenterName")<>"" then Conn.Execute("Insert into cost_centers ([Name]) VALUES (N'"&request.form("costCenterName")&"')")
	end if
elseif request("step")="driver" then
	response.write("<a href='costs.asp'>»«“ê‘  »Â ·Ì”  „—«ﬂ“</a>")
	if request("act")="show" then
		if isNumeric(request("id")) then 
			id=request("id")
			sql = "SELECT cost_drivers.id, cost_drivers.name, cost_drivers.cost_center_id, cost_drivers.rate, cost_drivers.is_sold, cost_drivers.has_order, cost_drivers.is_continuous,cost_unitSizes.name as unitSizeName, cost_unitSizes.id as unitSizeID FROM cost_drivers inner join cost_unitSizes on cost_drivers.unitsize=cost_unitSizes.id WHERE cost_center_id="&id
			'response.write sql
			set rs=Conn.Execute(sql)
			%>
			<table>
				<tr>
					<td>œ—«ÌÊ—</td>
					<td>÷—Ì»</td>
					<td>Ê«Õœ «‰œ«“Â êÌ—Ì</td>
					<td>is sold</td>
					<td>has order</td>
					<td>ÅÌÊ” êÌ “„«‰Ì œ«—œ</td>
					<td colspan=2></td>
				</tr>
				<%
				while not rs.eof
					%>
					<tr>
						<form method=post action='?act=edit&step=driver&centerID=<%=id%>'>
							<input type=hidden name='rowID' value='<%=rs("id")%>'>
							<td><input type=text name='driverName' value='<%=rs("name")%>'></td>
							<td><input type=text name='rate' value='<%=rs("rate")%>' style='direction:LTR;'></td>
							<td><%=echoUnitsize(rs("unitsizeID"))%></td>
							<td><input type=checkbox name='isSold' <% if rs("is_sold")="True" then response.write "checked" %>></td>
							<td><input type=checkbox name='hasOrder' <% if rs("has_order")="True" then response.write "checked" %>></td>
							<td><input type=checkbox name='isContinuous' <% if rs("is_continuous")="True" then response.write "checked"%>></td>
							<td><input type=submit value=' €ÌÌ—«  –ŒÌ—Â ‘Êœ' ></td>
							<td><input type=button value='Å«ﬂ ‘Êœ' onclick='if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „ÌùŒÊ«ÂÌœ «Ì‰ œ—«ÌÊ— —« Å«ﬂ ﬂ‰Ìœø\n\n")) window.location="costs.asp?act=del&step=driver&rowID=<%=rs("id")%>&id=<%=id%>";'></td>
						</form>
					</tr>
					<%
					rs.moveNext
				wend
				%>
				<tr>
					<form method=post action='?act=add&step=driver&centerID=<%=id%>'>
						<td><input type=text name='driverName'></td>
						<td><input type=text name='rate' style='direction:LTR;'></td>
						<td><%=echoUnitsize(0)%></td>
						<td><input type=checkbox name='isSold'></td>
						<td><input type=checkbox name='hasOrder'></td>
						<td><input type=checkbox name='isContinuous'></td>
						<td colspan=2 align=center><input type=submit name='action' value='«÷«›Â ‘Êœ'></td>
					</form>
				</tr>
			</table>
			
			<%
		
			rs.close
		end if
	elseif request("act")="add" then 
		isSold="0"
		hasOrder="0"
		isContinuous="0"
		if request.form("isSold")="on" then isSold="1"
		if request.form("hasOrder")="on" then hasOrder="1"
		if request.form("isContinuous")="on" then isContinuous="1"
		driverName=sqlsafe(request.form("driverName"))
		id=cint(request("centerID"))
		rate=cdbl(request("rate"))
		unitSize=sqlsafe(request.form("unitSize"))
		
		sql="INSERT INTO cost_drivers ([name],[cost_center_id],[rate],[unitsize],[is_sold],[has_order],is_Continuous) VALUES (N'"& driverName & "'," & id & "," & rate & "," & unitSize & "," & isSold & "," & hasOrder &","&isContinuous&")"
		'response.write(sql)
		Conn.Execute(sql)
		msg="œ—«ÌÊ— „Ê—œ ‰Ÿ— »Â „—ﬂ“ Â“Ì‰Â «÷«›Â ‘œ"
		response.redirect "?step=driver&act=show&id="&id&"&msg=" & Server.URLEncode(msg)
	elseif request("act")="edit" then 
		isSold="0"
		hasOrder="0"
		isContinuous="0"
		if request.form("isSold")="on" then isSold="1"
		if request.form("hasOrder")="on" then hasOrder="1"
		if request.form("isContinuous")="on" then isContinuous="1"
		driverName=sqlsafe(request.form("driverName"))
		id=cint(request("centerID"))
		rate=cdbl(request("rate"))
		unitSize=sqlsafe(request.form("unitSize"))
		rowID=request.form("rowID")
		sql="update cost_drivers SET [name]=N'"& driverName & "',[rate]=" & rate & ",[unitsize]=" & unitSize & ",[is_sold]=" & isSold & ",[has_order]=" & hasOrder &",is_Continuous="&isContinuous &" WHERE id="&rowID
		'response.write(sql)
		Conn.Execute(sql)
		msg="œ—«ÌÊ— „Ê—œ ‰Ÿ— »—Ê“ —”«‰Ì ‘œ"
		response.redirect "?step=driver&act=show&id="&id&"&msg=" & Server.URLEncode(msg)
	elseif request("act")="del" then 
		if isNumeric(request("rowID")) then
			Conn.Execute("DELETE cost_drivers WHERE id=" & cint(request("rowID")))
			msg="œ—«ÌÊ— „Ê—œ ‰Ÿ— Õ–› ‘œ"
			response.redirect "?step=driver&act=show&id="&request("id")&"&msg=" & Server.URLEncode(msg)
		end if
	end if
end if	
%>
<!--#include file="tah.asp" -->
