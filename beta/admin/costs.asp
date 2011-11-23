<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="„—ﬂ“ Â“Ì‰ÂùÂ«"
SubmenuItem=11
%>
<!--#include file="top.asp" -->
<%
function echoOperationType(operationID, driverID)
	result="<select name='operationType'>"
	set rss = Conn.Execute("select * from cost_operation_type where driver_id=" & driverID)
	while not rss.eof
		result = result & "<option value='" & rss("id") & "'"
		if cint(rss("id")) = operationID then result = result & " selected='selected'"
		result = result & ">" & rss("name") & "</option>"
		rss.moveNext
	wend
	result = result & "</select>"
	echoOperationType = result
	rss.close
end function
 
if request("act")="" or request("act")="" then
%>
	<form method=post action='?act=add&step=costCenters'>
		<span>„—ﬂ“ Â“Ì‰Â:</span>
		<input type=text name='costCenterName'>
		<input type=submit value='«÷«›Â ‘Êœ'>
	</form>
	<%
	set rs=Conn.Execute("SELECT cost_centers.id,cost_centers.name, count(cost_drivers.id) as driverCount FROM cost_centers left outer join cost_drivers on cost_centers.id=cost_drivers.cost_center_id group by cost_centers.id,cost_centers.name")
	while not rs.eof
		response.write("<a title='ÃÂ  ‰„«Ì‘ „—ﬂ“ Â“Ì‰Â ﬂ·Ìﬂ ﬂ‰Ìœ' href='?act=show&step=driver&id="&rs("id")&"'>"&rs("Name")&"(‘«„· "&rs("driverCount")&" œ—«ÌÊ—)</a>")
		if CInt(rs("driverCount"))=0 then response.write("<a style='margin:0 10 0 0;' title='»—«Ì Å«ﬂ ‘œ‰ «Ì‰ „—ﬂ“ ﬂ·Ìﬂ ﬂ‰Ìœ' href='?act=del&step=center&id="&rs("id")&"'>«Ì‰ „—ﬂ“ —« Å«ﬂ ﬂ‰Ìœ!</a>")
		response.write("<br>")	
		rs.moveNext
	wend
	rs.close
elseif request("act")="add" then  '------------------ ADD ACTION ----------------
	if request("step")="costCenters" and request.form("costCenterName")<>"" then 
		'--------------- add cost center ------------------
		Conn.Execute("Insert into cost_centers ([Name]) VALUES (N'"&request.form("costCenterName")&"')")
		msg="„—ﬂ“ Â“Ì‰Â " & "<b><i>" & request.form("costCenterName") & "</i></b>" & " «÷«›Â ‘œ."
		response.redirect "?msg=" & Server.URLEncode(msg)
	elseif request("step")="operationType" and request("opName")<>"" then
		'-------------- add operation type -----------------
		Conn.Execute("INSERT INTO cost_operation_type (driver_id,name) VALUES ("&request("driverID")&",N'"&request("opName")&"')")
		msg="⁄„·Ì«  " & "<b><i>" & request.form("opName") & "</i></b>" & " «÷«›Â ‘œ."
		response.redirect "?act=show&step=driver&id="&request("centerID")&"&msg=" & Server.URLEncode(msg)
	elseif request("step")="driver" then
		'------------------ add driver ---------------------
		isDirect="0"
		isContinuous="0"
		if request.form("isDirect")="on" then isDirect="1"
		if request.form("isContinuous")="on" then isContinuous="1"
		driverName=sqlsafe(request.form("driverName"))
		id=cint(request("centerID"))
		rate=cdbl(request("rate"))
		theType=cint(request("type"))
		sql="INSERT INTO cost_drivers ([name],[cost_center_id],[rate],[is_direct],is_countinuous,[type]) VALUES (N'"& driverName & "'," & id & "," & rate & "," & isDirect & ","&isContinuous&", " & theType & ")"
		response.write(sql)
		Conn.Execute(sql)
		msg="œ—«ÌÊ— „Ê—œ ‰Ÿ— »Â „—ﬂ“ Â“Ì‰Â «÷«›Â ‘œ"
		response.redirect "?step=driver&act=show&id="&id&"&msg=" & Server.URLEncode(msg)
	end if
elseif request("act")="del" then '----------------------- DELETE -------------------------------
	if request("step")="center" then 
		if cint(request("id"))>0 then  
			Conn.Execute("delete cost_centers where id="&request("id"))
			msg="„—ﬂ“ Â“Ì‰Â Õ–› ‘œ."
			response.redirect "?msg=" & Server.URLEncode(msg)
		end if
	elseif request("step")="diver" then 
		if isNumeric(request("rowID")) then
			set rs=Conn.Execute("SELECT * from costs where driver_id=" & cint(request("rowID")))
			if rs.eof then 
				Conn.Execute("DELETE cost_operation_type WHERE driver_id=" & cint(request("rowID")))
				Conn.Execute("DELETE cost_drivers WHERE id=" & cint(request("rowID")))
				msg="œ—«ÌÊ— „Ê—œ ‰Ÿ— Ê ⁄„·Ì« ùÂ«Ì „—»ÊÿÂ Õ–› ‘œ"
			else
				msg="çÊ‰ «“ «Ì‰ œ—«ÌÊ— «” ›«œÂ ‘œÂ «“ Å«ﬂ ﬂ—œ‰ ¬‰ „⁄–Ê—„"
			end if
			rs.close
			response.redirect "?step=driver&act=show&id="&request("id")&"&msg=" & Server.URLEncode(msg)
		end if
	elseif request("step")="operationType" then
		if isNumeric(request("rowID")) then
			set rs=Conn.Execute("SELECT * from costs where operation_type=" & cint(request("rowID")))
			if rs.eof then 
				Conn.Execute("DELETE cost_operation_type WHERE id=" & cint(request("rowID")))
				msg="⁄„·Ì«  „Ê—œ ‰Ÿ— Õ–› ‘œ"
			else
				msg="çÊ‰ «“ «Ì‰ ⁄„·Ì«  «” ›«œÂ ‘œÂ «“ Å«ﬂ ﬂ—œ‰ ¬‰ „⁄–Ê—„"
			end if
			response.redirect "?step=driver&act=show&id="&request("id")&"&msg=" & Server.URLEncode(msg)
		end if
	end if
elseif request("act")="edit" then '------------------- EDIT -----------------------
	if request("step")="driver" then 
		isDirect="0"
		isContinuous="0"
		if request.form("isDirect")="on" then isDirect="1"
		if request.form("isContinuous")="on" then isContinuous="1"
		driverName=sqlsafe(request.form("driverName"))
		id=cint(request("centerID"))
		rate=cdbl(request("rate"))
		unitSize=sqlsafe(request.form("unitSize"))
		rowID=request.form("rowID")
		sql="update cost_drivers SET [name]=N'"& driverName & "',[rate]=" & rate & ", [is_direct]=" & isDirect & ", is_countinuous="&isContinuous &",type="&request("type")&" WHERE id="&rowID
		'response.write(sql)
		Conn.Execute(sql)
		msg="œ—«ÌÊ— „Ê—œ ‰Ÿ— »—Ê“ —”«‰Ì ‘œ"
		response.redirect "?step=driver&act=show&id="&id&"&msg=" & Server.URLEncode(msg)
	elseif request("step")="operationType" then
		sql="UPDATE cost_operation_type set [name]=N'"&request("opName")&"' WHERE id=" & request("id")
		Conn.Execute(sql)
		msg="⁄„·Ì«  „Ê—œ ‰Ÿ— »—Ê“ —”«‰Ì ‘œ"
		response.redirect "?step=driver&act=show&id="&request("centerID")&"&msg=" & Server.URLEncode(msg)
	end if
elseif request("act")="show" then
	if request("step")="driver" then	
		if isNumeric(request("id")) then 
			id=request("id")
			set rsCost= Conn.Execute("SELECT * FROM cost_centers where id="&id)
			if rsCost.eof then 
				response.write("Œÿ«! «Ì‰ „—ﬂ“ Ÿ«Â—« Å«ﬂ ‘œÂ!!!!!!!")
				response.end
			end if
			costCenterName=rsCost("name")
			rsCost.close
			sql = "SELECT * FROM cost_drivers WHERE cost_center_id="&id
			'response.write sql
			set rs=Conn.Execute(sql)
			%>
			<a title="»«“ê‘  »Â ·Ì”  „—ﬂ“ Â“Ì‰ÂùÂ«" href="costs.asp">ÊÌ—«Ì‘ „—ﬂ“ Â“Ì‰Â <%=costCenterName%></a>
			<br><br>
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td align="center"><b>‰«„</b></td>
					<td align="center"><b>Ê«Õœ «‰œ«“Â êÌ—Ì</b></td>
					<td align="center"><b>÷—Ì»</b></td>
					<td align="center"><b>„” ﬁÌ„</b></td>
					<td align="center"><b>ÅÌÊ” êÌ</b></td>
					<td colspan=2></td>
				</tr>
				<%
				while not rs.eof
					%>
					<tr>
						<form method=post action='?act=edit&step=driver&centerID=<%=id%>'>							
							<td><input type="text" name="driverName" value="<%=rs("name")%>"></td>
							<td>
								<input type=hidden name='rowID' value='<%=rs("id")%>'>
								<select name="type">
									<option value="1" <%if cint(rs("type"))=1 then response.write(" selected='selected'")%>>⁄œœ</option>
									<option value="2" <%if cint(rs("type"))=2 then response.write(" selected='selected'")%>>“„«‰</option>
								</select>
							</td>
							<td><input type=text name='rate' value='<%=rs("rate")%>' style='direction:LTR;'></td>
							<td><input type=checkbox name='isDirect' <% if rs("is_direct")="True" then response.write "checked" %>></td>
							<td><input type=checkbox name='isContinuous' <% if rs("is_countinuous")="True" then response.write "checked"%>></td>
							<td><input type=submit value=' €ÌÌ—«  –ŒÌ—Â ‘Êœ' ></td>
							<td><input type=button value='Å«ﬂ ‘Êœ' onclick='if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „ÌùŒÊ«ÂÌœ «Ì‰ œ—«ÌÊ— Ê ﬂ·ÌÂ ⁄„·Ì« ùÂ«Ì „—»ÊÿÂ —« Å«ﬂ ﬂ‰Ìœø\n\n")) window.location="costs.asp?act=del&step=driver&rowID=<%=rs("id")%>&id=<%=id%>";'></td>
						</form>
					</tr>
					<tr bgcolor="#AAFFFF">
						<td colspan="7" align="center"><b>⁄„·Ì« ùÂ«</b></td>
					</tr>
					<%
					set rsop=Conn.Execute("select * from cost_operation_type where driver_id=" & rs("id"))
					if rsop.eof then 
					%>
					<tr bgcolor="#AAFFFF">
						<td colspan="7" align="center"><font color="red">ÂÌç ⁄„·Ì« Ì ÅÌœ« ‰‘œ! </font></td>
					</tr>
					
					<%
					end if
					while not rsop.eof
					%>
					<tr bgcolor="#AAFFFF">
						<form method="post" action="?act=edit&step=operationType&centerID=<%=id%>">
							<td colspan="5">
								<input type="hidden" name="id" value="<%=rsop("id")%>">
								<input name="opName" type="text" value="<%=rsop("name")%>">
							</td>
							<td><input type=submit value=' €ÌÌ—«  –ŒÌ—Â ‘Êœ' ></td>
							<td><input type=button value='Å«ﬂ ‘Êœ' onclick='if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „ÌùŒÊ«ÂÌœ «Ì‰ ⁄„·Ì«  —« Å«ﬂ ﬂ‰Ìœø\n\n")) window.location="costs.asp?act=del&step=operationType&rowID=<%=rsop("id")%>&id=<%=id%>";'></td>
						</form
					</tr>
					<%
						rsop.moveNext
					wend
					rsop.close
					%>
					<tr bgcolor="#AAAAFF">
						<form method="post" action="?act=add&step=operationType&centerID=<%=id%>">
							<td colspan="1" align="center">‰«„ ⁄„·Ì«  —« Ê«—œ ﬂ‰Ìœ</td>
							<td colspan="4">
								<input type="text" name="opName">
								<input type="hidden" name="driverID" value="<%=rs("id")%>">
							</td>
							<td colspan="2" align="center">
								<input type="submit" name="action" value="«÷«›Â ‘Êœ">
							</td>
						</form>
					</tr>
					<%
					rs.moveNext
				wend
				rs.close
				%>
				<tr>
					<form method=post action='?act=add&step=driver&centerID=<%=id%>'>
						<td><input type="text" name="driverName"></td>
						<td>
							<select name="type">
								<option value='1'>⁄œœ</option>
								<option value='2'>“„«‰</option>
							</select>
						</td>
						<td><input type=text name='rate' style='direction:LTR;'></td>
						<td><input type=checkbox name='isDirect'></td>
						<td><input type=checkbox name='isContinuous'></td>
						<td colspan=2 align=center><input type=submit name='action' value='«÷«›Â ‘Êœ'></td>
					</form>
				</tr>
			</table>
			
			<%
		end if	
	end if
end if	
%>
<!--#include file="tah.asp" -->
