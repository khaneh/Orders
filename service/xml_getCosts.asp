<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%
Response.Buffer = true
Response.ContentType = "text/xml;charset=windows-1256"
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
<%
select case request("act")
	case "orderCost":
		orderID = CDbl(Request("id"))
		mySQL="select costs.[order], costs.id, costs.gl_update, costs.insert_date, costs.description, isnull(DATEDIFF(mi,costs.start_time,costs.end_time),0) as theTime,end_counter - start_counter as theCount, cost_operation_type.name as operationName, cost_drivers.name as driverName, cost_centers.name as centerName, users.RealName from costs inner join cost_operation_type on costs.operation_type=cost_operation_type.id inner join cost_drivers on cost_operation_type.driver_id=cost_drivers.id inner join cost_centers on cost_drivers.cost_center_id=cost_centers.id inner join Users on costs.user_id=users.ID where [order]=" & orderID & " order by cost_centers.name,cost_drivers.name,cost_operation_type.name,costs.id desc"
		set rs=Conn.Execute(mySQL)
		set costs=server.createobject("MSXML2.DomDocument")
		costs.loadXML("<costs/>")
		while not rs.EOF
			
			set cost = costs.createElement("cost")
			set tmp = costs.createElement("id")
			tmp.text = rs("id")
			cost.AppendChild tmp
			set tmp = costs.createElement("insertDate")
			tmp.text = shamsiDate(rs("insert_date"))
			cost.AppendChild tmp
			set tmp = costs.createElement("order")
			if isnull(rs("order")) then 
				tmp.text = ""
			else
				tmp.text = rs("order")
			end if
			cost.AppendChild tmp			
			set tmp = costs.createElement("centerName")
			tmp.text = rs("centerName")
			cost.AppendChild tmp
			set tmp = costs.createElement("driverName")
			tmp.text = rs("driverName")
			cost.AppendChild tmp
			set tmp = costs.createElement("operationName")
			tmp.text = rs("operationName")
			cost.AppendChild tmp
			set tmp = costs.createElement("theCount")
			tmp.text = rs("theCount")
			cost.AppendChild tmp
			set tmp = costs.createElement("theTime")
			tmp.text = floor(cint(rs("theTime"))/60) & ":" & cint(rs("theTime")) mod 60
			cost.AppendChild tmp
			set tmp = costs.createElement("description")
			tmp.text = rs("description")
			cost.AppendChild tmp
			set tmp = costs.createElement("realName")
			tmp.text = rs("realName")
			cost.AppendChild tmp
			set tmp = costs.createElement("canEdit")
			if auth(3,9) and not CBool(rs("gl_update")) then
				tmp.text = "yes"
			else
				tmp.text = ""
			end if
			cost.AppendChild tmp
			
			costs.documentElement.AppendChild cost
			rs.MoveNext
		wend
		response.write(costs.xml)
		rs.close
		set rs = nothing
	case "costCenter":
		centerID = CDbl(Request("id"))
		mySQL="select costs.[order], costs.id, costs.insert_date, costs.gl_update, costs.description, isnull(DATEDIFF(mi,costs.start_time,costs.end_time),0) as theTime,end_counter - start_counter as theCount, cost_operation_type.name as operationName, cost_drivers.name as driverName, cost_centers.name as centerName, users.RealName from costs inner join cost_operation_type on costs.operation_type=cost_operation_type.id inner join cost_drivers on cost_operation_type.driver_id=cost_drivers.id inner join cost_centers on cost_drivers.cost_center_id=cost_centers.id inner join Users on costs.user_id=users.ID where cost_centers.id=" & centerID & " order by cost_centers.name,cost_drivers.name,cost_operation_type.name,costs.id desc"
		set rs=Conn.Execute(mySQL)
		set costs=server.createobject("MSXML2.DomDocument")
		costs.loadXML("<costs/>")
		while not rs.EOF
			set cost = costs.createElement("cost")
			set tmp = costs.createElement("id")
			tmp.text = rs("id")
			cost.AppendChild tmp
			set tmp = costs.createElement("insertDate")
			tmp.text = shamsiDate(rs("insert_date"))
			cost.AppendChild tmp
			set tmp = costs.createElement("order")
			if isnull(rs("order")) then 
				tmp.text = ""
			else
				tmp.text = rs("order")
			end if
			cost.AppendChild tmp
			set tmp = costs.createElement("centerName")
			tmp.text = rs("centerName")
			cost.AppendChild tmp
			set tmp = costs.createElement("driverName")
			tmp.text = rs("driverName")
			cost.AppendChild tmp
			set tmp = costs.createElement("operationName")
			tmp.text = rs("operationName")
			cost.AppendChild tmp
			set tmp = costs.createElement("theCount")
			if IsNull(rs("theCount")) then
				tmp.text=""
			else
				tmp.text = rs("theCount")
			end if
			cost.AppendChild tmp
			set tmp = costs.createElement("theTime")
			tmp.text = floor(cint(rs("theTime"))/60) & ":" & cint(rs("theTime")) mod 60
			cost.AppendChild tmp
			set tmp = costs.createElement("description")
			tmp.text = rs("description")
			cost.AppendChild tmp
			set tmp = costs.createElement("realName")
			tmp.text = rs("realName")
			cost.AppendChild tmp
			set tmp = costs.createElement("canEdit")
			if auth(3,9) and not CBool(rs("gl_update")) then
				tmp.text = "yes"
			else
				tmp.text = ""
			end if
			cost.AppendChild tmp
			
			costs.documentElement.AppendChild cost
			rs.MoveNext
		wend
		response.write(costs.xml)
		rs.close
		set rs = nothing
end select

function floor(x)
	dim temp	
	temp = Round(x)
	if temp > x then
	    temp = temp - 1
	end if
	floor = temp
end function
function ceil(x)
	dim temp	
	temp = Round(x)
	if temp < x then
	    temp = temp + 1
	end if
	ceil = temp
end function
%>