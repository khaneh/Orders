<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<%
if IsNumeric(request("centerID")) and CInt(request("centerID"))>0 then 
rowID=request("row")
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
	<script type="text/javascript">
		var rowID=<%=rowID%>;
		$("select#operationType-" + rowID).change(function (){
	    	$.ajax({
	    		type: 'POST',
	    		url: 'costCenterType.asp',
	    		data: {typeID: $("select#operationType-" + rowID).val()},
	    		dataType: "xml",
	    		success: function(xml) {
	    			$("input#driverID-" + rowID).val($(xml).find('id').text());
	    			$("input#type-" + rowID).val($(xml).find('type').text());
	    			$("input#isDirect-" + rowID).val($(xml).find('isDirect').text());
	    			$("input#isCountiuous-" + rowID).val($(xml).find('isCountiuous').text());
	    			$("input#driverID-" + rowID).change();
	    			$("input#type-" + rowID).change();
	    			$("input#isDirect-" + rowID).change();
	    			$("input#isCountiuous-" + rowID).change();
	    			
	    		}
	    	});
	    });
	    $("select#operationType-" + rowID).ready(function (){
	    	$("select#operationType-" + rowID).change();
	    	/*
$.ajax({
	    		type: 'POST',
	    		url: 'costCenterType.asp',
	    		data: {typeID: $("select#operationType-" + rowID).val()},
	    		dataType: "xml",
	    		success: function(xml){
	    			$("input#driverID-" + rowID).val($(xml).find('id').text());
	    			$("input#type-" + rowID).val($(xml).find('type').text());
	    			$("input#isDirect-" + rowID).val($(xml).find('isDirect').text());
	    			$("input#isCountiuous-" + rowID).val($(xml).find('isCountiuous').text());
	    			$("input#driverID-" + rowID).change();
	    			$("input#type-" + rowID).change();
	    			$("input#isDirect-" + rowID).change();
	    			$("input#isCountiuous-" + rowID).change();
	    		}
	    	});
*/
	    });

	</script>
</head>
	<body>
		<select name="operationType-<%=rowID%>" id='operationType-<%=rowID%>'>
	<%
	mySQL="select cost_drivers.*,cost_operation_type.name as typeName, cost_operation_type.id as typeID from  cost_drivers inner join cost_operation_type on cost_drivers.id=cost_operation_type.driver_id where cost_drivers.cost_center_id=" & request("centerID") & " order by cost_drivers.id, cost_operation_type.id"
	set rs=Conn.Execute(mySQL)
	while not rs.eof
		%>
			<option value="<%=rs("typeID")%>"><%=rs("name") & " - " & rs("typeName")%></option>
		<%
		rs.moveNext
	wend
	rs.close
	%>				
		</select>

	</body>
</html>
<%
elseif IsNumeric(request("typeID")) and CInt(request("typeID"))>0 then
	set rs=Conn.Execute("select cost_drivers.* from cost_drivers inner join cost_operation_type on cost_drivers.id=cost_operation_type.driver_id where cost_operation_type.id=" & request("typeID"))
	if not rs.eof then 
		response.write "<?xml version='1.0'?>"
%>
<driver>
	<id><%=rs("id")%></id>
	<type><%=rs("type")%></type>
	<isDirect><%=rs("is_direct")%></isDirect>
	<isCountiuous><%=rs("is_countinuous")%></isCountiuous>
</driver>
<%
	end if
end if
%>