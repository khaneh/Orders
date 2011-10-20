	<TABLE Class="OrderFile" cellspacing='1' align=center style='width:700;background-color:red;border:2 solid black;'>
		<tr>
			<td colspan="2" class="orderTableHeader" style="text-align:center;">فايلهاي مربوط به اين سفارش</td> 
			<td colspan='2' class="orderTableHeader" style="text-align:left;">
				<input class="GenButton" type="button" value="اضافه کردن فايل" onclick="window.location='file://///intelserver2500/OrdersStorage/<%=cstr(Order)%>';">
				<input type='file' value='اضافه کردن فايل'>
			</td>
		</tr>
		<%
		dim fso, fld, file
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		on error resume next
		set fld = fso.GetFolder("\\intelserver2500\OrdersStorage\" & cstr(Order))
		tmpCounter = 0
		for each file in fld.Files
			tmpCounter = tmpCounter + 1
			%>
			<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4"%>">
				<td>
					<a href='file://///intelserver2500/OrdersStorage/<%=cstr(Order)%>/<%=file.Name%>'>
						<%=left(file.Name,inStrRev(file.Name,".") - 1)%>
					</a>
				</td>
				<td>
					<%=file.Type%>
				</td> 
				<td>
					<%=separate(file.Size)%>
				</td>
				<td>
					<%=file.DateCreated%>
				</td>
			</tr>
			<%
		next
		if tempCounter = 0 then
%>
		<tr class="CusTD3">
			<td colspan="4">هيچ</td>
		</tr>
		<%
		end if
		set fso = nothing
		set fld = nothing
		set file = nothing
		%>
	</TABLE>