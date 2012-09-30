<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="áíÓÊ ãæÌæÏí ßÇÛĞ ÇİÓÊ ÏÑ ÇäÈÇÑ"
SubmenuItem=8
if not Auth(3 , 8) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<style>
	.w{background-color: white;}
	.w td{padding: 5px;}
	.g{background-color: #CCC;}
	.g td{padding: 5px;}
	th{background: #dd5;}
</style>
<br>
<br>
<br>
<center>
	<table cellspacing="0">
	<tr>
		<th>ßÏ ÇäÈÇÑ</th>
		<th>ÚäæÇä</th>
		<th>ãæÌæÏí</th>
	</tr>
	<%
	set rs=Conn.Execute("select * from InventoryItems inner join InventoryItemCategoryRelations on InventoryItemCategoryRelations.Item_ID=InventoryItems.id where Cat_ID=1 and enabled=1 and qtty>0 and owner=-1 order by InventoryItems.name ")
	cls="w"
	while not rs.eof 
		if cls="w" then
			cls="g"
		else
			cls="w"
		end if
	%>
	<tr class="<%=cls%>">
		<td><%=rs("oldItemID")%></td>
		<td><%=rs("name")%></td>
		<td><%=Separate(rs("qtty"))%></td>
	</tr>
	<%
		rs.moveNext
	wend
	%>
	</table>
</center>
<!--#include file="tah.asp" -->