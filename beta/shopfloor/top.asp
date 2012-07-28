<%
'shopfloor (3)
PageTitle= "ÊæáíÏ - " +PageTitle
menuItem=3

%>
<!--#include file="../menu.asp" -->
<TABLE cellspacing=0 cellpadding=0 width=100% height=450 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR height=20 bgcolor=<%=SelectedMenuColor%>>
	<TD>
		<TABLE cellspacing=0 cellpadding=0 width=100%>
		<TR class=alak height=25>

			<%if Auth(3 , 1) then %>
			<%if SubmenuItem="1" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='default.asp'>æÑæÏ ãÑÇÍá ÊæáíÏ </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='default.asp'>æÑæÏ ãÑÇÍá ÊæáíÏ </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(3 , 2) then %>
			<%if SubmenuItem="2" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='manageOrder.asp'>ÈÑäÇãå ÑíÒí</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='manageOrder.asp'>ÈÑäÇãå ÑíÒí</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(3 , 3) then %>
			<%if SubmenuItem="3" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='TraceOrder.asp?act=search'>ííÑí ÓİÇÑÔ</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='TraceOrder.asp?act=search'>ííÑí ÓİÇÑÔ</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(3 , 6) then %>
			<%if SubmenuItem="6" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='orderDesk.asp'>ÌÏæá ÊæáíÏ</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='orderDesk.asp'>ÌÏæá ÊæáíÏ</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			<%if Auth(3 , 7) then %>

			<%if SubmenuItem="7" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='costs.asp'>ÚãáßÑÏ</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='costs.asp'>ÚãáßÑÏ</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			
			<%if Auth(3 , 8) then %>

			<%if SubmenuItem="8" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='inStock.asp'><small>ãæÌæÏí ßÇÛĞ</small></A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='inStock.asp'><small>ãæÌæÏí ßÇÛĞ</small></A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<TD  width=200 align=left bgcolor=<%=SelectedMenuColor%>></TD>
			<TD  align=left bgcolor=<%=unSelectedSubMenuColor%>><A HREF="../logout.asp">ÎÑæÌ </A></TD>
		</TR>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD colspan=2 valign=top bgcolor=<%=AppFgColor%>>
	
	<%
	if request.queryString("errmsg")<>"" then
		response.write "<br>" 
		call showAlert (request.queryString("errmsg"),CONST_MSG_ERROR) 
		response.write "<br>" 
	end if
	if request.queryString("msg")<>"" then
		response.write "<br>" 
		call showAlert (request.queryString("msg"),CONST_MSG_INFORM) 
		response.write "<br>" 
	end if
	%>
