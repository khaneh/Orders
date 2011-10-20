<%
'Purchase (4)
PageTitle= "Œ—Ìœ  - " +PageTitle
menuItem=4

%>
<!--#include file="../menu.asp" -->
<TABLE cellspacing=0 cellpadding=0 width=100% height=450 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR height=20 bgcolor=<%=SelectedMenuColor%>>
	<TD>
		<TABLE cellspacing=0 cellpadding=0 width=100%>
		<TR class=alak height=25>

			<%if Auth(4 , 1) then %>
			<%if SubmenuItem="1" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='default.asp'> œ—ŒÊ«”  Œ—Ìœ ﬂ«·«Ì «‰»«—</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='default.asp'> œ—ŒÊ«”  Œ—Ìœ ﬂ«·«Ì «‰»«—</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(4 , 2) then %>
			<%if SubmenuItem="2" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='otherRequests.asp'>œ—ŒÊ«”  Œ—Ìœ ”«Ì—ﬂ«·« Â« Ê Œœ„« </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='otherRequests.asp'>œ—ŒÊ«”  Œ—Ìœ ”«Ì—ﬂ«·« Â« Ê Œœ„« </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(4 , 3) then %>
			<%if SubmenuItem="3" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='outServiceOrder.asp'>”›«—‘ Œ—Ìœ</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='outServiceOrder.asp'>”›«—‘ Œ—Ìœ</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(4 , 4) then %>
			<%if SubmenuItem="4" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='outServiceTrace.asp'>ÅÌêÌ—Ì Œ—Ìœ</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='outServiceTrace.asp'>ÅÌêÌ—Ì Œ—Ìœ</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(4 , 5) then %>
			<%if SubmenuItem="5" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='relateInvItemsToPurchOrders.asp'>œÊŒ ‰</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='relateInvItemsToPurchOrders.asp'>œÊŒ ‰</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<TD  width=110 align=left bgcolor=<%=SelectedMenuColor%>></TD>
			<TD  align=left bgcolor=<%=unSelectedSubMenuColor%>><A HREF="../logout.asp">Œ—ÊÃ </A></TD>
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
