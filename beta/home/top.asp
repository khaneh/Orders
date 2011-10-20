<%
' home (0)
PageTitle= "”·«„ - "& PageTitle
menuItem=0
%>
<!--#include file="../menu.asp" -->
<TABLE cellspacing=0 cellpadding=0 width=100% height=450 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR height=20 bgcolor=<%=SelectedMenuColor%>>
	<TD>
		<TABLE cellspacing=0 cellpadding=0 width="100%">
		<TR class='alak' height='25'>
			<%if Auth(0 , 1) then %>
			<%if SubmenuItem="1" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=80 align=center background="/images/MTB.gif" class='alak2'><A HREF='default.asp?sub=1'>ŒÊ«‰œ‰ ÅÌ«„ </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=80 align=center background="/images/MTS.gif" class='alak'><A HREF='default.asp?sub=1'>ŒÊ«‰œ‰ ÅÌ«„ </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(0 , 2) then %>
			<%if SubmenuItem="2" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=80 align=center background="/images/MTB.gif" class='alak2'><A HREF='message.asp'>‰Ê‘ ‰ ÅÌ«„</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=80 align=center background="/images/MTS.gif" class='alak'><A HREF='message.asp'>‰Ê‘ ‰ ÅÌ«„</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(0 , 3) then %>
			<%if SubmenuItem="3" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=80 align=center background="/images/MTB.gif" class='alak2'><A HREF='phoneBook.asp?sub=3'>œ› —  ·›‰</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=80 align=center background="/images/MTS.gif" class='alak'><A HREF='phoneBook.asp?sub=3'>œ› —  ·›‰</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(0 , 4) then %>
			<%if SubmenuItem="4" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=120 align=center background="/images/MTB.gif" class='alak2'><A HREF='furlough.asp'>œ—ŒÊ«”  „—Œ’Ì</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=120 align=center background="/images/MTS.gif" class='alak'><A HREF='furlough.asp'>œ—ŒÊ«”  „—Œ’Ì</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(0 , 5) then %>
			<%if SubmenuItem="5" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=100 align=center background="/images/MTB.gif" class='alak2'><A HREF='goodReq.asp'>œ—ŒÊ«”  ﬂ«·«</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=100 align=center background="/images/MTS.gif" class='alak'><A HREF='goodReq.asp'>œ—ŒÊ«”  ﬂ«·«</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth("D" , 0) then %>
			<%if SubmenuItem="6" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=60 align=center background="/images/MTB.gif" class='alak2'><A HREF='currentWorks.asp?panel=1'>„Ì“ ﬂ«—</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=60 align=center background="/images/MTS.gif" class='alak'><A HREF='currentWorks.asp?panel=1'>„Ì“ ﬂ«—</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			
			<%if Auth(0 , 8) then %>
			<%if SubmenuItem="7" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=60 align=center background="/images/MTB.gif" class='alak2'><A HREF='wiki.asp'>ﬂ «»Œ«‰Â</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=60 align=center background="/images/MTS.gif" class='alak'><A HREF='wiki.asp'>ﬂ «»Œ«‰Â</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<TD  width=* align=left bgcolor=<%=SelectedMenuColor%>></TD>
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
	
