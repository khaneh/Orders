<%
'CRM (1)
PageTitle= "„‘ —Ì«‰ - " & PageTitle
menuItem=1
%>
<!--#include file="../menu.asp" -->
<TABLE cellspacing=0 cellpadding=0 width=100% height=450 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR height=20 bgcolor=<%=SelectedMenuColor%>>
	<TD>
		<TABLE cellspacing=0 cellpadding=0 width=100%>
		<TR class=alak>

			<%if Auth(1 , 1) then %>
			<%if SubmenuItem="1" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='AccountInfo.asp'>‰„«Ì‘ «ÿ·«⁄« </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='AccountInfo.asp'>‰„«Ì‘ «ÿ·«⁄« </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(1 , 2) then %>
			<%if SubmenuItem="2" then %>
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='AccountEdit.asp?act=getAccount'>ÃœÌœ</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='AccountEdit.asp?act=getAccount'>ÃœÌœ</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(1 , 3) then %>
			<%if SubmenuItem="3" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='AccountEdit.asp'>«’·«Õ</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='AccountEdit.asp'>«’·«Õ</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			
			<%if Auth(1 , 6) then %>
			<%if SubmenuItem="6" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='reports.asp'>ê“«—‘ùÂ«</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='reports.asp'>ê“«—‘ùÂ«</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			
			<%if Auth(1 , 8) then %>
			<%if SubmenuItem="7" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='rfmModel.asp'>œ‘»Ê—œ „‘ —Ì«‰</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='rfmModel.asp'>œ‘»Ê—œ „‘ —Ì«‰</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			
			<%if Auth(1 , 9) then %>
			<%if SubmenuItem="4" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='campain.asp'>œ‘»Ê—œ ﬂ„ÅÌ‰</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='campain.asp'>œ‘»Ê—œ ﬂ„ÅÌ‰</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<TD  width=150 align=left bgcolor=<%=SelectedMenuColor%>></TD>
			<TD  align=left bgcolor=<%=unSelectedSubMenuColor%>><A HREF="../logout.asp">Œ—ÊÃ </A></TD>
		</TR>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD colspan="2" valign="top" bgcolor="<%=AppFgColor%>" style="padding:2px;">
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
