<%
'AR (6)
PageTitle= "«” ⁄·«„ - " +PageTitle
menuItem="F"

%>
<!--#include file="../menu.asp" -->
<TABLE cellspacing=0 cellpadding=0 width=100% height=380 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR height=20 bgcolor=<%=SelectedMenuColor%>>
	<TD>
		<TABLE cellspacing=0 cellpadding=0 width=100%>
		<TR class=alak height=25 >

			<%'if Auth(6 , 1) or Auth(6 , 3) then %>
			<%if SubmenuItem="1" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='Inquiry.asp'> «” ⁄·«„ </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='Inquiry.asp'> «” ⁄·«„ </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%'end if %>

			<%if Auth(6 , 6) then %>
			<%if SubmenuItem="6" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='AccountReport.asp?act=search'>ê“«—‘ Õ”«»</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='AccountReport.asp?act=search'>ê“«—‘ Õ”«»</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(6 , 8) then %>
			<%if SubmenuItem="8" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='CSRreport.asp?act=show'>ê“«—‘ «” ⁄·«„</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='CSRreport.asp?act=show'>ê“«—‘ «” ⁄·«„</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth("C" , 0) then %>
			<%if SubmenuItem="11" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='OtherReports.asp'>”«Ì— ê“«—‘ Â«</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='OtherReports.asp'>”«Ì— ê“«—‘ Â«</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<TD  width=50 align=left bgcolor=<%=SelectedMenuColor%>></TD>
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
