<%
'Bank (10 [=A])
PageTitle= "»«‰ﬂ - " +PageTitle
menuItem="A"
if not Auth("A" , 0) then response.end
%>
<!--#include file="../menu.asp" -->
<TABLE cellspacing=0 cellpadding=0 width=100% height=450 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR height=20 bgcolor=<%=SelectedMenuColor%>>
	<TD>
		<TABLE cellspacing=0 cellpadding=0 width=100%>
		<TR class=alak height=25>

			<%if Auth("A" , 1) then %>
			<%if SubmenuItem="1" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=70 align=center background="/images/MTB.gif" class='alak2'><A HREF='cheq.asp'>’œÊ— çﬂ</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=70 align=center background="/images/MTS.gif" class='alak'><A HREF='cheq.asp'>’œÊ— çﬂ</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth("A" , 4) then %>
			<%if SubmenuItem="4" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=70 align=center background="/images/MTB.gif" class='alak2'><A HREF='submitDraft.asp'>À»  ÕÊ«·Â</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=70 align=center background="/images/MTS.gif" class='alak'><A HREF='submitDraft.asp'>À»  ÕÊ«·Â</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth("A" , 5) then %>
			<%if SubmenuItem="5" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD width=100 align=center background="/images/MTB.gif" class='alak2'><A HREF='payment.asp'>Å—œ«Œ  ›«ﬂ Ê—</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD width=100  align=center background="/images/MTS.gif" class='alak'><A HREF='payment.asp'>Å—œ«Œ  ›«ﬂ Ê—</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth("A" , 6) then %>
			<%if SubmenuItem="6" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='CheqMoeen.asp'>œ—Ì«›  çﬂ „⁄Ì‰</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='CheqMoeen.asp'>œ—Ì«›  çﬂ „⁄Ì‰</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth("A" , 7) then %>
			<%if SubmenuItem="7" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='CheqMoeenSodoodr.asp'>’œÊ— çﬂ „⁄Ì‰</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='CheqMoeenSodoodr.asp'>’œÊ— çﬂ „⁄Ì‰</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth("A" , 8) then %>
			<%if SubmenuItem="8" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='CheqBook.asp'>œ› — çﬂ </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='CheqBook.asp'>œ› — çﬂ </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>


			<%if Auth("A" , 9) then %>
			<%if SubmenuItem="9" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='Contra.asp'>„€«Ì— </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='Contra.asp'>„€«Ì— </A></TD>
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
