<%
'Order (2)
PageTitle= "������� - " & PageTitle
menuItem=2

%>
<!--#include file="../menu.asp" -->
<TABLE cellspacing=0 cellpadding=0 width=100% height=450 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR height=20 bgcolor=<%=SelectedMenuColor%>>
	<TD>
		<TABLE cellspacing=0 cellpadding=0 width=100%>
		<TR class=alak>

			<%if Auth(2 , 1) then %>
			<%if SubmenuItem="1" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='order.asp?act=makeNewQoute'>���� ������� </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='order.asp?act=makeNewQoute'>���� ������� </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(2 , 2) then %>
			<%if SubmenuItem="2" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='order.asp?act=getEdit'>����� </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='order.asp?act=getEdit'>����� </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(2 , 3) then %>
			<%if SubmenuItem="3" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='order.asp?act=trace'>����� </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='order.asp?act=trace'>����� </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			
			<%if Auth(2 , 3) then %>
			<%if SubmenuItem="4" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='order.asp?act=getShow'>����� </A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='order.asp?act=getShow'>����� </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(2 , 6) then %>
			<%if SubmenuItem="5" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='reports.asp'>����� ���</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='reports.asp'>����� ��� </A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			<%if Auth(2 , 7) then %>
			<%if SubmenuItem="6" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='ManagerReports.asp'>����� ������</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='ManagerReports.asp'>����� ������</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			
			<%if Auth(2 , 3) then %>
			<%if SubmenuItem="7" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='orderFolow.asp'>������ �������</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='orderFolow.asp'>������ �������</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>
			
			<%if Auth(2 , 3) then %>
			<%if SubmenuItem="8" then %> 
				<TD width=10><img src='/images/RTB.gif'></td>
				<TD align=center background="/images/MTB.gif" class='alak2'><A HREF='inquiryFolow.asp'>������ �������</A></TD>
				<TD width=10><img src='/images/LTB.gif'></td>
			<%else %>  
				<TD width=10><img src='/images/RTS.gif'></td>
				<TD align=center background="/images/MTS.gif" class='alak'><A HREF='inquiryFolow.asp'>������ �������</A></TD>
				<TD width=10><img src='/images/LTS.gif'></td>
			<%end if %>
			<%end if %>

			
			<TD align=left bgcolor=<%=unSelectedSubMenuColor%>><A HREF="../logout.asp">���� </A></TD>
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
