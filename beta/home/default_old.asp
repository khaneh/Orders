<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Home (0)
session("id")=6
PageTitle= "ŒÊ«‰œ‰ ÅÌ«„Â«"
SubmenuItem=1
if not Auth(0 , 1) then NotAllowdToViewThisPage()

captionFGColor = "#C6C6D7"
captionBGColor = "#F4F4FE"
MsgBodyColor = "#F4F4FE" '"#FFFFEE"
UrgentMsgBodyColor = "#FFDDDD"
VeryUrgentMsgBodyColor = "yellow"
OrderReadyMsgBodyColor = "#33FF99" '#00CCFF
VeryUrgentMsgScrlColor = "#FFFFCC"
MsgMainColor = "#FFFFFF"
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>
.scrollOrderReady {
	scrollbar-3dlight-color:	<%=captionFGColor%>; 
	scrollbar-arrow-color:		<%=captionFGColor%>;
	scrollbar-base-color:		<%=captionFGColor%>;
	scrollbar-darkshadow-color:	<%=captionFGColor%>;
	scrollbar-face-color:		<%=OrderReadyMsgBodyColor%>;
	scrollbar-highlight-color:	<%=OrderReadyMsgBodyColor%>;
	scrollbar-shadow-color:		<%=captionFGColor%> 
	width:170; height:160; overflow:auto;
	}
.scrollNorm {
	scrollbar-3dlight-color:	<%=captionFGColor%>; 
	scrollbar-arrow-color:		<%=captionFGColor%>;
	scrollbar-base-color:		<%=captionFGColor%>;
	scrollbar-darkshadow-color:	<%=captionFGColor%>;
	scrollbar-face-color:		<%=MsgMainColor%>;
	scrollbar-highlight-color:	<%=MsgMainColor%>;
	scrollbar-shadow-color:		<%=captionFGColor%> 
	width:170; height:160; overflow:auto;
	}
.scrollUrg {
	scrollbar-3dlight-color:	<%=MsgMainColor%>; 
	scrollbar-arrow-color:		<%=MsgMainColor%>;
	scrollbar-base-color:		<%=MsgMainColor%>;
	scrollbar-darkshadow-color:	<%=MsgMainColor%>;
	scrollbar-face-color:		<%=UrgentMsgBodyColor%>;
	scrollbar-highlight-color:	<%=UrgentMsgBodyColor%>;
	scrollbar-shadow-color:		<%=MsgMainColor%> 
	width:170; height:160; overflow:auto;
	}
.scrollSpec {
	scrollbar-3dlight-color:	<%=MsgMainColor%>; 
	scrollbar-arrow-color:		<%=MsgMainColor%>;
	scrollbar-base-color:		<%=MsgMainColor%>;
	scrollbar-darkshadow-color:	<%=MsgMainColor%>;
	scrollbar-face-color:		<%=VeryUrgentMsgScrlColor%>;
	scrollbar-highlight-color:	<%=VeryUrgentMsgScrlColor%>;
	scrollbar-shadow-color:		<%=MsgMainColor%> 
	width:170; height:160; overflow:auto;
	}
.ArchiveDiv {
	scrollbar-3dlight-color:	<%=captionBGColor%>; 
	scrollbar-arrow-color:		<%=captionBGColor%>;
	scrollbar-base-color:		<%=captionBGColor%>;
	scrollbar-darkshadow-color:	<%=captionBGColor%>;
	scrollbar-face-color:		<%=captionFGColor%>;
	scrollbar-highlight-color:	<%=captionBGColor%>;
	scrollbar-shadow-color:		<%=captionBGColor%> 
	width:180; height:400; overflow:auto;
	}
.captionBut {
	border: solid 1pt white; 
	cursor:hand; 
	background-color: <%=captionFGColor%>; 
	width:10
	}
.captionBut a {	text-decoration:none; color: black;}
</style>
<center>
<BR><BR>
<TABLE>
<TR>
	<TD  valign=top>
		<TABLE width=530 border=0 align=center cellpadding=0 cellspacing=6>
		<TR>
			<TD colspan='3' height='1'></TD>
		</TR>
		<TR>
		<%
		if request("act") = "close" then
			MySQL = "UPDATE Messages SET IsRead = 1 WHERE (id = "& request("id") & ")"
			conn.Execute (MySQL)
		end if

		if request("act") = "minimize" then
			MySQL = "UPDATE Messages SET IsSmall = 1 WHERE (id = "& request("id") & ")"
			conn.Execute (MySQL)
		end if

		if request("act") = "mazimize" then
			MySQL = "UPDATE Messages SET IsSmall = 0 WHERE (id = "& request("id") & ")"
			conn.Execute (MySQL)
		end if
		'--------------------------------------------------------------------------------
		'--------------------------------------------------------------------------------
		'--------------------------------------------------------------------------------
		'--------------------------------------------------------------------------------
		MySQL = "SELECT * FROM Messages WHERE (MsgTo = "& session("id") & ") and  (IsRead = 0) and (IsSmall = 0) order by id DESC"
		set RSM = conn.Execute (MySQL)
		cnt = 0 
		Do While Not (RSM.eof) 
			set RSV=Conn.Execute ("SELECT RealName FROM Users where ID=" & RSM("MsgFrom")) 
			%>
			<TD valign=top bgcolor=<% if RSM("Urgent")=0 then %><%=MsgMainColor%><% elseif RSM("Urgent")=2 then %><%=VeryUrgentMsgBodyColor%><% elseif RSM("Urgent")=3 then %><%=OrderReadyMsgBodyColor%><% else %><%=UrgentMsgBodyColor%><% end if %> width=170 height=170 bgcolor=white style='border: solid 1pt black '>
				<div>
				<div  style=" height:10;position: relative; top:0; left:0">
				<TABLE cellpadding=0 cellspacing=0 width=100% >
				<TR height=5 bgcolor=<%=captionBGColor%>>
					<TD title="Õ–›" class="captionBut"><A class="" onclick="return confirm('¬Ì« ÕﬁÌﬁ « „Ì ŒÊ«ÂÌœ «Ì‰ —« Å«ﬂ ﬂ‰Ìœø')" HREF="default.asp?act=close&id=<%=RSM("id")%>"> X </A></TD>
					<TD title="¬—‘ÌÊ" class="captionBut"><A HREF="default.asp?act=minimize&id=<%=RSM("id")%>">&nbsp;v </A></TD>
					<TD align=left  valign=top width=15><FONT COLOR="<%=SelectedMenuColor%>">«“: </FONT></TD>
					<TD align=right valign=top width=75><div style="width:75px; text-overflow : ellipsis; overflow : hidden;"><NOBR>&nbsp; <%=RSV("RealName")%></NOBR></div></TD>
					<TD class="captionBut"><A HREF="message.asp?act=reply&id=<%=RSM("id")%>&retURL=<%=Server.URLEncode("default.asp")%>">Å«”Œ</a></TD>
					<TD class="captionBut"><A HREF="message.asp?act=forward&id=<%=RSM("id")%>&retURL=<%=Server.URLEncode("default.asp")%>">«—Ã«⁄</a></TD>
				</TR>
				<TR height=1 bgcolor=black>
					<td colspan=6></td>
				</TR>
				</TABLE>
				</div>
				<div  <% if RSM("Urgent")=0 then %>class="scrollNorm"<% elseif RSM("Urgent")=1 then  %>class="scrollUrg"<% elseif RSM("Urgent")=2 then  %>class="scrollSpec"<% elseif RSM("Urgent")=3 then  %>class="scrollOrderReady"<% end if %>">
				
				<TABLE cellpadding=0 cellspacing=0 width=100% border=0>
				<TR>
				</TR>
				<% if RSM("Urgent")=2 then %>
				<TR>
					<TD align=center colspan=2> <FONT size=5 COLOR="<%=SelectedMenuColor%>">ŒÌ·Ì ›Ê—Ì</FONT>
					<hr noshade style="height:1"></TD>
				</TR>
				<% end if %>
				<TR>
					<TD align=left  valign=top width=45> <FONT COLOR="<%=SelectedMenuColor%>"> «—ÌŒ: </FONT></TD>
					<TD align=right valign=top width=125>&nbsp; <span dir=ltr><%=RSM("MsgDate")%></span></TD>
				</TR>
				<TR>
					<TD align=left  valign=top width=45> <FONT COLOR="<%=SelectedMenuColor%>">”«⁄ : </FONT></TD>
					<TD align=right valign=top width=125>&nbsp; <%=RSM("MsgTime")%></TD>
				</TR>
				<% if NOT RSM("IsReply")=0 and not RSM("replyTo")=0 then %>
				<TR>
					<TD align=left class="" valign=top width=45> <FONT COLOR="<%=SelectedMenuColor%>">ﬁ»·Ì:</FONT></TD>
<% 
					MySQL = "SELECT MsgBody FROM Messages WHERE  (id = "& RSM("replyTo") & ")"

					set RSO = conn.Execute (MySQL)
					if not RSO.eof then
%>
					<TD align=right valign=top width=125>
					<TEXTAREA NAME="" ROWS="1" readonly COLS="25" style="font-family: tahoma; font-size: 7pt; border: none; background: transparent" title="<%=RSO("MsgBody")%>"><%=RSO("MsgBody")%></TEXTAREA>
					</TD>
				</TR>
<% 
					end if
					RSO.close
				end if 
				if RSM("RelatedTable")<>"NaN" and RSM("RelatedID")<>"0"  then %>
				<TR>
					<TD align=left  valign=top width=45> <FONT COLOR="<%=SelectedMenuColor%>">·Ì‰ﬂ: </FONT> </TD>
					<TD align=right valign=top width=125><%
							RelatedTable = "NaN"
							if trim(RSM("RelatedTable")) = "orders" then 
								RelatedTable="”›«—‘ ‘„«—Â "
								RelatedLink = "../order/TraceOrder.asp?act=show&order=" & RSM("RelatedID")
							elseif trim(RSM("RelatedTable")) = "accounts" then 
								RelatedTable="‘„«—Â Õ”«» "
								RelatedLink = "../CRM/AccountInfo.asp?act=show&SelectedCustomer="& RSM("RelatedID")
							elseif trim(RSM("RelatedTable")) = "invoices" then 
								RelatedTable="›«ﬂ Ê— ‘„«—Â "
								RelatedLink = "../AR/AccountReport.asp?act=showInvoice&invoice="  & RSM("RelatedID")
							end if
					%> <A HREF="<%=RelatedLink%>" target="_blank"><%=RelatedTable%> <%=RSM("RelatedID")%></A> </TD>
				</TR>
				<% end if %>
				<TR height=3>
					<TD align=right colspan=2> 
					<hr noshade style="height:1">
					</TD>
				</TR>
				<TR>
					<TD align=right colspan=2> 
					<%=replace(RSM("MsgBody"),chr(13),"<br>")%>
					</TD>
				</TR>
				</TABLE>
				</div>
				</div>
			</TD>
<%
			cnt = cnt + 1
			if cnt mod 3 = 0 then
				response.write "</tr><tr>"
			end if
			RSM.MoveNext
		Loop
		RSM.close
%>
		<td  valign=top>
		</td>
		</TR>
		</TABLE>
	</TD>
	<TD valign=top>
		<TABLE width=170 border=0 align=center cellpadding=0 cellspacing=6 >
		<TR height=100% >
			<TD width=170 valign=top height=100% style="border-right: 2pt dashed; border-top: 2pt dashed; border-color: white" >
		<CENTER style="background-color:<%=captionFGColor%>">¬—‘ÌÊ ÅÌ«„Â«</CENTER>
		<div class="ArchiveDiv">
		<table  cellpadding=0 cellspacing=4 >
		<%

		'--------------------------------------------------------------------------------
		'--------------------------------------------------------------------------------
		'--------------------------------------------------------------------------------
		'--------------------------------------------------------------------------------
		MySQL = "SELECT * FROM Messages WHERE (MsgTo = "& session("id") & ") and  (IsRead = 0) and (IsSmall = 1) order by id DESC"
		set RSM = conn.Execute (MySQL)
		cnt2 = 0 
		Do While Not (RSM.eof) 
			set RSV=Conn.Execute ("SELECT RealName FROM Users where ID=" & RSM("MsgFrom")) 
			%>
			<tr  valign=top><TD valign=top bgcolor=<% if RSM("Urgent")=0 then %><%=MsgMainColor%><% elseif RSM("Urgent")=2 then %><%=VeryUrgentMsgBodyColor%><% elseif RSM("Urgent")=3 then %><%=OrderReadyMsgBodyColor%><% else %><%=UrgentMsgBodyColor%><% end if %> width=170 height=40 bgcolor=white style='border: solid 1pt black '>
				<TABLE cellpadding=0 cellspacing=0 width=100%  title="<%=replace(RSM("MsgBody"),"""","`")%>" >
				<TR height=5 bgcolor=<%=captionBGColor%>>
					<A onclick="return confirm('¬Ì« ÕﬁÌﬁ « „Ì ŒÊ«ÂÌœ «Ì‰ —« Å«ﬂ ﬂ‰Ìœø')" HREF="default.asp?act=close&id=<%=RSM("id")%>"><TD class="captionBut"> X</TD></A>
					<A HREF="default.asp?act=mazimize&id=<%=RSM("id")%>"><TD class="captionBut">^</TD></A>
					<TD align=right valign=top width=200> &nbsp;<%=RSV("RealName")%> </TD>
					<TD align=left valign=top width=60><span dir=ltr><%=RSM("msgDate")%></span></TD>
				</TR>
				<TR height=1 bgcolor=black>
					<td colspan=5></td>
				</TR>
				<TR >
					<td colspan=5 ><INPUT readonly TYPE="text"   size=35 style="font-family:tahoma; font-size:10; background: transparent; border: solid 0pt	" value="<%=replace(left(RSM("MsgBody"),50),"""","`")%>"></td>
				</TR>
				</TABLE>
			</TD></tr>
			<%
			cnt2 = cnt2 + 1
			if cnt2 mod 4 = 0 then
				'response.write "</table></td><td  valign=top><table  cellpadding=0 cellspacing=4>"
				cnt = cnt + 1
			if cnt mod 4 = 0 then
				'response.write "</table></td></tr><tr><td  valign=top><table cellpadding=0 cellspacing=4>"
			end if
			end if
			RSM.MoveNext
		Loop

		RSM.close
		%>
		</table>
		</div>
			</TD>

		</td>
		</TR>
		</TABLE>
	
	</TD>
</TR>
</TABLE>
<!--#include file="tah.asp" -->