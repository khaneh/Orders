<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="������ �������"
SubmenuItem=1
%>
<!--#include file="top.asp" -->
<%
NumberOfCategories = 14
dim pagesPermissions(20,25)
dim pages(20,25)
'-------------------------------------
 pages ( 0 , 0 ) = "���� ���"         
	pages ( 0 , 1 ) = "������ ����"   
	pages ( 0 , 2 ) = "����� ����"    
	pages ( 0 , 3 ) = "���� ���� ����"     
	pages ( 0 , 4 ) = "������� �����" 
	pages ( 0 , 5 ) = "������� ����"  
	pages ( 0 , 6 ) = "������ ���� ����"     
	pages ( 0 , 7 ) = "����� ���� �� ���"     
	pages ( 0 , 8 ) = "��������"
'-------------------------------------
 pages ( 1 , 0 ) = "����� �����"
	pages ( 1 , 1 ) = "����� �������"
	pages ( 1 , 2 ) = "����"
	pages ( 1 , 3 ) = "�����"
	pages ( 1 , 4 ) = "����� ����� � ��� ������"
	pages ( 1 , 5 ) = "����� ����� ����"
	pages (1, 6)	= "����ԝ��"
	pages (1, 7)	= "����� ����� ����"
	'pages (1, 8)	= "����� ��� ��� ���"
'-------------------------------------
 pages ( 2 , 0 ) = "�������"
	pages ( 2 , 1 ) = "���� �����"
	pages ( 2 , 2 ) = "����� �����"
	pages ( 2 , 3 ) = "����� �����"
	pages ( 2 , 4 ) = "<s>����� ���� �������</s> deprecated"
	pages ( 2 , 5 ) = "������  �������"
	pages ( 2 , 6 ) = "����� ���"
	pages ( 2 , 7 ) = "����� ������"
	pages ( 2 , 8 ) = "����� ������ - ������ ����� �����"
	pages ( 2 , 9 ) = "�������"
'-------------------------------------
 pages ( 3 , 0 ) = "�����"
	pages ( 3 , 1 ) = "���� ����� �����"
	pages ( 3 , 2 ) = "������ ���� �����"
	pages ( 3 , 3 ) = "����� �����"
	pages ( 3 , 4 ) = "������  �������"
	pages ( 3 , 5 ) = "����� ������� ���� �� ����� ���� ������ �� ������ �� ���� ���� ��� ���"
	pages ( 3 , 6 ) = "���� �����"
'-------------------------------------
 pages ( 4 , 0 ) = "����"
	pages ( 4 , 1 ) = "������� ���� ����� �����"
	pages ( 4 , 2 ) = "������� ���� �������� �� � �����"
	pages ( 4 , 3 ) = "����� ����"
	pages ( 4 , 4 ) = "����� ����"
	pages ( 4 , 5 ) = "�����"
	pages ( 4 , 6 ) = "����� ����� ������"
'-------------------------------------
 pages ( 5 , 0 ) = "�����"
	pages ( 5 , 1 ) = "���� ����� ����� ��� ����� �� ����� �� � �������"
	pages ( 5 , 2 ) = "���� ����"
	pages ( 5 , 3 ) = "���� ����"
	pages ( 5 , 4 ) = "����� ����� ����"
	pages ( 5 , 5 ) = "�����"
	pages ( 5 , 6 ) = "����� ����"
	pages ( 5 , 7 ) = "����� ���� �� ������ "
	pages ( 5 , 8 ) = "����� ������"
	pages ( 5 , 9 ) = "���� ����� ����� ����� �� ����� ��"
	pages ( 5 , 10 ) = "��� ������� ����"					'A
	pages ( 5 , 11 ) = "������ ������"						'B
	pages ( 5 , 12 ) = "��� ����/���� �� ����� ������"		'C
	pages ( 5 , 13 ) = "��� ���� ���� �� �� ���� �����"		'D
'-------------------------------------
 pages ( 6 , 0 ) = "�������� ����"
	pages ( 6 , 1 ) = "���� ������ "
	pages ( 6 , 2 ) = "������� ������/��������"
	pages ( 6 , 3 ) = "����� ������"
	pages ( 6 , 4 ) = "���� ������ �ѐ��"
	pages ( 6 , 5 ) = "����� ������ �ѐ��"
	pages ( 6 , 6 ) = "����� ����"
	pages ( 6 , 7 ) = "�����"
	pages ( 6 , 8 ) = "����� ���� ����"
	pages ( 6 , 9 ) = "����� ���� ���"
	pages ( 6 , 10 ) = "������� ���� ������ ��� �� ����"	'A
	pages ( 6 , 11 ) = "<s>���� ���� �������</s> unused"	'B
	pages ( 6 , 12 ) = "����� ������"						'C
	pages ( 6 , 13 ) = "���� ������"						'D
	pages ( 6 , 14 ) = "�ǁ ������"							'E
	pages ( 6 , 15 ) = "����� ������"						'F
	pages ( 6 , 16 ) = "��� ��� ����"						'G
	pages ( 6 , 17 ) = "����� ������� ����"					'H
	pages ( 6 , 18 ) = "���� ������ �� ����� ������"		'I
	pages ( 6 , 19 ) = "�ǁ ����� ����"						'J
	pages ( 6 , 20 ) = "���� (������ / �ѐ�� ) ���� �����"	'K
	pages ( 6 , 21 ) = "߁� ��� ����"						'L
	pages ( 6 , 22 ) = "����� ������� �� ���� ��� �������� ������"						'M
	pages (6, 23)	 = "������� ��� ������ ��� �� ����"					'N
'-------------------------------------
 pages ( 7 , 0 ) = "�������� ����"
	pages ( 7 , 1 ) = "���� ������ ����"
	pages ( 7 , 2 ) = "����� ������ ����"
	pages ( 7 , 3 ) = "�����"
	pages ( 7 , 4 ) = "������� ������/��������"
	pages ( 7 , 5 ) = "����� ����"
	pages ( 7 , 6 ) = "�����"
	pages ( 7 , 7 ) = "����� ����"
	pages ( 7 , 8 ) = "����� ������� ����"
	pages ( 7 , 9 ) = "����� ������ ����"
	pages ( 7 , 10 ) = "�ǁ ����� ����"						'A

'-------------------------------------
 pages ( 8 , 0 ) = "�������� "
	pages ( 8 , 1 ) = "����  ���"
	pages ( 8 , 2 ) = "���� ��� ��� ��� ���"
	pages ( 8 , 3 ) = "[����] ����� �����"
	pages ( 8 , 4 ) = "���� ��"
	pages ( 8 , 5 ) = "���� ��� ��� ���� ���"
	pages ( 8 , 6 ) = "���� ��� ��� ����� ��� � ����� ����� ����"
	pages ( 8 , 7 ) = "���� ��� ��� ��� ����"
	pages ( 8 , 8 ) = "����� �� ����� ����� ��� ��� ��� ����"
	pages ( 8 , 9 ) = "���� ��� ������� ��"
	pages ( 8 , 10 ) = "����� ����"							'A
	pages ( 8 , 11 ) = "����� ������"						'B
	pages ( 8 , 12 ) = "������ ��� �����"					'C
	pages ( 8 , 13 ) = "����� ���� �� ��� ��� �����"		'D
	pages ( 8 , 14 ) = "���� ���� ����� ��"					'E
	pages ( 8 , 15 ) = "��� ���� (���� � ������) "			'F
	pages ( 8 , 16 ) = "����� ������ ���������"			'G
	'pages ( 8 , 17 ) = "������ʝ��"						'H
'-------------------------------------
 pages ( 9 , 0 ) = "�����"
	pages ( 9 , 1 ) = "������"
	pages ( 9 , 2 ) = "������ ���"
	pages ( 9 , 3 ) = "����� �����"
	pages ( 9 , 4 ) = "����"
	pages ( 9 , 5 ) = "�����"
	pages ( 9 , 6 ) = "����� �с��� �����"
	pages ( 9 , 9 ) = "����� �с��� ����� (���)"
	pages ( 9 , 7 ) = "����� �� ��� � ��� ������/������"
	pages ( 9 , 8 ) = "�ǁ ��� ������"
'-------------------------------------
 pages ( 10 , 0 ) = "����"									'A
	pages ( 10 , 1 ) = "���� ��"
	pages ( 10 , 2 ) = "����� �� �������"
	pages ( 10 , 3 ) = "����� �� �������"
	pages ( 10 , 4 ) = "��� �����"
	pages ( 10 , 5 ) = "������ ������"
	pages ( 10 , 6 ) = "������ �� ����"
	pages ( 10 , 7 ) = "���� �� ����"
	pages ( 10 , 8 ) = "���� ��"
	pages ( 10 , 9 ) = "������"
'-------------------------------------
 pages ( 11 , 0 ) = "�������� ����"							'B
	pages ( 11 , 1 ) = "����� ����"
	pages ( 11 , 2 ) = "���� ������� ��"
	pages ( 11 , 3 ) = "������"
	pages ( 11 , 4 ) = "�����"
	pages ( 11 , 5 ) = "����� ����"
	pages ( 11 , 6 ) = "����� ������� ����"
	pages ( 11 , 7 ) = "�ǁ ����� ����"						
	pages ( 11 , 8 ) = "����"						

'-------------------------------------
 pages ( 12 , 0 ) = "���� ����� ��� ��� �������� ����"		'C
	pages ( 12 , 1 ) = "����� ���� ������"
	pages ( 12 , 2 ) = "����� ���� ������ �� ����� ������"
	pages ( 12 , 3 ) = "����� ���� ��� ���� "
	pages ( 12 , 4 ) = "����� ������ ��� ���"
	pages ( 12 , 5 ) = "����� ���� ��� ���"
	pages ( 12 , 6 ) = "����� ���� ��� ���� ����� �������"
	pages ( 12 , 7 ) = "����� ���� ���� ���"
	pages ( 12 , 8 ) = "����� ���� ���� �����"
	pages ( 12 , 9 ) = "����� �������� �� ������ ���� ��� ����"

'-------------------------------------
 pages ( 13 , 0 ) = "��� ���"								'D
	pages ( 13 , 1 ) = "����� ��� ���"
	pages ( 13 , 2 ) = "������ ��� ����� ����"
	pages ( 13 , 3 ) = "������ ��� ����� ��� (���� ����)"
	pages ( 13 , 4 ) = "������ ��� ���� ��� (����� ����)"
	pages ( 13 , 5 ) = "���� ��� ����� ���"
	pages ( 13 , 6 ) = "�� ��� �ѐ���"
	pages ( 13 , 7 ) = "����� ��� ����� ����"

'-------------------------------------
 pages ( 14 , 0 ) = "������"								'E
	pages ( 14 , 1 ) = "Reserved - ���� �� ���� admin"
	pages ( 14 , 2 ) = "Reserved - ����� ���� �������"
	pages ( 14 , 3 ) = "Reserved"
	pages ( 14 , 4 ) = "Reserved"
	pages ( 14 , 5 ) = "���� ��� ��� ���� �������"
	pages ( 14 , 6 ) = "���� ����ԝ��� �����"


function Auth(menuID, subMenuID, permission)
	pr = permission

	st = inStr(pr,"#"&menuID)
	
	if subMenuID >= "A" then 
		subMenuID_int = cint(asc(subMenuID)-55)
	else
		subMenuID_int = cint(subMenuID)
	end if

	Auth = false
	if st > 0 then
		en = inStr(st+1, pr, "#")
		sm = inStr(st+2, pr, subMenuID)
		if subMenuID_int = 0 or ((sm <> 0) and (en > sm or en = 0 )) then
			Auth = true
		end if		
	end if
end function


%>
<style>
	.TABLE1 {font-family: tahoma; font-size: 8pt; Background-Color:navy; border:2 solid navy; cursor:pointer;}
	.TD_Rowspan {Background-Color:navy;}
	.textbox1 {border:none; width:100px; font-family: tahoma; font-size: 8pt; Background-Color:transparent; cursor:pointer;} 
	.TABLE1 TR {Background-Color:white; height:20px; }
	.TD10 {Background-Color:#CCCCCC;height:10px;}
	.TD00 {Background-Color:#BBBBFF;height:10px;}
	.TD11 {Background-Color:#FFFFFF;height:10px;}
	.TD01 {Background-Color:#DDDDFF;height:10px;}
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
function showUser(user){
	window.location="?act=edit&userID="+user;
}
//-->
</SCRIPT>
<%

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------ Users List
'-----------------------------------------------------------------------------------------------------
if request("act")="" then
	set RSc=Conn.Execute ("SELECT count(*) as c FROM Users WHERE ID<>0")
%>
	<BR>

	<TABLE class="table1" border=0 cellspacing=1 cellpadding=1 align=right>
	<TR>
		<TD rowspan="3" height=100>�����</TD>
	</tr>
	<TR>
		<%
		for i=0 to NumberOfCategories
			j=0
			if not pagesPermissions(i,j)="ok" then 
				disableChecks="disabled"
				rowBGColor="#F0F0F0"
				groupChecked=""
			else
				disableChecks=""
				rowBGColor="#33AACC"
				groupChecked="checked"
			end if

			Do While pages(i,j+1)<>"" 
				j=j+1
			Loop

			%>
			<td colspan=<%=j%> align=center ><%=pages(i,0)%></td>
 			<td rowspan=<%=RSc("c")+3%> class="TD_Rowspan" ></td>
<%
		next
%>
	</tr>
	<TR>
<%
		for i=0 to NumberOfCategories
			j=0
			while pages(i,j+1)<>"" 
				j=j+1

				if j <10 then
					jj = j
				else
					jj = chr(55+j)
				end if
				
				%>
				<td title="<%=pages(i,j)%>"><%response.write jj'pages(i,j)%></td>
				<%
			wend
		next

%>
	</TR>
<% 
	set RSV=Conn.Execute ("SELECT * FROM Users WHERE ID<>0 ORDER BY RealName") 
	tmpRowCounter=0
	Do while not RSV.eof
		Permission = RSV("Permission")
		RealName = RSV("RealName")
		tmpRowCounter = tmpRowCounter + 1

%>
		<TR height='10px' onclick="showUser(<%=RSV("ID")%>);">
			<TD height='10px' class="TD<%=tmpRowCounter Mod 2%>1"><INPUT TYPE="text" class="textbox1" Value="<%=RealName%>"><br></TD>
<%
			tmpColCounter=0
			for i=0 to NumberOfCategories
				j=0
				while pages(i,j+1)<>"" 
					tmpColCounter = tmpColCounter + 1
					j=j+1

					if i <10 then
						ii = i
					else
						ii = chr(55+i)
					end if

					if j <10 then
						jj = j
					else
						jj = chr(55+j)
					end if
					
					%>
					<td height='10px' class="TD<%=tmpRowCounter Mod 2%><%=tmpColCounter Mod 2%>" title="<%=RealName & ": " & vbCrLf & pages(i,j) %>"><%if Auth (ii,jj,Permission) then %><B>x</B><%else%>&nbsp;<%end if%></td>
					<%
				wend
			next

%>

		</TR>
<%
		RSV.moveNext

	Loop
	RSV.close
%>
		<TR height='10px' onclick="showUser('');">
			<TD height='10px' class="TD<%=(tmpRowCounter+1) Mod 2%>1"><INPUT TYPE="text" class="textbox1" Value="*����*"><br></TD>
			
			
			
		</TR>
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ Show Permissions
'-----------------------------------------------------------------------------------------------------
elseif request("act")="edit" then
%>
	<BR><BR>
<%
	userID = request("userID")
	if userID="" then 
		Password	= ""
%>
		<center>
		<H3>����� ����</H3>
		</center>
<%
	else
		userID = cint(request("userID"))
		set RSM = conn.Execute ("SELECT * from Users where ID="& userID & "")

		UserName	= RSM("UserName")
		Password	= "$aMe@sB4" 'Same as Before  instead of : RSM("Password") 
		if left(RSM("Password"),1)="'" then
			DisabledChecked="checked"
		else
			DisabledChecked=""
		end if 
		RealName	= RSM("RealName") 
		Account		= RSM("Account") 
		Permission	= RSM("Permission")
		Display		= RSM("Display")
		if Display then
			DisplayChecked="checked"
		else
			DisplayChecked=""
		end if

		sp = split (Permission,"#")
		for u=1 to ubound(sp)
			x = left(sp(u),1)
			if x >= "A" then 
				i=cint(asc(x)-55)
			else
				i=cint(x)
			end if

			pagesPermissions(i,0)="ok"
			for v = 2 to len(sp(u))
				tmpp=mid(sp(u),v,1)

				if tmpp >= "A" then 
					j=cint(asc(tmpp)-55)
				else
					j=cint(tmpp)
				end if
				'j=cint(mid(sp(u),v,1))
				pagesPermissions(i,j)="ok"
			next
		next
%>
		<center>
		<H3>������ �����: "<%=RealName%>"</H3>
		</center>
<%
	end if
%>

	<FORM METHOD=POST ACTION="?act=submit">

	<TABLE style="font-family:tahoma;font-size:9pt; " Cellspacing="0" Cellpadding="5" width=100%>
	<TR>
		<TD valign=top>

		������:<br><br>
		<TABLE border=1>
		<TR>
			<TD>�� �����:</TD>
			<TD dir=LTR><%=userID%>&nbsp;<INPUT TYPE="hidden" Name="userID" Value="<%=userID%>"></TD>
		</TR>
		<TR>
			<TD>��� ������:</TD>
			<TD><INPUT TYPE="text" NAME="UserName" Value="<%=UserName%>" dir=LTR></TD>
		</TR>
		<TR>
			<TD>��� ����</TD>
			<TD><INPUT TYPE="Password" NAME="Password" Value="<%=Password%>" dir=LTR></TD>
		</TR>
		<TR>
			<TD>��� ����� ���� ���:</TD>
			<TD><INPUT TYPE="text" NAME="RealName" Value="<%=RealName%>"></TD>
		</TR>
		<TR>
			<TD>����� ���� �� �����:</TD>
			<TD><INPUT TYPE="text" NAME="Account" Value="<%=Account%>"></TD>
		</TR>
		<TR>
			<TD>�� ���� ����� ���� ���</TD>
			<TD><INPUT TYPE="checkbox" NAME="Display" <%=DisplayChecked%> ></TD>
		</TR>
		<TR>
			<TD>��� ����� �� ����� ���� ���</TD>
			<TD><INPUT TYPE="checkbox" NAME="disable" <%=DisabledChecked%>></TD>
		</TR>
		<TR>
			<TD colspan=2 align=center><INPUT TYPE="submit" value="�����"></TD>
		</TR>
		</TABLE>

		</TD>
		<TD Width="350" >
		���� ��:<br><br>
		<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;"  Cellspacing="0" Cellpadding="5">
		<tbody id="PermissionsTable">
		<% 
			for i=0 to NumberOfCategories
				j=0
				if not pagesPermissions(i,j)="ok" then 
					disableChecks="disabled"
					rowBGColor="#F0F0F0"
					groupChecked=""
				else
					disableChecks=""
					rowBGColor="#33AACC"
					groupChecked="checked"
				end if
				%>
				<tr bgcolor='<%=rowBGColor%>'>
					<td width=20> <INPUT TYPE="checkbox" NAME="P<%=i%>" value="<%=j%>" onclick="activeGroup(this)" <%=groupChecked%>> </td>
					<td align=right><B><%=pages(i,0)%></B></td>
				</tr>
				<tr>
					<td width=20> </td>
					<td align=right> 
				<%
				while pages(i,j+1)<>"" 
					j=j+1

					if j <10 then
						jj = j
					else
						jj = chr(55+j)
					end if

					%>
					<INPUT TYPE="checkbox" <%=disableChecks%> NAME="P<%=i%>" value="<%=jj%>" <%if pagesPermissions(i,j)="ok" then %>checked<% end if %>> <%=pages(i,j)%> <BR>
					<%
				wend
			next
			%>
					</td>
				</tr>
		</table>
		</TD>
	</TR>
	</TABLE>
	</FORM>

<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------- Submit Changes
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submit" then
	userID		= request("userID")

	RealName	= request("RealName") 
	RealName	= "N'" & RealName & "'"

	Account		= request("Account") 

	UserName	= request("UserName")
	UserName	= "'" & UserName & "'"

	Password	= request("Password") 
	if Password="$aMe@sB4" then
		Password= "Password"
	else
		Password= "'" & Password & "'"
	end if

	if request("disable")="on" then 
		Password = "'''' + REPLACE(" & Password & ",'''','')"
	else
		Password = "REPLACE(" & Password & ",'''','')"
	end if

	Permission = ""
	for i=0 to NumberOfCategories 
		if request.form("P"& i ) <> "" then
			if i <10 then
				ii = i
			else

				ii = chr(55+i)
			end if
			alll = replace(request.form("P"& i ),", ","")
			Permission = Permission	& "#" & ii &  right( alll, len(alll)-1)
		end if
	next
	Permission	= "'" & Permission & "'"

	Display		= request("Display")
	if Display="on" then
		Display="1"
	else
		Display="0"
	end if


	if userID="" then 
		' Add New User
		mySQL="SELECT MAX(ID)+1 AS NewID FROM Users"
		set RS=Conn.Execute (mySQL) 
		userID = RS("NewID")
		RS.close
		mySQL="INSERT INTO Users (ID, UserName, Password, RealName, Account, Permission, Display) VALUES (" & userID & ", " & userName & ", " & Password & ", " & RealName & ", " & Account & ", " & Permission & "," & Display & ")"
		msg="����� ���� ����� ��."
	else
		' Update User Info
		userID = cint(request("userID"))

		mySQL="UPDATE Users SET UserName=" & UserName & ", Password=" & Password & ", RealName=" & RealName & ", Account=" & Account& ", Permission= " & Permission & ", Display=" & Display & " WHERE ID="& userID & ""
		msg="������� �� ��� ����� ����� �����"
	end if

	conn.Execute (mySQL)
	conn.close
	response.redirect "?act=edit&userID=" & userID & "&msg=" & Server.URLEncode(msg)
end if
%>
<SCRIPT LANGUAGE="JavaScript">
<!--

function activeGroup(src){
	rowNo=src.parentNode.parentNode.rowIndex;
	invTable=document.getElementById("PermissionsTable");
	theRowPr=invTable.getElementsByTagName("tr")[rowNo];
	theRow=invTable.getElementsByTagName("tr")[rowNo+1];
	boxCount=theRow.getElementsByTagName("INPUT").length;
	if (src.checked){
		theRowPr.bgColor= '#33AACC';
		for (i=0;i<boxCount;i++){
			theRow.getElementsByTagName("INPUT")[i].disabled=false;
		}
//		theRow.disabled=false;
	}
	else{
		for (i=0;i<boxCount;i++){
			theRow.getElementsByTagName("INPUT")[i].disabled=true;
		}
		theRowPr.bgColor= '#F0F0F0';
//		theRow.disabled=true;
	}

}
//-->
</SCRIPT>
<!--#include file="tah.asp" -->
