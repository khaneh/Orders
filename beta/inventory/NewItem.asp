<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= " ����� ����� ����"
SubmenuItem=4
if not Auth(5 , 4) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------- Submit a new Inventory Item
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="��� ������ ����" then
	catItem = request.form("catItem") 
	ItemName = request.form("ItemName") 
	OldItemID = request.form("OldItemID") 
	ownerAcc = request.form("ownerAcc") 
	costingMethod = request.form("cost") 
	Minim = request.form("Minim") 
	qtty = request.form("qtty") 
	Unit = request.form("Unit") 
	outByOrder = request.form("outByOrder")
	
	if outByOrder="on" then
		outByOrder = 1
	else
		outByOrder =0
	end if

	response.write "<BR><BR><TABLE align=center style='border: solid 2pt black'><TR><TD>"
	response.write  "<li> �� ���� ���� = " & catItem& "<br>"
	response.write  "<li> ��� ���� = " & ItemName& "<br>"
	'response.write  "<li> ���� ���� = " 
	'if ownerAcc=-1 then 
	'response.write 	"���� �ǁ � ���<br>"
	'else
	'response.write 	"����� ���� "& ownerAcc & "<br>"
	'end if
	response.write  "<li> �� ���� = " & OldItemID& "<br>"
	response.write  "<li> ���� ���� ����� = " & costingMethod& "<br>"
	response.write  "<li> ����� ������ = " & Minim& "<br>"
	response.write  "<li> ����� = " & qtty& "<br>"
	if outByOrder=1 then
		response.write  "<li> ���� �� ���� �����"
	end if
	response.write  "<li> ���� = " & Unit& "<br>"
	response.write "</TD></TR></TABLE>"

	mySql="INSERT INTO InventoryItems (OldItemID, owner, Name, Minim, Qtty, Unit, costingMethod,outByOrder) VALUES ("& OldItemID & ", "& ownerAcc & ",N'"& ItemName & "', "& Minim & ", 0, N'"& Unit & "','"& costingMethod & "', "& outByOrder & " )"	
	conn.Execute mySql

	set RS4 = conn.Execute ("SELECT * FROM InventoryItems where OldItemID=" & OldItemID & " and Name=N'" & ItemName & "'") 

	if RS4.EOF then
		response.write "<br><br>���! ���� ������ ���. ������ ��� ���� �� �� ���� ����� ���� Ȑ����."
		response.end
	end if

	ItemID = RS4("id")
	RS4.close

	mySql="INSERT INTO InventoryItemCategoryRelations (Item_ID, Cat_ID) VALUES ("& ItemID & ", "& catItem & ")"	
	conn.Execute mySql

	'if Qtty <> 0 then
		mySql="INSERT INTO InventoryLog (ItemID, RelatedID, logDate, Qtty, owner, CreatedBy, IsInput, type) VALUES ("& ItemID & ", -4 ,N'"& shamsiToday() & "', "& Qtty & ", -1 , "& session("id") & ", 1 , 4)"
		conn.Execute mySql
	'end if

	response.write "<center><br>����� ��� �� ������ ������ ��.</center>"
response.end

end if
%>

<%
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- New Item Form
'-----------------------------------------------------------------------------------------------------
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function hideIT()
{
//alert(document.all.aaa2.value)
if(document.all.aaa2.value==2) 
	{
		document.all.aaa1.style.visibility= 'visible'
		document.all.ownerAcc.value = ""
		document.all.ownerAcc.focus()
	}
	else
	{
		document.all.aaa1.style.visibility= 'hidden'
		document.all.ownerAcc.value = "-1"
	}
}
//-->
</SCRIPT>

<BR>
<FORM METHOD=POST ACTION="NewItem.asp">
<TABLE border=0 align=center>
<TR>
	<TD colspan=2 align=center><H3>���� ������� ����</H3></TD>
</TR>
<TR>
	<TD align=left>��� ����</TD>
	<TD align=right>
		<SELECT NAME="catItem" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1"  onchange='check(this);'>
		<option value="-1">���� ���� ���� </option>
		<option value="-1">------------------</option>
		<%
			set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories")
			while not (RS4.eof) %>
				<OPTION value="<%=RS4("ID")%>"<%
				'if trim(catItem1) = trim(RS4("ID")) then
				'response.write " selected "
				'end if
				%>>* <%=RS4("Name")%> </option>
				<%	
				RS4.MoveNext
			wend
			RS4.close
			%>
		</SELECT>
	</TD>
</TR>
<TR>
	<TD align=left>��� ����</TD>
	<TD align=right><INPUT TYPE="text" NAME="ItemName" size=64></TD>
</TR>
<TR>
	<TD align=left>�� ����</TD>
	<TD align=right><INPUT TYPE="text" NAME="OldItemID"></TD>
</TR>
<!--TR>
	<TD align=left>������</TD>
	<TD align=right>
	<SELECT NAME="aaa2"  onchange="hideIT()" >
	<option value=1>���� �ǁ � ���</option>
	<option value=2>����� (����� ����)</option>

	</SELECT></TD>
</TR>
<TR>
	<TD align=left></TD>
	<TD align=right><span name="aaa1" id="aaa1" style="visibility:'hidden'">
	<INPUT TYPE="text" NAME="ownerAcc" value="-1"></span></TD>
</TR>
<TR>
	<TD align=left></TD>
	<TD align=right height=20></TD>
</TR>
<TR>
	<TD align=left>���� �������</TD>
	<TD align=right><INPUT TYPE="radio" NAME="cost" checked value="AVRG">����� <BR>
	<INPUT TYPE="radio" NAME="cost"  value="LIFO">LIFO <BR>
	<INPUT TYPE="radio" NAME="cost" value="FIFO">FIFO</TD>
</TR>
<TR>
	<TD align=left></TD>
	<TD align=right height=20></TD>
</TR-->
<TR>
	<TD align=left>����� ������</TD>
	<TD align=right><INPUT TYPE="hidden" NAME="ownerAcc"  value="-1"><INPUT TYPE="hidden" NAME="cost" value="FIFO"><INPUT TYPE="text" NAME="Minim" value=0></TD>
</TR>
<TR>
	<TD align=left>����� ����</TD>
	<TD align=right><INPUT TYPE="text" NAME="qtty" value=0></TD>
</TR>
<TR>
	<TD align=left>���� ������ ����</TD>
	<TD align=right><INPUT TYPE="text" NAME="Unit" value="���"</TD>
</TR>
<TR>
	<TD align=left><INPUT TYPE="checkbox" NAME="outByOrder"></TD>
	<TD align=right>���� �� ���� �����</TD>
</TR>
<TR>
	<TD align=left></TD>
	<TD align=right height=20>
	<BR><INPUT TYPE="submit" NamE="submit" value="��� ������ ����"onclick="return validateForm()">
	</TD>
</TR>
</TABLE>
</FORM>

<SCRIPT LANGUAGE="JavaScript">
<!--
function validateForm(){ 
if (document.all.ItemName.value=="" || document.all.catItem.value=="" || document.all.OldItemID.value=="")
	{
		alert("���! ��� ���� �� ���� ���")
		return false
	}
}

function check(src){ 
	badCode = false;
	if (window.XMLHttpRequest) {
		var objHTTP=new XMLHttpRequest();
	} else if (window.ActiveXObject) {
		var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
	}
	objHTTP.open('GET','xml_MaxID4newItems.asp?id='+src.value,false)
	objHTTP.send()
	tmpStr = unescape(objHTTP.responseText)
	document.all.OldItemID.value=tmpStr;
}
//-->
</SCRIPT>


<!--#include file="tah.asp" -->
