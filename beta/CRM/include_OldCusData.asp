<%
' This include file needs OLD_ACID 

%>
<TABLE border="1" width="100%" cellspacing="1" cellpadding="5">
<%
	mySQL="SELECT * FROM [U82ACC] WHERE [Field2] = N'" & OLD_ACID & "' "
	Set RS1 = conn.Execute(mySQL)
	'response.write mySQL
	col1Color="#FFFFFF"
	col2Color="#CCCCCC"
	if (RS1.eof) then%>
		<tr><td dir='rtl' align='center'><h1>ÅÌœ« ‰‘œ...</h1></td></tr>
<%	else
		mySQL="SELECT Users.RealName as RealName FROM tblCSR_ACC_Relation INNER JOIN Users ON tblCSR_ACC_Relation.CSR = Users.ID WHERE (tblCSR_ACC_Relation.Account = '"& Request("acID")& "')"
		Set RS2 = conn.Execute(mySQL)
		if RS2.eof then
			CSRName="<B><FONT Color='red'>" & "‰«„‘Œ’" & "</FONT></B>"
		else
			CSRName="<B><FONT Color='blue'>" & RS2("RealName") & "</FONT></B>"
		end if
		RS2.close
%>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= CSRName %>&nbsp;
			</td>
			<td bgcolor="<%=col2Color%>" width="100">„”∆Ê· ÅÌêÌ—Ì :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field3") %>&nbsp;</td><td bgcolor="<%=col2Color%>">‰«„ Õ”«» :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field4") %>&nbsp;</td><td bgcolor="<%=col2Color%>"> Ì — ÕﬁÊﬁÌ :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field5") %>&nbsp;</td><td bgcolor="<%=col2Color%>">”„  :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field7") %>&nbsp;</td><td bgcolor="<%=col2Color%>">ﬂœ «ﬁ ’«œÌ :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field8") %>&nbsp;</td><td bgcolor="<%=col2Color%>">ﬂœ Å” Ì :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field9") %>&nbsp;</td><td bgcolor="<%=col2Color%>">ﬂœ „·Ì :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%=replace(RS1("Field10"),"‹","-") %>&nbsp;</td><td bgcolor="<%=col2Color%>"> ·›‰ :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%=replace(RS1("Field11"),"‹","-") %>&nbsp;</td><td bgcolor="<%=col2Color%>">Fax :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field12") %>&nbsp;</td><td bgcolor="<%=col2Color%>">¬œ—” :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field13") %>&nbsp;</td><td bgcolor="<%=col2Color%>">‘Â— :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field14") %>&nbsp;</td><td bgcolor="<%=col2Color%>">ŒÌ«»«‰ :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field15") %>&nbsp;</td><td bgcolor="<%=col2Color%>">ﬂÊçÂ :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field16") %>&nbsp;</td><td bgcolor="<%=col2Color%>">‘„«—Â :</td>
		</tr>
		<tr dir='rtl'>
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field17") %>&nbsp;</td><td rowspan="6" bgcolor="<%=col2Color%>"> Ê÷ÌÕ :</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field18") %>&nbsp;</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field19") %>&nbsp;</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field20") %>&nbsp;</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field21") %>&nbsp;</td>
		</tr>
		<tr dir='rtl'>
			<td bgcolor="<%=col1Color%>"><%= RS1("Field22") %>&nbsp;</td>
		</tr>
<%
	end if    
%>
</TABLE>