<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Purchase (4)
PageTitle="œ—ŒÊ«”  Œ—Ìœ ”«Ì—ﬂ«·« Â« Ê Œœ„« "
SubmenuItem=2
if not Auth(4 , 2) then NotAllowdToViewThisPage()

'OutService Page Request
'By Alix - Last changed: 81/01/13
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<SCRIPT LANGUAGE="JavaScript">
<!--
function hideIT()
{
if(document.all.tavafogh.checked) 
	{
		document.all.priceTavafoghi.style.visibility= 'visible'
	}
	else
	{
		document.all.price.value= ''
		document.all.priceComment.value= ''
		document.all.priceTavafoghi.style.visibility= 'hidden'
	}
}

function hideIT2()
{
if(document.all.tavafogh2.checked) 
	{
		document.all.priceTavafoghi2.style.visibility= 'visible'
	}
	else
	{
		document.all.orderID.value= ''
		document.all.priceTavafoghi2.style.visibility= 'hidden'
	}
}

//-->
</SCRIPT>

<%
catItem1 = request("catItem")
if catItem1="" then catItem1="-1"


'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------- Delete an Inventory Request for Buy from an order
'-----------------------------------------------------------------------------------------------------
if request("di")="y" then		
	myRequestID=request("i")
	set RSX=Conn.Execute ("SELECT * FROM purchaseRequests WHERE id = "& myRequestID )	
	if RSX("status")="new" then
	Conn.Execute ("update purchaseRequests SET status = 'del' where id = "& myRequestID )	
	end if
	response.redirect "otherRequests.asp?radif=" & request("r")
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------ Submit an Inventory Item request For Buy
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="À»  œ—ŒÊ«”   Œ—Ìœ ﬂ«·«" then
	priceComment = request.form("priceComment")
	otypeName = request.form("typeName")
	comment = request.form("comment")
	price = request.form("price")
	CreatedBy = session("id")
	qtty = request.form("qtty")
	DueDate = request.form("DueDate")
	orderID = request.form("orderID")

	if  comment="" then
		comment = "-"
	end if

	if  qtty="" then
		qtty = "0"
	end if

	if  price="" then
		price = "0"
	end if

	if  priceComment="" then
		priceComment = "-"
	end if

	if  otypeName="" then
		response.write "error"
		response.end
	end if

	if  orderID="" then
		orderID = "-1"
	end if

	mySql="INSERT INTO purchaseRequests (order_ID, typeName, typeID, comment, ReqDate, price,priceComment, CreatedBy, qtty, DueDate, IsService) VALUES ( "& orderID & " , N'"& otypeName & "', 0 , N'"& comment & "',N'"& shamsiToday() & "', "& price & ",N'"& priceComment & "', "& CreatedBy & ", "& qtty & " , N'"& DueDate & "', 1)"	
	conn.Execute mySql
	'RS1.close

	response.write "<center><br><br>œ—ŒÊ«”  À»  ‘œ </center><br>"
end if
%>
<center>
	<BR><BR>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
	<TR  >
		<TD align="right" colspan=2><H3>œ—ŒÊ«”  Œ—Ìœ ”«Ì— ﬂ«·«Â« Ê Œœ„« </H3></TD>
	</TR>
		<TR bgcolor="dddddd" ><td colspan=2>

		<FORM METHOD=POST ACTION="otherRequests.asp">
			‰Ê⁄ ﬂ«·« Ì« Œœ„ : <INPUT TYPE="text" name="typeName" size=31><br><br>
			 ⁄œ«œ: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="text" NAME="qtty" size=40 onKeyPress="return maskNumber(this);" dir="LTR"><br><br>
			<INPUT TYPE="checkbox" onclick="hideIT()" name="tavafogh"> Ê«›ﬁ ﬁÌ„  Œ«’Ì ’Ê—  ê—› Â «” <BR>
			<div name="priceTavafoghi" id="priceTavafoghi" style="visibility:'hidden'">ﬁÌ„ : &nbsp;<INPUT TYPE="text" NAME="price" ID="price" size=5 onKeyPress="return maskNumber(this);"> ÿ—›: <INPUT TYPE="text" NAME="priceComment" id="priceComment" size=26 ></div>

			<INPUT TYPE="checkbox" onclick="hideIT2()" name="tavafogh2">«Ì‰ œ—ŒÊ«”  „—»Êÿ »Â ”›«—‘ Œ«’Ì «” <BR>
			<div name="priceTavafoghi2" id="priceTavafoghi2" style="visibility:'hidden'">‘„«—Â ”›«—‘: &nbsp;<INPUT TYPE="text" NAME="orderID" ID="orderID" size=10 onKeyPress="return maskNumber(this);"> </div>

			 «—ÌŒÌ ﬂÂ ﬂ«·« „Ê—œ ‰Ì«“ «” : <INPUT dir=ltr TYPE="text" NAME="DueDate" size=15 value="<%=shamsiToday()%>" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"><br><br>
		 Ê÷ÌÕ« : <TEXTAREA NAME="comment" ROWS="7" COLS="32"></TEXTAREA>
			<br><center>
			<INPUT class=inputBut TYPE="submit" Name="Submit" Value="À»  œ—ŒÊ«”   Œ—Ìœ ﬂ«·«" style="width:125px;" tabIndex="14">
			</center>
		</FORM>

		</FONT></TD>
	</TR>
	<%
	'Gets Request for services list from DB
	set RS3=Conn.Execute ("SELECT * FROM purchaseRequests WHERE (status='new' and TypeID=0)")
	%>
		<%
		Do while not RS3.eof
		%>
		<TR bgcolor="#CCCCCC" title="<% 
			Comment = RS3("Comment")
			if Comment<>"-" then
				response.write " Ê÷ÌÕ: " & Comment
			else
				response.write " Ê÷ÌÕ ‰œ«—œ"
			end if
		%>">
			<TD align="right" valign=top><FONT COLOR="black">
			<INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RS3("id")%>" <%
			if RS3("status") = "new" then
				response.write " checked disabled "
			else 
				response.write " disabled "
			end if
			%>><B><%=RS3("TypeName")%></B>  <small dir=ltr>(<%=RS3("ReqDate")%>)</small></td>
			<td align=left width=5%><%
			if RS3("status") = "new" then
			%><a href="otherRequests.asp?di=y&i=<%=RS3("id")%>&r=<%=request("radif")%>"><b>Õ–›</b></a><%
			end if %></td>
		</tr>
		<% 
		RS3.moveNext
		Loop
		%>

	</table><br>
	

<!--#include file="tah.asp" -->
