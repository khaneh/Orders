	  <Tr><Td colspan="2" height="10px">
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
</STYLE>


	  &nbsp;</Td></Tr>
	  <Tr>
		<Td colspan="2" valign="top" align="center">

				<% if Auth("B" , 2) then %><input class="GenButton" type="button" value="æÑæÏ ÇÚáÇãíå" onclick="window.open('../AO/AO.asp');" <% if AccountIsDisabled then %> disabled <% end if %>> <% end if %> 

				<% if Auth("B" , 4) then %><input class="GenButton" type="button" value="ÏæÎÊä" onclick="window.open('../AO/ItemsRelation.asp?sys=AO&act=relate&selectedCustomer=<%=CusID%>');"> <% end if %> 

				<% if Auth("B" , 1) then %><input class="GenButton" type="button" value="ÒÇÑÔ ÍÓÇÈ" onclick="window.open('../AO/AccountReport.asp?sys=AO&act=show&selectedCustomer=<%=CusID%>');"> <% end if %> 

				<% if Auth(8 , "B") then %><input class="GenButton" type="button" value="ÒÇÑÔ ÊÝÕíáí" onclick="window.open('../accounting/tafsili.asp?accountID=<%=CusID%>&accountName=<%=escape(AccountTitle)%>&DateTo=<%=shamsitoday()%>&DateFrom=<%=fiscalYear & "/01/01"%>&moeenFrom=00000&moeenTo=99999&act=Show');"> <% end if %> 

				<BR><BR><BR>

		</Td>
	  </Tr>

  	  <Tr>
		<Td valign="top" align="left">
		</Td>
		<Td valign="top" align="center">
		</Td>
  	  </Tr>