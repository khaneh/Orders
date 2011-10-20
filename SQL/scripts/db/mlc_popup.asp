<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myElt, myForm
	myElt = Request.QueryString("elt")
	myForm = Request.QueryString("form")

	Dim mlcObj, myDate
	If Request.Form("mlcYear") <> "" Then
		myDate = DateSerial(Request.Form("mlcYear"), Request.Form("mlcMonth"), Request.Form("mlcDay"))
	ElseIf Request.QueryString("initialDate") <> "" Then
		myDate = Request.QueryString("initialDate")
	End If

	Set mlcObj = GetObject("script:" & Server.MapPath(".\..\inc\myLittleCalendar.wsc"))
	If isDate(myDate) Then mlcObj.mlcDate = myDate
	With mlcObj
		.mlcMonthName = myTObj.getTerm(442)
		.mlcDayAbbr = myTObj.getTerm(443)
		.mlcListPosition  = "Top"
		.mlcMonthAndYearPosition   = "Hidden"
		.mlcMonthPosition   = "Hidden"
		.mlcYearPosition   = "Hidden"
		.mlcEmptyCellStyle  = "font-family: Tahoma, Arial; font-size: 8pt; color: #000000; background-color: #DDDDDD;"
		.mlcCellStyle = "font-family: Tahoma, Arial; font-size: 8pt; color: #000000; background-color: #EEEEEE; text-decoration : none;"
		.mlcSelectedCellStyle = "font-family: Tahoma, Arial; font-size: 8pt; color: #000000; background-color: #7BBDBD; font-weight: Bold; text-decoration : none;"
		.mlcWeekDayStyle = "font-family: Tahoma, Arial; font-size: 8pt; color: #000000; background-color: #7BBDBD; font-weight: Bold;"
		.mlcListStyle = "font-family: Tahoma, Arial; color : #000000; font-size : 8pt; background-color: #EEEEEE;"
		.mlcListCellStyle = "font-family: Tahoma, Arial; color : #000000; font-size : 8pt; background-color: #7BBDBD;"
		.mlcFirstDayOfWeek = mla_cfg_firstdayofweek
		.mlcActionURL = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING") 
		.mlcYearRange = 50
	End With
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY <% If Request.Form("mlcDayChosen")  <> "" Then Response.Write "onLoad = ""return(setInfo(document.myLittleCalendar.mlcDate.value));""" End If %>>
	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
	<!--
		function setInfo(pStr)
		{
			if (window.opener.document.<% = myForm %>)
				window.opener.document.<% = myForm %>.<%=myElt%>.value = pStr;
			window.close();
		}
	//-->
	</SCRIPT>

	<DIV ALIGN=CENTER>
	<%
		mlcObj.displayCalendar()
		Set mlcObj = Nothing
	%>
	</DIV>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->
