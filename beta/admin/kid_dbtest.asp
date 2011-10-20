<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	Table { font-family:tahoma; font-size: 9pt;}
</style>
<TITLE>Test</TITLE>
</HEAD>
<BODY>
<%
    ' Open connection
'	strCnxn = "DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
strCnxn = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"
    Set Cnxn = Server.CreateObject("ADODB.Connection")
    Cnxn.Open strCnxn
    
    ' Open compound recordset
    Set rstCompound = Server.CreateObject("ADODB.Recordset")
    SQLCompound = "exec proc_Test"
    rstCompound.Open SQLCompound, Cnxn, 3 'adOpenStatic
    
    ' Display results from each SELECT statement
    errCount = 0
    catCount = 0
	allCats	= 0
	errCatName =""
	errCatCount=0
	Dim errCatNames (50)
	Dim errCatCounts (50)
	startTime = timer
    Do Until rstCompound Is Nothing
        if rstCompound.EOF then
			response.write " :: <FONT COLOR='green'>OK</FONT><br>"
        else
			if isnumeric(rstCompound.Fields(0)) then
				response.write "<FONT COLOR='red'>"
				catCount = catCount + 1
				errCatNames(catCount) = errCatName
				errCatCount=0
				Do Until rstCompound.EOF
					errCount = errCount + 1
					errCatCount = errCatCount + 1
					response.write "<br>" & rstCompound.Fields(0) 
					rstCompound.MoveNext
				Loop
				errCatCounts(catCount) = errCatCount
				response.write "<br></FONT>"
			else
				allCats = allCats + 1
				errCatName =	rstCompound.Fields(0)
				response.write "<br>" & errCatName 
			end if
        end if
        Set rstCompound = rstCompound.NextRecordset
    Loop
   
    ' clean up
    Cnxn.Close
    Set rstCompound = Nothing
    Set Cnxn = Nothing

	response.write "<br><hr><hr>"
	for i = 1 to catCount
		response.write errCatNames(i) & ": " & chr(9) & errCatCounts(i) & "<br>"
	Next 

	endTime = timer
	response.write "<hr><hr>End <br>"& errCount & " Errors Found in "& catCount & " Categories (Out of "& allCats & " Cat.s)<br>Operation took " & endTime - startTime & " secconds."
%>
</BODY>
</HTML>
