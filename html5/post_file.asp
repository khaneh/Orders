<%

'Server.ScriptTimeout = 6000

'caching
'response.addheader "Cache-Control", "no-store, no-cache, must-revalidate" 
'response.addheader "Cache-Control", "post-check=0, pre-check=0"
'response.addheader "Pragma", "no-cache"
response.addheader "Access-Control-Allow-Origin", "*"
response.addheader "Access-Control-Allow-Methods", "OPTIONS, HEAD, GET, POST, PUT, DELETE"
response.addheader "Access-Control-Allow-Headers", "X-File-Name, X-File-Type, X-File-Size"
%>
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<!-- #include virtual="/html5/freeASPUpload.asp" --> 
<%


' Response.write "orderID: " & Request("orderID")
' Response.end
set st = jsObject()

If Request.TotalBytes > 0 Then
    Dim ByteReceived
        ByteReceived = Request.TotalBytes
    dim dataReceived
    	dataReceived = Request.BinaryRead(ByteReceived)
    dim FileBinary
    dim signatureFile
    dim arrayContent
    dim orderID
    
    For i = 1 To lenB(dataReceived)
		FileBinary = FileBinary & chr(ascB(midB(dataReceived,i,1)))
	Next
	signatureFile = left(FileBinary,instr(FileBinary,"" & vbCrLf)-1)
	arrayContent  = split(FileBinary,signatureFile)
	dim i
		i=0
	dim start
		start=0
	do until i>= UBound(arrayContent) or start>0
		if instr (arrayContent(i),"" & vbCrLf & "Content-Disposition: form-data; name=""orderID""" & vbCrLf & "")>0 then 
			start = instr(arrayContent(i), "" & vbCrLf & "" & vbCrLf) + 4
			endPos = instr(start, arrayContent(i), "" & vbCrLf) 
			orderID = mid(arrayContent(i),start,endPos-start)
		end if
		i=i+1
	Loop

End If
st("orderID")=orderID

'st("orderID")=orderID

path = "C:\inetpub\jametest\html5\uploads"

path = path & "\" & orderID
'st("path")=path
' dim fs
' set fs=Server.CreateObject("Scripting.FileSystemObject")
' if Not fs.FolderExists(path) then
' 	fs.CreateFolder (path) 
' end if
' set fs=nothing

Set theForm = Server.CreateObject("ABCUpload4.XForm")
theForm.Overwrite = True
If theForm.Files("pic").FileExists Then
	set theFile = theForm.Files("pic")
' 	theFile.macBinary = true
'    theFile.Save path & "\" & theFile.SafeFileName
   theFile.Save theFile.SafeFileName
'    Response.write theFile.safeFile
End If

' Set Upload = New FreeASPUpload
' Upload.Save(path)
set Upload=Nothing
st("path")=path
'Response.end




st("status") = "File was uploaded successfuly!"
Response.Write toJSON(st)

%>

<%
	' ByteReceived = Request.TotalBytes
' dataReceived = Request.BinaryRead(ByteReceived)
' ' response.write length(Request.files)
' ' Response.end
' 
' For i = 1 To lenB(dataReceived)
'   FileBinary = FileBinary & chr(ascB(midB(dataReceived,i,1)))
' Next
' 
' signatureFile = left(FileBinary,instr(FileBinary,"" & vbCrLf)-1)
' arrayPieces  = split(FileBinary,signatureFile)
' 
' 
' for piece = 1 to ubound(arrayPieces)-1
'       start = instr(arrayPieces(piece),"" & vbCrLf & "" & vbCrLf)
'       heading = left(arrayPieces(piece),start-1)
'       start = start + len("" & vbCrLf) + len("" & vbCrLf)
'       contentFile = mid(arrayPieces(piece),start,len(arrayPieces(piece))-start-1)
' next
' 
' 'stop
' 'st("heading") = heading
' 
' if instr(heading,"file") > 0 then
' 
'     i = InStr(heading, "filename=")
' 	j = InStr(i + 10, heading, chr(34))
' 	
' 	uploadName = Mid(heading, i + 10, j - i - 10)
' 	i = InStrRev(uploadName,"\")
' 	
' 	if i <> 0 then
' 	fileName = mid(uploadName, i + 1)
' 	else    
' 	fileName = uploadName
' 	end if
'             
' 	if fileName <> "" then
' 	   dim FSO
' 	Set FSO = CreateObject("Scripting.FileSystemObject")
' 	Upload1 = True
' 	DimensioneFile1 = len(contentFile)
' 	EstensioneFile1 = right(contentFile,3)
' 	fileName1 = "uploads\" & fileName            'salvo i file nella directory uploads
' 	        
'  	if left(contentFile,3) = "ÿØÿ" then '"ÿØÿà"  then        'PRIMO SpieceNE DEL FILE .JPG
' 	        
' 	    Set textStream = FSO.CreateTextFile(server.mappath(fileName1), True, False)
' 	else        
' 	    Set textStream = FSO.OpenTextFile(server.mappath(fileName1), 8, True)                
' 	end if
' 	
' 	textStream.Write contentFile
' 	
' 	textStream.Close
' 	Set textStream = Nothing
' 	Set FSO = Nothing
' 	end if
' end if 
%>