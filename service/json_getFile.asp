<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
orderFolder = "\\big\orders"
select case request("act")
	case "list":
		set j = jsArray()
		dim orderID
		orderID = CDbl(Request("orderID"))
		orderFolder = orderFolder & "\" & orderID
		ListFolderContents(orderFolder)
		
end select
Response.Write toJSON(j)
set fs=nothing


sub ListFolderContents(path)
	dim fs
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if Not fs.FolderExists(path) then
		fs.CreateFolder (path) 
	end if
	'set fs = CreateObject("Scripting.FileSystemObject")
	set folder = fs.GetFolder(path)
	
' 	j(null)("folderName") = folder.Name
' 	j(null)("folderCount") = folder.Files.Count	
' 	j(null)("subFolderCount") = folder.SubFolders.Count	
' 	j(null)("folderSize") = Round(folder.Size / 1024)
	
	if folder.SubFolders.Count > 0 then 
	 	for each item in folder.SubFolders
	 		ListFolderContents(item.Path)
	 	next
	 end if
	'Display a list of files.
	
	for each item in folder.Files
		set j(null) = jsObject()
		j(null)("itemPath") = Right(item.path,len(item.path) - len(orderFolder))
		j(null)("realPath") = "/beta/order/file/" & orderID & Right(item.path,len(item.path) - len(orderFolder))
		j(null)("itemName") = item.Name
		j(null)("itemSize") = item.Size
		j(null)("itemDateLastModified") = item.DateLastModified
	next

end sub  
%>