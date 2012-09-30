<%
dim vntPostedData,strTargetDir,filename,objFSO

set objFSO =  server.CreateObject ("Scripting.FileSystemObject")
    strTargetDir = server.MapPath("uploads")&"\"
    filename = Request.QueryString("name")
    chunk = Request.QueryString("chunk")
    chunks = Request.QueryString("chunks")

        
        if chunk = 0 then 'test to see if this chunk  is the start of a file
        
            objFSO.CreateTextFile strTargetDir & filename,1

        end if

'save contents of binary read to a text file        
set file = objFSO.OpenTextFile (strTargetDir & filename,8,false,true)
        
        vntPostedData = Request.BinaryRead(Request.TotalBytes)

        file.Write vntPostedData
        file.close
        

            
        if cint(chunk) + 1 = cint(chunks) then 'test to see if this chunk is the end of a file 
            
'save transferd text file to a binary file then remove the BOM from the start of file and save again 

            set objFSO =  server.CreateObject ("Scripting.FileSystemObject")
            Set fromStream = CreateObject("ADODB.Stream")
            set toStream = CreateObject("ADODB.Stream")
            toStream.Charset = "utf-8"
            fromStream.Charset = "utf-8"
            fromStream.Type = 2 'text
            toStream.Type = 1   'binary
            'load text file to a stream
            fromStream.open
            fromStream.Loadfromfile(strTargetDir & filename)
            toStream.Open
           
            'copy text file stream to a binary file stream and save
            fromStream.copyto (toStream)
            fromStream.flush
            toStream.savetofile strTargetDir &"1" & filename,2
            toStream.flush
            fromStream.close

            'open the previous binary stream and set the stream position to 
            '5 which is after the BOM ( byte order mark ) and resave the stream 
            Set fromStream = CreateObject("ADODB.Stream")
            fromStream.Charset = "utf-8"
            fromStream.Type = 1 'binary
            fromstream.open
            fromStream.Loadfromfile(strTargetDir & "1" & filename)
            fromStream.position = 5
            fromstream.copyto (tostream)
            toStream.savetofile strTargetDir & filename,2
            toStream.close
            fromStream.close
            objFSO.deletefile(strTargetDir & "1" & filename)'delete tempfile

            end if 
            
     
%>






