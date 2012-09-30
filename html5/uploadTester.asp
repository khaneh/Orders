<%@ Language=VBScript %>
<% 
option explicit 
Response.Expires = -1
Server.ScriptTimeout = 600
' All communication must be in UTF-8, including the response back from the request
Session.CodePage  = 65001
%>
<!-- #include file="freeaspupload.asp" -->
<%


  ' ****************************************************
  ' Change the value of the variable below to the pathname
  ' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
  ' ****************************************************

  Dim uploadsDirVar
  uploadsDirVar = "C:\inetpub\jametest\html5\uploads" 
  

  ' Note: this file uploadTester.asp is just an example to demonstrate
  ' the capabilities of the freeASPUpload.asp class. There are no plans
  ' to add any new features to uploadTester.asp itself. Feel free to add
  ' your own code. If you are building a content management system, you
  ' may also want to consider this script: http://www.webfilebrowser.com/

function OutputForm()
%>
    <form name="frmSend" method="POST" enctype="multipart/form-data" accept-charset="utf-8" action="uploadTester.asp" onSubmit="return onSubmitForm();">
	<B>File names:</B><br>
    File 1: <input name="attach1" type="file" size=35 multiple><br>
    <br> 
	<!-- These input elements are obviously optional and just included here for demonstration purposes -->
	<B>Additional fields (demo):</B><br>
	Enter a number: <input type="text" name="enter_a_number"><br>
    Checkbox values: <input type="checkbox" value="1" name="checkbox_values"> 1 &nbsp;&nbsp;<input type="checkbox" value="2" name="checkbox_values"> 2<br>
    Drop-down list (with multiple selection): <br>	   
    <select name="list_values" class="TextBox" MULTIPLE>
        <option value='frist' > First</option>
        <option value='second' > Second</option>
        <option value='third' > Third</option>
    </select><br>
    <textarea rows="2" cols="20" name="t_area">Test text area</textarea><br>
	<!-- End of additional elements -->
    <input style="margin-top:4" type=submit value="Upload">
    </form>
<%
end function

function TestEnvironment()
    Dim fso, fileName, testFile, streamTest
    TestEnvironment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    if not fso.FolderExists(uploadsDirVar) then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not exist.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    fileName = uploadsDirVar & "\test.txt"
    on error resume next
    Set testFile = fso.CreateTextFile(fileName, true)
    If Err.Number<>0 then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have write permissions.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    Err.Clear
    testFile.Close
    fso.DeleteFile(fileName)
    If Err.Number<>0 then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have delete permissions</B>, although it does have write permissions.<br>Change the permissions for IUSR_<I>computername</I> on this folder."
        exit function
    end if
    Err.Clear
    Set streamTest = Server.CreateObject("ADODB.Stream")
    If Err.Number<>0 then
        TestEnvironment = "<B>The ADODB object <I>Stream</I> is not available in your server.</B><br>Check the Requirements page for information about upgrading your ADODB libraries."
        exit function
    end if
    Set streamTest = Nothing
end function

function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey

    Set Upload = New FreeASPUpload
    Upload.Save(uploadsDirVar)

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then Exit function

    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) then
        SaveFiles = "<B>Files uploaded:</B> "
        for each fileKey in Upload.UploadedFiles.keys
            SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
        next
    else
        SaveFiles = "No file selected for upload or the file name specified in the upload form does not correspond to a valid file in the system."
    end if
	SaveFiles = SaveFiles & "<br>Enter a number = " & Upload.Form("enter_a_number") & "<br>"
	SaveFiles = SaveFiles & "Checkbox values = " & Upload.Form("checkbox_values") & "<br>"
	SaveFiles = SaveFiles & "List values = " & Upload.Form("list_values") & "<br>"
	SaveFiles = SaveFiles & "Text area = " & Upload.Form("t_area") & "<br>"
end function
%>

<HTML>
<HEAD>
<TITLE>Test Free ASP Upload 2.0</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
BODY {background-color: white;font-family:arial; font-size:12}
</style>
<script>
function onSubmitForm() {
    var formDOMObj = document.frmSend;
    if (formDOMObj.attach1.value == "" && formDOMObj.attach2.value == "" && formDOMObj.attach3.value == "" && formDOMObj.attach4.value == "" )
        alert("Please press the Browse button and pick a file.")
    else
        return true;
    return false;
}
</script>

</HEAD>

<BODY>

<br><br>
<div style="border-bottom: #A91905 2px solid;font-size:16">Upload files to your server</div>
<%
Dim diagnostics
if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
    diagnostics = TestEnvironment()
    if diagnostics<>"" then
        response.write "<div style=""margin-left:20; margin-top:30; margin-right:30; margin-bottom:30;"">"
        response.write diagnostics
        response.write "<p>After you correct this problem, reload the page."
        response.write "</div>"
    else
        response.write "<div style=""margin-left:150"">"
        OutputForm()
        response.write "</div>"
    end if
else
    response.write "<div style=""margin-left:150"">"
    OutputForm()
    response.write SaveFiles()
    response.write "<br><br></div>"
end if

%>



</BODY>
</HTML>
