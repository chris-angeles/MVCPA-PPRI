<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Uploaded Files</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" />
</head>
<body>
<table styles="margin: auto; ">
<thead>
<tr>
	<th colspan="4"><%=Request.ServerVariables("SERVER_NAME")%> Uploaded Files</th>
</tr>
<tr>
	<th>File Name</th><th>File Date</th><th>File Size</th><th>link</th>
</tr>
</thead>
<tbody><%
Dim Debug, DocumentRoot, fso, files, file, DocumentRootLength
Debug = False

DocumentRoot = Application("DocumentRoot")
DocumentRootLength = Len(DocumentRoot)
If Debug = True Then
	Response.Write(DocumentRoot & "<br />" & vbCrLf)
End If

set fso = Server.CreateObject("Scripting.FileSystemOBject")
'Response.Write("<table styles=""margin: auto; "">" & vbCrLf)
ShowSubfolders fso.GetFolder(DocumentRoot), 3 
'Response.Write("</table>" & vbCrLf)

Sub ShowSubFolders(Folder, Depth)
	Dim Subfolder
	Set files = Folder.Files
	If files.count>0 Then 
		For Each file in files
			Response.Write("<tr><td>" & file.name & "</td><td>" & file.DateLastModified & "</td><td>" & _
				file.size & "</td><td><a href=""..\Documents" & Mid(file.path, DocumentRootLength) & """ target=""_blank"">link</a></td></tr>" & vbCrLf)
		Next
	End If

    If Depth > 0 then
        For Each Subfolder in Folder.SubFolders
			If Subfolder.Files.count > 0 Then
				Response.Write("<tr><td colspan=""4"" style=""background-color: PowderBlue; "">Folder: \Documents" & Mid(Subfolder.Path, DocumentRootLength) & "</td></tr>" & vbCrLf)
			End If
            ShowSubFolders Subfolder, Depth -1 
        Next
    End if
End Sub
%>
</tbody>
</table>
</body>
</html>
<!--#include file="../includes/prepDB.asp"-->