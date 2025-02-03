<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"-->
<!--#include file="../includes/clsUpload.asp"--><% 
Dim debug, i, j, Upload, DocumentRoot, FunctionalAreaID, UploadDirectory, ArchivesDirectory, fso, _
	filename, archivesfilename, extension, uploadfileName, file, _
	GrantID, AppID, MAGID, ISAID, AdjustmentID, MonitorID, Quarter, DirectoryID, Location, Prefix, Identifier, _
	NameToUse, DocumentTypeID, timestamp, failed
debug = False
DocumentRoot = Application("DocumentRoot")

If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLF)
	For each i in Request.Form
		Response.Write("Request.Form(""" & i & """)='" & Request.Form(i) & "'" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("Session(""" & i & """)='" & Session(i) & "'" & vbCrLf)
	Next
	for each i in Request.Cookies
		if Request.Cookies(i).HasKeys then
			for each j in Request.Cookies()
				response.write("Cookies(" & i & ":" & j & ")=" & Request.Cookies(i)(j))
			next
		else
			Response.Write("Cookies(""" & i & """)=" & Request.Cookies(i) & "<br>")
		end if
	next
	Response.Write("DocumentRoot=" & DocumentRoot)
	Response.Write("</pre>" & vbCrLF)
End If

Set Upload = New clsUpload
GrantID = Upload("GrantID").Value
AppID = Upload("AppID").Value
MAGID = Upload("MAGID").Value
AdjustmentID = Upload("AdjustmentID").Value
ISAID = Upload("ISAID").Value
MonitorID = Upload("MonitorID").Value
Quarter = Upload("Quarter").Value
FunctionalAreaID = Upload("FunctionalAreaID").Value
DocumentTypeID= Upload("DocumentTypeID").Value
uploadfileName = Upload("File1").FileName
extension = LCase(Mid(uploadfilename, InStrRev(uploadfilename, ".")+1))

If Len(FunctionalAreaID)>0 Then
	FunctionalAreaID = CInt(FunctionalAreaID)
Else
	FunctionalAreaID = 0
End If
If Len(DocumentTypeID)>0 Then
	DocumentTypeID = CInt(DocumentTypeID)
Else
	DocumentTypeID = 0
End If
If Len(GrantID)>0 Then
	GrantID = CInt(GrantID)
Else
	GrantID = 0
End If
If Len(AppID)>0 Then
	AppID = CInt(AppID)
Else
	AppID = 0
End If
If Len(AdjustmentID)>0 Then
	AdjustmentID = CInt(AdjustmentID)
Else
	AdjustmentID = 0
End If
If Len(ISAID)>0 Then
	ISAID = CInt(ISAID)
Else
	ISAID = 0
End If
If Len(MonitorID)>0 Then
	MonitorID = CInt(MonitorID)
Else
	MonitorID = 0
End If

If Debug = True Then
	Response.Write("<pre>")
	Response.Write("ISAID=" & ISAID & vbCrLf)
	Response.Write("AppID=" & AppID & vbCrLf)
	Response.Write("GrantID=" & GrantID & vbCrLf)
	Response.Write("AdjustmentID=" & AdjustmentID & vbCrLf)
	Response.Write("MonitorID=" & MonitorID & vbCrLf)
	Response.Write("Quarter=" & Quarter & vbCrLf)
	Response.Write("DocumentTypeID=" & DocumentTypeID & vbCrLf)
	Response.Write("Extension=" & Extension & vbCrLf)
	Response.Write("uploadfileName=" & uploadfileName & vbCrLf)
	Response.Write("fileName=" & fileName & vbCrLf)
	Response.Write("</pre>")
End If
Select Case extension
Case "pdf", "docx", "xlsx", "pptx", "doc", "xls", "ppt", "tif"
	failed = false
Case Else
	failed = true
End Select
If failed=True Then
	Response.Write("Error: Only the following file types are accepted: pdf, docx, xlsx, pptx, doc, xls, ppt.")
	SendMessage "Error: Only the following file types are accepted: pdf, docx, xlsx, pptx, doc, xls, ppt."
	Response.End
End If

sql = "SELECT A.FunctionalAreaID, B.DocumentTypeID, " & vbCrLf & _
	"	A.DirectoryID, A.Location, A.Prefix, A.Identifier, " & vbCrLF & _
	"	B.DocumentTypeDescription, B.NameToUse, B.EditableName, B.MVCPAOnly " & vbCrLf & _
	"FROM Lookup.FunctionalAreas AS A " & vbCrLf & _
	"JOIN Lookup.DocumentTypes AS B ON B.FunctionalAreaID=A.FunctionalAreaID " & vbCrLf & _
	"WHERE A.FunctionalAreaID=" & prepIntegerSQL(FunctionalAreaID) & " AND B.DocumentTypeID=" & prepIntegerSQL(DocumentTypeID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(SQL)
If rs.EOF = True Then
	Response.Write("Error: Documents Types " & DocumentTypeID & " Not Available for FunctionalAreaID=" & FunctionalAreaID)
	SendMessage "Error: Documents Types " & DocumentTypeID & " Not Available for FunctionalAreaID=" & FunctionalAreaID
	Response.End
Else
	FunctionalAreaID = rs.Fields("FunctionalAreaID")
	DirectoryID = rs.Fields("DirectoryID")
	Location = rs.Fields("Location")
	Prefix = rs.Fields("Prefix")
	Identifier = rs.Fields("Identifier")
	NameToUse = rs.Fields("NameToUse")
End If

If DirectoryID = "ISAID" Then
	UploadDirectory = DocumentRoot & Location & ISAID & "\"
ElseIf DirectoryID = "AppID" Then
	UploadDirectory = DocumentRoot & Location & AppID & "\"
ElseIf DirectoryID = "GrantID" Then
	UploadDirectory = DocumentRoot & Location & GrantID & "\"
ElseIf DirectoryID = "MonitorID" Then
	UploadDirectory = DocumentRoot & Location & MonitorID & "\"
ElseIf DirectoryID = "MAGID" Then
	UploadDirectory = DocumentRoot & Location & MAGID & "\"
End If
ArchivesDirectory = UploadDirectory & "Archives\"

If GrantID=0 And AppID=0 And ISAID=0 And MonitorID=0 And Len(MAGID)=0 Then
	Response.Write("A GrantID, AppID, MAGID, ISAID, or MonitorID must be provided!")
	SendMessage "A GrantID, AppID, MAGID, ISAID, or MonitorID must be provided!"
	Response.End
End If

If IsNull(Prefix) = False Then
	If Identifier = "{AdjustmentID}" Then
		filename = Prefix & Right("00000" & AdjustmentID,5) & " " & NameToUse & "." & extension
	ElseIf Identifier = "{Quarter}" Then
		filename = Prefix & CStr(Quarter) & " " & NameToUse & "." & extension
	Else
		filename = Prefix & " " & NameToUse & "." & extension
	End If
Else
	filename = NameToUse & "." & extension
End If

If Debug = True Then
	Response.Write("<pre>")
	Response.Write("UploadDirectory=""" & UploadDirectory & """" & vbCrLf)
	Response.Write("ArchivesDirectory=""" & ArchivesDirectory & """" & vbCrLf)
	Response.Write("filename=""" & filename & """" & vbCrLf)
	Response.Write("</pre>")
End If

set fso = Server.CreateObject("Scripting.FileSystemOBject")
If fso.FolderExists(UploadDirectory) Then
	If Debug = True Then
		Response.Write("<pre>UploadDirectory,""" & UploadDirectory & """, exists.</pre>")
	End If	
Else
	If Debug = True Then
		Response.Write("<pre>UploadDirectory, """ & UploadDirectory & """, does not exist.</pre>")
		fso.CreateFolder(UploadDirectory)
		If fso.FolderExists(UploadDirectory) Then
			If Debug = True Then
				Response.Write("<pre>UploadDirectory, """ & UploadDirectory & """, has been created.</pre>")
			End If
		Else
			Response.Write("Error: Unable to Create Directory")
			Response.End
		End If
	End If	
End If
If fso.FolderExists(ArchivesDirectory) Then
	If Debug = True Then
		Response.Write("<pre>UploadDirectory Archives,""" & ArchivesDirectory & """, exists.</pre>")
	End If	
Else
	If Debug = True Then
		Response.Write("<pre>UploadDirectory Archives, """ & ArchivesDirectory & """, does not exist.</pre>")
	End If	
	fso.CreateFolder(ArchivesDirectory)
	If fso.FolderExists(ArchivesDirectory) Then
		If Debug = True Then
			Response.Write("<pre>UploadDirectory Archives, """ & ArchivesDirectory & """, has been created.</pre>")
		End If
	Else
		Response.Write("Error: Unable to Create Archive Directory")
		Response.End
	End If
End If

If fso.FileExists(UploadDirectory & filename) Then
	Set file = fso.GetFile(UploadDirectory & filename)
	timestamp = file.DateLastModified
	archivesfilename = mid(filename, 1, InStrRev(filename, ".")-1) & " (" & year(timestamp) & right("00" & month(timestamp),2) & right("00" & day(timestamp),2) & ")." & extension
	If fso.FileExists(ArchivesDirectory & archivesfilename) Then
		If Debug = True Then
			Response.Write("<pre>Deleting exisiting " & ArchivesDirectory & archivesfilename & "</pre>" & vbCrLF)
		End If
		fso.DeleteFile(ArchivesDirectory & archivesfilename)
	End If
	fso.MoveFile UploadDirectory & filename, ArchivesDirectory & archivesfilename
	If Debug = True Then
		Response.Write("<pre>fso.MoveFile " & UploadDirectory & filename & ", " & ArchivesDirectory & archivesfilename & "</pre>" & vbCrLF)
	End If
End If

Upload("File1").SaveAs UploadDirectory & filename
Set Upload = Nothing

If Debug = True Then
	If FunctionalAreaID=2 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&AppID=" & AppID & """>Upload.asp?fid=" & FunctionalAreaID & "&AppID=" & AppID & "</a>")
	ElseIf FunctionalAreaID=3 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & """>Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "</a>")
	ElseIf FunctionalAreaID=1 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&ISAID=" & ISAID & """>Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "</a>")
	ElseIf FunctionalAreaID=4 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&AdjustmentID=" & AdjustmentID & """>Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&Quarter=" & Quarter & "</a>")
	ElseIf FunctionalAreaID=5 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&Quarter=" & Quarter & """>Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&Quarter=" & Quarter & "</a>")
	ElseIf FunctionalAreaID=6 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&Quarter=" & Quarter & """>Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&Quarter=" & Quarter & "</a>")
	ElseIf FunctionalAreaID=13 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&AppID=" & AppID & """>Upload.asp?fid=" & FunctionalAreaID & "&AppID=" & AppID & "</a>")
	ElseIf FunctionalAreaID=14 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&MAGID=" & MAGID & """>Upload.asp?fid=" & FunctionalAreaID & "&MAGID=" & MAGID & "</a>")
	ElseIf FunctionalAreaID=15 Then
		Response.Write("<a href=""Upload.asp?fid=" & FunctionalAreaID & "&MAGID=" & MAGID & """>Upload.asp?fid=" & FunctionalAreaID & "&MAGID=" & MAGID & "</a>")
	End If
	Response.End
End If
If FunctionalAreaID=2 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&AppID=" & AppID)
ElseIf FunctionalAreaID=3 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID)
ElseIf FunctionalAreaID=1 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&ISAID=" & ISAID)
ElseIf FunctionalAreaID=4 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&AdjustmentID=" & AdjustmentID)
ElseIf FunctionalAreaID=5 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&Quarter=" & Quarter)
ElseIf FunctionalAreaID=6 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID & "&Quarter=" & Quarter)
ElseIf FunctionalAreaID=11 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&GrantID=" & GrantID)
ElseIf FunctionalAreaID=12 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&MonitorID=" & MonitorID)
ElseIf FunctionalAreaID=13 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&AppID=" & MonitorID)
ElseIf FunctionalAreaID=14 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&MAGID=" & MAGID)
ElseIf FunctionalAreaID=15 Then
	Response.Redirect("Upload.asp?fid=" & FunctionalAreaID & "&MAGID=" & MAGID)
Else
	Response.Write("Error: Unable to redirect.")
	SendMessage "Error: Unable to redirect."
	Response.End
End If

%><!--#include file="../includes/prepDB.asp"-->