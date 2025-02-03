<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Uploaded Document List</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style>
	tr, td, th {padding: 5px;}
</style>
</head>
<body style="width: 100%">
<% 
Dim debug, i, GranteeID, FiscalYear, GranteeName, AppID, ApplicationName, GrantID, MagID, ProgramName, _
	ShowArchives, ShowDeleted,  _
	DocumentRoot, folder, LastAppID, fso, rsFSO 
', folders, files, file, LastGranteeID, LastGrantID
Debug = False

If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("<pre>Session(""" & i & """)='" & Session(i) & "'</pre>" & vbCrLf)
	Next
End If

If Len(Request.Form("GranteeID"))>0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID"))>0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
ElseIf Len(Session("GranteeID")) > 0 Then
	GranteeID = CInt(Session("GranteeID"))
Else
	GranteeID = 0
End If
If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
ElseIf Len(Session("FiscalYear")) > 0 Then
	FiscalYear = CInt(Session("FiscalYear"))
Else
	FiscalYear = 0
End If
If Request.Form("ShowArchives") = "1" Then
	ShowArchives = True
ElseIf Request.QueryString("ShowArchives") = "1" Then
	ShowArchives = True
Else
	ShowArchives = False
End If
If Request.Form("ShowDeleted") = "1" Then
	ShowDeleted = True
ElseIf Request.QueryString("ShowDeleted") = "1" Then
	ShowDeleted = True
Else
	ShowDeleted = False
End If

sql = "SELECT GranteeID, GranteeName " & vbCrLF & _
	"FROM Grantees AS A" & vbCrLF & _
	"ORDER BY REPLACE(A.GranteeName,'City Of ','') "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
%><form name="Selection" method="post">
Grantee: <select name="GranteeID" onchange="document.Selection.submit();">
	<option value="0">Select Grantee</option>
<%
	While rs.EOF = False
		If rs.Fields("GranteeID") = GranteeID Then
			Response.Write(vbTab & "<option value=""" & rs.Fields("GranteeID") & """ SELECTED>" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
		Else
			Response.Write(vbTab & "<option value=""" & rs.Fields("GranteeID") & """>" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
		End If
		rs.MoveNext()
	Wend
%>
  </select>&nbsp;&nbsp;&nbsp;
  Fiscal Year: <select name="FiscalYear" onchange="document.Selection.submit();">
	<option value="0">Select</option>
<%
	For i = 2017 to Application("CurrentFiscalYear")+1
		If i = FiscalYear Then
			Response.Write(vbTab & "<option value=""" & i & """ SELECTED>FY" & i & "</option>" & vbCrLf)
		Else
			Response.Write(vbTab & "<option value=""" & i & """>FY" & i & "</option>" & vbCrLf)
		End If
	Next
%>
     </select>&nbsp;&nbsp;&nbsp;
	 <input type="checkbox" name="ShowArchives" value="1" <% 
		If ShowArchives=True Then 
			Response.Write(" CHECKED")
		End If 
		%> onchange="document.Selection.submit();" />Show Archived Files&nbsp;&nbsp;&nbsp;
	 <input type="checkbox" name="ShowDeleted" value="1" <% 
		If ShowDeleted=True Then 
			Response.Write(" CHECKED")
		End If 
		%> onchange="document.Selection.submit();" />Show Deleted Files
</form>
<h1>Documents Uploaded to Website</h1>
<%
If GranteeID>0 and FiscalYear>0 Then
	DocumentRoot = Application("DocumentRoot")
	Set fso = Server.CreateObject("Scripting.FileSystemObject")

	sql = "SELECT A.GranteeID, GranteeName, COALESCE(B.FiscalYear, C.FiscalYear, " & prepIntegerSQL(FiscalYear) & ") AS FiscalYear, " & vbCrLf & _
		"	B.AppID, M.ProgramName AS ApplicationName, GrantID, D.MagID, ISNULL(C.ProgramName, 'Not Funded') AS ProgramName " & vbCrLf & _
		"FROM Grantees AS A " & vbCrLf & _
		"LEFT JOIN Application.IDs AS B ON B.GranteeID=A.GranteeID AND B.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"LEFT JOIN Application.Main AS M on M.AppID=B.AppID " & vbCrLf & _
		"LEFT JOIN [Grants].Main AS C ON C.GranteeID=A.GranteeID AND C.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"LEFT JOIN MAG.Main AS D ON D.GranteeID=A.GranteeID AND D.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"WHERE A.GranteeID=" & prepIntegerSQL(GranteeID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.Eof = False Then
		GranteeName = rs.Fields("GranteeName")
		GranteeID = rs.Fields("GranteeID")
	End If
	Response.Write("<h2>" & GranteeName & "</h2>" & vbCrLf)
	Response.Write("<h2>FY" & FiscalYear & "</h2>" & vbCrLf)
	Response.Write("<table style=""margin: auto; "">" & vbCrLf)
	While rs.EOF = False
		AppID = rs.Fields("AppID")
		GrantID = rs.Fields("GrantID")
		MAGID = rs.Fields("MAGID")
		ApplicationName = rs.Fields("ApplicationName")
		ProgramName = rs.Fields("ProgramName")

		ApplicationDocuments AppID, FiscalYear, ApplicationName, ShowDeleted

		GrantDocuments GrantID, FiscalYear, ProgramName, ShowDeleted

		MAGDocuments MAGID, FiscalYear, ShowDeleted

		' Monitoring Documents
		MonitoringDocuments GranteeID, GranteeName, FiscalYear, ShowDeleted

		rs.MoveNext

	Wend

	Response.Write("</table>" & vbCrLf)
End If
%>
</body>
</html>

<%
Function ApplicationDocuments (vAppID, vFiscalYear, vApplicationName, vShowDeleted)
	If ISNull(vAppID) = True Then
	Response.Write("<tr><th colspan=""3"">There is no application for " & vFiscalYear & ".</th></tr>" & vbCrLf)
		Exit Function
	End If
	' Application
	folder = DocumentRoot & "Application\" & vAppID
	If Debug = True Then
		Response.Write(folder & "<br />")
	End If
	Response.Write("<tr><th colspan=""3"">Application: " & vApplicationName & ", " & vFiscalYear & " (AppID=" & vAppID &  ")</th></tr>" & vbCrLf)
	If fso.FolderExists(folder) Then
		Set rsFSO = kc_fsoFiles(folder, "_")

		While Not rsFSO.EOF
			Response.Write("<tr><td><a href=""../Documents/Application/" & vAppID & "/" & rsFSO("Name") & """ target=""_blank"">" & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>" & vbCrLf)
			rsFso.MoveNext()
		Wend
  
		'finally, close out the recordset
		rsFSO.close()
		Set rsFSO = Nothing

		' Archives
		If ShowArchives = True Then
			folder = DocumentRoot & "Application\" & vAppID & "\Archives\"
			If Debug = True Then
				Response.Write(folder & "<br />")
			End If
			If fso.FolderExists(folder) Then
				Set rsFSO = kc_fsoFiles(folder, "_")

				While Not rsFSO.EOF
					Response.Write("<tr><td><a href=""../Documents/Application/" & vAppID & "/Archives/" & rsFSO("Name") & """ target=""_blank"">(Archived) " & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>" & vbCrLf)
					rsFso.MoveNext()
				Wend
  
				'finally, close out the recordset
				rsFSO.close()
				Set rsFSO = Nothing
			End If
		End If

		' Deleted
		If vShowDeleted = True Then
			folder = DocumentRoot & "Application\" & vAppID & "\_delete\"
			If Debug = True Then
				Response.Write(folder & "<br />")
			End If
			If fso.FolderExists(folder) Then
				Set rsFSO = kc_fsoFiles(folder, "_")

				While Not rsFSO.EOF
					Response.Write("<tr><td><a href=""../Documents/Application/" & vAppID & "/_delete/" & rsFSO("Name") & """ target=""_blank"">(Deleted) " & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>" & vbCrLf)
					rsFso.MoveNext()
				Wend
  
				'finally, close out the recordset
				rsFSO.close()
				Set rsFSO = Nothing
			End If
		End If

	Else
		Response.Write("<tr><td colspan=""3"" style=""text-align: center"">No Documents Found</td></tr>" & vbCrLf)
	End If
End Function

Function GrantDocuments (vGrantID, vFiscalYear, vProgramName, vShowDeleted)
		' Related grant documents or not funded message.
		folder = DocumentRoot & "Grant\" & vGrantID
		If Debug = True Then
			Response.Write(folder & "<br />")
		End If
		If IsNull(GrantID) Then
			Response.Write("<tr><th colspan=""3"">There is no taskforce grant for " & vFiscalYear & ".</th></tr>" & vbCrLf)
			Exit Function
		Else
			Response.Write("<tr><th colspan=""3"">Grant: " & vProgramName & ", " & vFiscalYear & " (GrantID=" & vGrantID &  ")</th></tr>" & vbCrLf)
			If fso.FolderExists(folder) Then
				Set rsFSO = kc_fsoFiles(folder, "_")

				If rsFSO.EOF = True Then
					Response.Write("<tr><td colspan=""3"" style=""text-align: center"">No Documents Found</td></tr>" & vbCrLf)
				Else
					While Not rsFSO.EOF
						Response.Write("<tr><td><a href=""../Documents/Grant/" & GrantID & "/" & rsFSO("Name") & """ target=""_blank"">" & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
						rsFso.MoveNext()
					Wend
				End If
  
				'finally, close out the recordset
				rsFSO.close()
				Set rsFSO = Nothing

				' Archives
				If ShowArchives = True Then
					folder = DocumentRoot & "Grant\" & vGrantID & "\Archives\"
					If Debug = True Then
						Response.Write(folder & "<br />")
					End If
					If fso.FolderExists(folder) Then
						Set rsFSO = kc_fsoFiles(folder, "_")

						While Not rsFSO.EOF
							Response.Write("<tr><td><a href=""../Documents/Grant/" & vGrantID & "/Archives/" & rsFSO("Name") & """ target=""_blank"">(Archives) " & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
							rsFso.MoveNext()
						Wend
  
						'finally, close out the recordset
						rsFSO.close()
						Set rsFSO = Nothing
					End If
				End If

				' Deleted
				If vShowDeleted = True Then
					folder = DocumentRoot & "Grant\" & vGrantID & "\_delete\"
					If Debug = True Then
						Response.Write(folder & "<br />")
					End If
					If fso.FolderExists(folder) Then
						Set rsFSO = kc_fsoFiles(folder, "_")

						While Not rsFSO.EOF
							Response.Write("<tr><td><a href=""../Documents/Grant/" & vGrantID & "/_delete/" & rsFSO("Name") & """ target=""_blank"">(Deleted) " & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
							rsFso.MoveNext()
						Wend
  
						'finally, close out the recordset
						rsFSO.close()
						Set rsFSO = Nothing
					End If
				End If

			End If
		End If
End Function

Function MAGDocuments (vMAGID, vFiscalYear, vShowDeleted)
		' Related grant documents or not funded message.
		folder = DocumentRoot & "MAG\" & vMAGID
		If Debug = True Then
			Response.Write(folder & "<br />")
		End If
		If IsNull(MAGID) Then
			Response.Write("<tr><th colspan=""3"">There is no auxilary grant for " & vFiscalYear & ".</th></tr>" & vbCrLf)
			Exit Function
		Else
			Response.Write("<tr><th colspan=""3"">Grant: Auxiliary Grant, " & vFiscalYear & " (MAGID=" & vMAGID &  ")</th></tr>" & vbCrLf)
			If fso.FolderExists(folder) Then
				Set rsFSO = kc_fsoFiles(folder, "_")

				If rsFSO.EOF = True Then
					Response.Write("<tr><td colspan=""3"" style=""text-align: center"">No Documents Found</td></tr>" & vbCrLf)
				Else
					While Not rsFSO.EOF
						Response.Write("<tr><td><a href=""../Documents/MAG/" & MAGID & "/" & rsFSO("Name") & """ target=""_blank"">" & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
						rsFso.MoveNext()
					Wend
				End If
  
				'finally, close out the recordset
				rsFSO.close()
				Set rsFSO = Nothing

				' Archives
				If ShowArchives = True Then
					folder = DocumentRoot & "MAG\" & vMAGID & "\Archives\"
					If Debug = True Then
						Response.Write(folder & "<br />")
					End If
					If fso.FolderExists(folder) Then
						Set rsFSO = kc_fsoFiles(folder, "_")

						While Not rsFSO.EOF
							Response.Write("<tr><td><a href=""../Documents/MAG/" & vMAGID & "/Archives/" & rsFSO("Name") & """ target=""_blank"">(Archives) " & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
							rsFso.MoveNext()
						Wend
  
						'finally, close out the recordset
						rsFSO.close()
						Set rsFSO = Nothing
					End If
				End If

				' Deleted
				If vShowDeleted = True Then
					folder = DocumentRoot & "MAG\" & vMAGID & "\_delete\"
					If Debug = True Then
						Response.Write(folder & "<br />")
					End If
					If fso.FolderExists(folder) Then
						Set rsFSO = kc_fsoFiles(folder, "_")

						While Not rsFSO.EOF
							Response.Write("<tr><td><a href=""../Documents/MAG/" & vMAGID & "/_delete/" & rsFSO("Name") & """ target=""_blank"">(Deleted) " & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
							rsFso.MoveNext()
						Wend
  
						'finally, close out the recordset
						rsFSO.close()
						Set rsFSO = Nothing
					End If
				End If

			End If
		End If
End Function


Function MonitoringDocuments (vGranteeID, vGranteeName, vFiscalYear, vShowDeleted)
	' Expected to show in a table with three columns so file listings of multiple types have uniform columns.
	Dim vsql, vrs, vMonitorDocumentCounter, vfolder

	vsql = "SELECT MonitorID, FiscalYear FROM Monitor.Main WHERE GranteeID=" & vGranteeID & " AND FiscalYear=" & FiscalYear
	If Debug = True Then
		Response.Write("<pre>" &vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = True Then
		Response.Write("<tr><th colspan=""3"">There is no monitoring for " & vFiscalYear & ".</th></tr>" & vbCrLf)
		Exit Function
	End  If

	while vrs.EOF = False
		folder = DocumentRoot & "Monitor\" & vrs.fields("MonitorID")
		If Debug = True Then
			Response.Write("</pre>Folder: " & folder & "</pre>")
		End If

		If fso.FolderExists(folder) Then
			Set rsFSO = kc_fsoFiles(folder, "_")

			If rsFSO.EOF = True Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center"">No Documents Found</td></tr>" & vbCrLf)
			Else
				'Response.Write("<tr><td colspan=""3"" style=""text-align: center"">FY" & vrs.Fields("FiscalYear") & "</td></tr>" & vbCrLf)
				Response.Write("<tr><td colspan=""3""></td></tr>" & vbCrLf)
				Response.Write("<tr><th colspan=""3"">Monitoring: " & GranteeName & " (GranteeID=" & GranteeID &  "), Fiscal Year = " & vFiscalYear & "</th></tr>" & vbCrLf)
				While Not rsFSO.EOF
					Response.Write("<tr><td><a href=""../Documents/Monitor/" & vrs.Fields("MonitorID") & "/" & rsFSO("Name") & """ target=""_blank"">" & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
					vMonitorDocumentCounter = vMonitorDocumentCounter + 1
					rsFso.MoveNext()
				Wend
  
				'finally, close out the recordset
				rsFSO.close()
				Set rsFSO = Nothing

				' Archives
				If ShowArchives = True Then
					folder = DocumentRoot & "Monitor\" & vrs.Fields("MonitorID") & "\Archives\"
					If Debug = True Then
						Response.Write(folder & "<br />")
					End If
					If fso.FolderExists(folder) Then
						Set rsFSO = kc_fsoFiles(folder, "_")

						While Not rsFSO.EOF
							Response.Write("<tr><td><a href=""../Documents/Monitor/" & vrs.Fields("MonitorID") & "/Archives/" & rsFSO("Name") & """ target=""_blank"">(Archives) " & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
							vMonitorDocumentCounter = vMonitorDocumentCounter + 1
							rsFso.MoveNext()
						Wend

  
						'finally, close out the recordset
						rsFSO.close()
						Set rsFSO = Nothing
					End If
				End If

				' Deleted
				If vShowDeleted = True Then
					folder = DocumentRoot & "Monitor\" & vrs.Fields("MonitorID") & "\_delete\"
					If Debug = True Then
						Response.Write(folder & "<br />")
					End If
					If fso.FolderExists(folder) Then
						Set rsFSO = kc_fsoFiles(folder, "_")

						While Not rsFSO.EOF
							vMonitorDocumentCounter = vMonitorDocumentCounter + 1
							Response.Write("<tr><td><a href=""../Documents/Grant/" & vrs.Fields("MonitorID") & "/_delete/" & rsFSO("Name") & """ target=""_blank"">(Deleted) " & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
							rsFso.MoveNext()
						Wend

  
						'finally, close out the recordset
						rsFSO.close()
						Set rsFSO = Nothing
					End If
				End If
			End If
		End If
		vrs.MoveNext
	Wend
	If vMonitorDocumentCounter = 0 Then
		Response.Write("<tr style=""vertical-align: top; ""><td colspan=""3"" style=""text-align: center"">No Documents in folder</td></tr>" & vbCrLf)
	End If

End Function

'**********
'kc_fsoFiles
'Purpose:
' 1. To create a recordset using the FSO object and ADODB
' 2. Allows you to exclude files from the recordset if needed
'Use:
' 1. Call the function when you're ready to open the recordset
' and output it onto the page.
' example:
' Dim rsFSO, strPath
' strPath = Server.MapPath("\PlayGround\FSO\Stuff\")
' Set rsFSO = kc_fsoFiles(strPath, "_")
' The "_" will exclude all files beginning with 
' an underscore 
'**********
Function kc_fsoFiles(theFolder, Exclude)
Dim rsFSO, objFSO, objFolder, File
  Const adInteger = 3
  Const adDate = 7
  Const adVarChar = 200
  
  'create an ADODB.Recordset and call it rsFSO
  Set rsFSO = Server.CreateObject("ADODB.Recordset")
  
  'Open the FSO object
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  
  'go get the folder to output it's contents
  Set objFolder = objFSO.GetFolder(theFolder)
  
  'Now get rid of the objFSO since we're done with it.
  Set objFSO = Nothing
  
  'create the various rows of the recordset
  With rsFSO.Fields
    .Append "Name", adVarChar, 200
    .Append "Type", adVarChar, 200
    .Append "DateCreated", adDate
    .Append "DateLastAccessed", adDate
    .Append "DateLastModified", adDate
    .Append "Size", adInteger
    .Append "TotalFileCount", adInteger
  End With
  rsFSO.Open()
	
  'Now let's find all the files in the folder
  For Each File In objFolder.Files
	
    'hide any file that begins with the character to exclude
    If (Left(File.Name, 1)) <> Exclude Then 
      rsFSO.AddNew
      rsFSO("Name") = File.Name
      rsFSO("Type") = File.Type
      rsFSO("DateCreated") = File.DateCreated
      rsFSO("DateLastAccessed") = File.DateLastAccessed
      rsFSO("DateLastModified") = File.DateLastModified
      rsFSO("Size") = File.Size
      rsFSO.Update
    End If

  Next
	
  'And finally, let's declare how we want the files 
  'sorted on the page. In this example, we are sorting 
  'by File Type in descending order,
  'then by Name in an ascending order.
  rsFSO.Sort = "Name ASC, DateCreated ASC "

  'Now get out of the objFolder since we're done with it.
  Set objFolder = Nothing

  'now make sure we are at the beginning of the recordset
  'not necessarily needed, but let's do it just to be sure.
  If rsFSO.BOF = False Then
	rsFSO.MoveFirst()
  End If
  Set kc_fsoFiles = rsFSO
	
End Function
%><!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->