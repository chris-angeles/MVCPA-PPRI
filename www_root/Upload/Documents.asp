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
<%
Dim debug, i, DocumentRoot, folders, folder, files, file, rsFSO, fso, LastGranteeID, LastGrantID, LastAppID, InitialYear
Debug = False

If Len(Request.Form("InitialYear")) > 0 Then
	InitialYear = CInt(Request.Form("InitialYear"))
ElseIf Len(Request.QueryString("InitialYear")) > 0 Then
	InitialYear = CInt(Request.QueryString("InitialYear"))
Else
	InitialYear = Year(Now())-3
End If

DocumentRoot = Application("DocumentRoot")
%>
<body style="width: 100%">
<h1>Documents Uploaded to the website</h1>
<form name="Selection" method="post" action="Documents.asp"><p style="text-align: center; width: 100%; "">Years since: 
<select name="InitialYear" onclick="document.Selection.submit();"><%
For i = 2017 to Year(Now()) + 1
	Response.Write(SelectOption(i, i, InitialYear))
Next
%>
</select></p></form>
<% 
Set fso = Server.CreateObject("Scripting.FileSystemObject")

sql = "SELECT A.GranteeID, B.GrantID, I.AppID, I.FiscalYear, A.GranteeName, B.ProgramName, B.GrantNumber " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLF & _
	"LEFT JOIN Application.IDs AS I ON I.GranteeID=A.GranteeID AND I.GrantClassID=1 " & vbCrLf & _
	"JOIN [Grants].Main AS B ON B.GranteeID=I.GranteeID " & vbCrLf & _
	"LEFT JOIN Application.Main AS C ON C.AppID=I.AppID " & vbCrLf & _
	"WHERE I.FiscalYear>=" & InitialYear & " " & vbCrLf & _
	"ORDER BY REPLACE(A.GranteeName,'City Of ',''), FiscalYear "

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

If rs.EOF Then
	Response.Write("Do Documents Found")
	Response.End
End If

LastGranteeID = 0
LastGrantID = 0
LastAppID = 0

Response.Write("<table style=""margin: auto; "">" & vbCrLf)
While rs.EOF = False
	If LastGranteeID <> rs.Fields("GranteeID") Then
		LastGranteeID = rs.Fields("GranteeID")
		Response.Write("<tr><th colspan=""3"">" & rs.Fields("GranteeName") & " (Grantee ID: " & rs.Fields("GranteeID") &  ")</th></tr>" & vbCrLf)

		folder = DocumentRoot & "Monitor\" & rs.Fields("GranteeID")
		If Debug = True Then
			Response.Write(folder & "<br />")
		End If
		If fso.FolderExists(folder) Then
			If LastGrantID <> rs.Fields("GranteeID") Then
				LastGrantID = rs.Fields("GranteeID")
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; "">Monitoring Documents</td></tr>")
				'Response.Write("<tr><th colspan=""3"">Monitoring Documents</th></tr>" & vbCrLf)
			End If
			Set rsFSO = kc_fsoFiles(folder, "_")

			While Not rsFSO.EOF
				Response.Write("<tr><td><a href=""../Documents/Monitor/" & rs.Fields("GranteeID") & "/" & rsFSO("Name") & """ target=""_blank"">" & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
				rsFso.MoveNext()
			Wend
  
			'finally, close out the recordset
			rsFSO.close()
			Set rsFSO = Nothing
		End If
		Response.Flush
	End If
	folder = DocumentRoot & "Application\" & rs.Fields("AppID")
	If Debug = True Then
		Response.Write(folder & "<br />")
	End If
	If fso.FolderExists(folder) Then
		If LastAppID <> rs.Fields("AppID") Then
			LastAppID = rs.Fields("AppID")
			Response.Write("<tr><th colspan=""3"">" & rs.Fields("ProgramName") & ", " & rs.Fields("FiscalYear") & " (Grant ID: " & rs.Fields("GrantID") & ", App ID=" & rs.Fields("AppID") & ")</th></tr>" & vbCrLf)
			Response.Write("<tr><td colspan=""3"" style=""text-align: center; "">Application</td></tr>")
		End If
		Set rsFSO = kc_fsoFiles(folder, "_")

		While Not rsFSO.EOF
			Response.Write("<tr><td><a href=""../Documents/Application/" & rs.Fields("AppID") & "/" & rsFSO("Name") & """ target=""_blank"">" & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
			rsFso.MoveNext()
		Wend
  
		'finally, close out the recordset
		rsFSO.close()
		Set rsFSO = Nothing
	End If

	folder = DocumentRoot & "Grant\" & rs.Fields("GrantID")
	If Debug = True Then
		Response.Write(folder & "<br />")
	End If
	If fso.FolderExists(folder) Then
		If LastGrantID <> rs.Fields("GrantID") Then
			LastGrantID = rs.Fields("GrantID")
			Response.Write("<tr><th colspan=""3"">" & rs.Fields("ProgramName") & ", " & rs.Fields("FiscalYear") & " (Grant ID: " & rs.Fields("GrantID") &  ")</th></tr>" & vbCrLf)
			Response.Write("<tr><td colspan=""3"" style=""text-align: center; "">Grant Documents</td></tr>")
		End If
		Set rsFSO = kc_fsoFiles(folder, "_")

		While Not rsFSO.EOF
			Response.Write("<tr><td><a href=""../Documents/Grant/" & rs.Fields("GrantID") & "/" & rsFSO("Name") & """ target=""_blank"">" & rsFSO("Name").Value & "</a></td><td>" & rsFSO("DateLastModified") & "</td><td>" & rsFSO("Type").Value & "</td></tr>")
			rsFso.MoveNext()
		Wend
  
		'finally, close out the recordset
		rsFSO.close()
		Set rsFSO = Nothing
	End If

	rs.MoveNext()
Wend

Response.Write("</table>" & vbCrLf)

%>
</body>
</html>

<%
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
%>
<!--#include file="../includes/InputHelpers.asp"-->