<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, FunctionalAreaID, FunctionalArea, DirectoryID, Location, Prefix, Identifier, _
	DocumentRoot, UploadDirectory, fso, filename, file, files, folder, _
	GrantID, AppID, ISAID, MAGID, GranteeID, AdjustmentID, MonitorID, GranteeName, ProgramName, _
	FiscalYear, Quarter, pdfonly, DocumentTypeID, GrantClass
debug = False
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
	Response.Write("</pre>" & vbCrLF)
End If
DocumentRoot = Application("DocumentRoot")

If Request.Form.Count > 0 Then
	FunctionalAreaID = Request.Form("FID")
	GrantID = Request.Form("GrantID")
	AppID = Request.Form("AppID")
	MAGID = Request.Form("MAGID")
	ISAID = Requet.Form("ISAID")
	DocumentTypeID = Request.Form("dt")
	AdjustmentID = Request.Form("AdjustmentID")
	MonitorID = Request.Form("MonitorID")
	Quarter = Request.Form("Quarter")
ElseIf Request.Querystring.Count > 0 Then
	FunctionalAreaID = Request.QueryString("FID")
	GrantID = Request.QueryString("GrantID")
	AppID = Request.QueryString("AppID")
	MAGID = Request.QueryString("MAGID")
	ISAID = Request.QueryString("ISAID")
	DocumentTypeID = Request.QueryString("dt")
	AdjustmentID = Request.Querystring("AdjustmentID")
	MonitorID = Request.QueryString("MonitorID")
	Quarter = Request.QueryString("Quarter")
End If
If Len(FunctionalAreaID)=0 Then
	Response.Write("A Functional Area ID must be provided!")
	SendMessage "A Functional Area ID must be provided!"
	Response.End
End If
If Len(GrantID)=0 And Len(AppID)=0 And Len(ISAID)=0 And Len(MonitorID)=0 And Len(MAGID)=0 Then
	Response.Write("A GrantID, AppID, MAGID, ISAID, or MonitorID must be provided!")
	SendMessage "A GrantID, AppID, MAGID, ISAID, or MonitorID must be provided!"
	Response.End
End If

sql = "SELECT A.FunctionalAreaID, A.FunctionalArea, B.DocumentTypeID, " & vbCrLf & _
	"	A.DirectoryID, A.Location, A.Prefix, A.Identifier, " & vbCrLF & _
	"	B.DocumentTypeDescription, B.NameToUse, B.EditableName, B.MVCPAOnly " & vbCrLf & _
	"FROM Lookup.FunctionalAreas AS A " & vbCrLf & _
	"JOIN Lookup.DocumentTypes AS B ON B.FunctionalAreaID=A.FunctionalAreaID " & vbCrLf
If DocumentTypeID > 0 Then
	sql = sql & "WHERE B.DocumentTypeID=" & prepIntegerSQL(DocumentTypeID)
Else
	sql = sql & "WHERE A.FunctionalAreaID=" & prepIntegerSQL(FunctionalAreaID)
End IF
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(SQL)
If rs.EOF = True Then
	Response.Write("Error: Unable to document records  for FunctionalAreaID=" & FunctionalAreaID)
	SendMessage "Error: Unable to document records  for FunctionalAreaID=" & FunctionalAreaID
	Response.End
Else
	FunctionalAreaID = rs.Fields("FunctionalAreaID")
	FunctionalArea = rs.Fields("FunctionalArea")
	DirectoryID = rs.Fields("DirectoryID")
	Location = rs.Fields("Location")
	Prefix = rs.Fields("Prefix")
	Identifier = rs.Fields("Identifier")
End If

If DirectoryID="ISAID" AND ISAID="" Then
	Response.Write("An ISAID must be provided to identify the directory!")
	SendMessage "An ISAID must be provided to identify the directory!"
	Response.End
ElseIf DirectoryID="ISAID" Then
	ISAID = CInt(ISAID)
End If
If DirectoryID="AppID" AND AppID="" Then
	Response.Write("An AppID must be provided to identify the directory!")
	SendMessage "An AppID must be provided to identify the directory!"
	Response.End
ElseIf DirectoryID="AppID" Then
	AppID = CInt(AppID)
End If
If DirectoryID="GrantID" AND GrantID="" Then
	Response.Write("A GrantID must be provided to identify the directory!")
	SendMessage "A GrantID must be provided to identify the directory!"
	Response.End
ElseIf DirectoryID="GrantID" Then
	GrantID = CInt(GrantID)
End If

If DirectoryID="MonitorID" AND MonitorID="" Then
	Response.Write("A MonitorID must be provided to identify the directory!")
	SendMessage "A MonitorID must be provided to identify the directory!"
	Response.End
ElseIf DirectoryID="MonitorID" Then
	MonitorID = CInt(MonitorID)
End If

If DirectoryID="MAG" AND MAGID="" Then
	Response.Write("A MAGID must be provided to identify the directory!")
	SendMessage "A MAGID must be provided to identify the directory!"
	Response.End
ElseIf DirectoryID="MAGID" Then
	MAGID = CInt(MAGID)
End If

If Len(DocumentTypeID)>0 Then
	DocumentTypeID = CInt(DocumentTypeID)
Else
	DocumentTypeID = 0
End If

If Instr(Identifier,"AdjustmentID") > 0 And Len(AdjustmentID)=0 Then
	Response.Write("An AdjustmentID must be provided to identify the appropriate Adjustment!")
	SendMessage "An AdjustmentID must be provided to identify the appropriate Adjustment!"
	Response.End
Else
	AdjustmentID = CInt(AdjustmentID)
End If
If Instr(Identifier,"FiscalYearID") > 0 Then
	Response.Write("The Quarter must be provided to identify the file!")
	SendMessage "The Quarter must be provided to identify the file!"
	Response.End
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

If Debug = True Then
	Response.Write("<pre>FunctionalAreaID=""" & FunctionalAreaID & """</pre>")
	Response.Write("<pre>DocumentTypeID=""" & DocumentTypeID & """</pre>")
	Response.Write("<pre>DocumentRoot = """ & DocumentRoot & """</pre>" & vbCrLf)
	Response.Write("<pre>Location = """ & Location & """</pre>" & vbCrLf)
	Response.Write("<pre>ISAID = """ & ISAID & """</pre>" & vbCrLf)
	Response.Write("<pre>AppID = """ & AppID & """</pre>" & vbCrLf)
	Response.Write("<pre>GrantID = """ & GrantID & """</pre>" & vbCrLf)
	Response.Write("<pre>MonitorID = """ & MonitorID & """</pre>" & vbCrLf)
	Response.Write("<pre>Upload Directory = """ & UploadDirectory & """</pre>" & vbCrLf)
	Response.Flush
End If

set fso = Server.CreateObject("Scripting.FileSystemOBject")
If fso.FolderExists(UploadDirectory) Then
	If Debug = True Then
		Response.Write("<pre>UploadDirectory,""" & UploadDirectory & """, exists.</pre>")
	End If	
Else
	If Debug = True Then
		Response.Write("<pre>UploadDirectory, """ & UploadDirectory & """, does not exist.</pre>")
	End If	
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
Set folder = fso.GetFolder(UploadDirectory)
Set files = folder.Files

If GrantID>0 Then
	sql = "SELECT G.GranteeID, G.GranteeName, A.GrantID, A.ProgramName, A.FiscalYear " & vbCrLF & _
		"FROM [Grants].Main AS A " & vbCrLF & _
		"LEFT JOIN Grantees AS G ON G.GranteeID=A.GranteeID " & vbCrLF & _
		"WHERE A.GrantID=" & prepIntegerSQL(GrantID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(SQL)
	If rs.EOF = True Then
		Response.Write("Error: Unable to retrieve Grant record for GrantID=" & GrantID)
		SendMessage "Error: Unable to retrieve Grant record for GrantID=" & GrantID
		Response.End
	Else
		GranteeName = rs.Fields("GranteeName")
		ProgramName = rs.Fields("ProgramName")
		FiscalYear = rs.Fields("FiscalYear")
		GranteeID = rs.Fields("GranteeID")
	End If
ElseIf AppID>0 Then
	sql = "SELECT I.GranteeID, G.GranteeName, I.AppID, I.FiscalYear, I.GrantClassID, GC.GrantClass, " & vbCrLf & _
		"	CASE WHEN I.GrantClassID=1 THEN A1.ProgramName WHEN I.GrantClassID=4 THEN A4.ProgramName ELSE 'Unknown Application' END AS ProgramName " & vbCrLf & _
		"FROM Application.IDs AS I " & vbCrLf & _
		"LEFT JOIN Application.Main AS A1 ON A1.AppID=I.AppID " & vbCrLF & _
		"LEFT JOIN CC.Application AS A4 ON A4.AppID=I.AppID " & vbCrLF & _
		"LEFT JOIN Grantees AS G ON G.GranteeID=I.GranteeID " & vbCrLF & _
		"LEFT JOIN Lookup.GrantClass AS GC ON GC.GrantClassID=I.GrantClassID " & vbCrLf & _
		"WHERE I.AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(SQL)
	If rs.EOF = True Then
		Response.Write("Error: Unable to retrieve Application record for AppID=" & AppID)
		SendMessage "Error: Unable to retrieve Application record for AppID=" & AppID
		Response.End
	Else
		GranteeName = rs.Fields("GranteeName")
		ProgramName = rs.Fields("ProgramName")
		FiscalYear = rs.Fields("FiscalYear")
		GrantClass = rs.Fields("GrantClass")
		If IsNull(GrantClass) - True Then
			GrantClass = ""
		End If
		GranteeID = rs.Fields("GranteeID")
	End If
ElseIf MonitorID>0 Then
	sql = "SELECT G.GranteeID, G.GranteeName, A.MonitorID, H.ProgramName, A.FiscalYear " & vbCrLF & _
		"FROM Monitor.Main AS A " & vbCrLF & _
		"LEFT JOIN Grantees AS G ON G.GranteeID=A.GranteeID " & vbCrLF & _
		"LEFT JOIN [Grants].Main AS H ON H.GranteeID=A.GranteeID AND H.FiscalYear=A.FiscalYear " & vbCrLf & _
		"WHERE A.MonitorID=" & prepIntegerSQL(MonitorID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(SQL)
	If rs.EOF = True Then
		Response.Write("Error: Unable to retrieve Application record for AppID=" & AppID)
		SendMessage "Error: Unable to retrieve Application record for AppID=" & AppID
		Response.End
	Else
		GranteeName = rs.Fields("GranteeName")
		ProgramName = rs.Fields("ProgramName")
		FiscalYear = rs.Fields("FiscalYear")
		GranteeID = rs.Fields("GranteeID")
	End If
ElseIf MAGID>0 Then
	sql = "SELECT G.GranteeID, G.GranteeName, A.MAGID, G.GranteeName + ' FY' + CAST(FiscalYear AS VARCHAR) + ' Auxiliary Grant' AS ProgramName, A.FiscalYear " & vbCrLf & _
		"FROM MAG.Main AS A " & vbCrLF & _
		"LEFT JOIN Grantees AS G ON G.GranteeID=A.GranteeID " & vbCrLF & _
		"WHERE A.MAGID=" & prepIntegerSQL(MAGID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(SQL)
	If rs.EOF = True Then
		Response.Write("Error: Unable to retrieve Application record for MAGID=" & AppID)
		SendMessage "Error: Unable to retrieve Application record for MAGID=" & AppID
		Response.End
	Else
		GranteeName = rs.Fields("GranteeName")
		ProgramName = rs.Fields("ProgramName")
		FiscalYear = rs.Fields("FiscalYear")
		GranteeID = rs.Fields("GranteeID")
	End If
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Document Upload</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function uploadDoc()
	{
		if (check() == false) {
			alert("You must select a document type to submit this page.");
			return false;
		}
		if (document.Upload.File1.value.length == 0) {
			alert("You must select a file first.");
			return false;
		}
		var re = /(?:\.([^.]+))?$/;
		var extension = re.exec(document.Upload.File1.value)[1].toLowerCase();
		if (extension == "pdf") {
			// do nothing;
		}
		else if (extension == 'docx' || extension == 'doc') {
			// do nothing
		}
		else if (extension == 'xlsx' || extension == 'xls') {
			// do nothing
		}
		else if (extension == 'pptx' || extension == 'ppt') {
			// do nothing
		}
		else {
			alert("You may only upload a pdf, docx, xlsx, or pptx file.");
			return false;
		}
		document.Upload.submit();
	}

	function check()
	{
		var radios = document.getElementsByName("DocumentTypeID");

		for (var i = 0, len = radios.length; i < len; i++) {
			if (radios[i].checked) {
				return true;
			}
		}

		return false;
	}

</script>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">This is the file upload page for the MVCPA.tamu.edu website.</div>

<div class="widecontent">

<h1>Document Upload for <%=GranteeName %></h1>
<%	If AppID>0 Then %>
<h2><%=GrantClass %> Application "<%=ProgramName %>" for Fiscal Year <%=FiscalYear %></h2>
<%	ElseIf GrantID>0 Then %>
<h2>Grant "<%=ProgramName %>" for Fiscal Year <%=FiscalYear %></h2>
<%	End If %>

<div style="width: 720px; margin: auto; ">
	<form name="Upload" method="post" enctype="multipart/form-data" action="UploadSubmit.asp">
	<input type="hidden" name="FunctionalAreaID" value="<%=FunctionalAreaID %>" />
	<input type="hidden" name="MAGID" value="<%=MAGID %>" />
	<input type="hidden" name="ISAID" value="<%=ISAID %>" />
	<input type="hidden" name="AppID" value="<%=AppID %>" />
	<input type="hidden" name="GrantID" value="<%=GrantID %>" />
	<input type="hidden" name="GranteeID" value="<%=GranteeID %>" />
	<input type="hidden" name="AdjustmentID" value="<%=AdjustmentID %>" />
	<input type="hidden" name="MonitorID" value="<%=MonitorID %>" />
	<input type="hidden" name="Quarter" value="<%=Quarter%>" />
	<p>Please select the type of document that you are uploading:</p>
<%
	If MVCPARights = True Then
		sql = "SELECT * FROM Lookup.DocumentTypes WHERE FunctionalAreaID=" & FunctionalAreaID & " ORDER BY DocumentTypeID "
	Else
		sql = "SELECT * FROM Lookup.DocumentTypes WHERE FunctionalAreaID=" & FunctionalAreaID & " AND MVCPAOnly=0 ORDER BY DocumentTypeID "
	End If
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error: No Documents Types Available for FunctionalAreaID=" & FunctionalAreaID)
		SendMessage "Error: No Documents Types Available for FunctionalAreaID=" & FunctionalAreaID
		Response.End
	Else
		While rs.EOF = False
			IF rs.Fields("MVCPAOnly") = True Then
				Response.Write("<input type=""radio"" name=""DocumentTypeID"" value=""" & rs.Fields("DocumentTypeID") & """ " & Checked(rs.Fields("DocumentTypeID"), DocumentTypeID) & ">" & rs.Fields("DocumentTypeDescription") & "<sup><font color=""red"">A</font></sup><br />" & vbCrLf)
			Else
				Response.Write("<input type=""radio"" name=""DocumentTypeID"" value=""" & rs.Fields("DocumentTypeID") & """ " & Checked(rs.Fields("DocumentTypeID"), DocumentTypeID) & ">" & rs.Fields("DocumentTypeDescription") & "<br />" & vbCrLf)
			End If
			rs.MoveNext()
		Wend
	End If
%>
	<br />
	<input type="file" name="File1" style="width: 600px; "
		title="Click on 'Browse' to select file. After selecting, click on 'Upload' to send file." 
		accept="application/pdf, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.openxmlformats-officedocument.presentationml.presentation, application/msword, application/vnd.ms-excel, application/vnd.ms-powerpoint" /><br />
	<br />
	<div style="text-align: center; margin: auto;">
		<input type="button" name="btnUpload" value="Upload" onclick="uploadDoc();" />&nbsp;&nbsp;&nbsp;
		<input type="button" name="btnClose" value="Close" onclick="JavaScript: window.close()" 
			title="Close will result in any selected being ignored and this window will be closed." />
	</div>
	<br />
	</form>
<%	If files.count>0 Then %>
	<div style="width: 600px">
	<h2>Current Documents in folder</h2>
<%
	For Each file in files
		If GrantID>0 Then
			Response.Write("<a href=""../Documents/Grant/" & GrantID & "/" & file.Name & _
				""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
		ElseIf AppID>0 Then
			Response.Write("<a href=""../Documents/Application/" & AppID & "/" & file.Name & _
				""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
		ElseIf MonitorID>0 Then
			Response.Write("<a href=""../Documents/Monitor/" & MonitorID & "/" & file.Name & _
				""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
		ElseIf MAGID>0 Then
			Response.Write("<a href=""../Documents/MAG/" & MAGID & "/" & file.Name & _
				""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
		End If
	Next
%><br />
	</div>
<%	End If %>
	</div>
</div>
<div style="width: 720px; margin: auto; margin-top: 20px; margin-bottom: 10px; font-style: italic;">Browsers have upgraded security. 
You may need to click on the "X" at the upper right hand corner of window to close this window. 
The close button does not work sometimes.</div>
<div style="width: 720px; margin: auto; margin-top: 10px; margin-bottom: 50px; font-style: italic;">Please ensure that any 
files uploaded are scanned for viruses. If your system has been setup by your IT department to scan all documents, this is 
probably done automatically. On systems where this does not happen, please ensure that your documents are free from viruses. 
This is especially important if you are working from home. Be sure that your virus software is up-to-date. Follow
your agencies cybersecurity processes. Know what to do if you discover that your computer has been hacked. If you find
out that your computer has been hacked and you recently uploaded document to GMTS, please let us know.</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->