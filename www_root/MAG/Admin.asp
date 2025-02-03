<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 

Dim debug, i, j, sql2, rs2, MAGID, ORI, GranteeID, FiscalYear, GranteeName, GrantResultID, _
	ResolutionConfirmedDate, ApplicationCertifiedCompleteDate, ApplicationConsideredDate, _
	GrantNumber, GrantAwardAmount, CashMatch, OfficialGrantAwardLetterDate, GrantAwardCertifiedComplete, _
	POIssueDate, GrantClosedDate, Notes, AdminUpdateID, AdminUpdateTimestamp

debug = False

If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLf)
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
			for each j in Request.Cookies(x)
				response.write("Cookies(" & i & ":" & j & ")=" & Request.Cookies(i)(j))
			next
		else
			Response.Write("Cookies(""" & i & """)=" & Request.Cookies(i) & "<br>")
		end if
	next
	Response.Write("Now=" & Now() & vbCrLf)
	Response.Write("</pre>" & vbCrLf)
End If

If Len(Request.Form("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	FiscalYear = 2022
End If

If Len(Request.Form("MAGID")) > 0 Then
	MAGID = CInt(Request.Form("MAGID"))
ElseIf Len(Request.QueryString("MAGID")) > 0 Then
	MAGID = CInt(Request.QueryString("MAGID"))
Else
	MAGID = 0
End If

If Len(Request.Form("GranteeID")) > 0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID")) > 0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	GranteeID = 0
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
	<title>MVCPA Auxiliary Grant Application Management</title>
	<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
	<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body>

<div class="pagetag">Auxiliary Grant Application Administrative Page for Fiscal Year <%=FiscalYear %></div>

<div class="widecontent">
<br />
<form name="Selection" method="post" action="Admin.asp">
Select Fiscal Year: 
<select name="FiscalYear" onchange="Selection.submit();">
	<option value="2022" <%=Selected(FiscalYear, 2022) %>>2022</option>
</select>
<%
sql = "SELECT A.MAGID, A.GranteeID, A.FiscalYear, B.GranteeNameSort, B.GranteeName, " & vbCrLf & _
	"	CASE WHEN OptionID=1 THEN 'Purchase' WHEN OptionID=2 THEN 'Lease' ELSE 'Unknown' END AS GrantOption " & vbCrLf & _
	"FROM MAG.Main AS A " & vbCrLf & _
	"JOIN dbo.Grantees AS B ON A.GranteeID=B.GranteeID " & vbCrLf & _
	"WHERE A.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
	"ORDER BY B.GranteeNameSort "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
%>
Select Application:
<select name="MAGID" onchange="Selection.submit();">
	<option value="0">Select MAG Grant Application</option>
<%
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write(vbTab & "<option value=""" & rs.Fields("MAGID") & """ " & Selected(MAGID, rs.Fields("MAGID")) & ">" & rs.Fields("GranteeName") & " (" & rs.Fields("GrantOption") & ")</option>" & vbCrLf)
	rs.MoveNext
Wend
%>
</select>
</form>
<%
If MAGID > 0 Or (GranteeID>0 And FiscalYear>0) Then
	sql = "SELECT G.GranteeID, G.GranteeName, ISNULL(G.ORI, 'None') AS ORI, O.Agency, " & vbCrLf & _
		"	M.FiscalYear, M.MAGID, M.OptionID, " & vbCrLf & _
		"	CASE WHEN OptionID=1 THEN 'Purchase' WHEN OptionID=2 THEN 'Lease' ELSE 'Unknown' END AS GrantOption, " & vbCrLf & _
		"	M.Certification, S.Name AS SubmitName, IsNull(M.SubmitID,0) AS SubmitID, M.SubmitTimestamp, " & vbCrLf & _
		"	G.AuthorizedOfficialID, AO.Name AS AuthorizedOfficialName, AO.Title AS AuthorizedOfficialTitle, " & vbCrLf & _
		"	PD.Name AS ProgramDirectorName, PD.Title AS ProgramDirectorTitle, " & vbCrLf & _
		"	FO.Name AS FinancialOfficerName, FO.Title AS FinancialOfficerTitle, " & vbCrLf & _
		"	CAST(CASE WHEN G.AuthorizedOfficialID>0 AND G.ProgramDirectorID>0 AND G.FinancialOfficerID>0 THEN 1 ELSE 0 END AS BIT) AS RequiredOfficials, " & vbCrLf & _
		"	StolenVehicles, StolenVehicleValue, " & vbCrLf & _
		"	TFGrant = ISNULL((SELECT ProgramName FROM [Grants].Main AS GM JOIN [Grants].ParticipatingAgencies AS PA ON PA.GrantID=GM.GrantID WHERE PA.ORI=G.ORI AND GM.FiscalYear=" & prepIntegerSQL(FiscalYear) & "), 'None'), " & vbCrLf & _
		"	A.GrantResultID, A.ResolutionConfirmedDate, A.ApplicationCertifiedCompleteDate, " & vbCrLf & _
		"	A.ApplicationConsideredDate, GrantAwardCertifiedComplete, " & vbCrLF & _
		"	A.GrantAwardAmount, A.CashMatch, A.GrantNumber, A.OfficialGrantAwardLetterDate, A.POIssueDate, " & vbCrLF & _
		"	A.GrantClosedDate, A.Notes, " & vbCrLf & _
		"	A.UpdateID AS AdminUpdateID, A.UpdateTimestamp AS AdminUpdateTimestamp " & vbCrLf & _
		"FROM MAG.Main AS M " & vbCrLf & _
		"LEFT JOIN MAG.Admin AS A ON A.MAGID=M.MAGID " & vbCrLf & _
		"LEFT JOIN Grantees AS G ON M.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN Lookup.ORI AS O ON O.ORI=G.ORI " & vbCrLf & _
		"LEFT JOIN System.Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf & _
		"LEFT JOIN System.Users AS PD ON PD.SystemID=G.ProgramDirectorID " & vbCrLf & _
		"LEFT JOIN System.Users AS FO ON FO.SystemID=G.FinancialOfficerID " & vbCrLf & _
		"LEFT JOIN System.Users AS S ON S.SystemID=M.SubmitID " & vbCrLf
	If MAGID > 0 Then
			sql = sql & "WHERE M.MAGID=" & prepIntegerSQL(MAGID)
	Else
			sql = sql & "WHERE M.GranteeID=" & prepIntegerSQL(GranteeID) & " AND FiscalYear=" & prepIntegerSQL(FiscalYear) & " "
	End If		
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		MAGID=rs.Fields("MagID")
		ORI = rs.Fields("ORI")
		GrantResultID = rs.Fields("GrantResultID")
		ResolutionConfirmedDate = rs.Fields("ResolutionConfirmedDate")
		ApplicationCertifiedCompleteDate = rs.Fields("ApplicationCertifiedCompleteDate")
		ApplicationConsideredDate = rs.Fields("ApplicationConsideredDate")
		GrantNumber = rs.Fields("GrantNumber")
		GrantAwardAmount = rs.Fields("GrantAwardAmount")
		CashMatch = rs.Fields("CashMatch")
		OfficialGrantAwardLetterDate = rs.Fields("OfficialGrantAwardLetterDate")
		GrantAwardCertifiedComplete = rs.Fields("GrantAwardCertifiedComplete")
		POIssueDate = rs.Fields("POIssueDate")
		GrantClosedDate = rs.Fields("GrantClosedDate")
		Notes = rs.Fields("Notes")
		AdminUpdateID = rs.Fields("AdminUpdateID")
		AdminUpdateTimestamp = rs.Fields("AdminUpdateTimestamp")
%>
<br />

<form name="Admin" method="post" action="AdminSubmit.asp">
<input type="hidden" name="MAGID" value="<%=MAGID %>" />

<table>

<tr>
	<td>Grantee:
	</td>
	<td><%=rs.Fields("GranteeName") %></td>
</tr>

<tr>
	<td>Grantee ID:</td>
	<td><%=rs.Fields("GranteeID") %></td>
</tr>

<tr>
	<td>Fiscal Year:</td>
	<td><%=rs.Fields("FiscalYear") %></td>
</tr>

<tr>
	<td>ORI:</td>
	<td><%=rs.Fields("ORI") %>&nbsp;&nbsp;<%=rs.Fields("Agency") %></td>
</tr>

<tr>
	<td>MAG ID:</td>
	<td><%=rs.Fields("MAGID") %></td>
</tr>

<tr>
	<td>Grant Option:</td>
	<td><%=rs.Fields("GrantOption") %> (<%=rs.Fields("OptionID") %>)</td>
</tr>

<tr>
	<td>Authorized Official:</td>
	<td><%=rs.Fields("AuthorizedOfficialName") %>, <%=rs.Fields("AuthorizedOfficialTitle") %></td>
</tr>

<tr>
	<td>Program Director:</td>
	<td><%=rs.Fields("ProgramDirectorName") %>, <%=rs.Fields("ProgramDirectorTitle") %></td>
</tr>

<tr>
	<td>Financial Officer:</td>
	<td><%=rs.Fields("FinancialOfficerName") %>, <%=rs.Fields("FinancialOfficerTitle") %></td>
</tr>

<tr>
	<td>Application Submission By:</td>
	<td><%
		If rs.Fields("SubmitID")>0 Then
			Response.Write(rs.Fields("SubmitName")  & ", " & rs.Fields("SubmitTimestamp"))
		Else
			Response.Write("Not submitted")
		End If 
	%></td>
</tr>

<tr>
	<td>Stolen Vehicles:</td>
	<td><%=prepIntegerWeb(rs.Fields("StolenVehicles")) %></td>
</tr>

<tr>
	<td>Stolen Vehicle Value:</td>
	<td><%=prepCurrencyWeb(rs.Fields("StolenVehicleValue")) %></td>
</tr>
<tr>
	<td>Task Force Grant:</td>
	<td><%=rs.Fields("TFGrant") %></td>
</tr>
<%
	If rs.Fields("TFGrant") <> "None" Then
		Response.Write("<tr><td></td><td style=""color: red; font-weight: bold; font-style: italic; "">Existing taskforce grant makes grantee ineligible for auxiliary grant.</td></tr>" & vbCrLf)
	End If
	TFCoverageAgencies rs.Fields("ORI"), rs.Fields("FiscalYear")
%>

<tr>
	<td colspan="2" style="text-align: center">&nbsp;</td>
</tr>

<tr style="vertical-align: top; ">
	<td>Date that Resolution Confirmed:</td>
	<td><%=DateField("ResolutionConfirmedDate", ResolutionConfirmedDate, True) %>&nbsp;<% DisplayResolution MAGID %></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Date Application Certified Complete:</td>
	<td><%=DateField("ApplicationCertifiedCompleteDate", ApplicationCertifiedCompleteDate, True) %></td>
</tr>

<tr>
	<td>Grant Result by MVCPA Board</td>
	<td><select name="GrantResultID" id="GrantResultID">
		<option value="0">Select grant result</option>
<%
	sql2 = "SELECT GrantResultID, GrantResult FROM Lookup.GrantResults ORDER BY GrantResultSort"
	If Debug = True Then
		Response.Write("<pre>" & sql2 & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs2 = Con.Execute(sql2)
	While rs2.EOF = False
		If GrantResultID = rs2.Fields("GrantResultID") Then
			Response.Write("<option value=""" & rs2.Fields("GrantResultID") & """ selected>" & rs2.Fields("GrantResult") & "</option>" & vbCrLf)
		Else
			Response.Write("<option value=""" & rs2.Fields("GrantResultID") & """>" & rs2.Fields("GrantResult") & "</option>" & vbCrLf)
		End If
		rs2.MoveNext()
	Wend
	rs2.Close
%></select></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Date that application is considered by MVCPA Board:</td>
	<td><%=DateField("ApplicationConsideredDate", ApplicationConsideredDate, True) %></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Grant Amount Awarded by MVCPA Board:</td>
	<td><%=CurrencyField("GrantAwardAmount", GrantAwardAmount, 15, 15, True, "") %></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Cash Match from <%=rs.Fields("GranteeName") %>:</td>
	<td><%=CurrencyField("CashMatch", CashMatch, 15, 15, True, "") %></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Grant Number for awarded grants: (dbl-click)</td>
	<td><%=TextFieldDblClick("GrantNumber", GrantNumber, 14, 16, True, "", "this.value='608-" & Mid(CStr(FiscalYear),3,2) & "-" & Mid(ORI,3) & "';") %></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Date that <b>Grant Award Certified Complete</b>:</td>
	<td><%=DateField("GrantAwardCertifiedComplete", GrantAwardCertifiedComplete, True) %>&nbsp;<% DisplaySGA MAGID %></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Date that Purchase Order is issued:</td>
	<td><%=DateField("POIssueDate", POIssueDate, True) %></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Grant Closed, Grantee not pursuing grant:</td>
	<td><%=DateField("GrantClosedDate", GrantClosedDate, True) %></td>
</tr>

<tr style="vertical-align: top; ">
	<td colspan="2">Notes:
	<%=TextArea("Notes", Notes, 4, 120, 2000, True, "") %></td>
</tr>

<tr>
	<td colspan="2" style="text-align: center">&nbsp;</td>
</tr>

<tr>
	<td colspan="2" style="text-align: center"><input type="submit" value="Save" name="Save" id="Save" title="Submit values" />&nbsp;&nbsp;
		<input type="button" name="Close" id="Close" value="Close" onclick="window.close();" />
	</td>
</tr>

<tr>
	<td colspan="2" style="text-align: center">&nbsp;</td>
</tr>
<% DisplayFolder MAGID %>
</table>
</form>
<br />
<br />
<div style="width: 100%; text-align: center">Application Form</div>
<iframe src="MAGApplication.asp?GranteeID=<%=rs.Fields("GranteeID")%>&FiscalYear=<%=rs.Fields("FiscalYear")%>&NoButtons=1" width="100%" height="400"></iframe>
<%
	End If
End If
%>
</div>
</body>
</html>
<%
Function DisplayResolution(vMAGID)
	Dim vfso, vfile, vfilespec, vDocumentFolder
	set vfso = Server.CreateObject("Scripting.FileSystemOBject")

	vDocumentFolder = Application("DocumentRoot") & "\MAG\" & vMAGID & "\"
	If vfso.FolderExists(vDocumentFolder) = False Then
		vfso.CreateFolder(vDocumentFolder)
	End If
	vfilespec = vDocumentFolder & "Resolution.pdf"

	If vfso.FileExists(vfilespec) Then
		Set vfile = vfso.GetFile(vfilespec)
			Response.Write("<a href=""../Documents/MAG/" & vMAGID & "/" & vfile.Name & _
				""" target=""_blank"">" & vfile.Name & "</a> (" & vfile.DateLastModified & ")<br />" & vbCrLf)
	End If
End Function

Function DisplaySGA(vMAGID)
	Dim vfso, vfile, vfilespec, vDocumentFolder
	set vfso = Server.CreateObject("Scripting.FileSystemOBject")

	vDocumentFolder = Application("DocumentRoot") & "\MAG\" & vMAGID & "\"
	If vfso.FolderExists(vDocumentFolder) = False Then
		vfso.CreateFolder(vDocumentFolder)
	End If
	vfilespec = vDocumentFolder & "Signed Statement of Grant Award.pdf"

	If vfso.FileExists(vfilespec) Then
		Set vfile = vfso.GetFile(vfilespec)
			Response.Write("<a href=""../Documents/MAG/" & vMAGID & "/" & vfile.Name & _
				""" target=""_blank"">" & vfile.Name & "</a> (" & vfile.DateLastModified & ")<br />" & vbCrLf)
	End If
End Function

Function DisplayFolder(vMAGID)
	Dim fso, folder, files, file, vDocumentFolder
	Response.Write("<tr><td colspan=""2"" style=""text-align: center; ""><a href=""../Upload/Upload.asp?FID=14&MAGID=" & vMAGID & """ target=""_blank"">File Upload</a></td></tr>" & vbCrLf)

	set fso = Server.CreateObject("Scripting.FileSystemOBject")

	vDocumentFolder = Application("DocumentRoot") & "\MAG\" & vMAGID & "\"
	If fso.FolderExists(vDocumentFolder) = False Then
		fso.CreateFolder(vDocumentFolder)
	End If

	If fso.FolderExists(vDocumentFolder) Then
		Set folder = fso.GetFolder(vDocumentFolder)
		Set files = folder.Files
		Response.Write("<tr><th colspan=""2"">Current Documents in folder</th></tr>" & vbCrLf)
		If files.count>0 Then 
			Response.Write("<tr><th colspan=""2"" style=""text-align: center; "">")
			For Each file in files
					Response.Write("<a href=""../Documents/MAG/" & vMAGID & "/" & file.Name & _
						""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
			Next
			Response.Write("</td></tr>" & vbCrLf)
		Else
			Response.Write("<tr><td colspan=""2"" style=""text-align: center; "">There are no documents in the folder.</td></tr>")
		End If
	End If
End Function

Function TFCoverageAgencies(vORI, vFiscalYear)
	Dim vsql, vrs, vcount
	vcount = 0
	vsql = "SELECT C.ORI, O.Agency, G.GranteeName, M.ProgramName " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLF & _
		"JOIN [Grants].Main AS M ON M.GranteeID=G.GranteeID AND M.FiscalYear=" & vFiscalYear & " " & vbCrLf & _
		"JOIN [Grants].CoverageAgencies AS C ON C.GrantID=M.GrantID AND C.ORI='" & vORI & "' " & vbCrLf & _
		"JOIN Lookup.ORI AS O ON O.ORI=C.ORI " & vbCrLf 
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf & vbTab & "<td>Common Coverage Agencies</td>" & vbCrLf & vbTab & "<td>")
		While vrs.EOF = False
			If vcount > 0 Then
				Response.Write("<br />" & vbCrLf)
			End If
			Response.Write(vrs.Fields("ORI") & ": " & vrs.Fields("Agency") & " in common with " & vrs.Fields("GranteeName") & " " & vrs.Fields("ProgramName"))
			vcount = vcount + 1
			vrs.MoveNext
		Wend
		Response.Write("</td>" & vbCrLf & "</tr>")
	End If
End Function
%>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/InputValidation.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->