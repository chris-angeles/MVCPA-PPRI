<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, PermitEdit, ShowExcel, Columns, GrantID, FiscalYear, GranteeID, GranteeName, _
	ProgramName,  AdministrativeComments, CanSubmit, Version, ViewDocuments, Quarter, _
	SubmitID, SubmitName, SubmitTimestamp, ApprovalID, ApprovalTimestamp, ApprovalName
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
	Response.Write("</pre>" & vbCrLf)
End If

If Request.Form.Count>0 Then
	GrantID = Request.Form("GrantID")
Else
	GrantID = Request.QueryString("GrantID")
End If

If Len(GrantID)>0 Then
	GrantID = CInt(GrantID)
Else
	GrantID=0
End If

If Request.Querystring("ShowExcel")="1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "SELECT H.GrantID, H.FiscalYear, G.GranteeID, G.GranteeName, H.ProgramName,  " & vbCrLf & _
	"	ISNULL(Y.SubmitID,0) AS SubmitID, S.Name AS SubmitName, Y.SubmitTimestamp, " & vbCrLF & _
	"	ISNULL(Y.ApprovalID,0) AS ApprovalID, A.Name AS ApprovalName, Y.ApprovalTimestamp, " & vbCrLf & _
	"	Y.AdministrativeComments, " & vbCrLf & _
	"	CAST(CASE WHEN " & UserSystemID & " IN (G.ProgramDirectorID, G.ProgramManagerID) THEN 1 ELSE 0 END AS BIT) AS CanSubmit " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS H ON G.GranteeID=H.GranteeID " & vbCrLf & _
	"LEFT JOIN YE.Main AS Y ON Y.GrantID=H.GrantID " & vbCrLf & _
	"LEFT JOIN [System].Users AS S ON S.SystemID=Y.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS A ON A.SystemID=Y.ApprovalID " & vbCrLf & _
	"WHERE H.GrantID=" & prepIntegerSQL(GrantID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF Then
	Response.Write("Error: Grant not found.")
	Response.End
Else
	GrantID = rs.Fields("GrantID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	ProgramName = rs.Fields("ProgramName")
	SubmitID = rs.Fields("SubmitID")
	SubmitName = rs.Fields("SubmitName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	ApprovalID = rs.Fields("ApprovalID")
	ApprovalTimestamp = rs.Fields("ApprovalTimestamp")
	ApprovalName = rs.Fields("ApprovalName")
	AdministrativeComments = rs.Fields("AdministrativeComments")
	CanSubmit = rs.Fields("CanSubmit")
End If

If FiscalYear>2020 Then
	Version = 2
Else
	Version = 1
End If

If GranteeID>0 Then
	If SubmitID = 0 Then
		PermitEdit = CheckPermissions(UserSystemID, GranteeID, True)
	ElseIf SubmitID > 0 Then
		PermitEdit = False
	Else
		PermitEdit = False
	End If
Else
	PermitEdit = False
End If
'PermitEdit = True ' For testing.
'CanSubmit = True ' For Testing.

sql = "SELECT A.QuestionID, A.Version, A.Identifier, A.Section, REPLACE(A.Question, '{FiscalYear}','FY" & (FiscalYear MOD 100) & "') AS Question, " & vbCrLF & _
	"	A.QuestionSort, B.Response " & vbCrLf & _
	"FROM YE.Questions AS A " & vbCrLf & _
	"LEFT JOIN YE.Responses AS B ON B.QuestionID=A.QuestionID AND B.GrantID=" & prepIntegerSQL(GrantID) & " " & " AND B.Version=A.Version " & vbCrLf & _
	"WHERE A.Version=" & Version
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ProgressReport" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr><th colspan=""" & columns & """>" & GranteeName & " MVCPA Progress Report for Fiscal Year " & FiscalYear & ", Quarter " & Quarter & "</th></tr>" & vbCrLf)
	If SubmitID>0 Then 
		Response.Write("<tr><td colspan=""" & columns & """ style=""text-align: center; font-weight: bold; "">The progress report was submitted by " & SubmitName & " at " & SubmitTimestamp & " and is now locked.</td></tr>" & vbCrLf)
	End If
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Performance Report for <%=GranteeName %> <%=ProgramName %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style>
	tr, td, th {padding: 5px;}
</style>
<script type="text/javascript">
	function validateForm()
	{
		return true;
	}

	function submitForm(action)
	{
		validateForm();
		if (action == "submit") {
			if (document.PR["Confirmed"].checked == false) {
				alert("You must read certification and check box to submit progress report.");
				return false;
			}
			document.PR["action"].value = "submit";
		}
		document.PR.submit();
	}

	function changedCurrencyField(field)
	{
		if (checkCurrency(field) == false)
			return false;
		return true;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body style="width: 90%; margin: auto;">
<div style="text-align: center; margin: auto;">Year End Performance Report FY<%=(FiscalYear MOD 100) %></div>
<div style="text-align: center; margin: auto;"><%=GranteeName %></div>
<div style="text-align: center; margin: auto;"><%=ProgramName %></div>
<%	
If SubmitID > 0 Then
	Response.Write("<div style=""text-align: center; margin: auto;"">Submitted by " & SubmitName & " at " & SubmitTimestamp & "</div>" & vbCrLf)
End If
If ApprovalID > 0 Then
	Response.Write("<div style=""text-align: center; margin: auto;"">MVCPA Approval by " & ApprovalName & " at " & ApprovalTimestamp & "</div>" & vbCrLf)
End If
%>
<br />
<%	If FiscalYear>2020 Then %>
<p>The Final or Year End Performance Progress Report is due on October 15th of each year. 
This report will provide the grantees' own analysis of the just completed grant year and activities. 
The year-end progress report is critical to accurately obtain a summary of all the quarterly progress 
reports and grant activities. <b>The information is used directly in the annual report to the Texas 
Legislature.</b> The report should be written by the taskforce commander or unit supervisor most familiar 
with grant funded activities. Please provide overall summary perspective and lessons learned in the 
responses. Provide answers that funders <b>(MVCPA and the Legislature)</b> and policy makers can gain 
law enforcement perspective on what effect the program has and how state resources can be used to 
support the economic motor vehicle theft enforcement teams.</p>
<%	ElseIf FiscalYear>2019 Then %>
<p>The following year-end progress report is critical to accurately obtain a summation of the 
progress report and grant activities of all grantees. <b>The information is used directly in the 
annual report to the Texas Legislature.</b> The report should be written by the taskforce commander 
or unit supervisor most familiar with grant funded activities. Please provide overallsummary perspective 
and lessons learned in the responses below. Provide answers that funders and policy makers can 
gain perspective on what effect the program has and how state resources can be used to support 
the economic motor vehicle theft enforcement teams.</p>
<%	Else %>
<p>The due date is October 1, <%=FiscalYear %> for the year-end progress report.  
	Please provide answers to these questions for the FY<%=(FiscalYear MOD 100) %>  grant year. 
	Provide insight into how this grant affected the community you serve and your law 
	enforcement organization.</p>
<%	End If %>
<br />
<form name="YearEnd" action="YearEndSubmit.asp" method="post">
<%
	Response.Write(HiddenField("GrantID", GrantID))
	Response.Write(HiddenField("Version", Version))
End If
%>
<table style="margin: auto;">
<%
	While rs.EOF = False 
		If IsNull(rs.Fields("Section")) = False Then
			Response.Write("<tr><td colspan=""2"" style=""font-weight: bold"">" & rs.Fields("Section") & "</td></tr>")
		End If
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right;"">" & rs.Fields("Identifier") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("Question") & "</td>" & vbCrLF)
		Response.Write("</tr>" & vbCrLf)
		Response.Write("<tr>" & vbCrLf)
		Response.Write(vbTab & "<td></td><td>" & TextArea2("Response_" & rs.Fields("QuestionID"), rs.Fields("Response"), 12, 920, 10000, PermitEdit, "") & "</td>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend 
%>
</table>
<%
Quarter = "YE"
ViewDocuments = True
If ViewDocuments = True Then
	Dim Folder, file, files, DocumentFolder, fso, counter
	counter=0
	DocumentFolder = Application("DocumentRoot") & "\Grant\" & GrantID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	Response.Write("<table style=""margin: auto; "">" & vbCrLf)
	Response.Write("<tr><td>Current Documents in folder: ")
	If PermitEdit = True Then
		Response.Write("<a href=""../Upload/Upload.asp?fid=5&quarter=" & Quarter & "&GrantID=" & GrantID & """ class=""plainlink"" target=""_blank"">Upload</a>")
	End If
	Response.Write("</td>" & vbCrLf)
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		If files.count>0 Then 
			Response.Write("<tr><td>")
			For Each file in files
				If Left(file.Name,4)="PR"&Quarter Then
					Response.Write("<a href=""../Documents/Grant/" & GrantID & "/" & file.Name & _
						""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
					counter = counter + 1
				End If
			Next
			Response.Write("</td></tr>" & vbCrLf)
		End If
	End If
	If counter = 0 Then
		Response.Write("<tr style=""vertical-align: top; ""><td style=""text-align: center; "">No Documents in folder</td></tr>")
	End If
	Response.Write("</table>" & vbCrLf)
	Response.Write("<hr />" & vbCrLf)
End If


If MVCPARights = True Then
	Response.Write("<div style=""text-align: center; margin: auto;"">Administrative Section</div>")
	If SubmitID>0 Then
		If ApprovalID>0 Then
			Response.Write("<div>" & CheckBoxField("Approval", True) & " MVCPA Approval by " & ApprovalName & " at " & ApprovalTimestamp & "</div>")
		Else
			Response.Write("<div>" & CheckBoxField("Approval", False ) & " MVCPA Approval</div>")
		End If
		If SubmitID > 0 Then
			Response.Write("<div>" & CheckBoxField("Unsubmit", False ) & " Unsubmit Report</div>")
		End If
	End If
	Response.Write("<div>Comments: <br />" &  TextArea2("AdministrativeComments", AdministrativeComments, 6, 920, 1990, MVCPARights, "") & "</div>" & vbCrLf)
End If
%>
<br />
<div style="margin: auto; text-align: center; "><input type="submit" name="Save" value="Save" />&nbsp;&nbsp;
<%	If CanSubmit=True and SubmitID = 0 Then %>
	<input type="submit" name="Submit" value="Submit" />&nbsp;&nbsp;
<%	End If %>
	<input type="button" name="Close" value="Close" onclick="window.close();" />
</div>


</form>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->