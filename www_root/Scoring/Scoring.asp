<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, RoundCurrency, LastSectionID, LastQuestionID, Columns, ScoringGroups, PermitEdit, PermitEditScores, _
	FiscalYear, GrantClassID, GrantTypeID, AppID, QuestionID, GSAVersion, TextSectionVersion, GrantTypeVersion, ScoringVersion
debug = False
RoundCurrency = True

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

If Len(Request.Form("FiscalYear"))>0 Then 
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then 
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	FiscalYear=2025
End If

If Len(Request.Form("GrantClassID"))>0 Then 
	GrantClassID = CInt(Request.Form("GrantClassID"))
ElseIf Len(Request.QueryString("GrantClassID"))>0 Then 
	GrantClassID = CInt(Request.QueryString("GrantClassID"))
Else
	GrantClassID=1
End If

If GrantClassID = 1 Then
	If FiscalYear>= 2022 Then
		GSAVersion = 5
	ElseIf FiscalYear>= 2021 Then
		GSAVersion = 4
	ElseIf FiscalYear>= 2020 Then
		GSAVersion = 2
	ElseIf FiscalYear>= 2018 Then
		GSAVersion = 2
	Else
		GSAVersion = 1
	End If

	If FiscalYear=2022 Then
		GrantTypeVersion = 2
	Else
		GrantTypeVersion = 1
	End If

ElseIf GrantClassID = 4 Then
	GSAVersion = 1001
	GrantTypeVersion = 3
End If


If Len(Request.Form("GrantTypeID"))>0 Then
	GrantTypeID = CInt(Request.Form("GrantTypeID"))
ElseIf Len(Request.QueryString("GrantTypeID"))>0 Then 
	GrantTypeID = CInt(Request.QueryString("GrantTypeID"))
Else
	GrantTypeID = 0
End If

If Len(Request.Form("AppID"))>0 Then 
	AppID = CInt(Request.Form("AppID"))
ElseIf Len(Request.QueryString("AppID"))>0 Then 
	AppID = CInt(Request.QueryString("AppID"))
Else
	AppID = 0
End If

If Len(Request.Form("QuestionID")) > 0 Then
	QuestionID = CInt(Request.Form("QuestionID"))
ElseIf Len(Request.QueryString("QuestionID")) > 0 Then
	QuestionID = CInt(Request.QueryString("QuestionID"))
Else
	QuestionID = 0
End If

If GrantClassID=1 Then
	If FiscalYear >= 2024 Then
		TextSectionVersion = 2
		ScoringVersion = 3
	ElseIf FiscalYear = 2022 Then
		TextSectionVersion = 2
		ScoringVersion = 2
	Else
		TextSectionVersion = 1
		ScoringVersion = 1
	End If
ElseIf GrantClassID=4 Then
	TextSectionVersion = 3
	ScoringVersion = 1001
End If

If UserSystemID=1 Then
	PermitEdit = True
	PermitEditScores = True
ElseIf FiscalYear < 2024 Then
	PermitEdit = False
	PermitEditScores = False
ElseIf MVCPAScorer = False Then 
	PermitEdit = False
	PermitEditScores = False
ElseIf FiscalYear=2024 And MVCPAScorer = True Then
	PermitEdit = True
	PermitEditScores = True
ElseIf FiscalYear=2025 And MVCPAScorer = True Then
	PermitEdit = True
	PermitEditScores = True
ElseIf FiscalYear=2026 And MVCPAScorer = True Then
	PermitEdit = True
	PermitEditScores = True
Else
	PermitEdit = False
	PermitEditScores = False
End If
If Debug = True Then
	Response.Write("<pre>PermitEdit=" & PermitEdit & "; PermitEditScores=" & PermitEditScores & "</pre>" & vbCrLf)
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Grant Application Scoring</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm()
	{
		saving = true;
		SaveChanges();
		document.Scoring.submit();
	}

	function validateScore(field, maxScore)
	{
		var score;
		if (field.value == "") {
			return true;
		}
		score = parseInt(field.value)
		if (isNaN(score)) {
			alert("Invalid Score, Score must be numeric!");
			field.focus();
			return false;
		}
		if (maxScore == -1) {
			if (score > 0) {
				alert("Invalid Score, Score cannot be greater than than zero!");
				field.focus();
				return false;
			}
		} else {
			if (score < 0) {
				alert("Invalid Score, Score cannot be less than than zero!");
				field.focus();
				return false;
			}
			if (score > maxScore) {
				alert("Invalid Score, Score cannot be greater than " + maxScore + "!");
				field.focus();
				return false;
			}
		}
		return true;
	}
</script>
</head>
<body style="width: 98%; margin: auto; ">
<h1>Grant Application Scoring</h1>
<form name="Scoring" id="Scoring" method="post" action="ScoringSubmit.asp">
<input type="hidden" name="Changes" id="Changes" />
<input type="hidden" name="TextSectionVersion" id="TextSectionVersion" value="<%=TextSectionVersion %>" />
<input type="hidden" name="ScoringVersion" id="ScoringVersion" value="<%=ScoringVersion %>" />
<%
Response.Write("Fiscal Year: <select name=""FiscalYear"" onchange=""submitForm();"">" & vbCrLf)
Response.Write(vbTab & "<option value=""0"">Select Fiscal Year</option>" & vbCrLf)
For i = 2018 to 2024 Step 2
	If i = FiscalYear Then
		Response.Write(vbTab & "<option value=""" & i & """ selected>" & i & "</option>" & vbCrLf)
	Else
		Response.Write(vbTab & "<option value=""" & i & """>" & i & "</option>" & vbCrLf)
	End If
Next
i = 2025
If i = FiscalYear Then
	Response.Write(vbTab & "<option value=""" & i & """ selected>" & i & "</option>" & vbCrLf)
Else
	Response.Write(vbTab & "<option value=""" & i & """>" & i & "</option>" & vbCrLf)
End If

Response.Write("</select>&nbsp;&nbsp;" & vbCrLf)

Response.Write("Grant Class: <select name=""GrantClassID"" onchange=""submitForm();"">" & vbCrLf)
If GrantClassID = 0 Then
	Response.Write(vbTab & "<option value=""0"" selected>Select Grant Class</option>" & vbCrLf)
Else
	Response.Write(vbTab & "<option value=""0"">Select Grant Class</option>" & vbCrLf)
End If
If GrantClassID = 1 Then
	Response.Write(vbTab & "<option value=""1"" selected>Taskforce Grant</option>" & vbCrLf)
Else
	Response.Write(vbTab & "<option value=""1"">Taskforce Grant</option>" & vbCrLf)
End If
IF GrantClassID = 4 Then
	Response.Write(vbTab & "<option value=""4"" selected>Catalytic Converter Grant</option>" & vbCrLf)
Else
	Response.Write(vbTab & "<option value=""4"">Catalytic Converter Grant</option>" & vbCrLf)
End If

Response.Write("</select>&nbsp;&nbsp;" & vbCrLf)


sql = "SELECT GrantTypeID, GrantType FROM Lookup.GrantType WHERE Version=1 ORDER BY 1"
If Debug = True Then
	Response.Write("<pre>Dropdown query for picking grant type to score: " & vbCrLf & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("Grant Types: <select name=""GrantTypeID"" onchange=""submitForm();"">" & vbCrLf)
	Response.Write(vbTab & "<option value=""0"">All Grant Types</option>" & vbCrLf)
	While rs.EOF = False
		If rs.Fields("GrantTypeID") = GrantTypeID Then
			Response.Write(vbTab & "<option value=""" & rs.Fields("GrantTypeID") & """ selected>" & rs.Fields("GrantType") & "</option>" & vbCrLf)
		Else
			Response.Write(vbTab & "<option value=""" & rs.Fields("GrantTypeID") & """>" & rs.Fields("GrantType") & "</option>" & vbCrLf)
		End If
		rs.MoveNext()
	Wend
	Response.Write("</select>&nbsp;&nbsp;" & vbCrLf)
End If

Response.Write("<br />" & vbCrLf)
sql = "SELECT I.AppID, A.ProgramName, C.GranteeName, A.GrantTypeID, B.GrantType " & vbCrLf & _
	"FROM Application.IDs AS I " & vbCrLf
If GrantClassID=1 Then
	sql = sql & "LEFT JOIN Application.Main AS A ON A.AppID=I.AppID " & vbCrLf
ElseIf GrantClassID = 4 Then
	sql = sql & "LEFT JOIN CC.Application AS A ON A.AppID=I.AppID " & vbCrLf
End If
sql = sql & "LEFT JOIN Lookup.GrantType AS B ON B.GrantTypeID=A.GrantTypeID AND B.Version=" & prepIntegerSQL(TextSectionVersion) & " " & vbCrLf & _
	"LEFT JOIN Grantees AS C ON C.GranteeID=I.GranteeID " & vbCrLf & _
	"WHERE I.GrantClassID=" & prepIntegerSQL(GrantClassID) & " AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND A.SubmitTimestamp IS NOT NULL" & vbCrLf
If GrantTypeID>0 Then
	sql = sql & " AND A.GrantTypeID=" & prepIntegerSQL(GrantTypeID) & " " & vbCrLf
End If
If MVCPARights = False And MVCPAScorer=False Then
	sql = sql & "	AND C.GranteeID IN (SELECT GranteeID FROM [System].[GranteePermissions] WHERE SystemID=" & prepIntegerSQL(UserSystemID) & ")"
End If
sql = sql &	"ORDER BY C.GranteeNameSort, A.GrantTypeID "
If Debug = True Then
	Response.Write("<pre>Dropdown query for picking application to score: " & vbCrLf & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

Response.Write("Applications: <select name=""AppID"" onchange=""submitForm();"">" & vbCrLf)
Response.Write("<option value=""0"">Select Application</option>" & vbCrLf)
If rs.EOF = False Then
	While rs.EOF = False
		If rs.Fields("AppID") = AppID Then
			Response.Write(vbTab & "<option value=""" & rs.Fields("AppID") & """ selected>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</option>" & vbCrLf)
		Else
			Response.Write(vbTab & "<option value=""" & rs.Fields("AppID") & """>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</option>" & vbCrLf)
		End If
		rs.MoveNext()
	Wend
	If AppID=-1 Then
		Response.Write(vbTab & "<option value=""-1"" selected>Show Instructions only</option>" & vbCrLf)
	Else
		Response.Write(vbTab & "<option value=""-1"">Show Instructions only</option>" & vbCrLf)
	End If
End If
Response.Write("</select><br />" & vbCrLf)

sql = "SELECT QuestionID, QuestionName, QuestionNumber FROM Scoring.Questions WHERE QuestionID<>14 AND Version=" & prepIntegerSQL(ScoringVersion) & " ORDER BY QuestionSort"
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("Questions: <select name=""QuestionID"" onchange=""submitForm();"">" & vbCrLf)
	Response.Write(vbTab & "<option value=""0"">All Questions</option>" & vbCrLf)
	While rs.EOF = False
		If rs.Fields("QuestionID") = QuestionID Then
			Response.Write(vbTab & "<option value=""" & rs.Fields("QuestionID") & """ selected>" & rs.Fields("QuestionNumber") & "." & rs.Fields("QuestionName") & "</option>" & vbCrLf)
		Else
			Response.Write(vbTab & "<option value=""" & rs.Fields("QuestionID") & """>" & rs.Fields("QuestionNumber") & "." & rs.Fields("QuestionName") & "</option>" & vbCrLf)
		End If
		rs.MoveNext()
	Wend
	Response.Write("</select>" & vbCrLf)
	If AppID=0 and QuestionID=0 Then
		Response.Write(" Results will not be displayed for all applications and all questions. Only Executive Summary will be shown.")
	End If
End If
%>
<br />

<table style="border-collapse: collapse; ">
<%
If AppID=-1 Then
	sql = "SELECT A.SectionName, A.SectionPoints, A.SectionID, A.Version, " & vbCrLf & _
	"	B.QuestionID, B.QuestionNumber, B.QuestionSort, B.QuestionName, B.Question, B.AdditonalText, " & vbCrLf & _
	"	B. QuestionType, B.QuestionPoints, B.ScoringGroups, B.Objective, B.ShowRatingName, " & vbCrLf & _
	"	B.Group1MaximumPoints, B.Group1MinimumPoints, B.Group2MaximumPoints, B.Group2MinimumPoints, " & vbCrLf & _
	"	B.Group3MaximumPoints, B.Group3MinimumPoints, B.Group4MaximumPoints, B.Group4MinimumPoints, " & vbCrLf & _
	"	B.Group5MaximumPoints, Group5MinimumPoints, B.Group6MaximumPoints, B.Group6MinimumPoints, " & vbCrLf & _
	"	B.Group1Criteria, B.Group2Criteria, B.Group3Criteria, B.Group4Criteria, B.Group5Criteria, " & vbCrLf & _
	"	B.Group6Criteria, -1 AS AppID, NULL AS ProgramName, NULL AS GrantType, " & vbCrLf & _
	"	NULL AS GranteeName, NULL AS Score, NULL as Comments, " & prepIntegerSQL(FiscalYear-2) & " AS HistoricalDataYear, NULL AS BudgetCashMatch " & vbCrLf & _
	"FROM Scoring.vwSections AS A " & vbCrLf & _
	"LEFT JOIN Scoring.Questions AS B ON A.SectionID=B.SectionID AND A.Version=B.Version " & vbCrLf
	If QuestionID>0 Then
		sql = sql & "WHERE B.QuestionID=" & prepIntegerSQL(QuestionID) & " AND A.Version=" & prepIntegerSQL(TextSectionVersion) & " " & vbCrLf
	ELSE
		sql = sql & "WHERE  AND A.Version=" & prepIntegerSQL(ScoringVersion) & " " & vbCrLf
	End If
	sql = sql & "ORDER BY B.QuestionSort"
Else
	sql = "SELECT A.SectionName, A.SectionPoints, A.SectionID, A.Version, " & vbCrLf & _
		"	B.QuestionID, B.QuestionNumber, B.QuestionSort, B.QuestionName, B.Question, B.AdditonalText,B. QuestionType, B.QuestionPoints, B.ScoringGroups, B.Objective, B.ShowRatingName, B.Group1MaximumPoints, B.Group1MinimumPoints, B.Group2MaximumPoints, B.Group2MinimumPoints, B.Group3MaximumPoints, B.Group3MinimumPoints, B.Group4MaximumPoints, B.Group4MinimumPoints, B.Group5MaximumPoints, Group5MinimumPoints, B.Group6MaximumPoints, B.Group6MinimumPoints, B.Group1Criteria, B.Group2Criteria, B.Group3Criteria, B.Group4Criteria, B.Group5Criteria, B.Group6Criteria, C.AppID, C.ProgramName, D.GrantType, " & vbCrLf & _
		"	E.GranteeName, F.Score, F.Comments, C.HistoricalDataYear, C.BudgetCashMatch " & vbCrLf & _
		"FROM Scoring.Sections AS A " & vbCrLf & _
		"LEFT JOIN Scoring.Questions AS B ON A.SectionID=B.SectionID AND A.Version=B.Version " & vbCrLf & _
		"CROSS JOIN Application.IDs AS I " & vbCrLf
	If GrantClassID=1 Then
		sql = sql & "LEFT JOIN Application.Main AS C ON C.AppID=I.AppID " & vbCrLf
	Else
		sql = sql & "LEFT JOIN CC.Application AS C ON C.AppID=I.AppID " & vbCrLf
	End If
	sql = sql & "LEFT JOIN Lookup.GrantType AS D ON D.GrantTypeID=C.GrantTypeID AND D.Version=" & prepIntegerSQL(GrantTypeVersion) & " " & vbCrLf & _
		"LEFT JOIN Grantees AS E ON E.GranteeID=I.GranteeID " & vbCrLf & _
		"LEFT JOIN Scoring.Scores AS F ON F.AppID=I.AppID AND F.QuestionID=B.QuestionID AND F.Version=" & prepIntegerSQL(ScoringVersion) & " AND F.SystemID=" & prepIntegerSQL(UserSystemID) & " " & vbCrLf
	If QuestionID>0 And AppID=0 And GrantTypeID>0 Then
		sql = sql & "WHERE I.GrantClassID=" & prepIntegerSQL(GrantClassID) & " AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND A.Version=" & prepIntegerSQL(ScoringVersion)  & " AND B.QuestionID=" & prepIntegerSQL(QuestionID) & " AND C.GrantTypeID=" & prepIntegerSQL(GrantTypeID) & " AND C.SubmitTimestamp IS NOT NULL " & vbCrLf
	ElseIf QuestionID>0 And AppID=0 Then
		sql = sql & "WHERE I.GrantClassID=" & prepIntegerSQL(GrantClassID) & " AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND A.Version=" & prepIntegerSQL(ScoringVersion)  & " AND B.QuestionID=" & prepIntegerSQL(QuestionID) & " --AND C.SubmitTimestamp IS NOT NULL " & vbCrLf
	ElseIf QuestionID=0 And AppID<>0 Then
		sql = sql & "WHERE I.GrantClassID=" & prepIntegerSQL(GrantClassID) & " AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND A.Version=" & prepIntegerSQL(ScoringVersion)  & " AND C.AppID=" & prepIntegerSQL(AppID) & " --AND C.SubmitTimestamp IS NOT NULL " & vbCrLf
	ElseIf QuestionID>0 And AppID<>0 Then
		sql = sql & "WHERE I.GrantClassID=" & prepIntegerSQL(GrantClassID) & " AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND A.Version=" & prepIntegerSQL(ScoringVersion)  & " AND B.QuestionID=" & prepIntegerSQL(QuestionID) & " AND C.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf
	ElseIf QuestionID=0 And AppID=0 And GrantTypeID>0 Then
		sql = sql & "WHERE I.GrantClassID=" & prepIntegerSQL(GrantClassID) & " AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND A.Version=" & prepIntegerSQL(ScoringVersion)  & " AND B.QuestionID=14 AND C.GrantTypeID=" & GrantTypeID & " --AND C.SubmitTimestamp IS NOT NULL " & vbCrLf
	ElseIf QuestionID=0 And AppID=0 Then
		sql = sql & "WHERE I.GrantClassID=" & prepIntegerSQL(GrantClassID) & " AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND A.Version=" & prepIntegerSQL(ScoringVersion)  & " AND B.QuestionID=14 --AND C.SubmitTimestamp IS NOT NULL " & vbCrLf
	End If
	If MVCPARights = False and AppID<>-1 And MVCPAScorer=False Then
		sql = sql & "	AND C.GranteeID IN (SELECT GranteeID FROM [System].[GranteePermissions] WHERE SystemID=" & prepIntegerSQL(UserSystemID) & ")"
	End If
	sql = sql & "ORDER BY E.GranteeNameSort, C.GrantTypeID, B.QuestionSort, I.AppID "
End If
If Debug = True Then
	Response.Write("<pre>GrantClassID=" & GrantClassID & "</pre>")
	Response.Write("<pre>ScoringVersion=" & ScoringVersion & "</pre>")
	Response.Write("<pre>TextSectionVersion=" & TextSectionVersion & "</pre>")
		Response.Write("<pre>Query to retrieve sections to loop through in display: " & vbCrLf & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
while rs.EOF = False
	ScoringGroups = rs.Fields("ScoringGroups")
	If ScoringGroups = 0 Then 
		Columns = 4
	ElseIf IsNull(ScoringGroups) Then 
		Columns = 4
	Else
		Columns = ScoringGroups
	End If
	If LastSectionID <> rs.Fields("SectionID") Then
		LastSectionID = rs.Fields("SectionID")
		Response.Write("<tr style=""background-color: Aqua; ""><th colspan=""" & Columns & """>" & rs.Fields("SectionName") & " (" & rs.Fields("SectionPoints") & " points)</th></tr>" & vbCrLf)
	End If
	If LastQuestionID=rs.Fields("QuestionID") Then
		' Do nothing
	Else
		LastQuestionID=rs.Fields("QuestionID")
		If IsNull(rs.Fields("QuestionNumber")) = True Then
			Response.Write("<tr style=""background-color: LightCyan; ""><td colspan=""" & Columns & """>" & rs.Fields("Question") & "</td></tr>" & vbCrLf)
		Else
			Response.Write("<tr style=""background-color: LightCyan; ""><td colspan=""" & Columns & """>" & rs.Fields("QuestionNumber") & ". " & rs.Fields("Question") & "</td></tr>" & vbCrLf)
		End If
		If rs.Fields("QuestionPoints")>0 Then
			If ScoringGroups > 0 Then
				Response.Write("<tr style=""background-color: LightCyan; "">" & vbCrLf)
				For i = 1 to ScoringGroups
					If IsNull(rs.Fields("Group" & i & "MinimumPoints")) Then
						Response.Write("<td style=""text-align: center; border: 1px solid black; "">" & rs.Fields("Group" & i & "MaximumPoints") & " points</td>" & vbCrLf)
					ElseIf rs.Fields("Group" & i & "MinimumPoints") = rs.Fields("Group" & i & "MaximumPoints") Then
						Response.Write("<td style=""text-align: center; border: 1px solid black; "">" & rs.Fields("Group" & i & "MaximumPoints") & " points</td>" & vbCrLf)
					Else
						Response.Write("<td style=""text-align: center; border: 1px solid black; "">" & rs.Fields("Group" & i & "MaximumPoints") & "-" & rs.Fields("Group" & i & "MinimumPoints") & " points</td>" & vbCrLf)
					End If
				Next
				Response.Write("</tr>" & vbCrLf)
					Response.Write("<tr style=""background-color: LightCyan; "">" & vbCrLf)
					For i = 1 to ScoringGroups
						Response.Write("<td style=""border: 1px solid black; "">" & rs.Fields("Group" & i & "Criteria") & "</td>" & vbCrLf)
					Next
					Response.Write("</tr>" & vbCrLf)
				If rs.Fields("ShowRatingName") = True Then
					Response.Write("<tr style=""background-color: LightCyan; "">" & vbCrLf)
					Response.Write(vbTab & "<td style=""color: black; text-align: center; border: 1px solid black; "">(Excellent)</td>" & vbCrLf)
					Response.Write(vbTab & "<td style=""color: green; text-align: center; border: 1px solid black; "">(Good)</td>" & vbCrLf)
					Response.Write(vbTab & "<td style=""color: #d0d000; text-align: center; border: 1px solid black; "">(Marginal)</td>" & vbCrLf)
					Response.Write(vbTab & "<td style=""color: red; text-align: center; border: 1px solid black; "">(Poor)</td>" & vbCrLf)
					'Response.Write(vbTab & "<td style=""border: 1px solid black; ""></td>" & vbCrLf)
					Response.Write("</tr>" & vbCrLf)
				End If
			End If
		End If
		Response.Write("<tr><th colspan=""" & Columns & """>&nbsp;</th></tr>" & vbCrLf)
	End If
	If (AppID>0 Or QuestionID>0) And AppID<>-1 Then
		Response.Write("<tr style=""background-color: wheat; "">" & vbCrLf)
		Response.Write("<td colspan=""" & (Columns+1) & """ style=""font-weight: bold; "">" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</td>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
		If rs.Fields("QuestionPoints")>0 Or rs.Fields("QuestionType") = "ScoreOnly" Then
			Response.Write("<tr style=""vertical-align: top; background-color: wheat; ""><td colspan=""" & (Columns+1) & """>" & vbCrLf)
			If PermitEdit = True Then
				Response.Write("<textarea name=""Comments_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ id=""Comments_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ cols=""100"" rows=""2"" style=""background-color: Bisque; "">" & rs.Fields("Comments") & "</textarea>" & vbCrLf)
			Else
				Response.Write("<input type=""hidden"" name=""Comments_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ id=""Comments_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ value=""" & prepStringWeb(rs.Fields("Comments")) & """><div style=""background-color: wheat; "">" & rs.Fields("Comments") & "</div>" & vbCrLf)
			End If
			If rs.Fields("Objective") = True Or PermitEdit=False Or PermitEditScores=False Then
				Response.Write(vbTab & " Score: <input type=""text"" name=""score_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ id=""score_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ size=""2"" value=""" & rs.Fields("Score") & """ style=""background-color: wheat; text-align: right; border: none"" readonly tabindex=""-1"">" & vbCrLf)
			ElseIf rs.Fields("Objective") = True Or PermitEdit=False Or PermitEditScores=False Then
				Response.Write(vbTab & " Score: <input type=""text"" name=""score_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ id=""score_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ size=""2"" value=""" & rs.Fields("Score") & """ style=""background-color: wheat; text-align: right; border: none"" readonly tabindex=""-1"">" & vbCrLf)
			Else
				Response.Write(vbTab & " Score: <input type=""text"" name=""score_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ id=""score_" & rs.Fields("AppID") & "_" & rs.Fields("QuestionID") & """ size=""2"" value=""" & rs.Fields("Score") & """ style=""background-color: Bisque; text-align: right;"" onchange=""validateScore(this, " & rs.Fields("QuestionPoints") & ");"">" & vbCrLf)
			End If
			Response.Write(vbTab & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		Response.Write("<tr><td colspan=""" & Columns & """>" & vbCrLf)

		If rs.Fields("QuestionType") = "Budget" Then
			'Response.Write("<pre>Budget</pre>")
			ShowBudget rs.Fields("AppID")
		ElseIf rs.Fields("QuestionType") = "BudgetDetail" Then
			'Response.Write("<pre>BudgetDetail</pre>")
			ShowBudget rs.Fields("AppID")
			ShowBudgetDetail rs.Fields("AppID")
			ShowBudgetNarrative rs.Fields("AppID")
		ElseIf rs.Fields("QuestionType") = "BudgetDetailStatistic" Then
			'Response.Write("<pre>BudgetDetailStatistic</pre>")
			ShowBudget rs.Fields("AppID")
			ShowBudgetDetail rs.Fields("AppID")
			ShowStatistics2 rs.Fields("AppID"), rs.Fields("HistoricalDataYear")
			ShowBudgetNarrative rs.Fields("AppID")
		ElseIf rs.Fields("QuestionType") = "Matching" Then
			ShowSectionText rs.Fields("AppID"), rs.Fields("QuestionID")
			ShowBudget rs.Fields("AppID")
			ShowMatching rs.Fields("AppID")
			ShowBudgetNarrative rs.Fields("AppID")
		ElseIf rs.Fields("QuestionType") = "SectionText" Then
			'Response.Write("<pre>SectionText</pre>")
			ShowSectionText rs.Fields("AppID"), rs.Fields("QuestionID")
		ElseIf rs.Fields("QuestionType") = "Informational" Then
			'Response.Write("<pre>Informational</pre>")
			ShowSectionText rs.Fields("AppID"), rs.Fields("QuestionID")
			ShowBudgetCashMatch rs.Fields("BudgetCashMatch")
		ElseIf rs.Fields("QuestionType") = "Calculate" Then
			' Do nothing. This is a prepared value.
		ElseIf rs.Fields("QuestionType") = "BMVStatistics" Then
			BMVStatistics rs.Fields("AppID")
		ElseIf rs.Fields("QuestionType") = "Statistics2" Then
			ShowStatistics2 rs.Fields("AppID"), rs.Fields("HistoricalDataYear")
			ShowSectionText rs.Fields("AppID"), rs.Fields("QuestionID")
		ElseIf rs.Fields("QuestionType") = "Statistics3" Then
			ShowStatistics2 rs.Fields("AppID"), rs.Fields("HistoricalDataYear")
			ShowSectionText rs.Fields("AppID"), rs.Fields("QuestionID")
			ShowIndexCrimes2024 rs.Fields("AppID"), rs.Fields("QuestionID")
		ElseIf rs.Fields("QuestionType") = "MVTStatistics" Then
			MVTStatistics rs.Fields("AppID")
		ElseIf rs.Fields("QuestionType") = "PreviousPerformance2024" Then
			ShowPreviousPerformance2024 rs.Fields("AppID"), rs.Fields("QuestionID")
		ElseIf rs.Fields("QuestionType") = "ScoreOnly" Then
			ShowScoreOnly rs.Fields("AppID"), rs.Fields("QuestionID")
		End If

		Response.Write("</td></tr>" & vbCrLf)
		Response.Write("<tr><td colspan=""" & Columns & """>&nbsp;</td></tr>" & vbCrLf)
	ElseIf AppID=-1 Then
		ShowTextSectionQuestions rs.Fields("QuestionID")
	End If
	rs.MoveNext
Wend
%>
</table>
<%
If AppID > 0 Then
	Response.Write("<div style=""width: 500px; margin: auto; ""><div style=""text-align: center; font-weight: bold; ""><a href=""/Application/GSADisplay.asp?AppID=" & prepIntegerSQL(AppID) & """ target=""_blank"">Link</a> to Goals, Targets, and Activities Targets for Application</div>" & vbCrLf)
	Response.Write("<br />" & vbCrLf)
End If

If AppID > 0 Or QuestionID>0 Then
	Dim DocumentFolder, fso, folder, files, file
	DocumentFolder = Application("DocumentRoot") & "\Application\" & AppID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemOBject")
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		If files.count>0 Then 
			Response.Write("<div style=""width: 500px; margin: auto; ""><div style=""text-align: center; font-weight: bold; "">Current Documents in folder</div>" & vbCrLf)
		For Each file in files
				Response.Write("<a href=""../Documents/Application/" & AppID & "/" & file.Name & _
					""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
		Next
			Response.Write("<br /></div>" & vbCrLf)
		End If
	End If
End If
%>
<div style="text-align: center; margin: auto; ">
	<input type="button" value="save" onclick="return submitForm();" />
	<input type="button" value="close" onclick="window.close();" /></div>
</form>
<script src="../includes/formchanges.js"></script>
<script type="text/javascript">
	var saving = false;
	var form = document.getElementById("Scoring");

	// form being updated
	form.onsubmit = function () { saving = true; };

	// form not saved warning
	window.onunload = function() {
		if (!saving) {
			var f = FormChanges(form);
			if (f.length > 0) 
			{
				if (window.confirm("Your form updates have not be saved. Do you wish to continue without saving?"))
					return true;
				else
					submitForm();
					return false;
			}
		}
	};

	// show changed messages
	function DetectChanges()
	{
		var f = FormChanges(form), msg = "";
		for (var e = 0, el = f.length; e < el; e++) msg += "\n" + f[e].id;
		alert((msg ? "Elements changed:" : "No changes made.") + msg);
	}

	// Save changes
	function SaveChanges()
	{
		var f = FormChanges(form), msg = "";
		for (var e = 0, el = f.length; e < el; e++) msg += f[e].id + "\n";
		if (msg.length > 0)
			msg = msg.substring(0, msg.length - 1);
		document.Scoring.Changes.value = msg;
	}

</script>
</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"--><%

Sub BMVStatistics(vAppID)
	dim vsql, vrs 

	vsql = "SELECT A.AppID, HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, " & vbCrLf & _
		"	LarcenyFromMVParts1, LarcenyFromMVParts2, LarcenyFromMVParts3, LarcenyJurisdiction, " & vbCrLf & _
		"	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, " & vbCrLf & _
		"	DataProblems " & vbCrLf & _
		"FROM Application.Main AS A " & vbCrLf & _
		"WHERE A.AppID=" & prepIntegerSQL(vAppID)
	Set vrs=Con.Execute(vsql)
	Response.Write("<table style=""width: 550px; margin: auto; "">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr><th colspan=""4"">Statistics to Support Grant Problem Statement</th></tr>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("	<td style=""text-align: center; "">Use UCR data</td>" & vbCrLf)
	Response.Write("	<th>" & (vrs.Fields("HistoricalDataYear")-2) & "</th>" & vbCrLf)
	Response.Write("	<th>" & (vrs.Fields("HistoricalDataYear")-1) & "</th>" & vbCrLf)
	Response.Write("	<th>" & (vrs.Fields("HistoricalDataYear")) & "</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("	<td style=""font-weight: bold; "">Larceny from a motor vehicle</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("LarcenyFromMV1")) & "</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("LarcenyFromMV2")) & "</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("LarcenyFromMV3")) & "</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("	<td style=""font-weight: bold; "">Larceny from a motor vehicle - Parts</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("LarcenyFromMVParts1")) & "</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("LarcenyFromMVParts2")) & "</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("LarcenyFromMVParts3")) & "</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("	<td style=""font-weight: bold; "">Jurisdictions included in totals</td>" & vbCrLf)
	Response.Write("	<td colspan=""3"" style=""text-align: center; "">" & vbCrLf)

	If vrs.Fields("LarcenyJurisdiction")= 0 Then
		Response.Write("Select Jurisdiction")
	ElseIf vrs.Fields("LarcenyJurisdiction")=1 Then
		Response.Write("Statistics for Taskforce Only")
	ElseIf vrs.Fields("LarcenyJurisdiction")=2 Then
		Response.Write("Statistics for Area of Jurisdiction")
	ElseIf vrs.Fields("LarcenyJurisdiction")=3 Then
		Response.Write("Statistics a combination of Taskforce and Jurisdiction")
	End If

	Response.Write("</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)

	Response.Write("<br>" & vbCrLf)

	vsql = "SELECT A.TextSectionID, A.Section, A.SubSection, A.Question, B.SectionText " & vbCrLf & _
		"FROM Lookup.TextSections AS A " & vbCrLf & _
		"LEFT JOIN Application.SectionText AS B ON A.TextSectionID=B.TextSectionID AND B.AppID=" & prepIntegerSQL(vAppID) & vbCrLf & _
		"WHERE A.TextSectionID IN (2, 5) AND A.Version=" & TextSectionVersion & " " & vbCrLf & _
		"ORDER BY A.Section, A.SubSection "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	while vrs.EOF = False 
		Response.Write("<div style=""font-weight: bold; "">" & vrs.Fields("Section") & "." & vrs.Fields("SubSection") & ": " & vrs.Fields("Question") & "</div>" & vbCrLf)
		Response.Write(textarea2html(vrs.Fields("SectionText")))
		vrs.MoveNext()
	Wend
End Sub

Sub MVTStatistics(vAppID)
	dim vsql, vrs 

	vsql = "SELECT A.AppID, HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, " & vbCrLf & _
		"	LarcenyFromMVParts1, LarcenyFromMVParts2, LarcenyFromMVParts3, LarcenyJurisdiction, " & vbCrLf & _
		"	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, " & vbCrLf & _
		"	DataProblems " & vbCrLf & _
		"FROM Application.Main AS A " & vbCrLf & _
		"WHERE A.AppID=" & prepIntegerSQL(vAppID)
	Set vrs=Con.Execute(vsql)
	Response.Write("<table style=""width: 550px; margin: auto; "">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr><th colspan=""4"">Statistics to Support Grant Problem Statement</th></tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("	<td style=""font-weight: bold; "">Theft of a Motor Vehicle</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("MVT1")) & "</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("MVT2")) & "</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("MVT3")) & "</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("	<td style=""font-weight: bold; "">Recoveries of Motor Vehicles</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("RecoveryMVT1")) & "</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("RecoveryMVT2")) & "</td>" & vbCrLf)
	Response.Write("	<td style=""text-align: right; "">" & prepIntegerWeb(vrs.Fields("RecoveryMVT3")) & "</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("	<td style=""font-weight: bold; "">Jurisdictions included in totals</td>" & vbCrLf)
	Response.Write("	<td colspan=""3"" style=""text-align: center; "">" & vbCrLf)

	If vrs.Fields("MVTJurisdiction")= 0 Then
		Response.Write("Select Jurisdiction")
	ElseIf vrs.Fields("MVTJurisdiction")=1 Then
		Response.Write("Statistics for Taskforce Only")
	ElseIf vrs.Fields("MVTJurisdiction")=2 Then
		Response.Write("Statistics for Area of Jurisdiction")
	ElseIf vrs.Fields("MVTJurisdiction")=3 Then
		Response.Write("Statistics a combination of Taskforce and Jurisdiction")
	End If

	Response.Write("</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)

	Response.Write("<br>" & vbCrLf)

	vsql = "SELECT A.TextSectionID, A.Section, A.SubSection, A.Question, B.SectionText " & vbCrLf & _
		"FROM Lookup.TextSections AS A " & vbCrLf & _
		"LEFT JOIN Application.SectionText AS B ON A.TextSectionID=B.TextSectionID AND B.AppID=" & prepIntegerSQL(vAppID) & vbCrLf & _
		"WHERE A.TextSectionID IN (3, 5) AND A.Version=" & TextSectionVersion & " " & vbCrLf & _
		"ORDER BY A.Section, A.SubSection "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	while vrs.EOF = False 
		Response.Write("<div style=""font-weight: bold; "">" & vrs.Fields("Section") & "." & vrs.Fields("SubSection") & ": " & vrs.Fields("Question") & "</div>" & vbCrLf)
		Response.Write(textarea2html(vrs.Fields("SectionText")))
		vrs.MoveNext()
	Wend
End Sub

Sub ShowStatistics2(vAppID, vHistoricalDataYear)
	Dim vrs, vsql
	Response.Write("<div style=""text-align: center; font-weight: bold; "">Statistics to Support Grant Problem Statement</div>" & vbCrLf)
	Response.Write("<table class=""bordertable"">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("	<th>Reported Cases</th>" & vbCrLf)
	Response.Write("	<th colspan=""3"" style=""border: solid black thin; "">" & (vHistoricalDataYear-1) &"</th>" & vbCrLf)
	Response.Write("	<th colspan=""3"" style=""border: solid black thin; "">" & (vHistoricalDataYear) & "</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
	Response.Write("	<th style=""width: 175px; "">Jurisdiction</th>" & vbCrLf)
	Response.Write("	<th style=""width: 115px; "">Motor Vehicle Theft<br />(MVT)</th>" & vbCrLf)
	Response.Write("	<th style=""width: 115px; "" title=""Burglary from Motor Vehicle including theft of parts"">Burglary from Motor Vehicle<br />(BMV)</th>" & vbCrLf)
	Response.Write("	<th style=""width: 115px; "">Fraud-Related Motor Vehicle Crime<br />(FRMVC)</th>" & vbCrLf)
	Response.Write("	<th style=""width: 115px; "">Motor Vehicle Theft<br />(MVT)</th>" & vbCrLf)
	Response.Write("	<th style=""width: 115px; "" title=""Burglary from Motor Vehicle including theft of parts"">Burglary from Motor Vehicle<br />(BMV)</th>" & vbCrLf)
	Response.Write("	<th style=""width: 115px; "">Fraud-Related Motor Vehicle Crime<br />(FRMVC)</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)

	vsql = "WITH CTE AS ( " & vbCrLf & _
		"SELECT A.GranteeID, REPLACE(A.GranteeName,'City of ','') AS Grantee_Name, B.ProgramName,  " & vbCrLf & _
		"	C.StatisticsID, B.AppID, Jurisdiction, C.MVT1, C.BMV1, C.FRMVC1, C.MVT2, C.BMV2, C.FRMVC2 " & vbCrLf & _
		"FROM Grantees AS A " & vbCrLf & _
		"LEFT JOIN Application.IDs AS I ON I.GranteeID=A.GranteeID " & vbCrLf & _
		"JOIN Application.Main AS B ON B.AppID=I.AppID " & vbCrLf & _
		"JOIN Application.[Statistics] AS C ON C.AppID=I.AppID " & vbCrLf & _
		"WHERE I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND I.AppID=" & prepIntegerSQL(vAppID) & " " & vbCrLf & _
		") " & vbCrLf & _
		"SELECT * FROM CTE " & vbCrLf & _
		"UNION " & vbCrLf & _
		"SELECT GranteeID, Grantee_Name, ProgramName,  " & vbCrLf & _
		"	999999 AS StatisticsID, AppID, 'Total' AS Jurisdiction,  " & vbCrLf & _
		"	SUM(MVT1) AS MVT1, SUM(BMV1) AS BMV1, SUM(FRMVC1) AS FRMV1, SUM(MVT2) AS MVT2, SUM(BMV2) AS BMV2, SUM(FRMVC2) AS FRMVC2 " & vbCrLf & _
		"FROM CTE " & vbCrLf & _
		"GROUP BY GranteeID, Grantee_Name, ProgramName, AppID " & vbCrLf & _
		"ORDER BY Grantee_Name, AppID, StatisticsID "	
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = True Then
		Response.Write(vbTab & "<tr><td colspan=""7"">&nbsp;</td></tr>" & vbCrLf)
		Response.Write(vbTab & "<tr><th colspan=""7""><i>No Statistical Data has been entered yet.</i></th></tr>" & vbCrLf)
	Else
		While vrs.EOF = False
			Response.Write(vbtab & "<tr style=""vertical-align: top; "">" & vbCrLf)
			Response.Write(vbTab & vbTab & "<td>" & vrs.Fields("Jurisdiction") & "</td>" & vbCrLf)
			Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(vrs.Fields("MVT1")) & "</td>" & vbCrLf)
			Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(vrs.Fields("BMV1")) & "</td>" & vbCrLf)
			Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(vrs.Fields("FRMVC1")) & "</td>" & vbCrLf)
			Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(vrs.Fields("MVT2")) & "</td>" & vbCrLf)
			Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(vrs.Fields("BMV2")) & "</td>" & vbCrLf)
			Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(vrs.Fields("FRMVC2")) & "</td>" & vbCrLf)
			Response.Write(vbtab & "<tr>" & vbCrLf)
			vrs.MoveNext
		Wend
	End If
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
	Response.Write("<br />" & vbCrLf)
End Sub

Sub ShowSectionText(vAppID, vQuestionID)
	Dim vrs, vsql

	vsql = "SELECT A.TextSectionID, A.Section, A.SubSection, A.Question, B.SectionText " & vbCrLf & _
		"FROM Lookup.TextSections AS A " & vbCrLf & _
		"LEFT JOIN Application.SectionText AS B ON A.TextSectionID=B.TextSectionID AND B.AppID=" & prepIntegerSQL(vAppID) & vbCrLf & _
		"LEFT JOIN Scoring.QuestionTextSectionID AS C ON C.TextSectionID=A.TextSectionID --AND C.Version=A.Version " & vbCrLf & _
		"WHERE C.QuestionID=" & prepIntegerSQL(vQuestionID) & " AND A.Version=" & TextSectionVersion & " " & vbCrLf & _
		"ORDER BY A.Section, A.SubSection "
	If Debug = True Then
		Response.Write("<pre>ShowSectionText(vAppID=" & vAppID & ", vQuestionID=" & vQuestionID & "): " & vbCrLf & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	While vrs.EOF = False
		Response.Write("<div style=""font-weight: bold; "">" & vrs.Fields("Section") & "." & vrs.Fields("SubSection") & ": " & vrs.Fields("Question") & "</div><br />" & vbCrLf)
		Response.Write(textarea2html(vrs.Fields("SectionText")) & "<br /><br />")
		vrs.MoveNext()
	Wend
End Sub

Sub ShowPreviousPerformance2024(vAppID, vQuestionID)
	If Debug = True Then
		Response.Write("<pre>ShowSectionText(vAppID=" & vAppID & ", vQuestionID=" & vQuestionID & "): " & vbCrLf & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Response.Write("<a href=""PreviousPerformance2024.xlsx"" target=""_blank"">Open Previous Performance Spreadsheet for review</a><br /><br />")
End Sub

Sub ShowIndexCrimes2024(vAppID, vQuestionID)
	If Debug = True Then
		Response.Write("<pre>ShowSectionText(vAppID=" & vAppID & ", vQuestionID=" & vQuestionID & "): " & vbCrLf & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Response.Write("<a href=""IndexCrimes2024.xlsx"" target=""_blank"">Open Index Crimes for 2022 Spreadsheet for review</a><br /><br />")
End Sub

Sub ShowScoreOnly(vAppID, vQuestionID)
	If Debug = True Then
		Response.Write("<pre>ShowSectionText(vAppID=" & vAppID & ", vQuestionID=" & vQuestionID & "): " & vbCrLf & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
End Sub

Sub ShowBudget(vAppID)
	Dim vrs, vsql, TotalMVCPAFunds, TotalCashMatch, GrandTotal, TotalInKindMatch, PctMVCPA, PctCashMatch
	vsql = "SELECT SUM(MVCPAFunds) AS TotalMVCPAFunds, SUM(CashMatch) AS TotalCashMatch, SUM(LineTotal) AS GrandTotal, SUM(InKindMatch) AS TotalInKindMatch " & vbCrLf & _
		"FROM Application.BudgetDetails " & vbCrLf & _
		"WHERE AppID=" & prepIntegerSQL(vAppID)
	If Debug = True Then
		Response.Write("<pre>ShowBudgetDetail, get Totals(vAppID=" & vAppID & "): " & vbCrLf & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		TotalMVCPAFunds = vrs.Fields("TotalMVCPAFunds")
		TotalCashMatch = vrs.Fields("TotalCashMatch")
		GrandTotal = vrs.Fields("GrandTotal")
		TotalInKindMatch = vrs.Fields("TotalInKindMatch")
		PctMVCPA = 100.0*TotalMVCPAFunds / GrandTotal
		PctCashMatch = 100.0*TotalCashMatch / TotalMVCPAFunds
	Else
		Response.Write("Error retrieving Budget Totals")
		Response.End
	End If
	Response.Write("<table style=""width: 896px; margin: auto"">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom"">" & vbCrLf)
	Response.Write("<th>Budget Category</th>" & vbCrLf)
	Response.Write("<th>MVCPA<br />Expenditures</th>" & vbCrLf)
	Response.Write("<th>Cash<br />Match<br />Expenditures</th>" & vbCrLf)
	Response.Write("<th>Total<br />Expenditures</th>" & vbCrLf)
	Response.Write("<th>In-Kind<br />Match</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)

	vsql = "SELECT ISNULL(A.BudgetCategoryID,99) AS BudgetCategoryID, ISNULL(A.BudgetCategory, 'Total') As BudgetCategory, " & vbCrLf & _
		"	SUM(LineTotal) AS LineTotal, SUM(MVCPAFunds) AS [MVCPAFunds], " & vbCrLf & _
		"	SUM(CashMatch) AS [CashMatch], SUM(InKindMatch) AS [InKindMatch] " & vbCrLf & _
		"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
		"LEFT JOIN Application.BudgetDetails AS B ON A.BudgetCategoryID=B.BudgetCategoryID AND B.AppID=" & _
			prepIntegerSQL(vAppID) & " " & vbCrLf & _
		"GROUP BY GROUPING SETS ((A.BudgetCategoryID,A.BudgetCategory),()) " & vbCrLf & _
		"ORDER BY ISNULL(A.BudgetCategoryID,99) "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	While vrs.EOF = False
		Response.Write(vbTab & "<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td>" & vrs.Fields("BudgetCategory") & "</td>" & vbCrLf)  
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(vrs.Fields("MVCPAFunds"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(vrs.Fields("CashMatch"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(vrs.Fields("LineTotal"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(vrs.Fields("InKindMatch"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>")
		vrs.MoveNext
	Wend
	If TotalMVCPAFunds>0 Then
		Response.Write("<tr><td style=""text-align: center;"">Cash Match Percentage</td><td style=""text-align: right; ""><!--" & prepNumberWeb(PctMVCPA, 2) & _
			"%--></td><td style=""text-align: right; "">" & prepNumberWeb(PctCashMatch, 2) & "%</td><td></td><td></td></tr>" & vbCrLf)
	End If

	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
	Response.Write("<br />" & vbCrLf)
End Sub

Sub ShowBudgetDetail(vAppID)
	Dim vrs, vsql, vLastCategory
	vsql = "SELECT B.BudgetItemID, A.BudgetCategoryID, A.BudgetCategory, B.Description, SubCategory, PctTime, LineTotal, MVCPAFunds, CashMatch, InKindMatch " & vbCrLf & _
		"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
		"LEFT JOIN Application.BudgetDetails AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND AppID=" & prepIntegerSQL(vAppID) & " " & vbCrLf & _
		"LEFT JOIN Lookup.BudgetSubcategories AS C ON C.BudgetCategoryID=B.BudgetCategoryID AND C.SubCategoryID=B.SubCategoryID " & vbCrLf & _
		"UNION " & vbCrLf & _
		"SELECT 2147483647 AS BudgetItemID, A.BudgetCategoryID, A.BudgetCategory, null AS PctTime, 'Total ' + A.BudgetCategory AS Description, null, SUM(LineTotal) AS LineTotal, SUM(MVCPAFunds) AS MVCPAFunds, SUM(CashMatch) AS CashMatch, Sum(InKindMatch) AS InKindMatch " & vbCrLf & _
		"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
		"LEFT JOIN Application.BudgetDetails AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND AppID=" & prepIntegerSQL(vAppID) & " " & vbCrLf & _
		"GROUP BY A.BudgetCategoryID, A.BudgetCategory" & vbCrLf & _
		"ORDER BY 2, 1 "
	vLastCategory=0
	If Debug = True Then
		Response.Write("<pre>ShowBudgetDetail(vAppID=" & vAppID & "): " & vbCrLf & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		Response.Write("<table style=""margin: auto; "">" & vbCrLf)
		Response.Write("<thead><tr>" & vbCrLf)
		Response.Write(vbTab & "<th>Description</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Subcategory</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Pct Time</th>" & vbCrLf)
		Response.Write(vbTab & "<th>MVCPA Funds</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Cash Match</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Total</th>" & vbCrLf)
		Response.Write(vbTab & "<th>In-Kind Match</th>" & vbCrLf)
		Response.Write("</tr></thead>" & vbCrLf)
		Response.Write("<tbody>" & vbCrLf)
		While vrs.EOF = False
			If vLastCategory <> vrs.Fields("BudgetCategoryID") Then
				vLastCategory = vrs.Fields("BudgetCategoryID")
				Response.Write("<tr><td colspan=""6"">&nbsp;</td></tr>" & vbCrLf)
				Response.Write("<tr><th colspan=""6"">" & vrs.Fields("BudgetCategory") & "</th></tr>" & vbCrLf)
			End If
			Response.Write("<tr>" & vbCrLf)
			Response.Write(vbTab & "<td>" & vrs.Fields("Description") & "</td>")
			Response.Write(vbTab & "<td>" & vrs.Fields("SubCategory") & "</td>")
			If vrs.Fields("BudgetCategoryID")=1 And IsNull(vrs.Fields("PctTime")) = False Then
				Response.Write(vbTab & "<td style=""text-align: right; "">" & prepNumberWeb(vrs.Fields("PctTime"),2) & "%</td>")
			Else
				Response.Write(vbTab & "<td></td>")
			End If
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("MVCPAFunds"), RoundCurrency) & "</td>")
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("CashMatch"), RoundCurrency) & "</td>")
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("LineTotal"), RoundCurrency) & "</td>")
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("InKindMatch"), RoundCurrency) & "</td>")
			Response.Write("</tr>" & vbCrLf)
			vrs.MoveNext
		Wend
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
		Response.Write("<br />" & vbCrLf)
	End If
End Sub

Sub ShowBudgetNarrative(vAppID)
	Dim vrs, vsql
	vsql = "SELECT A.BudgetCategoryID, A.BudgetCategory, REPLACE(B.Narrative,CHAR(13)+CHAR(10),'<br />') AS Narrative " & vbCrLf & _
		"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
		"JOIN Application.BudgetDetailNarrative AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND B.AppID=" & prepIntegerSQL(vAppID) & " " & vbCrLF & _
		"ORDER BY A.BudgetCategoryID "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		Response.Write("<br /><table>" & vbCrLf)
		Response.Write("<thead><tr><th>Budget Narrative</th></tr></thead>" & vbCrLf)
		While vrs.EOF = False
			Response.Write("<tr><td><b>" & vrs.Fields("BudgetCategory") & "</b>: " & textarea2html(vrs.Fields("Narrative")) & "</td></tr>" & vbCrLf)
			vrs.MoveNext()
		Wend
		Response.Write("</table>" & vbCrLf)
		Response.Write("<br />" & vbCrLf)
	End If
End Sub

Sub ShowMatching(vAppID)
	Dim vsql, vrs, TotalMVCPAFunds, TotalCashMatch, GrandTotal, TotalInKindMatch, PctMVCPA, PctCashMatch
	vsql = "SELECT SUM(MVCPAFunds) AS TotalMVCPAFunds, SUM(CashMatch) AS TotalCashMatch, SUM(LineTotal) AS GrandTotal, SUM(InKindMatch) AS TotalInKindMatch " & vbCrLf & _
		"FROM Application.BudgetDetails " & vbCrLf & _
		"WHERE AppID=" & prepIntegerSQL(vAppID)
	If Debug = True Then
		Response.Write("<pre>ShowBudgetDetail, get Totals(vAppID=" & vAppID & "): " & vbCrLf & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		TotalMVCPAFunds = vrs.Fields("TotalMVCPAFunds")
		TotalCashMatch = vrs.Fields("TotalCashMatch")
		GrandTotal = vrs.Fields("GrandTotal")
		TotalInKindMatch = vrs.Fields("TotalInKindMatch")
		PctMVCPA = 100.0*TotalMVCPAFunds / GrandTotal
		PctCashMatch = 100.0*TotalCashMatch / TotalMVCPAFunds
	Else
		Response.Write("Error retrieving Budget Totals")
		Response.End
	End If
	vsql = "SELECT A.Source, B.MatchSource, A.Amount " & vbCrLf & _
		"FROM Application.Matches AS A " & vbCrLf & _
		"LEFT JOIN Lookup.MatchSources AS B ON B.MatchSourceID=A.MatchSourceID " & vbCrLf & _
		"WHERE A.MatchTypeID=1 AND A.AppID=" & prepIntegerSQL(vAppID) & " " & vbCrLf & _
		"ORDER BY A.MatchID "
	If Debug = True Then
		Response.Write("<pre>ShowMatching(AppID=" & vAppID & "): " & vbCrLf & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		Response.Write(vbTab & "<div style=""text-align: center"">Cash Match</div>" & vbCrLf)  
		Response.Write("<table style=""margin: auto; width: 650px; "">" & vbCrLf)
		Response.Write("<thead>" & vbCrLf)
		Response.Write(vbTab & "<tr><th colspan=""3"">Source of Cash Match</th></tr>" & vbCrLf)
		Response.Write("</thead>" & vbCrLf)
		Response.Write("<tbody>" & vbCrLf)
		While vrs.EOF = False
			Response.Write("<tr>" & vbCrLf)
			Response.Write(vbTab & "<td>" & vrs.Fields("Source") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & vrs.Fields("MatchSource") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("Amount"), RoundCurrency) & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
			vrs.MoveNext
		Wend
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("<tfoot><tr><td style=""font-weight: bold; "">Total Cash Match</td><td></td><td style=""text-align: right; "">" & prepCurrencyWebRound(TotalCashMatch, RoundCurrency) & "</td></tr></tfoot>")
		Response.Write("</table>" & vbCrLf)
	End If

	vsql = "SELECT A.Source, B.MatchSource, A.Amount " & vbCrLf & _
		"FROM Application.Matches AS A " & vbCrLf & _
		"LEFT JOIN Lookup.MatchSources AS B ON B.MatchSourceID=A.MatchSourceID " & vbCrLf & _
		"WHERE A.MatchTypeID=2 AND A.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
		"ORDER BY A.MatchID "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		Response.Write("<br />" & vbCrLf)
		Response.Write("<div style=""text-align: center"">In-Kind Match</div>" & vbCrLf)  
		Response.Write("<table style=""margin: auto; width: 650px; "">" & vbCrLf)
		Response.Write("<thead><tr><th colspan=""3"">Source of In-Kind Match</th></tr></thead>" & vbCrLf)
		Response.Write("<tbody>" & vbCrLf)
		While vrs.EOF = False
			Response.Write("<tr>" & vbCrLf)
			Response.Write(vbTab & "<td>" & vrs.Fields("Source") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & vrs.Fields("MatchSource") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("Amount"), RoundCurrency) & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
			vrs.MoveNext
		Wend
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("<tfoot><tr><td style=""font-weight: bold; "">Total In-Kind Match</td><td></td><td style=""text-align: right; "">" & prepCurrencyWebRound(TotalInKindMatch, RoundCurrency) & "</td></tr></tfoot>")
		Response.Write("</table>" & vbCrLf)
	End If
End Sub

Sub ShowTextSectionQuestions(vQuestionID)
	Dim vrs, vsql
	vsql = "SELECT A.QuestionID, B.TextSectionID, B.Section, B.SubSection, B.Question " & vbCrLf & _
	"FROM Scoring.QuestionTextSectionID AS A " & vbCrLf & _
	"JOIN Lookup.TextSections AS B ON B.TextSectionID=A.TextSectionID AND B.Version=A.Version " & vbCrLf & _
	"WHERE A.Version=" & prepIntegerSQL(TextSectionVersion) & " AND A.QuestionID=" & prepIntegerSQL(vQuestionID)
	If Debug = True Then
		Response.Write("<pre>ShowTextSectionQuestions(" & vQuestionID & "): " & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		While vrs.EOF = False
			Response.Write("<tr><td colspan=""4"">Show <b>" & vrs.Fields("Section") & "." & vrs.Fields("SubSection") & ": " & vrs.Fields("Question") & "</b></td></tr>" & vbCrLf)
			vrs.MoveNext
		Wend
		Response.Write("<tr><td colspan=""4"">&nbsp;</td></tr>" & vbCrLf)
	End If
End Sub

Sub ShowBudgetCashMatch(vBudgetCashMatch)
	If IsNull(vBudgetCashMatch) Then
		Response.Write("<tr><td colspan=""4""><b>Budget Entry Option</b>: Enter MVCPA and Cash Match Amounts</td></tr>" & vbCrLf)
	Else
		Response.Write("<tr><td colspan=""4""><b>Budget Entry Option</b>: Let system calculate MVCPA Funds and Cash Match using a Match Percentage of " & formatPercent(vBudgetCashMatch/100.0) & ".</td></tr>" & vbCrLf)
	End If
End Sub

Function outputGSA(vAppID)
	Response.Write("<table style=""margin: auto""><thead><tr><th>ID</th><th>Activity</th><th>Measure</th><th>Target</th></tr></thead>" & vbCrLf)
	Dim vrs, vsql, LastMandatory, LastGoal, LastStrategy
	vsql = "SELECT G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
		"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
		"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
		"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.ResponseTypeID, " & vbCrLf & _
		"	T.IntegerResponse, T.DecimalResponse " & vbCrLf & _
		"FROM Lookup.Goals AS G " & vbCrLf & _
		"LEFT JOIN Lookup.Strategies AS S ON S.GoalID=G.GoalID AND S.Version=G.Version " & vbCrLf & _
		"LEFT JOIN Lookup.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID AND S.Version=G.Version " & vbCrLf & _
		"LEFT JOIN Application.GSATargets AS T ON T.AppID=" & prepIntegerSQL(vAppID) & " AND T.GoalID=G.GoalID AND T.StrategyID=S.StrategyID AND T.ActivityID=A.ActivityID AND T.MeasureID=A.MeasureID AND T.Version=A.Version " & vbCrLf & _
		"WHERE G.Version=" & GSAVersion & " AND T.IntegerResponse IS NOT NULL OR T.DecimalResponse IS NOT NULL " & vbCrLf & _
		"ORDER BY A.Mandatory DESC, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	LastMandatory = True
	LastGoal=0
	LastStrategy=0
	Set vrs=Con.Execute(vsql)
	While vrs.EOF = False
		If LastMandatory <> vrs.Fields("Mandatory") Then
			LastMandatory = vrs.Fields("Mandatory")
			If LastMandatory = False Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: YellowGreen; "">Measures for Grantees. Add Target values for those that you will measure.</th></tr>" & vbCrLf)
			End If
		End If
		If LastGoal <> vrs.Fields("GoalID") And vrs.Fields("Mandatory") = False Then
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			LastGoal = vrs.Fields("GoalID")
			Response.Write("<td style=""text-align: right; "">" & vrs.Fields("GoalID") & "</td>" & vbCrLf)
			Response.Write("<th colspan=""3"" style=""background-color: PowderBlue;"">Goal " & vrs.Fields("GoalID") & ": " & vrs.Fields("Goal") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		ElseIf LastGoal <> vrs.Fields("GoalID") And vrs.Fields("Mandatory") = True Then
			LastGoal = vrs.Fields("GoalID")
			If vrs.Fields("GoalID") = 1 Then
				Response.Write("<tr style=""vertical-align: top; ""><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Mandatory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLf)
			ElseIf vrs.Fields("GoalID")=2 Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Mandatory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLf)
			End If
		End If
		If LastStrategy <> vrs.Fields("StrategyID") And vrs.Fields("Mandatory") = False  Then
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			LastStrategy = vrs.Fields("StrategyID")
			Response.Write("<td style=""text-align: right; "">" & vrs.Fields("GoalID") & "." & vrs.Fields("StrategyID") & "</td>" & vbCrLf)
			Response.Write("<th colspan=""3"" style=""background-color: PeachPuff; "">Strategy " & vrs.Fields("StrategyID") & ": " & vrs.Fields("Strategy") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & vrs.Fields("MeasureNumber") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & vrs.Fields("Activity") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & vrs.Fields("Measure") & "</td>" & vbCrLf)
		If vrs.Fields("Mandatory") Then
			Response.Write(vbTab & "<td class=""usertext"">Mandatory. Reporting for ")
			If vrs.Fields("IntegerResponse") = 0 Then
				Response.Write("Select Jurisdiction")
			ElseIf vrs.Fields("IntegerResponse") = 1 Then
				Response.Write("Taskforce Only")
			ElseIf vrs.Fields("IntegerResponse") = 2 Then
				Response.Write("Area of Jurisdiction")
			ElseIf vrs.Fields("IntegerResponse") = 3 Then
				Response.Write("Combination of TF and Jurisdiction")
			End If
			Response.Write("</td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=1 Then
			Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & vrs.Fields("IntegerResponse") & "</td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=2 Then
			Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & formatnumber(vrs.Fields("DecimalResponse")) & "</td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=3 Then
				Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & formatnumber(vrs.Fields("DecimalResponse")) & "</td>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
		vrs.MoveNext()
	Wend
	Response.Write("</table>" & vbCrLf)
End Function

function textarea2html(vText)
	if IsNull(vText) = true Then
		textarea2html = null
	ElseIf Len(vText)=0 Then
		textarea2html = ""
	Else
		'textarea2html = Replace(vText, vbCrLf&vbCrLf, "<br /><br />")
		textarea2html = Replace(vText, vbCrLf, "<br />")
	End If
end function

%>