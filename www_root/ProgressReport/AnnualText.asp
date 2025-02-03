<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, ShowExcel, FiscalYear, OrderBy, Columns, _
	LastQuestion, LastGoal, LastStrategy, LastGrant, SearchWords, words, TextResponse
Columns = 4

debug = False
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("<pre>Session(""" & i & """)='" & Session(i) & "'</pre>" & vbCrLf)
	Next
End If

LastQuestion = 0
LastGoal = 0
LastGrant = 0
If Request.QueryString("ShowExcel") = "1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	If Month(Date()) > 9 Then
		FiscalYear = Year(Date)+1
	Else
		FiscalYear = Year(Date)
	End If
End If

If Len(Request.Form("SearchWords")) > 0 Then
	SearchWords = Request.Form("SearchWords")
ElseIf Len(Request.QueryString("SearchWords")) > 0 Then
	SearchWords = Request.QueryString("SearchWords")
Else
	SearchWords = ""
End If
If Len(SearchWords)>0 Then
	Words = Split(SearchWords," ")
End If

sql = "SELECT * FROM vwProgressReportAnnualText " & vbCrLF & _
	"WHERE Fiscal_Year=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf
'If Len(SearchWords)>0 Then
'	sql = sql & vbTab & "AND TextResponse LIKE " & prepStringSQL("%" & SearchWords & "%") & " " & vbCrLf
'End If
If Len(SearchWords)>0 Then
	For i = 0 to UBound(Words)
		'sql = sql & vbTab & "AND TextResponse LIKE " & prepStringSQL("%" & Words(i) & "%") & " " & vbCrLf
		sql = sql & vbTab & "AND TextResponse LIKE " & prepStringSQL("%" & REPLACE(Words(i),"~"," ") & "%") & " " & vbCrLf
	Next
End If
sql = sql & "ORDER BY Fiscal_Year, Mandatory DESC, GoalID, StrategyID, ActivityID, MeasureID, CASE WHEN GrantID IS NULL THEN 1 ELSE 0 END, Grantee, Quarter "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs=Con.Execute(sql)

If ShowExcel = True and Debug = False Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ApplicationStatus" & FiscalYear & ".xls"
ElseIf Debug = False Then ' Start of Web only code
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Progress Report Annual Text Responses</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2018 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;Containing words <input type="text" name="SearchWords" size="20" maxlength="30" value="<%=SearchWords %>" /><input type="submit" value="GO" style="width: 45px;" />&nbsp;&nbsp;
<a href="AnnualText.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&SearchWords=<%=Server.URLEncode(SearchWords) %>" target="_blank">Excel</a>
</form>

<br />
<%	End If %>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr><th colspan=""" & (Columns) & """ >" & FiscalYear & " Progress Report Text Responses</th></tr>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	Response.Write(vbTab & "<th>Grant ID</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Grantee</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Program and Responses</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Award</th>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		If LastGoal <> rs.Fields("GoalID") And rs.Fields("Mandatory") = True Then
			LastGoal = rs.Fields("GoalID")
			If rs.Fields("GoalID") = 1 Then
				Response.Write("<tr><th colspan=""" & (Columns) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Mandatory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLf)
			ElseIf rs.Fields("GoalID")=2 Then
				Response.Write("<tr><th colspan=""" & (Columns) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Mandatory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLf)
			ElseIf rs.Fields("GoalID")=8 Then
					Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Fraud-Related Motor Vehicle Crime Measures must be reported during the grant term if awarded."">Mandatory Fraud-Related Motor Vehicle Crime Measures Required for all Grantees</th></tr>" & vbCrLf)
				End If
		End If
		If LastGoal <> rs.Fields("GoalID") Then
			LastGoal = rs.Fields("GoalID")
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th colspan=""" & Columns & """>Goal " & rs.Fields("GoalID") & ": " & rs.Fields("Goal") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		If LastStrategy <> rs.Fields("StrategyID") Then
			LastStrategy = rs.Fields("StrategyID")
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th colspan=""" & Columns & """>Strategy " & rs.Fields("StrategyID") & ": " & rs.Fields("Strategy") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		If LastQuestion <> rs.Fields("QuestionID") Then
			LastQuestion = rs.Fields("QuestionID")
			Response.Write("<tr style=""background-color: PowderBlue; vertical-align: top; "">" & vbCrLf)
			Response.Write(vbTab & "<th>" & rs.Fields("GoalID") & "." & rs.Fields("StrategyID") & "." & rs.Fields("ActivityID") & "</th>" & vbCrLf)
			Response.Write(vbTab & "<td colspan=""" & (columns - 1) & """ style=""font-weight: bold; text-align: left; "">" & rs.Fields("Activity") & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
			If IsNull(rs.Fields("Measure")) = False Then
				Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
				Response.Write(vbTab & "<th></th>" & vbCrLf)
				Response.Write(vbTab & "<td colspan=""" & (columns - 1) & """ style=""font-weight: bold; text-align: left; "">" & rs.Fields("Measure") & "</td>" & vbCrLf)
				Response.Write("</tr>" & vbCrLf)
			End If
		End If
		If LastGrant <> rs.Fields("GrantID") Then
			LastGrant = rs.Fields("GrantID")
			Response.Write("<tr style=""vertical-align: top; background-color: Lavender; "">" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right"">" & rs.Fields("GrantID").value & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: left; white-space: nowrap;"">" & rs.Fields("Grantee").value & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: left; white-space: nowrap;"">" & rs.Fields("ProgramName").value & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; white-space: nowrap;"">" & prepCurrencyWeb(rs.Fields("AwardAmount").value) & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		If IsNull(rs.Fields("TextResponse").value) = False Then
			Response.Write(vbTab & "<tr style=""vertical-align: top; ""><td></td><td style=""text-align: center; "">Q" & rs.Fields("Quarter") & "</td>")
			TextResponse = Replace(rs.Fields("TextResponse").value, vbCrLf, "<br />")
			If Len(SearchWords)>0 Then
				For i = 0 to UBound(Words)
					TextResponse = Replace(TextResponse, Replace(Words(i),"~"," "),"<span style=""background-color: yellow;"">" & Replace(Words(i),"~"," ") & "</span>")
				Next
			End If
			Response.Write("<td style=""text-align: left; "" colspan=""2"">" & TextResponse & "</td>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("<tbody>" & vbCrLf)
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%	If ShowExcel = False Then %>
<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>
</body>
</html>
<%	End If %>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->