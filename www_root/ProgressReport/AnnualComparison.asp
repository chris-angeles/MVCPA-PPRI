<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, ShowExcel, OrderBy, Columns, _
	ShowQuestions, ShowQuestionsDescription, ShowQuestionsClause, _
	ShowGrantees, Hide20182019, HideResponses, LastQuestion, LastGoal, LastStrategy

ShowQuestionsDescription = Array("All", "Mandatory", "Border", "Mandatory and Border", "Exclude No Targets")
ShowQuestionsClause = Array("", " AND Mandatory=1 ", "AND BorderReport=1 "," AND (Mandatory=1 Or BorderReport=1) ", " AND (ISNULL(Mandatory,0)=0 AND ISNULL(NoTarget,0)=0) ")
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
If Request.QueryString("ShowExcel") = "1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

If Len(Request.Form("ShowQuestions")) > 0 Then
	ShowQuestions = CInt(Request.Form("ShowQuestions"))
ElseIf Len(Request.QueryString("ShowQuestions")) > 0 Then
	ShowQuestions = CInt(Request.QueryString("ShowQuestions"))
Else
	ShowQuestions = 0
End If

If Len(Request.Form("ShowGrantees")) > 0 Then
	ShowGrantees = CInt(Request.Form("ShowGrantees"))
ElseIf Len(Request.QueryString("ShowGrantees")) > 0 Then
	ShowGrantees = CInt(Request.QueryString("ShowGrantees"))
Else
	ShowGrantees = 0
End If

If Request.Form("Hide20182019") = "1" Then
	Hide20182019 = True
ElseIf Request.QueryString("Hide20182019") = "1" Then
	Hide20182019 = True
Else
	Hide20182019 = False
End If

If Request.Form("HideResponses") = "1" Then
	HideResponses = True
ElseIf Request.QueryString("HideResponses") = "1" Then
	HideResponses = True
Else
	HideResponses = False
End If

If Hide20182019 = False Then
	If HideResponses = True Then
		Columns = 9
	Else
		Columns = 12
	End If
Else
	If HideResponses = True Then
		Columns = 7
	Else
		Columns = 8
	End If
End If

If ShowExcel = True and Debug = False Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=AnnualComparison.xls"
ElseIf Debug = False Then ' Start of Web only code
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Progress Report Annual Comparisons</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<form name="Selection" id="Selection" method="post" >
<label for="ShowQuestions">Questions to Show: </label> <select name="ShowQuestions" id="ShowQuestions" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowQuestionsDescription)
	Response.Write("<option value=""" & i & """" & Selected(ShowQuestions, i) & ">" & ShowQuestionsDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<label for="ShowGrantees">Grantees to Show: </label> <select name="ShowGrantees" id="ShowGrantees" onchange="Selection.submit();">
<%
Response.Write("<option value=""0""" & Selected(ShowGrantees, 0) & ">All</option>" & vbCrLf)
Response.Write("<option value=""-5""" & Selected(ShowGrantees, -5) & ">Border</option>" & vbCrLf)
Response.Write("<option value=""-4""" & Selected(ShowGrantees, -4) & ">Port</option>" & vbCrLf)
Response.Write("<option value=""-3""" & Selected(ShowGrantees, -3) & ">Port 2</option>" & vbCrLf)
Response.Write("<option value=""-2""" & Selected(ShowGrantees, -2) & ">Border and Port</option>" & vbCrLf)
Response.Write("<option value=""-1""" & Selected(ShowGrantees, -1) & ">Border, Port, and Port 2</option>" & vbCrLf)

sql = "SELECT GranteeID, REPLACE(GranteeName, 'City of ','') AS Grantee " & vbCrLf & _
	"FROM Grantees " & vbCrLf & _
	"WHERE GranteeID IN (SELECT GranteeID FROM [Grants].Main) " & vbCrLf & _
	"ORDER BY 2"
Set rs=Con.Execute(sql)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("GranteeID") & """" & Selected(ShowGrantees, rs.Fields("GranteeID")) & ">" & rs.Fields("grantee") & "</option>" & vbCrLf)
	rs.MoveNext
Wend
%>
</select>&nbsp;&nbsp;
<input type="checkbox" name="Hide20182019" value="1" <%=Checked(Hide20182019, True) %> onchange="Selection.submit();" /> Hide 2018 / 2019&nbsp;&nbsp;
<input type="checkbox" name="HideResponses" value="1" <%=Checked(HideResponses, True) %> onchange="Selection.submit();" /> Hide Number of Responses&nbsp;&nbsp;
<a href="AnnualComparison.asp?ShowExcel=1&ShowQuestions=<%=ShowQuestions %>&ShowGrantees=<%=ShowGrantees %>&HideResponses=<%=HideResponses %>&Hide20182019=<%=Hide20182019 %>" target="_blank">Excel</a>
</form>

<br />
<%	End If 

sql = "WITH CTE AS ( " & vbCrLf & _
	"SELECT A.GranteeID, REPLACE(A.GranteeName,'City Of ','') AS Grantee, A.BorderCounty, A.PortCounty, " & vbCrLf & _
	"	B.GoalID, B.StrategyID, B.ActivityID, B.MeasureID, B.LatestQID, " & vbCrLf & _
	"	C.QuestionID, C.Goal, C.Strategy, C.Activity, C.Measure, C.ResponseTypeID, " & vbCrLf & _
	"	C.Mandatory, C.BorderReport, C.NoTarget, " & vbCrLf
If Hide20182019 = False Then
	sql = sql & "	CASE WHEN Q2018.Integer_Total IS NOT NULL THEN CAST(Q2018.Integer_Total AS decimal) " & vbCrLf & _
	"		ELSE Q2018.Decimal_Total END AS [2018_Total], " & vbCrLf & _
	"	Q2018.Responses AS [2018_Responses], " & vbCrLf & _
	"	CASE WHEN Q2019.Integer_Total IS NOT NULL THEN CAST(Q2019.Integer_Total AS decimal) " & vbCrLf & _
	"		ELSE Q2019.Decimal_Total END AS [2019_Total], " & vbCrLf & _
	"	Q2019.Responses AS [2019_Responses], " & vbCrLf
End If
	sql = sql & "	CASE WHEN Q2020.Integer_Total IS NOT NULL THEN CAST(Q2020.Integer_Total AS decimal) " & vbCrLf & _
	"		ELSE Q2020.Decimal_Total END AS [2020_Total], " & vbCrLf & _
	"	Q2020.Responses AS [2020_Responses], " & vbCrLf & _
	"	CASE WHEN Q2021.Integer_Total IS NOT NULL THEN CAST(Q2021.Integer_Total AS decimal) " & vbCrLf & _
	"		ELSE Q2021.Decimal_Total END AS [2021_Total], " & vbCrLf & _
	"	Q2021.Responses AS [2021_Responses], " & vbCrLf & _
	"	CASE WHEN Q2022.Integer_Total IS NOT NULL THEN CAST(Q2022.Integer_Total AS decimal) " & vbCrLf & _
	"		ELSE Q2022.Decimal_Total END AS [2022_Total], " & vbCrLf & _
	"	Q2022.Responses AS [2022_Responses] " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"CROSS JOIN PR.YearMatrix AS B " & vbCrLf & _
	"LEFT JOIN PR.vwQuestions AS C ON C.QuestionID=B.LatestQID " & vbCrLf
If Hide20182019 = False Then
	sql = sql & "LEFT JOIN PR.vwGrantQIDTotal AS Q2018 ON Q2018.QuestionID=B.QID2018 AND Q2018.GranteeID=A.GranteeID AND Q2018.Fiscal_Year=2018 " & vbCrLf & _
	"LEFT JOIN PR.vwGrantQIDTotal AS Q2019 ON Q2019.QuestionID=B.QID2019 AND Q2019.GranteeID=A.GranteeID AND Q2019.Fiscal_Year=2019 " & vbCrLf
End If
	sql = sql & "LEFT JOIN PR.vwGrantQIDTotal AS Q2020 ON Q2020.QuestionID=B.QID2020 AND Q2020.GranteeID=A.GranteeID AND Q2020.Fiscal_Year=2020 " & vbCrLf & _
	"LEFT JOIN PR.vwGrantQIDTotal AS Q2021 ON Q2021.QuestionID=B.QID2021 AND Q2021.GranteeID=A.GranteeID AND Q2021.Fiscal_Year=2021 " & vbCrLf & _
	"LEFT JOIN PR.vwGrantQIDTotal AS Q2022 ON Q2022.QuestionID=B.QID2022 AND Q2022.GranteeID=A.GranteeID AND Q2022.Fiscal_Year=2022 " & vbCrLf & _
	"WHERE ("
If Hide20182019 = False Then
	sql = sql & "Q2018.Responses>0 OR Q2019.Responses>0 OR "
End If
	sql = sql & "Q2020.Responses>0 OR Q2021.Responses>0 OR Q2022.Responses>0) AND B.[QID2022] IS NOT NULL " & vbCrLf & _
	ShowQuestionsClause(ShowQuestions) & vbCrLf
If ShowGrantees = -5 Then
	sql = sql & "AND BorderCounty=1 " & vbCrLf
ElseIf ShowGrantees = -4 Then
	sql = sql & "AND PortCounty=1 " & vbCrLf
ElseIf ShowGrantees = -3 Then
	sql = sql & "AND Port2County=1 " & vbCrLf
ElseIf ShowGrantees = -2 Then
	sql = sql & "AND (BorderCounty=1 OR PortCounty=1) " & vbCrLf
ElseIf ShowGrantees = -1 Then
	sql = sql & "AND (BorderCounty=1 OR PortCounty=1 OR Port2County=1) " & vbCrLf
ElseIf ShowGrantees>0 Then
	sql = sql & " AND A.GranteeID=" & ShowGrantees & vbCrLf
End If
sql = sql & ") " & vbCrLf & _
	"SELECT GranteeID, Grantee, BorderCounty, PortCounty, " & vbCrLf & _
	"	GoalID, StrategyID, ActivityID, MeasureID, " & vbCrLf & _
	"	LatestQID, QuestionID, Goal, Strategy, Activity, Measure, " & vbCrLf & _
	"	ResponseTypeID, Mandatory, BorderReport, NoTarget, " & vbCrLf
If Hide20182019 = False Then
	sql = sql & "	[2018_Total], [2018_Responses], " & vbCrLf & _
	"	[2019_Total], [2019_Responses], " & vbCrLf
End If
	sql = sql & "	[2020_Total], [2020_Responses], " & vbCrLf & _
	"	[2021_Total], [2021_Responses], " & vbCrLf & _
	"	[2022_Total], [2022_Responses] " & vbCrLf & _
	"FROM CTE " & vbCrLf & _
	"UNION " & vbCrLf & _
	"SELECT NULL AS GranteeID, ' Total' AS Grantee, NULL AS BorderCounty, NULL AS PortCounty, " & vbCrLf & _
	"	GoalID, StrategyID, ActivityID, MeasureID, " & vbCrLf & _
	"	LatestQID, QuestionID, Goal, Strategy, Activity, Measure, " & vbCrLf & _
	"	ResponseTypeID, Mandatory, BorderReport, NoTarget, " & vbCrLf
If Hide20182019 = False Then
	sql = sql & "	SUM([2018_Total]) AS [2018_Total], SUM([2018_Responses]) AS [2018_Responses], " & vbCrLf & _
	"	SUM([2019_Total]) AS [2019_Total], SUM([2019_Responses]) AS [2019_Responses], " & vbCrLf
End If
	sql = sql & "	SUM([2020_Total]) AS [2020_Total], SUM([2020_Responses]) AS [2020_Responses], " & vbCrLf & _
	"	SUM([2021_Total]) AS [2021_Total], SUM([2021_Responses]) AS [2021_Responses], " & vbCrLf & _
	"	SUM([2022_Total]) AS [2022_Total], SUM([2022_Responses]) AS [2022_Responses] " & vbCrLf & _
	"FROM CTE " & vbCrLf & _
	"GROUP BY GoalID, StrategyID, ActivityID, MeasureID, LatestQID, QuestionID, Goal, Strategy, Activity, Measure, ResponseTypeID, Mandatory, BorderReport, NoTarget " & vbCrLf & _
	"ORDER BY GoalID, StrategyID, ActivityID, MeasureID, LatestQID, Grantee "

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs=Con.Execute(sql)

%>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	If HideResponses = True Then
		Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
		Response.Write(vbTab & "<th>Grantee ID</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Grantee</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">2018 Total</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">2019 Total</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">2020 Total</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">2021 Total</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">2022 Total</th>" & vbCrLf)
		Response.Write(vbTab & "</tr>" & vbCrLf)
	Else
		Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
		Response.Write(vbTab & "<th rowspan=""2"">Grantee ID</th>" & vbCrLf)
		Response.Write(vbTab & "<th rowspan=""2"">Grantee</th>" & vbCrLf)
		If Hide20182019 = False Then
			Response.Write(vbTab & "<th colspan=""2"" style=""width: 100px; "">2018</th>" & vbCrLf)
			Response.Write(vbTab & "<th colspan=""2"" style=""width: 100px; "">2019</th>" & vbCrLf)
		End If
		Response.Write(vbTab & "<th colspan=""2"" style=""width: 100px; "">2020</th>" & vbCrLf)
		Response.Write(vbTab & "<th colspan=""2"" style=""width: 100px; "">2021</th>" & vbCrLf)
		Response.Write(vbTab & "<th colspan=""2"" style=""width: 100px; "">2022</th>" & vbCrLf)
		Response.Write(vbTab & "</tr>" & vbCrLf)
		Response.Write(vbTab & "<tr>" & vbCrLf)
		If Hide20182019 = False Then
			Response.Write(vbTab & "<th style=""width: 50px; "">Total</th>" & vbCrLf)
			Response.Write(vbTab & "<th style=""width: 50px; "">Responses</th>" & vbCrLf)
			Response.Write(vbTab & "<th style=""width: 50px; "">Total</th>" & vbCrLf)
			Response.Write(vbTab & "<th style=""width: 50px; "">Responses</th>" & vbCrLf)
		End If
		Response.Write(vbTab & "<th style=""width: 50px; "">Total</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">Responses</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">Total</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">Responses</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">Total</th>" & vbCrLf)
		Response.Write(vbTab & "<th style=""width: 50px; "">Responses</th>" & vbCrLf)
		Response.Write(vbTab & "</tr>" & vbCrLf)
	End If
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False

		If LastGoal <> rs.Fields("GoalID") Then
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
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th></th>" & vbCrLf)
			Response.Write(vbTab & "<td colspan=""" & (columns - 1) & """ style=""font-weight: bold; text-align: left; "">" & rs.Fields("Measure") & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If

		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & rs.Fields("GranteeID").value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; white-space: nowrap;"">" & rs.Fields("Grantee").value & "</td>" & vbCrLf)

		If Hide20182019 = False Then
			If IsNull(rs.Fields("2018_Total").value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf rs.Fields("ResponseTypeID") = 1 Then
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2018_Total").value,0, True, True, True) & "</td>" & vbCrLf)
			ElseIf rs.Fields("ResponseTypeID") = 2 Then
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2018_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
			ElseIf rs.Fields("ResponseTypeID") = 3 Then
				Response.Write(vbTab & "<td style=""text-align: right; "">$" & formatnumber(rs.Fields("2018_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
			End If

			If HideResponses = False Then
				If IsNull(rs.Fields("2018_Responses").value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2018_Responses").value,0, True, True, True) & "</td>" & vbCrLf)
				End If
			End If

			If IsNull(rs.Fields("2019_Total").value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf rs.Fields("ResponseTypeID") = 1 Then
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2019_Total").value,0, True, True, True) & "</td>" & vbCrLf)
			ElseIf rs.Fields("ResponseTypeID") = 2 Then
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2019_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
			ElseIf rs.Fields("ResponseTypeID") = 3 Then
				Response.Write(vbTab & "<td style=""text-align: right; "">$" & formatnumber(rs.Fields("2019_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
			End If

			If HideResponses = False Then
				If IsNull(rs.Fields("2019_Responses").value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2019_Responses").value,0, True, True, True) & "</td>" & vbCrLf)
				End If
			End If
		End If

		If IsNull(rs.Fields("2020_Total").value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 1 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2020_Total").value,0, True, True, True) & "</td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 2 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2020_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 3 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">$" & formatnumber(rs.Fields("2020_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
		End If

		If HideResponses = False Then
			If IsNull(rs.Fields("2020_Responses").value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2020_Responses").value,0, True, True, True) & "</td>" & vbCrLf)
			End If
		End If

		If IsNull(rs.Fields("2021_Total").value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 1 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2021_Total").value,0, True, True, True) & "</td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 2 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2021_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 3 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">$" & formatnumber(rs.Fields("2021_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
		End If

		If HideResponses = False Then
			If IsNull(rs.Fields("2021_Responses").value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2021_Responses").value,0, True, True, True) & "</td>" & vbCrLf)
			End If
		End If

		If IsNull(rs.Fields("2022_Total").value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 1 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2022_Total").value,0, True, True, True) & "</td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 2 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2022_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 3 Then
			Response.Write(vbTab & "<td style=""text-align: right; "">$" & formatnumber(rs.Fields("2022_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
		End If

		If HideResponses = False Then
			If IsNull(rs.Fields("2022_Responses").value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("2022_Responses").value,0, True, True, True) & "</td>" & vbCrLf)
			End If
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