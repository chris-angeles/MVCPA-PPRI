<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, ShowExcel, FiscalYear, OrderBy, Quarter, StartQuarter, EndQuarter, Months, Months2, _
	OrderByDescription, QuarterDescription, OrderByField, Show, ShowDescription, ShowClause, _
	IncludeStatutory, LastQuestion, LastGoal, LastStrategy, Total, GrantClassID, Version
OrderByDescription = Array("GrantID", "Grantee Name")
QuarterDescription = Array("", "September 1 - November 30", "December 1 - February 28", "March 1 - May 31", "June 1 - August 31", "September 1 - November 30", "September 1 - February 28", "September 1 - May 31", "September 1 - August 31", "March 1 - August 31")
OrderByField = Array("D.GrantID", "REPLACE(E.GranteeName,'City of ','')")
ShowDescription = Array ("All", "Border", "Port", "Port 2", "Border and Port", "Border, Port, and Port 2")
ShowClause = Array ("1=1", "E.BorderCounty=1", "E.PortCounty=1", "E.Port2County=1", "(E.BorderCounty=1 OR E.PortCounty=1)", "(E.BorderCounty=1 OR E.PortCounty=1 OR E.Port2County=1)")

Months = Array("Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug")
Months2 = Array("September", "October", "November", "December", "January", "February", "March", "April", "May", "June", "July", "August")

GrantClassID=1

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

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
End If

If Len(Request.Form("Quarter"))>0 Then
	Quarter = CInt(Request.Form("Quarter"))
ElseIf Len(Request.QueryString("Quarter"))>0 Then
	Quarter = CInt(Request.QueryString("Quarter"))
Else
	Quarter = 1
End If

If Len(Request.Form("Show")) > 0 Then
	Show = CInt(Request.Form("Show"))
ElseIf Len(Request.QueryString("Show")) > 0 Then
	Show = CInt(Request.QueryString("Show"))
Else
	Show = 4
End If

If Request.Form("IncludeStatutory")="1" Then
	IncludeStatutory = True
ElseIf Request.QueryString("IncludeStatutory")="1" Then
	IncludeStatutory = True
Else
	IncludeStatutory = False
End If

Version = PRVersion(GrantClassID, FiscalYear)

Months = Array("Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug")
Months2 = Array("September", "October", "November", "December", "January", "February", "March", "April", "May", "June", "July", "August")

If Quarter = 1 Then
	StartQuarter = 1
	EndQuarter = 1
ElseIf Quarter = 2 Then
	StartQuarter = 2
	EndQuarter = 2
ElseIf Quarter = 3 Then
	StartQuarter = 3
	EndQuarter = 3
ElseIf Quarter = 4 Then
	StartQuarter = 4
	EndQuarter = 4
ElseIf Quarter = 5 Then
	StartQuarter = 1
	EndQuarter = 1
ElseIf Quarter = 6 Then
	StartQuarter = 1
	EndQuarter = 2
ElseIf Quarter = 7 Then
	StartQuarter = 1
	EndQuarter = 3
ElseIf Quarter = 8 Then
	StartQuarter = 1
	EndQuarter = 4
ElseIf Quarter = 9 Then
	StartQuarter = 3
	EndQuarter = 4
End If

sql = "SELECT A.GoalID, B.StrategyID, C.ActivityID, C.Activity, C.Measure, C.QuestionID, C.ResponseTypeID, " & vbCrLf & _
	"	A.Goal, B.Strategy, " & vbCrLf & _
	"	D.GrantID, E.GranteeName, D.ProgramName, D.FiscalYear, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Sep), CONVERT(VARCHAR,DecimalResponse_Sep)) AS September, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Oct), CONVERT(VARCHAR,DecimalResponse_Oct)) AS October, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Nov), CONVERT(VARCHAR,DecimalResponse_Nov)) AS November, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Dec), CONVERT(VARCHAR,DecimalResponse_Dec)) AS December, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Jan), CONVERT(VARCHAR,DecimalResponse_Jan)) AS January, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Feb), CONVERT(VARCHAR,DecimalResponse_Feb)) AS February, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Mar), CONVERT(VARCHAR,DecimalResponse_Mar)) AS March, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Apr), CONVERT(VARCHAR,DecimalResponse_Apr)) AS April, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_May), CONVERT(VARCHAR,DecimalResponse_May)) AS May, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Jun), CONVERT(VARCHAR,DecimalResponse_Jun)) AS June, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Jul), CONVERT(VARCHAR,DecimalResponse_Jul)) AS July, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,IntegerResponse_Aug), CONVERT(VARCHAR,DecimalResponse_Aug)) AS August, " & vbCrLf & _
	"	ISNULL(CONVERT(VARCHAR,H.Integer_Total), ISNULL(CONVERT(VARCHAR,H.Decimal_Total),CONVERT(VARCHAR,H.Dollar_Total,1))) AS Total, " & vbCrLf & _
	"	TextResponse_Q1, TextResponse_Q2, TextResponse_Q3, TextResponse_Q4" & vbCrLf & _
	"FROM PR.Goals AS A " & vbCrLf & _
	"LEFT JOIN PR.Strategies AS B ON B.GoalID=A.GoalID And A.Version=B.Version " & vbCrLf & _
	"LEFT JOIN PR.Activities AS C ON C.StrategyID=B.StrategyID AND C.GoalID=A.GoalID AND C.Version=B.Version" & vbCrLf & _
	"LEFT JOIN PR.GrantQuestions AS G ON G.QuestionID=C.QuestionID " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS D ON D.GrantID=G.GrantID " & vbCrLf & _
	"LEFT JOIN Grantees AS E ON E.GranteeID=D.GranteeID " & vbCrLf & _
	"LEFT JOIN PR.Responses AS F ON F.QuestionID=C.QuestionID AND F.GrantID=D.GrantID " & vbCrLf & _
	"LEFT JOIN vwProgressReportTotals AS H ON H.GrantID=D.GrantID AND H.QuestionID=G.QuestionID " & vbCrLf & _
	"WHERE A.Version=" & prepIntegerSQL(Version) & " AND "
If IncludeStatutory = True Then
	sql = sql & "((C.BorderReport=1) Or C.Mandatory=1) "
Else
	sql = sql & "(C.BorderReport=1) "
End If
sql = sql & " AND D.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND " & ShowClause(Show) & " " & vbCrLf & _
	"ORDER BY A.GoalID, B.StrategyID, C.ActivityID, " & OrderByField(OrderBy)
Set rs=Con.Execute(sql)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

If ShowExcel = True and Debug = False Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ApplicationStatus" & FiscalYear & ".xls"
Else ' Start of Web only code
	If Debug = False Then 
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Border Operations Report</title>
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
</select>&nbsp;&nbsp;
<label for="Quarter">Quarter:</label> <select name="Quarter" id="Quarter" onchange="Selection.submit();">
<%
	For i = 1 to 9
		Response.Write("<option value=""" & i & """" & selected(Quarter, i) & ">" & QuarterDescription(i) & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="Show">Show:</label> <select name="Show" id="Show" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowDescription)
	Response.Write("<option value=""" & i & """" & Selected(Show, i) & ">" & ShowDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<input type="checkbox" name="IncludeStatutory" id="IncludeStatutory" value="1" <%=Checked(IncludeStatutory, True) %> onclick="Selection.submit();" />
<label for="IncludeStatutory">Include Statutory Requirements</label>&nbsp;&nbsp;
<label for="OrderBy">Order By:</label>
<select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;<a href="BorderOperations.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&OrderBy=<%=OrderBy %>&Quarter=<%=Quarter %>&Show=<%=Show %>&IncludeStatutory=<%
	If IncludeStatutory = True Then 
		Response.Write("1")
	Else
		Response.Write("0")
	End If
%>" target="_blank">Excel</a>
</form>

<br />
<%	End If %>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	Response.Write(vbTab & "<th>Grant ID</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Grantee</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Program</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Fiscal Year</th>" & vbCrLf)
	For i = 3*(StartQuarter-1) to (3*(EndQuarter)-1)
		Response.Write(vbTab & "<th>" & Months2(i) & "</th>" & vbCrLf)
	Next
	Response.Write(vbTab & "<th>Total</th>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Total = 0.0
		If LastGoal <> rs.Fields("GoalID") Then
			LastGoal = rs.Fields("GoalID")
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th colspan=""" & (3*(EndQuarter-StartQuarter+1)+5) & """>Goal " & rs.Fields("GoalID") & ": " & rs.Fields("Goal") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		If LastStrategy <> rs.Fields("StrategyID") Then
			LastStrategy = rs.Fields("StrategyID")
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th colspan=""" & (3*(EndQuarter-StartQuarter+1)+5) & """>Strategy " & rs.Fields("StrategyID") & ": " & rs.Fields("Strategy") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		If LastQuestion <> rs.Fields("QuestionID") Then
			LastQuestion = rs.Fields("QuestionID")
			Response.Write("<tr style=""background-color: PowderBlue; vertical-align: top; "">" & vbCrLf)
			Response.Write(vbTab & "<th>" & rs.Fields("GoalID") & "." & rs.Fields("StrategyID") & "." & rs.Fields("ActivityID") & "</th>" & vbCrLf)
			Response.Write(vbTab & "<td colspan=""" & (3*(EndQuarter-StartQuarter+1)+4) & """ style=""font-weight: bold; test-align: left; "">" & rs.Fields("Activity") & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th></th>" & vbCrLf)
			Response.Write(vbTab & "<td colspan=""" & (3*(EndQuarter-StartQuarter+1)+4) & """ style=""font-weight: bold; test-align: left; "">" & rs.Fields("Measure") & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & rs.Fields("GrantID").value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("GranteeName").value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("ProgramName").value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & rs.Fields("FiscalYear").value & "</td>" & vbCrLf)
		If rs.Fields("ResponseTypeID") = 5 Then
			For i = StartQuarter To EndQuarter
				Response.Write(vbTab & "<td colspan=""3"" style=""text-align: left; "">" & rs.Fields("TextResponse_Q" & i).value & "</td>" & vbCrLf)
			Next
		ElseIf rs.Fields("ResponseTypeID") = 6 Then
			For i = (StartQuarter-1)*3 To (EndQuarter-1)*3 '0 to ((Quarter-1)*3) Step 3
				If rs.Fields(Months2(i)).value = "1" Then
					Response.Write(vbTab & "<td colspan=""3"" style=""text-align: center; "">yes</td>" & vbCrLf)
				ElseIf rs.Fields(Months2(i)).value = 0 Then
					Response.Write(vbTab & "<td colspan=""3"" style=""text-align: center; "">no</td>" & vbCrLf)
				End If
			Next
		ElseIf rs.Fields("ResponseTypeID") = 7 Then
			For i = 3*(StartQuarter-1) to (3*(EndQuarter)-1)
				If rs.Fields(Months2(i)).value = 1 Then
					Response.Write(vbTab & "<td style=""text-align: center; "">yes</td>" & vbCrLf)
				ElseIf rs.Fields(Months2(i)).value = 0 Then
					Response.Write(vbTab & "<td style=""text-align: center; "">no</td>" & vbCrLf)
				End If
			Next
		Else
			For i = 3*(StartQuarter-1) to (3*(EndQuarter)-1)
				If rs.Fields("ResponseTypeID") = 1 Then
					Response.Write(vbTab & "<td style=""text-align: right"">" & prepIntegerWeb(rs.Fields(Months2(i)).value) & "</td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWeb(rs.Fields(Months2(i)).value) & "</td>" & vbCrLf)
				End If
				If IsNull(rs.Fields(Months2(i)).value) = False Then
					Total = Total + rs.Fields(Months2(i)).value
				End If
			Next
		End If
		If rs.Fields("ResponseTypeID") = 5 Or rs.Fields("ResponseTypeID") = 6 Or rs.Fields("ResponseTypeID") = 7 Then
			Response.Write(vbTab & "<td style=""text-align: right""></td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 1 Then
			Response.Write(vbTab & "<td style=""text-align: right"">" & prepIntegerWeb(Total) & "</td>" & vbCrLf)
		ElseIf rs.Fields("ResponseTypeID") = 3 Then
			Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWeb(Total) & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right""></td>" & vbCrLf)
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
<!--#include file="../ProgressReport/PRVersionInclude.asp"-->