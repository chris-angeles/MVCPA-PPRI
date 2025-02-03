<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, k, ViewDocuments, PermitEdit, ShowExcel, Columns, CurrentDate, _
	Quarter, MaxQuarter, ShowOneQuarter, QuarterDescription, Version, GrantID, _
	StartDate, LastGoal, LastStrategy, LastMandatory,  Confirmed, FiscalYear
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

If Request.Querystring("ShowExcel")="1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

If Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	FiscalYear=2022
End If

Columns = 6

If FiscalYear>2021 Then
	Version = 5
ElseIf FiscalYear>2020 Then
	Version = 4
ElseIf FiscalYear>2019 Then
	Version = 3
Else
	Version = 2
End If

PermitEdit = False

ViewDocuments = True


sql = "SELECT A.QuestionID, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.NoTarget, A.ResponseTypeID " & vbCrLf & _
	"FROM PR.Goals AS G " & vbCrLf & _
	"LEFT JOIN PR.Strategies AS S ON S.GoalID=G.GoalID AND S.Version=G.Version " & vbCrLf & _
	"LEFT JOIN PR.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID AND A.Version=G.Version " & vbCrLf & _
	"WHERE G.Version=" & prepIntegerSQL(Version) & " " & vbCrLF & _
	"ORDER BY A.Mandatory DESC, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ProgressReport" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr><th colspan=""" & columns & """>MVCPA Progress Report for Fiscal Year " & FiscalYear & ", Quarter " & Quarter & "</th></tr>" & vbCrLf)
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Progress Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style>
	tr, td, th {padding: 5px;}
</style>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body style="width: 100%">
<div style="text-align: center;"><form name="Selection" method="get" action="Version.asp">
<select name="FiscalYear" onchange="document.Selection.submit();">
<%
For i = 2018 to 2022
	Response.Write(SelectOption(i, "FiscalYear " & i, FiscalYear))
Next
%>
	</select>
</form></div>
<H1>MVCPA Progress Report for Fiscal Year <%=FiscalYear %> (Version <%=Version %>)</H1>
<%	End If %>
<table style="margin: auto">
<thead>
	<tr>
		<th>ID</th>
		<th>Activity</th>
		<th>Measure</th>
		<th>No Target</th>
		<th>Response Type</th>
		<th title="QuestionID">QID</th>
	</tr>
</thead>
<%
LastMandatory = True
LastGoal=0
LastStrategy=0
Set rs=Con.Execute(sql)
While rs.EOF = False
	'If LastMandatory <> rs.Fields("Mandatory") Then
	'	LastMandatory = rs.Fields("Mandatory")
	'	If LastMandatory = False Then
	'		Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: YellowGreen; "">Measures for Grantees. Add Target values for those that you will measure.</th></tr>" & vbCrLf)
	'	End If
	'End If
	If LastGoal <> rs.Fields("GoalID") And rs.Fields("Mandatory") = False Then
		Response.Write("<tr>" & vbCrLf)
		LastGoal = rs.Fields("GoalID")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("GoalID") & "</td>" & vbCrLf)
		If rs.Fields("GoalID") < 4 Then
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PowderBlue;"">Goal " & rs.Fields("GoalID") & ": " & rs.Fields("Goal") & "</th>" & vbCrLf)
		Else
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PowderBlue;"">Section " & rs.Fields("GoalID") & ": " & rs.Fields("Goal") & "</th>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
	ElseIf LastGoal <> rs.Fields("GoalID") And rs.Fields("Mandatory") = True Then
		LastGoal = rs.Fields("GoalID")
		If rs.Fields("GoalID") = 1 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Mandatory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLf)
		ElseIf rs.Fields("GoalID")=2 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Mandatory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLf)
		ElseIf rs.Fields("GoalID")=8 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Fraud-Related Motor Vehicle Crime Measures must be reported during the grant term if awarded."">Mandatory Fraud-Related Motor Vehicle Crime Measures Required for all Grantees</th></tr>" & vbCrLf)
		End If
	End If

	' Strategy row
	If LastStrategy <> rs.Fields("StrategyID") And rs.Fields("Mandatory") = False  Then
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		LastStrategy = rs.Fields("StrategyID")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("GoalID") & "." & rs.Fields("StrategyID") & "</td>" & vbCrLf)
		If rs.Fields("GoalID") = 6 Then
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PeachPuff; "">Subsection " & rs.Fields("StrategyID") & "</th>" & vbCrLf)
		ElseIf rs.Fields("GoalID") = 4 Or rs.Fields("GoalID") = 5 Then
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PeachPuff; "">Subsection " & rs.Fields("StrategyID") & ": " & rs.Fields("Strategy") & "</th>" & vbCrLf)
		Else
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PeachPuff; "">Strategy " & rs.Fields("StrategyID") & ": " & rs.Fields("Strategy") & "</th>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
	End If

	' Question row
	If rs.Fields("QuestionID") = 112 or rs.Fields("QuestionID") = 113 or rs.Fields("QuestionID") = 114 Then
		Response.Write("<tr style=""vertical-align: top; "" title=""Specific LBB Border Security Requirement"">" & vbCrLf)
	Else
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
	End If
	Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("MeasureNumber") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("Activity") & "</td>" & vbCrLf)
	' Measure cell adds the description of the reporting area for the required reporting.
	Response.Write(vbTab & "<td>" & rs.Fields("Measure")) 
	Response.Write(vbTab & "<td>" & rs.Fields("NoTarget")) 
	Response.Write(vbTab & "<td>" & rs.Fields("ResponseTypeID")) 
	Response.Write("<td>" & rs.Fields("QuestionID") & "</td>")
	Response.Write("</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	rs.MoveNext()
Wend
Response.Write("</table>" & vbCrLf)
If ShowExcel = False Then
%>
<br />

<div style="text-align: center; margin: auto; ">
	<input type="button" name="Close" value="Close" title="Ignore any pending changes and close window." 
		onclick="window.close();"/>
</div>
<div style="text-align: right"><a href="Version.asp?FiscalYear=<%=FiscalYear %>&ShowExcel=1" target="_blank">Excel</a></div>
</body>
</html>
<%
End If
%>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->