<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, k, ViewDocuments, PermitEdit, ShowExcel, Columns, CurrentDate, _
	Quarter, MaxQuarter, ShowOneQuarter,  Version, GrantID, GrantClassID, _
	StartDate, LastGoal, LastStrategy, LastMandatory,  Confirmed, _
	FiscalYear, DisplayQuarterOffset, GranteeID, GranteeName, ProgramName, SubmitID, SubmitName, SubmitTimestamp, _
	AdministrativeComments, ApprovalID, ApprovalDate, ApprovalName, CanSubmit
debug = False

CurrentDate = Date() 
'CurrentDate = cdate("3/2/2018")

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

If Len(Request.Form("GrantID"))>0 Then
	GrantID = Request.Form("GrantID")
Else
	GrantID = Request.QueryString("GrantID")
End If

IF Len(GrantID)>0 Then
	GrantID = CInt(GrantID)
Else
	GrantID=0
End If

If Request.Querystring("ShowExcel")="1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

If Request.Form("ShowOneQuarter")="1" Then
	ShowOneQuarter = True
ElseIf Request.Querystring("ShowOneQuarter")="1" Then
	ShowOneQuarter = True
ElseIf Request.Form("ShowOneQuarter")="0" Then
	ShowOneQuarter = False
ElseIf Request.Form.Count=0 Then
	ShowOneQuarter = True
Else
	ShowOneQuarter = True
End If

sql = "SELECT H.GrantID, H.FiscalYear, G.GranteeID, G.GranteeName, H.ProgramName " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS H ON G.GranteeID=H.GranteeID " & vbCrLf & _
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
End If

' Determine Reporting Period
If Len(Request.Form("Quarter"))>0 Then
	Quarter = CInt(Request.Form("Quarter"))
ElseIf Len(Request.QueryString("Quarter"))>0 Then
	Quarter = CInt(Request.QueryString("Quarter"))
ElseIf CurrentDate < CDate("2/1/" & FiscalYear) Then
	Quarter = 1
ElseIf CurrentDate < CDate("4/1/" & FiscalYear) Then
	Quarter = 2
ElseIf CurrentDate < CDate("7/1/" & FiscalYear) Then
	Quarter = 3
Else
	Quarter = 4
End If
' Determine Max Reporting Period to select
If CurrentDate >= CDate("6/1/" & FiscalYear) Then
	MaxQuarter = 4
ElseIf CurrentDate >= CDate("3/1/" & FiscalYear) Then
	MaxQuarter = 3
ElseIf CurrentDate > CDate("12/1/" & (FiscalYear-1)) Then
	MaxQuarter = 2
Else
	MaxQuarter = 1
End If

If Debug = True Then
	Response.Write("<pre>Quarter=" & Quarter & "</pre>")
	Response.Write("<pre>CurrentDate=" & CurrentDate & "</pre>")
	Response.Write("<pre>MaxQuarter=" & MaxQuarter & "</pre>")
	Response.Flush
End If	
If Quarter = 1 Then
	StartDate = CDate("12/1/" & (FiscalYear-1))
ElseIf Quarter = 2 Then
	StartDate = CDate("3/1/" & FiscalYear)
ElseIf Quarter = 3 Then
	StartDate = CDate("6/1/" & FiscalYear)
ElseIf Quarter = 4 Then
	StartDate = CDate("9/1/" & FiscalYear)
End If
Columns = 4 + Quarter*3

sql = "SELECT A.GrantID, A.GrantClassID, A.DisplayQuarterOffset, ISNULL(B.Quarter," & Quarter & ") AS Quarter, " & vbCrLf & _
	"	CAST(CASE WHEN SubmitID IS NOT NULL THEN ISNULL(B.Confirmed,0) ELSE 0 END AS BIT) AS Confirmed, " & vbCrLf & _
	"	B.SubmitID, B.SubmitTimestamp, C.Name AS SubmitName, " & vbCrLF & _
	"	B.AdministrativeComments, B.ApprovalID, B.ApprovalDate, D.Name AS ApprovalName, " & vbCrLf & _
	"	CAST(CASE WHEN " & UserSystemID & " IN (E.ProgramDirectorID, E.ProgramManagerID) THEN 1 ELSE 0 END AS BIT) AS CanSubmit " & vbCrLf & _
	"FROM [Grants].Main AS A " & vbCrLf & _
	"LEFT JOIN PR.Main AS B ON B.GrantID=A.GrantID AND Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLF & _
	"LEFT JOIN [System].Users AS C ON C.SystemID=B.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS D ON D.SystemID=B.ApprovalID " & vbCrLf & _
	"LEFT JOIN Grantees AS E ON E.GranteeID=A.GranteeID " & vbCrLf & _
	"WHERE A.GrantID=" & prepIntegerSQL(GrantID) 
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF Then
	Response.Write("Error: Progress Report Record not found.")
	Response.End
Else
	GrantClassID = rs.Fields("GrantClassID")
	DisplayQuarterOffset = rs.Fields("DisplayQuarterOffset")
	Quarter = rs.Fields("Quarter")
	Confirmed = rs.Fields("Confirmed")
	SubmitID = rs.Fields("SubmitID")
	SubmitName = rs.Fields("SubmitName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	AdministrativeComments = rs.Fields("AdministrativeComments")
	ApprovalID = rs.Fields("ApprovalID")
	ApprovalDate = rs.Fields("ApprovalDate")
	ApprovalName = rs.Fields("ApprovalName")
	CanSubmit = rs.Fields("CanSubmit")
End If

Version = PRVersion(GrantClassID, FiscalYear)

If GranteeID>0 Then
	'If FiscalYear=2020 Then
	'	PermitEdit = CheckPermissions(UserSystemID, GranteeID, True)
	'Else
	If IsNull(SubmitID) = True Then
		PermitEdit = CheckPermissions(UserSystemID, GranteeID, True)
	ElseIf IsNull(SubmitID) = False Then
		PermitEdit = False
	Else
		PermitEdit = False
	End If
Else
	PermitEdit = False
End If
' PermitEdit=True ' For testing.
'ViewDocuments = CheckPermissions(UserSystemID, GranteeID, True)
ViewDocuments = True

If Debug = True Then
	Response.Write("<pre>PermitEdit=" & PermitEdit & ", CheckPermissions=" & _
	CheckPermissions(UserSystemID, GranteeID, True) & "; CanSubmit=" & CanSubmit & _
	"; StartDate=" & StartDate & "; CurrentDate=" & CurrentDate & "; ShowOneQuarter=" & ShowOneQuarter & "</pre>")
	Response.Flush
End If	

sql = "SELECT A.QuestionID, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.NoTarget, A.ResponseTypeID, " & vbCrLf & _
	"	Q.IntegerTarget, Q.DecimalTarget, " & vbCrLf & _
	"	IntegerResponse_Sep, IntegerResponse_Oct, IntegerResponse_Nov, " & vbCrLf & _
	"	IntegerResponse_Dec, IntegerResponse_Jan, IntegerResponse_Feb, " & vbCrLf & _
	"	IntegerResponse_Apr, IntegerResponse_May, IntegerResponse_Jun, " & vbCrLf & _
	"	IntegerResponse_Mar, IntegerResponse_Jul, IntegerResponse_Aug, " & vbCrLf & _
	"	DecimalResponse_Sep, DecimalResponse_Oct, DecimalResponse_Nov, " & vbCrLf & _
	"	DecimalResponse_Dec, DecimalResponse_Jan, DecimalResponse_Feb, " & vbCrLf & _
	"	DecimalResponse_Mar, DecimalResponse_Apr, DecimalResponse_May, " & vbCrLf & _
	"	DecimalResponse_Jun, DecimalResponse_Jul, DecimalResponse_Aug, " & vbCrLf & _
	"	TextResponse_Q1, TextResponse_Q2, TextResponse_Q3, TextResponse_Q4 " & vbCrLf & _
	"FROM PR.Goals AS G " & vbCrLf & _
	"LEFT JOIN PR.Strategies AS S ON S.GoalID=G.GoalID AND S.Version=G.Version " & vbCrLf & _
	"LEFT JOIN PR.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID AND A.Version=G.Version " & vbCrLf & _
	"LEFT JOIN PR.GrantQuestions AS Q ON Q.GrantID=" & prepIntegerSQL(GrantID) & " AND Q.QuestionID=A.QuestionID " & vbCrLf & _
	"LEFT JOIN PR.Responses AS R ON R.GrantID=" & prepIntegerSQL(GrantID) & " AND R.QuestionID=A.QuestionID " & vbCrLf & _
	"WHERE Q.GrantID IS NOT NULL AND G.Version=" & prepIntegerSQL(Version) & " " & vbCrLF & _
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
<title>MVCPA Progress Report for <%=GranteeName %> <%=ProgramName %></title>
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
<body style="width: 100%">
<div style="text-align: center;"><form name="Selection" method="post" action="Report.asp"><%=HiddenField("GrantID",GrantID) %><%=HiddenField("FiscalYear",FiscalYear) %>
<select name="Quarter" onchange="document.Selection.submit();">
<%
For i = 1 to MaxQuarter
	If GrantClassID=4 And FiscalYear=2024 Then
		Response.Write(SelectOption(i, "Q" & i, Quarter))
	Else
		Response.Write(SelectOption(i, "Quarter " & i & ": " & ReportingPeriodDates(FiscalYear, (i+DisplayQuarterOffset)), Quarter))
	End If
Next
%>
	</select>&nbsp;&nbsp;&nbsp;<input type="checkbox" name="ShowOneQuarter" value="1" <%=Checked(ShowOneQuarter, True) %> onclick="document.Selection.submit();">Show only one quarter
</form></div>
<h1><%=GranteeName %> MVCPA Progress Report for Fiscal Year <%=FiscalYear %>, Quarter <%=Quarter %></h1>
<!--<h2>Goals, Strategies, and Activities</h2>-->
<%	If SubmitID>0 Then %>
<p style="text-align: center; font-weight: bold; ">The progress report was submitted by <%=SubmitName%> at <%=SubmitTimestamp %> and is now locked.</p>
<%	End If %>
<form name="PR" method="post" action="ReportSubmit.asp">
<%=HiddenField("GrantID", GrantID) %><%=HiddenField("Quarter", Quarter) %><%=HiddenField("action","save") %><%=HiddenField("ShowOneQuarter",ShowOneQuarter) %><%=HiddenField("Version",Version) %>
<table style="margin: auto">
<thead>
<%	End If %>
	<tr>
		<th>ID</th>
		<th>Activity</th>
		<th>Measure</th>
		<th>Target</th>
<%	If Quarter=1 Or ShowOneQuarter=False Then %>
		<th>September</th>
		<th>October</th>
		<th>November</th>
<%	End If
	If (Quarter>1 And ShowOneQuarter=False) Or Quarter=2 Then %>
		<th>December</th>
		<th>January</th>
		<th>February</th>
<%	End If 
	If (Quarter>2 And ShowOneQuarter=False) Or Quarter=3 Then %>
		<th>March</th>
		<th>April</th>
		<th>May</th>
<%	End If 
	If (Quarter>3 And ShowOneQuarter=False) Or Quarter=4 Then %>
		<th>June</th>
		<th>July</th>
		<th>August</th>
<%	End If %>
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
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Statutory Motor Vehicle Theft (MVT) Measures Required for ALL Grantees.</th></tr>" & vbCrLf)
		ElseIf rs.Fields("GoalID")=2 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Statutory Burglary of a Motor Vehicle (BMV) Measures Required for ALL Grantees</th></tr>" & vbCrLf)
		ElseIf rs.Fields("GoalID")=5 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "">Statutory Apprehension Questions Required for ALL Grantees</th></tr>" & vbCrLf)
		ElseIf rs.Fields("GoalID")=6 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "">Quarterly Summary Questions Required for ALL Grantees</th></tr>" & vbCrLf)
		ElseIf rs.Fields("GoalID")=8 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Fraud-Related Motor Vehicle Crime Measures must be reported during the grant term if awarded."">Statutory Fraud-Related Motor Vehicle Crime (FRMVC) Measures Required for ALL Grantees</th></tr>" & vbCrLf)
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
	'If rs.Fields("Mandatory") Then
	'	If rs.Fields("IntegerTarget") = 1 Then
	'		Response.Write(" (Taskforce Only)")
	'	ElseIf rs.Fields("IntegerTarget") = 2 Then
	'		Response.Write(" (Area of Jurisdiction)")
	'	ElseIf rs.Fields("IntegerTarget") = 3 Then
	'		Response.Write( " (Combination of Taskforce and Jurisdiction)")
	'	End If
	'End If
	Response.Write("</td>" & vbCrLf)

	' Target cell
	If rs.Fields("Mandatory") Then
		Response.Write("<td>Required</td>")
	ElseIf rs.Fields("NoTarget") Then
		Response.Write("<td style=""background-color: #e6e6e6; text-align: center; "">No Target</td>" & vbCrLf)
	ElseIf rs.Fields("ResponseTypeID")=1 Then ' Integer
		If IsNull(rs.Fields("IntegerTarget")) Then
			Response.Write("<td></td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right;"">" & formatnumber(rs.Fields("IntegerTarget"),0,true, false, true) & "</td>" & vbCrLf)
		End If
	ElseIf rs.Fields("ResponseTypeID")=2 Then ' Decimal
		If IsNull(rs.Fields("DecimalTarget")) Then
			Response.Write("<td></td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right;"">" & formatnumber(rs.Fields("DecimalTarget"),2) & "</td>" & vbCrLf)
		End If
	ElseIf rs.Fields("ResponseTypeID")=3 Then ' Money
		If IsNull(rs.Fields("DecimalTarget")) Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right;"">" & formatcurrency(rs.Fields("DecimalTarget"),2, True, False, True) & "</td>" & vbCrLf)
		End If
	Else
		Response.Write(vbTab & "<td></td>" & vbCrLf)
	End If

	' Response Cells
	If rs.Fields("ResponseTypeID")=1 Then ' Integer
		If ShowExcel = True Then
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Sep")) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Oct")) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Nov")) & "</td>" & vbCrLf)
		ElseIf Quarter = 1 Then
			Response.Write(vbTab & "<td>" & IntegerField("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Sep"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & IntegerField("Response_Oct_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Oct"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & IntegerField("Response_Nov_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Nov"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
		ElseIf ShowOneQuarter=False And Quarter>1 Then
			Response.Write(vbTab & "<td>" & IntegerField("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Sep"), 8, 12, False, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & IntegerField("Response_Oct_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Oct"), 8, 12, False, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & IntegerField("Response_Nov_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Nov"), 8, 12, False, "") & "</td>" & vbCrLf)
		End If
		If Quarter > 1 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Dec")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Jan")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Feb")) & "</td>" & vbCrLf)
			ElseIf Quarter = 2 Then
				Response.Write(vbTab & "<td>" & IntegerField("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Dec"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Jan_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jan"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Feb_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Feb"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & IntegerField("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Dec"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Jan_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jan"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Feb_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Feb"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
		If Quarter > 2 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Mar")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Apr")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_May")) & "</td>" & vbCrLf)
			ElseIf Quarter = 3 Then
				Response.Write(vbTab & "<td>" & IntegerField("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Mar"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Apr_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Apr"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_May_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_May"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & IntegerField("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Mar"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Apr_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Apr"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_May_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_May"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
		If Quarter > 3 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Jun")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Jul")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_Aug")) & "</td>" & vbCrLf)
			ElseIf Quarter = 4 Then
				Response.Write(vbTab & "<td>" & IntegerField("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jun"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Jul_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jul"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Aug_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Aug"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & IntegerField("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jun"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Jul_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jul"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & IntegerField("Response_Aug_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Aug"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
	ElseIf rs.Fields("ResponseTypeID")=2 Then ' Decimal
		If ShowExcel = True Then
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Sep"),2) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Oct"),2) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Nov"),2) & "</td>" & vbCrLf)
		ElseIf Quarter = 1 Then
			Response.Write(vbTab & "<td>" & NumberField("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Sep"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & NumberField("Response_Oct_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Oct"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & NumberField("Response_Nov_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Nov"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
		ElseIf ShowOneQuarter = False Then
			Response.Write(vbTab & "<td>" & NumberField("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Sep"), 8, 12, False, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & NumberField("Response_Oct_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Oct"), 8, 12, False, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & NumberField("Response_Nov_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Nov"), 8, 12, False, "") & "</td>" & vbCrLf)
		End If
		If Quarter > 1 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Dec"),2) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Jan"),2) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Feb"),2) & "</td>" & vbCrLf)
			ElseIf Quarter = 2 Then
				Response.Write(vbTab & "<td>" & NumberField("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Dec"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Jan_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jan"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Feb_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Feb"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & NumberField("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Dec"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Jan_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jan"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Feb_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Feb"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
		If Quarter > 2 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Mar"),2) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Apr"),2) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_May"),2) & "</td>" & vbCrLf)
			ElseIf Quarter = 3 Then
				Response.Write(vbTab & "<td>" & NumberField("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Mar"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Apr_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Apr"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_May_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_May"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & NumberField("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Mar"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Apr_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Apr"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_May_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_May"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
		If Quarter > 3 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Jun"),2) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Jul"),2) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_Aug"),2) & "</td>" & vbCrLf)
			ElseIf Quarter = 4 Then
				Response.Write(vbTab & "<td>" & NumberField("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jun"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Jul_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jul"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Aug_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Aug"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & NumberField("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jun"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Jul_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jul"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & NumberField("Response_Aug_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Aug"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
	ElseIf rs.Fields("ResponseTypeID")=3 Then ' dollars and cents
		If ShowExcel = True Then
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Sep")) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Oct")) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Nov")) & "</td>" & vbCrLf)
		ElseIf Quarter = 1 Then
			Response.Write(vbTab & "<td>" & CurrencyField("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Sep"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & CurrencyField("Response_Oct_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Oct"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & CurrencyField("Response_Nov_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Nov"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
		ElseIf ShowOneQuarter=False Then
			Response.Write(vbTab & "<td>" & CurrencyField("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Sep"), 8, 12, False, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & CurrencyField("Response_Oct_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Oct"), 8, 12, False, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & CurrencyField("Response_Nov_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Nov"), 8, 12, False, "") & "</td>" & vbCrLf)
		End If
		If Quarter > 1 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Dec")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Jan")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Feb")) & "</td>" & vbCrLf)
			ElseIf Quarter = 2 Then
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Dec"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Jan_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jan"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Feb_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Feb"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Dec"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Jan_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jan"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Feb_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Feb"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
		If Quarter > 2 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Mar")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Apr")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_May")) & "</td>" & vbCrLf)
			ElseIf Quarter = 3 Then
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Mar"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Apr_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Apr"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_May_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_May"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Mar"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Apr_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Apr"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_May_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_May"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
		If Quarter > 3 Then
			If ShowExcel = True Then
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Jun")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Jul")) & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_Aug")) & "</td>" & vbCrLf)
			ElseIf Quarter = 4 Then
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jun"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Jul_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jul"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Aug_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Aug"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jun"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Jul_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Jul"), 8, 12, False, "") & "</td>" & vbCrLf)
				Response.Write(vbTab & "<td>" & CurrencyField("Response_Aug_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_Aug"), 8, 12, False, "") & "</td>" & vbCrLf)
			End If
		End If
	ElseIf rs.Fields("ResponseTypeID")=5 Then ' Text area response. One per quarter.
		If ShowExcel = True Then
			Response.Write("<td colspan=""3"" style=""text-align: left; border: 1px solid black;"">" & prepStringWeb(rs.Fields("TextResponse_Q1")) & "</td>" & vbCrLf)
		ElseIf Quarter=1 Then
			Response.Write("<td colspan=""3"">" & TextArea("Response_Q1_" & rs.Fields("QuestionID"), rs.Fields("TextResponse_Q1"), 20, 60, 32000, PermitEdit, "") & "</td>" & vbCrLf)
		ElseIf ShowOneQuarter=False Then
			Response.Write("<td colspan=""3"">" & TextArea("Response_Q1_" & rs.Fields("QuestionID"), rs.Fields("TextResponse_Q1"), 20, 60, 32000, False, "") & "</td>" & vbCrLf)
		End If
		If Quarter > 1 Then
			If ShowExcel = True Then
				Response.Write("<td colspan=""3"" style=""text-align: left; border: 1px solid black;"">" & prepStringWeb(rs.Fields("TextResponse_Q2")) & "</td>" & vbCrLf)
			ElseIf Quarter=2 Then
				Response.Write("<td colspan=""3"">" & TextArea("Response_Q2_" & rs.Fields("QuestionID"), rs.Fields("TextResponse_Q2"), 20, 60, 32000, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write("<td colspan=""3"">" & TextArea("Response_Q2_" & rs.Fields("QuestionID"), rs.Fields("TextResponse_Q2"), 20, 60, 32000, False, "") & "</td>" & vbCrLf)
			End If
		End If
		If Quarter > 2 Then
			If ShowExcel = True Then
				Response.Write("<td colspan=""3"" style=""text-align: left; border: 1px solid black;"">" & prepStringWeb(rs.Fields("TextResponse_Q3")) & "</td>" & vbCrLf)
			ElseIf Quarter=3 Then
				Response.Write("<td colspan=""3"">" & TextArea("Response_Q3_" & rs.Fields("QuestionID"), rs.Fields("TextResponse_Q3"), 20, 60, 32000, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write("<td colspan=""3"">" & TextArea("Response_Q3_" & rs.Fields("QuestionID"), rs.Fields("TextResponse_Q3"), 20, 60, 32000, False, "") & "</td>" & vbCrLf)
			End If
		End If
		If Quarter > 3 Then
			If ShowExcel = True Then
				Response.Write("<td colspan=""3"" style=""text-align: left; border: 1px solid black;"">" & prepStringWeb(rs.Fields("TextResponse_Q4")) & "</td>" & vbCrLf)
			ElseIf Quarter=4 Then
				Response.Write("<td colspan=""3"">" & TextArea("Response_Q4_" & rs.Fields("QuestionID"), rs.Fields("TextResponse_Q4"), 20, 60, 32000, PermitEdit, "") & "</td>" & vbCrLf)
			ElseIf ShowOneQuarter=False Then
				Response.Write("<td colspan=""3"">" & TextArea("Response_Q4_" & rs.Fields("QuestionID"), rs.Fields("TextResponse_Q4"), 20, 60, 32000, False, "") & "</td>" & vbCrLf)
			End If
		End If
	ElseIf rs.Fields("ResponseTypeID") = 6 Then ' one yes/no question
		If Quarter = 1 Then
			Response.Write(vbTab & yesnoradio("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Sep"), 3, PermitEdit, ShowExcel) & vbCrLf)
		ElseIf ShowOneQuarter=False Then
			Response.Write(vbTab & yesnoradio("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Sep"), 3, False, ShowExcel) & vbCrLf)
		End If
		If Quarter = 2 Then
			Response.Write(vbTab & yesnoradio("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Dec"), 3, PermitEdit, ShowExcel) & vbCrLf)
		ElseIf Quarter > 2 And ShowOneQuarter=False Then
			Response.Write(vbTab & yesnoradio("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Dec"), 3, False, ShowExcel) & vbCrLf)
		End If
		If Quarter = 3 Then
			Response.Write(vbTab & yesnoradio("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Mar"), 3, PermitEdit, ShowExcel) & vbCrLf)
		ElseIf Quarter > 3 And ShowOneQuarter=False Then
			Response.Write(vbTab & yesnoradio("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Mar"), 3, False, ShowExcel) & vbCrLf)
		End If
		If Quarter = 4 Then
			Response.Write(vbTab & yesnoradio("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jun"), 3, PermitEdit, ShowExcel) & vbCrLf)
		ElseIf Quarter > 3 And ShowOneQuarter=False Then
			Response.Write(vbTab & yesnoradio("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jun"), 3, False, ShowExcel) & vbCrLf)
		End If
	ElseIf rs.Fields("ResponseTypeID") = 7 Then ' a set of three yes/no questions
		If Quarter = 1 Then
			Response.Write(vbTab & yesnoradio("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Sep"), 1, PermitEdit, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Oct_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Oct"), 1, PermitEdit, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Nov_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Nov"), 1, PermitEdit, ShowExcel) & vbCrLf)
		ElseIf ShowOneQuarter=False Then
			Response.Write(vbTab & yesnoradio("Response_Sep_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Sep"), 1, False, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Oct_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Oct"), 1, False, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Nov_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Nov"), 1, False, ShowExcel) & vbCrLf)
		End If
		If Quarter = 2 Then
			Response.Write(vbTab & yesnoradio("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Dec"), 1, PermitEdit, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Jan_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jan"), 1, PermitEdit, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Feb_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Feb"), 1, PermitEdit, ShowExcel) & vbCrLf)
		ElseIf Quarter > 2 and ShowOneQuarter = False Then
			Response.Write(vbTab & yesnoradio("Response_Dec_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Dec"), 1, False, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Jan_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jan"), 1, False, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Feb_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Feb"), 1, False, ShowExcel) & vbCrLf)
		End If
		If Quarter = 3 Then
			Response.Write(vbTab & yesnoradio("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Mar"), 1, PermitEdit, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Apr_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Apr"), 1, PermitEdit, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_May_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_May"), 1, PermitEdit, ShowExcel) & vbCrLf)
		ElseIf Quarter > 3 And ShowONeQuarter=False Then
			Response.Write(vbTab & yesnoradio("Response_Mar_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Mar"), 1, False, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Apr_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Apr"), 1, False, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_May_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_May"), 1, False, ShowExcel) & vbCrLf)
		End If
		If Quarter = 4 Then
			Response.Write(vbTab & yesnoradio("Response_Jun_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jun"), 1, PermitEdit, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Jul_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Jul"), 1, PermitEdit, ShowExcel) & vbCrLf)
			Response.Write(vbTab & yesnoradio("Response_Aug_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_Aug"), 1, PermitEdit, ShowExcel) & vbCrLf)
		End If
	End If
	Response.Write("</tr>" & vbCrLf)
	rs.MoveNext()
Wend
Response.Write("</table>" & vbCrLf)
If ShowExcel = False Then
	Response.Write("<hr />" & vbCrLf)

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
					If Left(file.Name,3)="PR"&Quarter Then
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
%>

<br />
<p><%=CheckBoxField2("Confirmed", Confirmed, PermitEdit) %>
I have reviewed and confirmed the information in this report and I attest that this 
report is correct and complete and supported by documentation for purposes set forth 
in the Statement of Grant Award. I am aware that any false, fictitious, or fraudulent 
information may subject me to criminal, civil, or administrative penalties. </p>

<%
	If MVCPARights = True Or MVCPAViewer = True Then
		Response.Write("<table style=""margin: auto; "">" & vbCrLF)
		Response.Write("<tr><td colspan=""2"">Administrative Comments:<br />" & vbCrLf)
		Response.Write(TextArea2("AdministrativeComments", AdministrativeComments, 6, 920, 8000, MVCPARights, "") & "</td></tr>" & vbCrLf)
		If IsNull(SubmitID) = False Then
			Response.Write("<tr><td>MVCPA Approval Date:</td>" & vbCrLf)
			Response.Write("<td>" & DateField("ApprovalDate", ApprovalDate, MVCPARights))
			If IsNull(ApprovalName) = False Then
				Response.Write(" by " & ApprovalName)
			End If
			Response.Write("</td>")
			If IsNull(ApprovalDate)= True Or MVCPAAdministrator=True Then
				Response.Write("<tr><td colspan=""2"">" &  CheckBoxField("Unsubmit", False) & " Unsubmit Progress Report (Clears submission and approval.)</td></tr>" & vbCrLf)
			End If
		End If
		Response.Write("</table>" & vbCrLf)
	End If
'CanSubmit = True

If MVCPARights = True Or MVCPAViewer = True Then
	sql = "SELECT A.SubmitTimestamp, B.Name AS SubmitName " & vbCrLf & _
		"FROM PR.Main AS A " & vbCrLf & _
		"LEFT JOIN [System].Users AS B ON A.SubmitID=B.SystemID " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND SubmitTimestamp IS NOT NULL " & vbCrLf & _
		"UNION " & vbCrLf & _
		"SELECT A.SubmitTimestamp, B.Name AS SubmitName " & vbCrLf & _
		"FROM PR.Main_Log AS A " & vbCrLf & _
		"LEFT JOIN [System].Users AS B ON A.SubmitID=B.SystemID " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND SubmitTimestamp IS NOT NULL " & vbCrLf & _
		"ORDER BY 1 DESC "
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		Response.Write("<br><table style=""margin: auto; "">" & vbCrLf)
		Response.WRite("<thead><tr><th>Submission History</th></tr></thead>" & vbCrLf & "<tbody>" & vbCrLf)
		While rs.EOF = False
			Response.Write("<tr><td style=""text-align: center"">Submitted By " & rs.Fields("SubmitName") & ", " & rs.Fields("SubmitTimestamp") & "</td></tr>" & vbCrLf)
			rs.MoveNext()
		Wend
		Response.Write("</tbody>" & vbCrLf & "</table><br />" & vbCrLf)
	Else
		Response.Write("<div style=""text-align: center; margin: auto; "">No Submissions</div>" & vbCrLf)
	End If
End If
%><br />

<div style="text-align: center; margin: auto; ">
<%	If PermitEdit = True Or MVCPARights = True Then %>
	<input type="submit" name="Save" value="Save" title="Save changes and return to progress report" onclick="return submitForm('save');" />
<%	End If %>
<%	If CanSubmit = True And CurrentDate>=StartDate And IsNull(SubmitID) = True Then %>
	<input type="button" name="Submit" value="Submit" title="Save changes and submit the progress report. After submission, the report will no longer be editable." onclick="return submitForm('submit');"/>
<%	End If %>
	<input type="button" name="Close" value="Close" title="Ignore any pending changes and close window." 
		onclick="window.close();"/>
</div>
<div style="text-align: right"><a href="Report.asp?GrantID=<%=GrantID %>&Quarter=<%=Quarter %>&ShowExcel=1" target="_blank">Excel</a></div>
</form>
</body>
</html>
<%
End If

Function yesnoradio(vname, vvalue, vcolumns, veditable, vexcel)
	If vexcel = True Then
		If vvalue = 1 Then 
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: 1px solid black; "">Yes</td>"
		ElseIf vvalue = 2 Then
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: 1px solid black; "">No</td>"
		Else
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: 1px solid black; "">Yes or No</td>"
		End If
	ElseIf veditable = True Then
		If vvalue = 1 Then
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: ""><input type=""radio"" name=""" & vname & """ value=""1"" checked>Yes " & vbCrLf & _
				"<input type=""radio"" name=""" & vname & """ value=""2"">No</td>"
		ElseIf vvalue = 2 Then
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: ""><input type=""radio"" name=""" & vname & """ value=""1"">Yes " & vbCrLf & _
				"<input type=""radio"" name=""" & vname & """ value=""2"" checked>No</td>"
		Else
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: ""><input type=""radio"" name=""" & vname & """ value=""1"">Yes " & vbCrLf & _
				"<input type=""radio"" name=""" & vname & """ value=""2"">No</td>"
		End If
	Else
		If vvalue = 1 Then
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: "">Yes" & HiddenField(vname, vvalue) & "</td>"
		ElseIf vvalue = 2 Then
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: "">No" & HiddenField(vname, vvalue) & "</td>"
		Else
			yesnoradio = "<td colspan=""" & vcolumns & """ style=""text-align: center;border: "">Yes or No" & HiddenField(vname, vvalue) & "</td>"
		End If
	End If
End Function

Function ReportingPeriodDates(vFiscalYear, vReportingPeriod)
	If vReportingPeriod = 1 Then
		ReportingPeriodDates = "September 1, " & (vFiscalYear-1) & " - November 30, " & (vFiscalYear-1)
	ElseIf vReportingPeriod = 2 and (vFiscalYear Mod 4 = 0) Then
		ReportingPeriodDates = "December 1, " & (vFiscalYear-1) & " - February 29, " & vFiscalYear
	ElseIf vReportingPeriod = 2 Then
		ReportingPeriodDates = "December 1, " & (vFiscalYear-1) & " - February 28, " & vFiscalYear
	ElseIf vReportingPeriod = 3 Then
		ReportingPeriodDates = "March 1, " & vFiscalYear & " - May 31, " & vFiscalYear
	ElseIf vReportingPeriod = 4 Then
		ReportingPeriodDates = "June 1, " & vFiscalYear & " - August 31, " & vFiscalYear
	ElseIf vReportingPeriod = 5 Then
		ReportingPeriodDates = "September 1, " & vFiscalYear & " - November 30, " & vFiscalYear
	ElseIf vReportingPeriod = 6 and (vFiscalYear Mod 4 = 0) Then
		ReportingPeriodDates = "December 1, " & vFiscalYear & " - February 29, " & (vFiscalYear+1)
	ElseIf vReportingPeriod = 6 Then
		ReportingPeriodDates = "December 1, " & vFiscalYear & " - February 28, " & (vFiscalYear+1)
	ElseIf vReportingPeriod = 7 Then
		ReportingPeriodDates = "March 1, " & (vFiscalYear+1) & " - May 31, " & (vFiscalYear+1)
	ElseIf vReportingPeriod = 8 Then
		ReportingPeriodDates = "June 1, " & (vFiscalYear+1) & " - August 31, " & (vFiscalYear+1)
	Else
		ReportingPeriodDates = "Error in Reporting Period"
	End If
End Function
%>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../ProgressReport/PRVersionInclude.asp"-->