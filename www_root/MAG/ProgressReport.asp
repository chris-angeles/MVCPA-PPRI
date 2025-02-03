<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, k, ViewDocuments, PermitEdit, ShowExcel, Columns, CurrentDate, _
	Quarter, MaxQuarter, QuarterDescription, Version, MAGID, _
	StartDate, LastGoal, LastStrategy, LastMandatory,  Confirmed, _
	FiscalYear, GranteeID, GranteeName, BorderCounty,PortCounty, ProgramName, _
	SubmitID, SubmitName, SubmitTimestamp, _
	AdministrativeComments, ApprovalID, ApprovalDate, ApprovalName, CanSubmit
debug = False

CurrentDate = Date() 
'CurrentDate = cdate("3/2/2018")
QuarterDescription = Array("", "September 1 - November 30, 2022", "December 1, 2022 - February 28, 2023", _
	"March 1 - May 31, 2023", "June 1 - August 31, 2023", "September 1 - November 30, 2023", _
	"December 1, 2023 - February 28, 2024", "March 1 - May 31, 2024", "June 1 - August 31, 2024")

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
	MAGID = Request.Form("MAGID")
Else
	MAGID = Request.QueryString("MAGID")
End If

IF Len(MAGID)>0 Then
	MAGID = CInt(MAGID)
Else
	MAGID=0
End If

If Request.Querystring("ShowExcel")="1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "SELECT H.MAGID, H.FiscalYear, G.GranteeID, G.GranteeName, 'MVCPA Auxiliary Grant' AS ProgramName, " & vbCrLf & _
	"	ISNULL(G.BorderCounty,0) AS BorderCounty, ISNULL(G.PortCounty,0) AS PortCounty " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN [MAG].Main AS H ON G.GranteeID=H.GranteeID " & vbCrLf & _
	"WHERE H.MAGID=" & prepIntegerSQL(MAGID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF Then
	Response.Write("Error: Grant not found.")
	Response.End
Else
	MAGID = rs.Fields("MAGID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	ProgramName = rs.Fields("ProgramName")
	BorderCounty = rs.Fields("BorderCounty")
	PortCounty = rs.Fields("PortCounty")
End If

' Determine current Reporting Period - Keep old quarter for one month.
If Len(Request.Form("Quarter"))>0 Then
	Quarter = CInt(Request.Form("Quarter"))
ElseIf Len(Request.QueryString("Quarter"))>0 Then
	Quarter = CInt(Request.QueryString("Quarter"))
ElseIf CurrentDate < CDate("1/1/" & (FiscalYear+1)) Then
	Quarter = 1
ElseIf CurrentDate < CDate("4/1/" & (FiscalYear+1)) Then
	Quarter = 2
ElseIf CurrentDate < CDate("7/1/" & (FiscalYear+1)) Then
	Quarter = 3
ElseIf CurrentDate < CDate("10/1/" & (FiscalYear+1)) Then
	Quarter = 4
ElseIf CurrentDate < CDate("1/1/" & (FiscalYear+2)) Then
	Quarter = 5
ElseIf CurrentDate < CDate("4/1/" & (FiscalYear+2)) Then
	Quarter = 6
ElseIf CurrentDate < CDate("7/1/" & (FiscalYear+2)) Then
	Quarter = 7
Else
	Quarter = 8
End If

' Determine Max Reporting Period to select
If CurrentDate >= CDate("6/1/" & (FiscalYear+2)) Then
	MaxQuarter = 8
ElseIf CurrentDate >= CDate("3/1/" & (FiscalYear+2)) Then
	MaxQuarter = 7
ElseIf CurrentDate >= CDate("12/1/" & (FiscalYear+1)) Then
	MaxQuarter = 6
ElseIf CurrentDate >= CDate("9/1/" & (FiscalYear+1)) Then
	MaxQuarter = 5
ElseIf CurrentDate >= CDate("6/1/" & (FiscalYear+1)) Then
	MaxQuarter = 4
ElseIf CurrentDate >= CDate("3/1/" & (FiscalYear+1)) Then
	MaxQuarter = 3
ElseIf CurrentDate >= CDate("3/1/" & (FiscalYear+1)) Then
	MaxQuarter = 3
ElseIf CurrentDate > CDate("12/1/" & (FiscalYear)) Then
	MaxQuarter = 2
Else
	MaxQuarter = 1
End If


If Debug = True Then
	Response.Write("<pre>Quarter=" & Quarter & "</pre>")
	Response.Write("<pre>CurrentDate=" & CurrentDate & "</pre>")
	Response.Write("<pre>FiscalYear=" & FiscalYear & "</pre>")
	Response.Write("<pre>MaxQuarter=" & MaxQuarter & "</pre>")
	Response.Write("<pre>BorderCounty=" & BorderCounty & "</pre>")
	Response.Write("<pre>PortCounty=" & PortCounty & "</pre>")
	Response.Flush
End If	
If Quarter = 1 Then
	StartDate = CDate("9/1/" & (FiscalYear-1))
ElseIf Quarter = 2 Then
	StartDate = CDate("12/1/" & (FiscalYear-1))
ElseIf Quarter = 3 Then
	StartDate = CDate("3/1/" & FiscalYear)
ElseIf Quarter = 4 Then
	StartDate = CDate("6/1/" & FiscalYear)
End If
Columns = 3 + 3

sql = "SELECT A.MAGID, ISNULL(B.Quarter," & Quarter & ") AS Quarter, " & vbCrLf & _
	"	B.Confirmed, " & vbCrLf & _
	"	B.SubmitID, B.SubmitTimestamp, C.Name AS SubmitName, " & vbCrLF & _
	"	B.AdministrativeComments, B.ApprovalID, B.ApprovalDate, D.Name AS ApprovalName, " & vbCrLf & _
	"	CAST(CASE WHEN " & UserSystemID & " IN (E.ProgramDirectorID, E.ProgramManagerID) THEN 1 ELSE 0 END AS BIT) AS CanSubmit " & vbCrLf & _
	"FROM [MAG].Main AS A " & vbCrLf & _
	"LEFT JOIN [MAG].ProgressReportSubmissions AS B ON B.MAGID=A.MAGID AND Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLF & _
	"LEFT JOIN [System].Users AS C ON C.SystemID=B.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS D ON D.SystemID=B.ApprovalID " & vbCrLf & _
	"LEFT JOIN Grantees AS E ON E.GranteeID=A.GranteeID " & vbCrLf & _
	"WHERE A.MAGID=" & prepIntegerSQL(MAGID) 
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF Then
	Response.Write("Error: Progress Report Record not found.")
	Response.End
Else
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

If FiscalYear>2021 Then 
	Version = 5
ElseIf FiscalYear>2020 Then
	Version = 4
ElseIf FiscalYear>2019 Then
	Version = 3
Else
	Version = 2
End If

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
	"; StartDate=" & StartDate & "; CurrentDate=" & CurrentDate & "</pre>")
	Response.Flush
End If	

sql = "WITH CTE AS (" & vbCrLf & _
	"	SELECT A.QuestionID, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"		CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"			CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"		G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.MAGSpecial, A.ResponseTypeID, " & vbCrLf & _
	"		IntegerResponse_M1, IntegerResponse_M2, IntegerResponse_M3, " & vbCrLf & _
	"		DecimalResponse_M1, DecimalResponse_M2, DecimalResponse_M3, " & vbCrLf & _
	"		TextResponse, " & vbCrLf & _
	"		CAST(CASE WHEN R.MAGID IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS RecordPresent " & vbCrLf & _
	"	FROM PR.Goals AS G " & vbCrLf & _
	"	LEFT JOIN PR.Strategies AS S ON S.GoalID=G.GoalID AND S.Version=G.Version " & vbCrLf & _
	"	LEFT JOIN PR.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID AND A.Version=G.Version " & vbCrLf & _
	"	LEFT JOIN MAG.ProgressReportResponses AS R ON R.MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND R.QuestionID=A.QuestionID " & vbCrLf & _
	"	WHERE G.Version=" & prepIntegerSQL(Version) & " " & vbCrLf & _
	") SELECT * FROM CTE " & vbCrLf
	If BorderCounty = True Or PortCounty = True Then
		sql = sql & "WHERE (ISNULL(Mandatory,0)=1 OR GoalID IN (4,7) OR ISNULL(MAGSpecial,0)=1)" & vbCrLf & _
			"	AND MeasureNumber NOT IN ('7.1.1', '7.1.4') " & vbCrLf
	Else
		sql = sql & "WHERE (ISNULL(Mandatory,0)=1 OR GoalID IN (7) OR ISNULL(MAGSpecial,0)=1) " & vbCrLf & _
			"	AND MeasureNumber NOT IN ('7.1.1', '7.1.4') " & vbCrLf
	End If
	sql = sql & "ORDER BY Mandatory DESC, GoalID, StrategyID, ActivityID, MeasureID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ProgressReport" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr><th colspan=""" & columns & """>" & GranteeName & " MVCPA Auxiliary Grant Program Progress Report for Fiscal Year " & FiscalYear & ", Quarter " & Quarter & " of Grant</th></tr>" & vbCrLf)
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
<div style="text-align: center;"><form name="Selection" method="post" action="ProgressReport.asp"><%=HiddenField("MAGID",MAGID) %><%=HiddenField("FiscalYear",FiscalYear) %>
<select name="Quarter" onchange="document.Selection.submit();">
<%
For i = 1 to MaxQuarter
	Response.Write(SelectOption(i, "Quarter " & i & ": " & QuarterDescription(i), Quarter))
Next
%>
	</select>
</form></div>
<h1><%=GranteeName %> MVCPA Progress Report for Fiscal Year <%=FiscalYear %>, Quarter <%=Quarter %></h1>
<!--<h2>Goals, Strategies, and Activities</h2>-->
<%	If SubmitID>0 Then %>
<p style="text-align: center; font-weight: bold; ">The progress report was submitted by <%=SubmitName%> at <%=SubmitTimestamp %> and is now locked.</p>
<%	End If %>
<form name="PR" method="post" action="ProgressReportSubmit.asp">
<%=HiddenField("MAGID", MAGID) %><%=HiddenField("Quarter", Quarter) %><%=HiddenField("action","save") %><%=HiddenField("Version",Version) %>
<table style="margin: auto">
<thead>
<%	End If %>
	<tr>
		<th>ID</th>
		<th>Activity</th>
		<th>Measure</th>
		<th><%=MonthDescription(Quarter,1) %></th>
		<th><%=MonthDescription(Quarter,2) %></th>
		<th><%=MonthDescription(Quarter,3) %></th>
	</tr>
</thead>
<%
LastMandatory = True
LastGoal=0
LastStrategy=0
Set rs=Con.Execute(sql)
While rs.EOF = False
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
	Response.Write("</td>" & vbCrLf)

	' Response Cells
	If rs.Fields("ResponseTypeID")=1 Then ' Integer
		If ShowExcel = True Then
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_M1")) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_M2")) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepIntegerWeb(rs.Fields("IntegerResponse_M3")) & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td>" & IntegerField("Response_M1_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_M1"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & IntegerField("Response_M2_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_M2"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & IntegerField("Response_M3_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_M3"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
		End If
	ElseIf rs.Fields("ResponseTypeID")=2 Then ' Decimal
		If ShowExcel = True Then
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_M1"),2) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_M2"),2) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepNumberWeb(rs.Fields("DecimalResponse_M3"),2) & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td>" & NumberField("Response_M1_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_M1"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & NumberField("Response_M2_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_M2"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & NumberField("Response_M3_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_M3"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
		End If
	ElseIf rs.Fields("ResponseTypeID")=3 Then ' dollars and cents
		If ShowExcel = True Then
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_M1")) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_M2")) & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; border: 1px solid black;"">" & prepCurrencyWeb(rs.Fields("DecimalResponse_M3")) & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td>" & CurrencyField("Response_M1_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_M1"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & CurrencyField("Response_M2_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_M2"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td>" & CurrencyField("Response_M3_" & rs.Fields("QuestionID"), rs.Fields("DecimalResponse_M3"), 8, 12, PermitEdit, "") & "</td>" & vbCrLf)
		End If
	ElseIf rs.Fields("ResponseTypeID")=5 Then ' Text area response. One per quarter.
		If ShowExcel = True Then
			Response.Write("<td colspan=""3"" style=""text-align: left; border: 1px solid black;"">" & prepStringWeb(rs.Fields("TextResponse")) & "</td>" & vbCrLf)
		Else
			Response.Write("<td colspan=""3"">" & TextArea("Response_" & rs.Fields("QuestionID"), rs.Fields("TextResponse"), 20, 60, 32000, PermitEdit, "") & "</td>" & vbCrLf)
		End If
	ElseIf rs.Fields("ResponseTypeID") = 6 Then ' one yes/no question
		Response.Write(vbTab & yesnoradio("Response_M1_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_M1"), 3, PermitEdit, ShowExcel) & vbCrLf)
	ElseIf rs.Fields("ResponseTypeID") = 7 Then ' a set of three yes/no questions
		Response.Write(vbTab & yesnoradio("Response_M1_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_M1"), 1, PermitEdit, ShowExcel) & vbCrLf)
		Response.Write(vbTab & yesnoradio("Response_M2_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_M2"), 1, PermitEdit, ShowExcel) & vbCrLf)
		Response.Write(vbTab & yesnoradio("Response_M3_" & rs.Fields("QuestionID"), rs.Fields("IntegerResponse_M3"), 1, PermitEdit, ShowExcel) & vbCrLf)
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
		DocumentFolder = Application("DocumentRoot") & "\MAG\" & MAGID & "\"
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		Response.Write("<table style=""margin: auto; "">" & vbCrLf)
		Response.Write("<tr><td>Current Documents in folder: ")
		If PermitEdit = True Then
			Response.Write("<a href=""../Upload/Upload.asp?fid=16&quarter=" & Quarter & "&MAGID=" & MAGID & """ class=""plainlink"" target=""_blank"">Upload</a>")
		End If
		Response.Write("</td>" & vbCrLf)
		If fso.FolderExists(DocumentFolder) Then
			Set folder = fso.GetFolder(DocumentFolder)
			Set files = folder.Files
			If files.count>0 Then 
				Response.Write("<tr><td>")
				For Each file in files
					If Left(file.Name,3)="PR"&Quarter Then
						Response.Write("<a href=""../Documents/MAG/" & MAGID & "/" & file.Name & _
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
		"FROM MAG.ProgressReportSubmissions AS A " & vbCrLf & _
		"LEFT JOIN [System].Users AS B ON A.SubmitID=B.SystemID " & vbCrLf & _
		"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND SubmitTimestamp IS NOT NULL " & vbCrLf & _
		"UNION " & vbCrLf & _
		"SELECT A.SubmitTimestamp, B.Name AS SubmitName " & vbCrLf & _
		"FROM MAG.ProgressReportSubmissions_Log AS A " & vbCrLf & _
		"LEFT JOIN [System].Users AS B ON A.SubmitID=B.SystemID " & vbCrLf & _
		"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND SubmitTimestamp IS NOT NULL " & vbCrLf & _
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
<%	If PermitEdit = True Or MVCPARights = True Or MVCPAViewer = True Then %>
	<input type="submit" name="Save" value="Save" title="Save changes and return to progress report" onclick="return submitForm('save');" />
<%	End If %>
<%	If CanSubmit = True And CurrentDate>=StartDate And IsNull(SubmitID) = True Then %>
	<input type="button" name="Submit" value="Submit" title="Save changes and submit the progress report. After submission, the report will no longer be editable." onclick="return submitForm('submit');"/>
<%	End If %>
	<input type="button" name="Close" value="Close" title="Ignore any pending changes and close window." 
		onclick="window.close();"/>
</div>
<div style="text-align: right"><a href="ProgressReport.asp?MAGID=<%=MAGID %>&Quarter=<%=Quarter %>&ShowExcel=1" target="_blank">Excel</a></div>
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

Function MonthDescription(vquarter, vmonth)
	If vquarter = 1 and vmonth = 1 Then 
		MonthDescription = "September 2022"
	ElseIf vquarter = 1 and vmonth = 2 Then 
		MonthDescription = "October 2022"
	ElseIf vquarter = 1 and vmonth = 3 Then 
		MonthDescription = "November 2022"
	ElseIf vquarter = 2 and vmonth = 1 Then 
		MonthDescription = "December 2022"
	ElseIf vquarter = 2 and vmonth = 2 Then 
		MonthDescription = "January 2023"
	ElseIf vquarter = 2 and vmonth = 3 Then 
		MonthDescription = "February 2023"
	ElseIf vquarter = 3 and vmonth = 1 Then 
		MonthDescription = "March 2023"
	ElseIf vquarter = 3 and vmonth = 2 Then 
		MonthDescription = "April 2023"
	ElseIf vquarter = 3 and vmonth = 3 Then 
		MonthDescription = "May 2023"
	ElseIf vquarter = 4 and vmonth = 1 Then 
		MonthDescription = "June 2023"
	ElseIf vquarter = 4 and vmonth = 2 Then 
		MonthDescription = "July 2023"
	ElseIf vquarter = 4 and vmonth = 3 Then 
		MonthDescription = "August 2023"
	Else
		MonthDescription = "unknown"
	End If
End Function

%>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->