<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><%
Dim debug, i, columnname, ShowExcel, FiscalYear, ResolutionStatus, _
	OrderBy, OrderByDescription, OrderByField, FilterBy, FilterByDescription, _
	ShowDatesAndLinks, ShowAwardInfo, ShowAppInfo, ShowER, ShowPR, ShowPRDescription, ShowPayee
Dim DocumentFolder, fso, folder, file, files
debug = False

ShowPRDescription = Array("None", "1", "2", "3", "4", "5", "6", "7", "8", "All")

OrderByDescription = Array("MAG ID", _
	"Grantee Name", _
	"Grantee ID", _
	"Resolution Confirmed", _
	"Award Letter Sent", _
	"Award Complete", _
	"ER Status")

OrderByField = Array("M.[MAGID]", _
	"G.[GranteeNameSort]", _
	"[Grantee ID]", _
	"A.[ResolutionConfirmedDate] DESC, G.[GranteeNameSort] ASC", _
	"A.[OfficialGrantAwardLetterDate] DESC, G.[GranteeNameSort] ASC", _
	"A.[GrantAwardCertifiedComplete] DESC, G.[GranteeNameSort] ASC", _
	"[ER Status]")

FilterByDescription = Array("Show All", "Grantee Created", "Started (not submitted)", "Started or Submitted", _
	"Submitted", "Closed", "Awarded", "Awarded, not Closed", "Paid", "ER Submitted, Not Paid")

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
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
End If

If Len(Request.Form("FilterBy"))>0 Then
	FilterBy = CInt(Request.Form("FilterBy"))
ElseIf Len(Request.QueryString("FilterBy"))>0 Then
	FilterBy = CInt(Request.QueryString("FilterBy"))
Else
	FilterBy = 6
End If

If Request.Form("ShowDatesAndLinks")="1" Then
	ShowDatesAndLinks = True
	ResolutionStatus = True
ElseIf Request.QueryString("ShowDatesAndLinks")="1" Then
	ShowDatesAndLinks = True
	ResolutionStatus = True
Else
	ShowDatesAndLinks = False
	ResolutionStatus = False
End If

If Request.Form("ShowAppInfo")="1" Then
	ShowAppInfo = True
ElseIf Request.QueryString("ShowAppInfo")="1" Then
	ShowAppInfo = True
Else
	ShowAppInfo = False
End If

If Request.Form("ShowAwardInfo")="1" Then
	ShowAwardInfo = True
ElseIf Request.QueryString("ShowAwardInfo")="1" Then
	ShowAwardInfo = True
Else
	ShowAwardInfo = False
End If

If Request.Form("ShowER")="1" Then
	ShowER = True
ElseIf Request.QueryString("ShowER")="1" Then
	ShowER = True
Else
	ShowER = False
End If

If Len(Request.Form("ShowPR"))>0 Then
	ShowPR = CInt(Request.Form("ShowPR"))
ElseIf Len(Request.QueryString("ShowPR"))>0 Then
	ShowPR = CInt(Request.QueryString("ShowPR"))
Else
	ShowPR = 0
End If

If Request.Form("ShowPayee")="1" Then
	ShowPayee = True
ElseIf Request.QueryString("ShowPayee")="1" Then
	ShowPayee = True
Else
	ShowPayee = False
End If

If Request.QueryString("ShowExcel")="1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

' To avoid an error, don't sort by ER information when ER information is not included.
If OrderByDescription(OrderBy) = "ER Status" And ShowER = False Then
	OrderBy = 1
End If

sql = "SELECT G.[GranteeID] AS [Grantee ID], G.[GranteeName] AS [Grantee Name], M.MAGID AS [MAG ID], " & vbCrLf & _
	"	CASE WHEN OptionID=1 THEN 'Purchase' WHEN OptionID=2 THEN 'Lease' ELSE '' END AS [Option], " & vbCrLf & _
	"	CASE WHEN ISNULL(G.BorderCounty,0)=1 THEN 'B' ELSE '' END + " & vbCrLf & _
	"		CASE WHEN ISNULL(G.PortCounty,0)=1 THEN 'P' ELSE '' END AS [B/P] "

If ShowAppInfo = True and ShowPayee = False Then
	sql = sql & "," & vbCrLf & "	CASE WHEN M.SubmitID IS NOT NULL THEN 'Application Submitted' " & vbCrLf & _
	"		WHEN M.MAGID IS NOT NULL THEN 'Application started' " & vbCrLf & _
	"		WHEN M.MAGID IS NULL THEN 'Grantee Created' ELSE '' END AS [Status], " & vbCrLf & _
	"	StolenVehicles AS [Stolen Vehicles], StolenVehicleValue AS [Stolen Vehicles Value], " & vbCrLf & _
	"	G.StatePayeeIDNo AS [Payee ID], " & vbCrLf & _
	"	U.Name AS 'Submit By', M.SubmitTimestamp AS [Submit Timestamp]"
ElseIf ShowAppInfo = True and ShowPayee = True Then
	sql = sql & "," & vbCrLf & "	CASE WHEN M.SubmitID IS NOT NULL THEN 'Application Submitted' " & vbCrLf & _
	"		WHEN M.MAGID IS NOT NULL THEN 'Application started' " & vbCrLf & _
	"		WHEN M.MAGID IS NULL THEN 'Grantee Created' ELSE '' END AS [Status], " & vbCrLf & _
	"	StolenVehicles AS [Stolen Vehicles], StolenVehicleValue AS [Stolen Vehicles Value], " & vbCrLf & _
	"	U.Name AS 'Submit By', M.SubmitTimestamp AS [Submit Timestamp]"
End If
If ShowAwardInfo = True Then
	sql = sql & ", " & vbCrLf & _
		"	A.ApplicationConsideredDate AS [Board Date], R.GrantResult AS Result, " & vbCrLf & _
		"	A.GrantAwardAmount AS [Grant Award Amount], A.CashMatch AS [Cash Match], " & vbCrLf & _
		"	A.GrantNumber AS [Grant Number]" & vbCrLf
End If
IF ShowDatesAndLinks = True Then
	sql = sql & ", " & vbCrLf & "	A.[ResolutionConfirmedDate] AS [Resolution Confirmed], " & vbCrLf & _
	"	A.[OfficialGrantAwardLetterDate] AS [Award Letter Sent], " & vbCrLf & _
	"	A.[GrantAwardCertifiedComplete] AS [Award Complete], " & vbCrLf & _
	"	A.[GrantClosedDate] AS [Grant Closed Date] " & vbCrLf
End If
If ShowER = True Then
	sql = sql & ", " & vbCrLf & _
		"	ER.ReviewDate AS [Review Date], ER.AuditApprovalDate AS [Audit Date], " & vbCrLf & _
		"	ER.DirectorApprovalDate AS [Director Approval Date], " & vbCrLf & _
		"	ER.AmountPaid AS [Amount To Pay], ER.DatePaid AS [Date Paid], " & vbCrLf & _
		"	CASE WHEN ER.MAGID IS NULL THEN '' " & vbCrLf & _
		"		WHEN ER.DatePaid IS NOT NULL THEN 'Paid' " & vbCrLf & _
		"		WHEN ER.DirectorApprovalID IS NOT NULL THEN 'Director Approval' " & vbCrLf & _
		"		WHEN ER.AuditApprovalID IS NOT NULL THEN 'Audit Approval' " & vbCrLf & _
		"		WHEN ER.ReviewID IS NOT NULL THEN 'Reviewed' " & vbCrLf & _
		"		WHEN ER.SubmitID IS NOT NULL THEN 'Submitted' " & vbCrLf & _
		"		ELSE 'Started' END AS [ER Status] "
End If
If ShowPayee = True Then
	sql = sql & ", " & vbCrLf & "	G.StatePayeeIDNo AS [State Payee ID No], " & vbCrLF & _
		"	G.Address1 AS [Address 1], G.Address2 AS [Address 2], G.City, G.State, G.ZIP "
End If
If ShowPR = 0 Then
	' Do not include progress report in query.
ElseIf ShowPR = 9 Then
	sql = sql & ", " & vbCrLf & _
		" CASE WHEN P1.ApprovalDate IS NOT NULL THEN 'Approved' WHEN P1.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status Qtr 1], Q1.[PR Questions Qtr 1], " & vbCrLf & _
		" CASE WHEN P2.ApprovalDate IS NOT NULL THEN 'Approved' WHEN P2.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status Qtr 2], Q2.[PR Questions Qtr 2], " & vbCrLf & _
		" CASE WHEN P3.ApprovalDate IS NOT NULL THEN 'Approved' WHEN P3.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status Qtr 3], Q3.[PR Questions Qtr 3], " & vbCrLf & _
		" CASE WHEN P4.ApprovalDate IS NOT NULL THEN 'Approved' WHEN P4.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status Qtr 4], Q4.[PR Questions Qtr 4], " & vbCrLf & _
		" CASE WHEN P5.ApprovalDate IS NOT NULL THEN 'Approved' WHEN P5.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status Qtr 5], Q5.[PR Questions Qtr 5], " & vbCrLf & _
		" CASE WHEN P6.ApprovalDate IS NOT NULL THEN 'Approved' WHEN P6.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status Qtr 6], Q6.[PR Questions Qtr 6], " & vbCrLf & _
		" CASE WHEN P7.ApprovalDate IS NOT NULL THEN 'Approved' WHEN P7.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status Qtr 7], Q7.[PR Questions Qtr 7], " & vbCrLf & _
		" CASE WHEN P8.ApprovalDate IS NOT NULL THEN 'Approved' WHEN P8.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status Qtr 8], Q8.[PR Questions Qtr 8], " & vbCrLf & _
		"	NULL AS [PR] " & vbCrLf
Else
	sql = sql & ", " & vbCrLf & _
		"P.ApprovalDate AS [Approval Date], " & vbCrLf & _
		" CASE WHEN P.SubmitID>0 THEN 'Submitted' ELSE '' END AS [PR Status], Q.[PR Questions], " & _
	"	NULL AS [PR] " & vbCrLf
End If
sql = sql & vbCrLf & _
	"FROM [Grantees] AS G " & vbCrLf & _
	"LEFT JOIN MAG.Main AS M ON M.GranteeID = G.GranteeID " & vbCrLf & _
	"LEFT JOIN [System].[Users] AS U ON U.SystemID=M.SubmitID " & vbCrLf & _
	"LEFT JOIN [MAG].Admin AS A ON A.MAGID=M.MAGID " & vbCrLf & _
	"LEFT JOIN [MAG].ExpenditureReport AS ER ON ER.MAGID=M.MAGID " & vbCrLf & _
	"LEFT JOIN Lookup.GrantResults AS R ON R.GrantResultID=A.GrantResultID " & vbCrLf
If ShowPR = 0 Then
	' Do not include progress report in query.
ElseIf ShowPR = 9 Then
	sql = sql & "LEFT JOIN [MAG].ProgressReportSubmissions AS P1 ON P1.MAGID=M.MAGID AND P1.Quarter=1 " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions Qtr 1] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q1 ON Q1.MAGID=M.MAGID AND Q1.Quarter=1 " & vbCrLf & _
		"LEFT JOIN [MAG].ProgressReportSubmissions AS P2 ON P2.MAGID=M.MAGID AND P2.Quarter=2 " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions Qtr 2] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q2 ON Q2.MAGID=M.MAGID AND Q2.Quarter=2 " & vbCrLf & _
		"LEFT JOIN [MAG].ProgressReportSubmissions AS P3 ON P3.MAGID=M.MAGID AND P3.Quarter=3 " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions Qtr 3] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q3 ON Q3.MAGID=M.MAGID AND Q3.Quarter=3 " & vbCrLf & _
		"LEFT JOIN [MAG].ProgressReportSubmissions AS P4 ON P4.MAGID=M.MAGID AND P4.Quarter=4 " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions Qtr 4] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q4 ON Q4.MAGID=M.MAGID AND Q4.Quarter=4 " & vbCrLf & _
		"LEFT JOIN [MAG].ProgressReportSubmissions AS P5 ON P5.MAGID=M.MAGID AND P5.Quarter=5 " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions Qtr 5] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q5 ON Q5.MAGID=M.MAGID AND Q5.Quarter=5 " & vbCrLf & _
		"LEFT JOIN [MAG].ProgressReportSubmissions AS P6 ON P6.MAGID=M.MAGID AND P6.Quarter=6 " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions Qtr 6] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q6 ON Q6.MAGID=M.MAGID AND Q6.Quarter=6 " & vbCrLf & _
		"LEFT JOIN [MAG].ProgressReportSubmissions AS P7 ON P7.MAGID=M.MAGID AND P7.Quarter=7 " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions Qtr 7] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q7 ON Q7.MAGID=M.MAGID AND Q7.Quarter=7 " & vbCrLf & _
		"LEFT JOIN [MAG].ProgressReportSubmissions AS P8 ON P8.MAGID=M.MAGID AND P8.Quarter=8 " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions Qtr 8] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q8 ON Q8.MAGID=M.MAGID AND Q8.Quarter=8 " & vbCrLf
Else
	sql = sql & "LEFT JOIN [MAG].ProgressReportSubmissions AS P ON P.MAGID=M.MAGID AND P.Quarter=" & prepIntegerSQL(ShowPR) & " " & vbCrLf & _
		"LEFT JOIN (SELECT MAGID, Quarter, COUNT(*) AS [PR Questions] " & vbCrLf & _
		"	FROM MAG.ProgressReportResponses " & vbCrLf & _
		"	WHERE IntegerResponse_M1>1 OR IntegerResponse_M2>0 OR IntegerResponse_M3>0 OR DecimalResponse_M1>1 OR DecimalResponse_M2>0 OR DecimalResponse_M3>0 OR LEN(TextResponse)>0 " & vbCrLf & _
		"	GROUP BY MAGID, Quarter) AS Q ON Q.MAGID=M.MAGID AND Q.Quarter=" & prepIntegerSQL(ShowPR) & " " & vbCrLf
End If
sql = sql & "WHERE (G.AuxiliaryGrant=1 OR M.MAGID IS NOT NULL) "
If FilterBy = 1 Then
	sql = sql & " AND M.MAGID IS NULL " & vbCrLf
ElseIf FilterBy = 2 Then
	sql = sql & " AND M.MAGID IS NOT NULL AND M.SubmitID IS NULL " & vbCrLf
ElseIf FilterBy = 3 Then
	sql = sql & " AND M.MAGID IS NOT NULL " & vbCrLf
ElseIf FilterBy=4 Then
	sql = sql & " AND M.SubmitID IS NOT NULL " & vbCrLf
ElseIf FilterBy=5 Then
	sql = sql & " AND A.GrantClosedDate IS NOT NULL " & vbCrLf
ElseIf FilterBy=6 Then
	sql = sql & " AND GrantAwardAmount>0 " & vbCrLf
ElseIf FilterBy=7 Then
	sql = sql & " AND GrantAwardAmount>0 AND GrantClosedDate IS NULL " & vbCrLf
ElseIf FilterBy=8 Then
	sql = sql & " AND ER.DatePaid IS NOT NULL " & vbCrLf
ElseIf FilterBy=9 Then
	sql = sql & " AND ER.SubmitID IS NOT NULL AND ER.DatePaid IS NULL " & vbCrLf
End If
sql = sql & "ORDER BY " & OrderByField(OrderBy)

If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=MAGStatusReport" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title><%=FiscalYear%> Auxiliary Grant Status Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<div class="sectiontitle" style="white-space: nowrap;"><%=FiscalYear%> Auxiliary Grant Status Report</div>
<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2022 to 2022
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="OrderBy">Order By:</label> <select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	If OrderByDescription(i) = "ER Status" And ShowER = False Then
		' Skip: To avoid an error, don't sort by ER information when ER information is not included.
	Else
		Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
	End If
Next
%>
</select>&nbsp;&nbsp;
<label for="FilterBy">Filter By:</label> <select name="FilterBy" id="FilterBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(FilterByDescription)
	Response.Write("<option value=""" & i & """" & Selected(FilterBy, i) & ">" & FilterByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select><br />
<input name="ShowDatesAndLinks" type="checkbox" <%=Checked(ShowDatesAndLinks, True) %> value="1" onchange="Selection.submit();" /> Show Dates and Links&nbsp;&nbsp;
<input name="ShowAppInfo" type="checkbox" <%=Checked(ShowAppInfo, True) %> value="1" onchange="Selection.submit();" /> Show App Info&nbsp;&nbsp;
<input name="ShowAwardInfo" type="checkbox" <%=Checked(ShowAwardInfo, True) %> value="1" onchange="Selection.submit();" /> Show Award Info&nbsp;&nbsp;
<input name="ShowER" type="checkbox" <%=Checked(ShowER, True) %> value="1" onchange="Selection.submit();" /> Show ER Info&nbsp;&nbsp;
<input name="ShowPayee" type="checkbox" <%=Checked(ShowPayee, True) %> value="1" onchange="Selection.submit();" /> Show Payee&nbsp;&nbsp;
<label for="ShowPR">Show PR:</label> <select name="ShowPR" id="ShowPR" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowPRDescription)
	Response.Write("<option value=""" & i & """" & Selected(ShowPR, i) & ">" & ShowPRDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>
&nbsp;&nbsp;<a href="StatusReport.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&ShowDatesAndLinks=<%
	If ShowDatesAndLinks = True Then 
		Response.Write("1") 
	Else 
		Response.Write("0") 
	End If
%>&ShowAwardInfo=<%
	If ShowAwardInfo = True Then 
		Response.Write("1") 
	Else 
		Response.Write("0") 
	End If
%>&ShowER=<%
	If ShowER = True Then 
		Response.Write("1") 
	Else 
		Response.Write("0") 
	End If
%>&ShowPayee=<%
	If ShowPayee = True Then 
		Response.Write("1") 
	Else 
		Response.Write("0") 
	End If
%>&ShowPR=<%=ShowPR %>&OrderBy=<%=OrderBy %>&FilterBy=<%=FilterBy %>" target="_blank">Excel</a>
</form>

<br />
<table class="reporttable">
<%
End If

' Create once to use repeatedly for each grant.
set fso = Server.CreateObject("Scripting.FileSystemObject")

If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
	For i = 0 To (rs.Fields.Count-1)
		If rs.Fields(i).Name="B/P" Then
			Response.Write("<th title=""Border/Port"">" & Replace(rs.Fields(i).Name," ","<br />") & "</th>")
		ElseIf rs.Fields(i).Name="PR" Then
			Response.Write("<th title=""Link to Progress Report"">PR</th>")
		ElseIf InStr(rs.Fields(i).Name, "Qtr ") > 0 Then
			columnname = Replace(rs.Fields(i).Name," ","<br />")
			columnname = Replace(columnname,"Qtr<br />","Qtr ")
			Response.Write("<th>" & columnname & "</th>")
		Else
			Response.Write("<th>" & Replace(rs.Fields(i).Name," ","<br />") & "</th>" & vbCrLf)
		End If
	Next
	If ResolutionStatus = True Then
		Response.Write("<th title=""Resolution"">Res</th>")
		Response.Write("<th title=""Signed Statement of Grant Award"">SGA</th>")
	End If
	Response.Write(vbCrLf & "</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top; white-space: nowrap; "">" & vbCrLf)
		For i = 0 To (rs.Fields.Count-1)
			If rs.Fields(i).Name="PR" Then
				If ShowPR > 0 And ShowPR < 5 Then
					Response.Write("<td title=""Link to Progress Report""><a href=""../MAG/ProgressReport.asp?MAGID=" & rs.Fields("Mag ID") & "&Quarter=" & ShowPR & """ target=""_blank"">PR</a></td>" & vbCrLf)
				ElseIf ShowPR = 5 Then
					Response.Write("<td title=""Link to Progress Report""><a href=""../MAG/ProgressReport.asp?MAGID=" & rs.Fields("Mag ID") & """ target=""_blank"">PR</a></td>" & vbCrLf)
				End If
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "Grantee ID" Then
				If MVCPARights = True And ShowExcel = False Then
					Response.Write("<td style=""text-align: right""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\Grantees\Grantee.asp?GranteeID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "MAG ID" Then
				If MVCPARights = True And ShowExcel = False Then
					Response.Write("<td style=""text-align: right""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\MAG\Admin.asp?MAGID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="Fiscal Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Name="Cash Match Pct" Or rs.Fields(i).Name="Revised Cash Match Pct"  Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, false, false) & "%</td>")
			ElseIf rs.Fields(i).Name = "Submission Date" Then
				Response.Write("<td style=""text-align: right"">" & formatdatetime(rs.Fields(i).value, vbGeneralDate) & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & prepCurrencyWeb(rs.Fields(i).value) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adDate Or rs.Fields(i).Type=adDBTimeStamp Then
				Response.Write("<td style=""text-align: right"">" & rs.Fields(i).value & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		If ResolutionStatus = True Then
			DocumentFolder = Application("DocumentRoot") & "\MAG\" & rs.Fields("MAG ID") & "\"
			If fso.FolderExists(DocumentFolder) Then
				If fso.FileExists(DocumentFolder & "Resolution.pdf") Then
					Response.Write("<td><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "/Documents/MAG/" & rs.Fields("MAG ID") & "/Resolution.pdf"" target=""_blank"">Res</a></td>")
				Else
					Response.Write("<td>" & "No" & "</td>")
				End If
			Else
				Response.Write("<td title=""Directory does not exist"">No</td>")
			End If
			If fso.FolderExists(DocumentFolder) Then
				If fso.FileExists(DocumentFolder & "Signed Statement of Grant Award.pdf") Then
					Response.Write("<td><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "/Documents/MAG/" & rs.Fields("MAG ID") & "/Signed Statement of Grant Award.pdf"" target=""_blank"">SGA</a></td>")
				Else
					Response.Write("<td>" & "No" & "</td>")
				End If
			Else
				Response.Write("<td title=""Directory does not exist"">No</td>")
			End If
		End If
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%	If ShowExcel = False Then %>
<div style="width: 100%; text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%	End If %>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->