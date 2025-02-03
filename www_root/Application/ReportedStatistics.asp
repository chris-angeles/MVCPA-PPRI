<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, HistoricalDataYear, ShowColumns, KeyColumns, ShowExcel, PreviousAppID
Debug = False
PreviousAppID=0

If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLF)
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
	Response.Write("</pre>" & vbCrLF)
End If

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = Request.Form("FiscalYear")
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = Request.QueryString("FiscalYear")
Else
	FiscalYear=2022
End If

If Request.QueryString("ShowExcel")="1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

If ShowExcel = False Then
	ShowColumns = 7
	KeyColumns = 1
Else
	ShowColumns = 9
	KeyColumns = 3
End If


HistoricalDataYear = FiscalYear - 2

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ReportedStatistics" & FiscalYear & ".xls"
Else ' Start of Web only code
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Reported Statistics</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">
<%
End If
%>
<table class="bordertable">
<thead>
	<tr><th  colspan="<%=ShowColumns %>">Statistics to Support Grant Problem Statement</th></tr>
	<tr>
		<th colspan="<%=KeyColumns %>">Reported Cases Category</th>
		<th colspan="3" style="border: solid black thin; "><%=(HistoricalDataYear-1) %></th>
		<th colspan="3" style="border: solid black thin; "><%=(HistoricalDataYear) %></th>
	</tr>
<%	If ShowExcel = False Then %>
	<tr style="vertical-align: bottom; ">
		<th style="width: 175px; ">Jurisdiction</th>
		<th style="width: 115px; ">Motor Vehicle Theft<br />(MVT)</th>
		<th style="width: 115px; " title="Burglary from Motor Vehicle including theft of parts">Burglary from Motor Vehicle<br />(BMV)</th>
		<th style="width: 115px; ">Fraud-Related Motor Vehicle Crime<br />(FRMVC)</th>
		<th style="width: 115px; ">Motor Vehicle Theft<br />(MVT)</th>
		<th style="width: 115px; " title="Burglary from Motor Vehicle including theft of parts">Burglary from Motor Vehicle<br />(BMV)</th>
		<th style="width: 115px; ">Fraud-Related Motor Vehicle Crime<br />(FRMVC)</th>
	</tr>
<%	Else %>
	<tr style="vertical-align: bottom; ">
		<th>AppID</th>
		<th>Grantee</th>
		<th style="width: 175px; ">Jurisdiction</th>
		<th style="width: 115px; ">Motor Vehicle Theft (MVT)</th>
		<th style="width: 115px; " title="Burglary from Motor Vehicle including theft of parts">Burglary from Motor Vehicle (BMV)</th>
		<th style="width: 115px; ">Fraud-Related Motor Vehicle Crime (FRMVC)</th>
		<th style="width: 115px; ">Motor Vehicle Theft (MVT)</th>
		<th style="width: 115px; " title="Burglary from Motor Vehicle including theft of parts">Burglary from Motor Vehicle (BMV)</th>
		<th style="width: 115px; ">Fraud-Related Motor Vehicle Crime (FRMVC)</th>
	</tr>
<%	End If %>
</thead>
<tbody>
<%
sql = "WITH CTE AS ( " & vbCrLf & _
	"SELECT A.GranteeID, REPLACE(A.GranteeName,'City of ','') AS Grantee_Name, B.ProgramName,  " & vbCrLf & _
	"	C.StatisticsID, I.AppID, Jurisdiction, C.MVT1, C.BMV1, C.FRMVC1, C.MVT2, C.BMV2, C.FRMVC2 " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"LEFT JOIN Application.IDs AS I ON I.GranteeID=A.GranteeID and I.GrantClassID=1 " & vbCrLf & _
	"JOIN Application.Main AS B ON B.AppID=I.AppID " & vbCrLf & _
	"JOIN Application.[Statistics] AS C ON C.AppID=B.AppID " & vbCrLf & _
	"WHERE I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
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
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write(vbTab & "<tr><td colspan=""" & ShowColumns & """>&nbsp;</td></tr>" & vbCrLf)
	Response.Write(vbTab & "<tr><th colspan=""" & ShowColumns & """><i>No Statistical Data has been entered yet.</i></th></tr>" & vbCrLf)
Else
While rs.EOF = False
	If rs.Fields("AppID") = PreviousAppID Then
		' Do nothing
	ElseIf ShowExcel = False Then
		Response.Write(vbtab & "<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<th colspan=""" & ShowColumns & """>Grantee: " & rs.Fields("Grantee_Name") & ", Program Name: " & rs.Fields("ProgramName") & " (" & rs.Fields("AppID") & ")</td>")
		Response.Write(vbtab & "</tr>" & vbCrLf)
		PreviousAppID = rs.Fields("AppID")
	End If
	Response.Write(vbtab & "<tr style=""vertical-align: top; "">" & vbCrLf)
	If ShowExcel = True Then
		Response.Write(vbTab & vbTab & "<td>" & rs.Fields("AppID") & "</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td>" & rs.Fields("Grantee_Name") & "</td>" & vbCrLf)
	End If
	Response.Write(vbTab & vbTab & "<td>" & rs.Fields("Jurisdiction") & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("MVT1")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("BMV1")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("FRMVC1")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("MVT2")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("BMV2")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("FRMVC2")) & "</td>" & vbCrLf)
	Response.Write(vbtab & "</tr>" & vbCrLf)
	rs.MoveNext
Wend
End If

%>
</tbody>
</table>
<%
If ShowExcel = False Then
%>
<div style="width: 100%; text-align: right;"><a href="ReportedStatistics.asp?FiscalYear=<%=FiscalYear %>&ShowExcel=1" target="_blank">Excel</a></div>
</body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->