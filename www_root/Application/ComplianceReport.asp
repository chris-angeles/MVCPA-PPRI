<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, OrderByDescription, OrderByField, ShowExcel, _
	ShowOnlySubmitted, ThisYearSource, ThisYearSourceDescription, LastYearSource, LastYearSourceDescription
OrderByDescription = Array("App ID", "Grantee Name")
OrderByField = Array("[App ID]", "REPLACE(A.[Grantee Name],'City of ','')")
ThisYearSourceDescription = Array("Application", "Negotiation")
LastYearSourceDescription = Array("Grant", "Application", "Negotiation")
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

If Len(Request.Form("ThisYearSource"))>0 Then
	ThisYearSource = CInt(Request.Form("ThisYearSource"))
ElseIf Len(Request.QueryString("ThisYearSource"))>0 Then
	ThisYearSource = CInt(Request.QueryString("ThisYearSource"))
Else
	ThisYearSource = 0
End If


If Len(Request.Form("LastYearSource"))>0 Then
	LastYearSource = CInt(Request.Form("LastYearSource"))
ElseIf Len(Request.QueryString("LastYearSource"))>0 Then
	LastYearSource = CInt(Request.QueryString("LastYearSource"))
Else
	LastYearSource = 0
End If

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
End If

If Request.QueryString("ShowExcel")="1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

If True = False Then
	If FiscalYear = 2018 Then
		ShowOnlySubmitted = True
		ThisYearSource = "Negotiation"
	ElseIf FiscalYear = 2019 Then
		ShowOnlySubmitted = False
		ThisYearSource = "Application"
	ElseIf FiscalYear = 2020 Then
		ShowOnlySubmitted = False
		ThisYearSource = "Negotiation"
	ElseIf FiscalYear = 2021 Then
		ShowOnlySubmitted = False
		ThisYearSource = "Negotiation"
	ElseIf FiscalYear = 2022 Then
		ShowOnlySubmitted = False
		ThisYearSource = "Negotiation"
	ElseIf FiscalYear = 2023 Then
		ShowOnlySubmitted = False
		ThisYearSource = "Application"
	Else
		ShowOnlySubmitted = True
		ThisYearSource = "Application"
	End If
End If

If ThisYearSourceDescription(ThisYearSource) = "Negotiation" And LastYearSourceDescription(LastYearSource) = "Grant" Then
	sql = "SELECT A.[App ID], A.[Grantee Name], A.[Revised Program Name], A.[Revised Grant Type], " & vbCrLf & _
		"	A.[Revised Submitted By], A.[Revised Submitted], A.[Revised Status], " & vbCrLf & _
		"	A.[Revised MVCPA Funds Requested] AS [" & (FiscalYear) & " Revised MVCPA Funds Requested], " & vbCrLf & _
		"	G.AwardAmount AS [" & (FiscalYear-1) & " MVCPA Funds], " & vbCrLf & _
		"	A.[Revised Cash Match] AS [" & FiscalYear & " Revised Cash Match Requested], " & vbCrLf & _
		"	G.MatchAmount AS [" & (FiscalYear-1) & " Cash Match], " & vbCrLf & _
		"	ROUND(A.[Revised Cash Match],2) - ROUND(G.MatchAmount,2) AS [Change in Cash Match], " & vbCrLf & _
		"	A.[Revised Cash Match Pct] AS [" & FiscalYear & " Revised Cash Match Percent], " & vbCrLf & _
		"	100.0*G.MatchAmount/G.AwardAmount AS [" & (FiscalYear-1) & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(A.[Revised Cash Match Pct],4) - ROUND(100.0*G.MatchAmount/G.AwardAmount,4) AS [Change in Cash Match Percent], " & vbCrLf & _
		"	CASE WHEN G.AwardAmount>0 AND G.MatchAmount>0 THEN 100.0*(ROUND(A.[Revised Cash Match Pct],4) - ROUND(100.0*(G.MatchAmount/G.AwardAmount),4))/ROUND((100*G.MatchAmount/G.AwardAmount),4) ELSE NULL END AS [Cash Match Percent Change], " & vbCrLf & _
		"	CASE WHEN G.AwardAmount>0 THEN 100.0*(A.[Revised MVCPA Funds Requested]/G.AwardAmount) ELSE NULL END AS [Revised MVCPA Pct of Last Year], " & vbCrLf & _
		"	CASE WHEN G.MatchAmount>0 THEN 100.0*A.[Revised Cash Match]/G.MatchAmount ELSE NULL END AS [Revised Cash Match Pct of Last Year], " & vbCrLf & _
		"	A.[Revised In-Kind Match] AS [" & FiscalYear & " Revised In-Kind Match] " & vbCrLf & _
		"FROM Application.vwSummary AS A " & vbCrLf & _
		"LEFT JOIN [Grants].Main AS G ON G.GranteeID=A.[Grantee ID] AND G.[FiscalYear]=A.[Fiscal Year]-1 AND G.GrantClassID=A.GrantClassID " & vbCrLf & _
		"WHERE A.GrantClassID=1 AND [Fiscal Year]=" & prepIntegerSQL(FiscalYear) & vbCrLf & _
		"	AND A.Negotiation=1 " & vbCrLf
ElseIf ThisYearSourceDescription(ThisYearSource) = "Application" And LastYearSourceDescription(LastYearSource) = "Grant" Then
	sql = "SELECT A.[App ID], A.[Grantee Name], A.[Program Name], A.[Grant Type], " & vbCrLf & _
		"	A.[Submitted By], A.[Submitted], A.[Status], " & vbCrLf & _
		"	A.[MVCPA Funds Requested] AS [" & (FiscalYear) & " MVCPA Funds Requested], " & vbCrLf & _
		"	G.AwardAmount AS [" & (FiscalYear-1) & " MVCPA Funds], " & vbCrLf & _
		"	A.[Cash Match] AS [" & FiscalYear & " Cash Match Requested], " & vbCrLf & _
		"	G.MatchAmount AS [" & (FiscalYear-1) & " Cash Match], " & vbCrLf & _
		"	ROUND(A.[Cash Match],2) - ROUND(G.MatchAmount,2) AS [Change in Cash Match], " & vbCrLf & _
		"	ROUND(A.[Cash Match Pct],4) AS [" & FiscalYear & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(100.0*G.MatchAmount/G.AwardAmount,4) AS [" & (FiscalYear-1) & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(A.[Cash Match Pct],4) - ROUND(100.0*G.MatchAmount/G.AwardAmount,4) AS [Change in Cash Match Percent], " & vbCrLf & _
		"	CASE WHEN G.AwardAmount>0 AND G.MatchAmount>0 THEN 100.0*(ROUND(A.[Cash Match Pct],4) - ROUND(100.0*(G.MatchAmount/G.AwardAmount),4))/ROUND((100*G.MatchAmount/G.AwardAmount),4) ELSE NULL END AS [Cash Match Percent Change], " & vbCrLf & _
		"	CASE WHEN G.AwardAmount>0 THEN 100.0*(A.[MVCPA Funds Requested]/G.AwardAmount) ELSE NULL END AS [MVCPA Pct of Last Year], " & vbCrLf & _
		"	CASE WHEN G.MatchAmount>0 THEN 100.0*A.[Cash Match]/G.MatchAmount ELSE NULL END AS [Cash Match Pct of Last Year], " & vbCrLf & _
		"	A.[In-Kind Match] AS [" & FiscalYear & " In-Kind Match] " & vbCrLf & _
		"FROM Application.vwSummary AS A " & vbCrLf & _
		"LEFT JOIN [Grants].Main AS G ON G.GranteeID=A.[Grantee ID] AND G.[FiscalYear]=A.[Fiscal Year]-1 AND G.GrantClassID=A.GrantClassID " & vbCrLf & _
		"WHERE A.GrantClassID=1 AND A.[Fiscal Year]=" & prepIntegerSQL(FiscalYear) & vbCrLf
ElseIf ThisYearSourceDescription(ThisYearSource) = "Application" And LastYearSourceDescription(LastYearSource) = "Negotiation" Then
	sql = "SELECT A.[App ID], A.[Grantee Name], A.[Program Name], A.[Grant Type], " & vbCrLf & _
		"	A.[Submitted By], A.[Submitted], A.[Status], " & vbCrLf & _
		"	A.[MVCPA Funds Requested] AS [" & (FiscalYear) & " MVCPA Funds Requested], " & vbCrLf & _
		"	G.[Revised MVCPA Funds Requested] AS [" & (FiscalYear-1) & " MVCPA Funds], " & vbCrLf & _
		"	A.[Cash Match] AS [" & FiscalYear & " Cash Match Requested], " & vbCrLf & _
		"	G.[Revised Cash Match] AS [" & (FiscalYear-1) & " Cash Match], " & vbCrLf & _
		"	ROUND(A.[Cash Match],2) - ROUND(G.[Revised Cash Match],2) AS [Change in Cash Match], " & vbCrLf & _
		"	ROUND(A.[Cash Match Pct],4) AS [" & FiscalYear & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(G.[Revised Cash Match Pct],4) AS [" & (FiscalYear-1) & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(A.[Cash Match Pct],4) - ROUND(G.[Revised Cash Match Pct],4) AS [Change in Cash Match Percent], " & vbCrLf & _
		"	CASE WHEN G.[Revised Cash Match Pct]>0 THEN 100.0*(ROUND(A.[Cash Match Pct],4) - ROUND(G.[Revised Cash Match Pct],4))/ROUND(G.[Revised Cash Match Pct],4) ELSE NULL END AS [Cash Match Percent Change], " & vbCrLf & _
		"	CASE WHEN G.[Revised MVCPA Funds Requested]>0 THEN 100.0*(A.[MVCPA Funds Requested]/G.[Revised MVCPA Funds Requested]) ELSE NULL END AS [MVCPA Pct of Last Year], " & vbCrLf & _
		"	CASE WHEN G.[Revised Cash Match]>0 THEN 100.0*A.[Cash Match]/G.[Revised Cash Match] ELSE NULL END AS [Cash Match Pct of Last Year], " & vbCrLf & _
		"	A.[In-Kind Match] AS [" & FiscalYear & " In-Kind Match] " & vbCrLf & _
		"FROM Application.vwSummary AS A " & vbCrLf & _
		"LEFT JOIN Application.vwSummary AS G ON G.[Grantee ID]=A.[Grantee ID] AND G.[Fiscal Year]=A.[Fiscal Year]-1 AND G.GrantClassID=A.GrantClassID " & vbCrLf & _
		"WHERE A.GrantClassID=1 AND A.[Fiscal Year]=" & prepIntegerSQL(FiscalYear) & vbCrLf
ElseIf ThisYearSourceDescription(ThisYearSource) = "Negotiation" And LastYearSourceDescription(LastYearSource) = "Negotiation" Then
	sql = "SELECT A.[App ID], A.[Grantee Name], A.[Program Name], A.[Grant Type], " & vbCrLf & _
		"	A.[Submitted By], A.[Submitted], A.[Status], " & vbCrLf & _
		"	A.[MVCPA Funds Requested] AS [" & (FiscalYear) & " MVCPA Funds Requested], " & vbCrLf & _
		"	G.[Revised MVCPA Funds Requested] AS [" & (FiscalYear-1) & " MVCPA Funds], " & vbCrLf & _
		"	A.[Revised Cash Match] AS [" & FiscalYear & " Cash Match Requested], " & vbCrLf & _
		"	G.[Revised Cash Match] AS [" & (FiscalYear-1) & " Cash Match], " & vbCrLf & _
		"	ROUND(A.[Revised Cash Match],2) - ROUND(G.[Revised Cash Match],2) AS [Change in Cash Match], " & vbCrLf & _
		"	ROUND(A.[Revised Cash Match Pct],4) AS [" & FiscalYear & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(G.[Revised Cash Match Pct],4) AS [" & (FiscalYear-1) & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(A.[Revised Cash Match Pct],4) - ROUND(G.[Revised Cash Match Pct],4) AS [Change in Cash Match Percent], " & vbCrLf & _
		"	100.0*(ROUND(A.[Revised Cash Match Pct],4) - ROUND(G.[Revised Cash Match Pct],4))/ROUND(G.[Revised Cash Match Pct],4) AS [Cash Match Percent Change], " & vbCrLf & _
		"	100.0*(A.[Revised MVCPA Funds Requested]/G.[Revised MVCPA Funds Requested]) AS [MVCPA Pct of Last Year], " & vbCrLf & _
		"	100*A.[Revised Cash Match]/G.[Revised Cash Match] AS [Cash Match Pct of Last Year], " & vbCrLf & _
		"	A.[Revised In-Kind Match] AS [" & FiscalYear & " In-Kind Match] " & vbCrLf & _
		"FROM Application.vwSummary AS A " & vbCrLf & _
		"LEFT JOIN Application.vwSummary AS G ON G.[Grantee ID]=A.[Grantee ID] AND G.[Fiscal Year]=A.[Fiscal Year]-1 AND G.GrantClassID=A.GrantClassID " & vbCrLf & _
		"WHERE A.GrantClassID=1 AND A.[Fiscal Year]=" & prepIntegerSQL(FiscalYear) & _
		"	AND A.Negotiation=1 " & vbCrLf
ElseIf ThisYearSourceDescription(ThisYearSource) = "Application" And LastYearSourceDescription(LastYearSource) = "Application" Then
	sql = "SELECT A.[App ID], A.[Grantee Name], A.[Program Name], A.[Grant Type], " & vbCrLf & _
		"	A.[Submitted By], A.[Submitted], A.[Status], " & vbCrLf & _
		"	A.[MVCPA Funds Requested] AS [" & (FiscalYear) & " MVCPA Funds Requested], " & vbCrLf & _
		"	G.[MVCPA Funds Requested] AS [" & (FiscalYear-1) & " MVCPA Funds], " & vbCrLf & _
		"	A.[Cash Match] AS [" & FiscalYear & " Cash Match Requested], " & vbCrLf & _
		"	G.[Cash Match] AS [" & (FiscalYear-1) & " Cash Match], " & vbCrLf & _
		"	ROUND(A.[Cash Match],2) - ROUND(G.[Cash Match],2) AS [Change in Cash Match], " & vbCrLf & _
		"	ROUND(A.[Cash Match Pct],4) AS [" & FiscalYear & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(G.[Cash Match Pct],4) AS [" & (FiscalYear-1) & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(A.[Cash Match Pct],4) - ROUND(G.[Cash Match Pct],4) AS [Change in Cash Match Percent], " & vbCrLf & _
		"	100.0*(ROUND(A.[Cash Match Pct],4) - ROUND(G.[Cash Match Pct],4))/ROUND(G.[Cash Match Pct],4) AS [Cash Match Percent Change], " & vbCrLf & _
		"	100.0*(A.[MVCPA Funds Requested]/G.[MVCPA Funds Requested]) AS [MVCPA Pct of Last Year], " & vbCrLf & _
		"	100*A.[Cash Match]/G.[Revised Cash Match] AS [Cash Match Pct of Last Year], " & vbCrLf & _
		"	A.[In-Kind Match] AS [" & FiscalYear & " In-Kind Match] " & vbCrLf & _
		"FROM Application.vwSummary AS A " & vbCrLf & _
		"LEFT JOIN Application.vwSummary AS G ON G.[Grantee ID]=A.[Grantee ID] AND G.[Fiscal Year]=A.[Fiscal Year]-1 " & vbCrLf & _
		"WHERE A.GrantClassID=1 AND A.[Fiscal Year]=" & prepIntegerSQL(FiscalYear) 
ElseIf ThisYearSourceDescription(ThisYearSource) = "Negotiation" And LastYearSourceDescription(LastYearSource) = "Application" Then
	sql = "SELECT A.[App ID], A.[Grantee Name], A.[Program Name], A.[Grant Type], " & vbCrLf & _
		"	A.[Submitted By], A.[Submitted], A.[Status], " & vbCrLf & _
		"	A.[MVCPA Funds Requested] AS [" & (FiscalYear) & " MVCPA Funds Requested], " & vbCrLf & _
		"	G.[Revised MVCPA Funds Requested] AS [" & (FiscalYear-1) & " MVCPA Funds], " & vbCrLf & _
		"	A.[Cash Match] AS [" & FiscalYear & " Cash Match Requested], " & vbCrLf & _
		"	G.[Revised Cash Match] AS [" & (FiscalYear-1) & " Cash Match], " & vbCrLf & _
		"	ROUND(A.[Cash Match],2) - ROUND(G.[Revised Cash Match],2) AS [Change in Cash Match], " & vbCrLf & _
		"	ROUND(A.[Cash Match Pct],4) AS [" & FiscalYear & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(G.[Revised Cash Match Pct],4) AS [" & (FiscalYear-1) & " Cash Match Percent], " & vbCrLf & _
		"	ROUND(A.[Cash Match Pct],4) - ROUND(G.[Revised Cash Match Pct],4) AS [Change in Cash Match Percent], " & vbCrLf & _
		"	100.0*(ROUND(A.[Cash Match Pct],4) - ROUND(G.[Revised Cash Match Pct],4))/ROUND(G.[Revised Cash Match Pct],4) AS [Cash Match Percent Change], " & vbCrLf & _
		"	100.0*(A.[MVCPA Funds Requested]/G.[Revised MVCPA Funds Requested]) AS [MVCPA Pct of Last Year], " & vbCrLf & _
		"	100*A.[Cash Match]/G.[Revised Cash Match] AS [Cash Match Pct of Last Year], " & vbCrLf & _
		"	A.[In-Kind Match] AS [" & FiscalYear & " In-Kind Match] " & vbCrLf & _
		"FROM Application.vwSummary AS A " & vbCrLf & _
		"LEFT JOIN Application.vwSummary AS G ON G.[Grantee ID]=A.[Grantee ID] AND G.[Fiscal Year]=A.[Fiscal Year]-1 " & vbCrLf & _
		"WHERE A.GrantClassID=1 AND A.[Fiscal Year]=" & prepIntegerSQL(FiscalYear) & _
		"	AND A.Negotiation=1 " & vbCrLf
Else
	Response.Write("This is an unexpected comparison and a query has not been constructed..")
	Response.End
End If

sql = sql & "ORDER BY " & OrderByField(OrderBy)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

Set rs=Con.Execute(sql)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ComplianceReport" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Compliance Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2017 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;&nbsp;
Compare selected year's <label for="ThisYearSource"></label><select name="ThisYearSource" id="ThisYearSource" onchange="Selection.submit();">
<%
For i = 0 to UBound(ThisYearSourceDescription)
	Response.Write("<option value=""" & i & """" & Selected(ThisYearSource, i) & ">" & ThisYearSourceDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>

to previous year's 

<label for="LastYearSource"></label><select name="LastYearSource" id="LastYearSource" onchange="Selection.submit();">
<%
For i = 0 to UBound(LastYearSourceDescription)
	Response.Write("<option value=""" & i & """" & Selected(LastYearSource, i) & ">" & LastYearSourceDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>
&nbsp;&nbsp;&nbsp;
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>
&nbsp;&nbsp;&nbsp;<a href="ComplianceReport.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&ThisYearSource=<%=ThisYearSource %>&LastYearSource=<%=LastYearSource %>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a>
</form>
</div>
<table class="reporttable">
<%
End if

If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom"">" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top"">" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write(vbTab & "<td></td>")
			ElseIf rs.Fields(i).Name = "GranteeID" Then
				If MVCPARights = True Then
					Response.Write(vbTab & "<td style=""text-align: right""><a href=""..\Grantees\Grantee.asp?GranteeID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "ISAID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Application\ISA.asp?ISAID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf Instr(1, rs.Fields(i).Name, "percent", vbTextCompare)>0 Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,4, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adNumeric Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, false, true) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		'Response.Write("<td>" & rs.Fields("Cash Match Pct Chg").Type & "</td>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
Else
	Response.WRite("<tr><td>Nothing to show</td></tr>" & vbCrLf)
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