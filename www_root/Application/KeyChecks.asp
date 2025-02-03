<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, OrderByDescription, OrderByField, ShowExcel, _
	ShowOnlySubmitted, ApplicationSchema
OrderByDescription = Array("App ID", "Grantee Name")
OrderByField = Array("[App ID]", "REPLACE([Grantee Name],'City of ','')")
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

ApplicationSchema = getApplicationSchema(FiscalYear)

If ApplicationSchema="Negotiation " THEN
	sql = "SELECT [App ID], [Grantee Name], [Revised Program Name], [Revisted Grant Type], [Revised Submitted By], [Submitted], [Status], " & vbCrLf & _
		"	[Revised MVCPA Funds Requested] AS [" & (FiscalYear) & " MVCPA Funds Requested], " & vbCrLF & _
		"	G.AwardAmount AS [" & (FiscalYear-1) & " MVCPA Funds], " & vbCrLF & _
		"	[Revised Cash Match] AS [" & FiscalYear & " Cash Match Requested], " & vbCrLf & _
		"	G.MatchAmount AS [" & (FiscalYear-1) & " Cash Match], " & vbCrLf & _
		"	[Revised Cash Match Pct] AS [" & FiscalYear & " Cash Match Percent], " & vbCrLf & _
		"	100.0*G.MatchAmount/G.AwardAmount AS [" & (FiscalYear-1) & " Cash Match Percent], " & vbCrLf & _
		"	[Revised Cash Match Pct] - 100.0*G.MatchAmount/G.AwardAmount AS [Change in Cash Match Percent], " & vbCrLf & _
		"	([Revised Cash Match Pct] - 100.0*(G.MatchAmount/G.AwardAmount))/(G.MatchAmount/G.AwardAmount) AS [Cash Match Percent Change], " & vbCrLf & _
		"	100.0*([Revised MVCPA Funds Requested]/G.AwardAmount) AS [MVCPA Pct of Last Year], " & vbCrLf & _
		"	100*[Cash Match]/G.MatchAmount AS [Cash Match Pct of Last Year], " & vbCrLf & _
		"	[Revised In-Kind Match] AS [" & FiscalYear & " In-Kind Match] " & vbCrLf & _
		"FROM Application.vwSummary AS A " & vbCrLf & _
		"LEFT JOIN [Grants].Main AS G ON G.GranteeID=A.[Grantee ID] AND G.[FiscalYear]=A.[Fiscal Year]-1 " & vbCrLf & _
		"WHERE [Fiscal Year]=" & prepIntegerSQL(FiscalYear) & vbCrLf & _
		"	AND Negotiation=1 " & vbCrLf
Else
	sql = "SELECT [App ID], [Grantee Name], [Program Name], [Grant Type], [Submitted By], [Submitted], [Status], " & vbCrLf & _
		"	[MVCPA Funds Requested] AS [" & (FiscalYear) & " MVCPA Funds Requested], " & vbCrLF & _
		"	G.AwardAmount AS [" & (FiscalYear-1) & " MVCPA Funds], " & vbCrLF & _
		"	[Cash Match] AS [" & FiscalYear & " Cash Match Requested], " & vbCrLf & _
		"	G.MatchAmount AS [" & (FiscalYear-1) & " Cash Match], " & vbCrLf & _
		"	[Cash Match Pct] AS [" & FiscalYear & " Cash Match Percent], " & vbCrLf & _
		"	100.0*G.MatchAmount/G.AwardAmount AS [" & (FiscalYear-1) & " Cash Match Percent], " & vbCrLf & _
		"	[Cash Match Pct] - 100.0*G.MatchAmount/G.AwardAmount AS [Change in Cash Match Percent], " & vbCrLf & _
		"	([Cash Match Pct] - 100.0*(G.MatchAmount/G.AwardAmount))/(G.MatchAmount/G.AwardAmount) AS [Cash Match Percent Change], " & vbCrLf & _
		"	100.0*([MVCPA Funds Requested]/G.AwardAmount) AS [MVCPA Pct of Last Year], " & vbCrLf & _
		"	100*[Cash Match]/G.MatchAmount AS [Cash Match Pct of Last Year], " & vbCrLf & _
		"	[In-Kind Match] AS [" & FiscalYear & " In-Kind Match] " & vbCrLf & _
		"FROM Application.vwSummary AS A " & vbCrLf & _
		"LEFT JOIN [Grants].Main AS G ON G.GranteeID=A.[Grantee ID] AND G.[FiscalYear]=A.[Fiscal Year]-1 " & vbCrLf & _
		"WHERE [Fiscal Year]=" & prepIntegerSQL(FiscalYear) & vbCrLf
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
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;&nbsp;<a href="KeyChecks.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a>
&nbsp;&nbsp;&nbsp;<span style="font-size: smaller; ">(Using <%=ApplicationSchema %> tables.)</span></form>
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
<!--#include file="../includes/getApplicationSchema.asp"-->