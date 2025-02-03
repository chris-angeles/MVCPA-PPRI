<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, OrderByDescription, OrderByField, ShowExcel, Columns, _
	ShowOnlySubmitted, ApplicationSchema
OrderByDescription = Array("App ID", "Grantee Name")
OrderByField = Array("[App ID]", "REPLACE([GranteeName],'City of ','')")
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

sql = "SELECT I.GranteeId AS [Grantee ID], A.GranteeName AS [Grantee Name], B.AppID AS [App ID], I.FiscalYear AS [Fiscal Year], " & vbCrLf & _
	"	Personnel, Overtime, CASE WHEN [Personnel]>0 THEN 100.0*[Overtime]/[Personnel] ELSE NULL END AS [Percent Overtime], " & vbCrLf & _
	"	--[NICB Personnel], [NICB Overtime], CASE WHEN [NICB Personnel]>0 THEN 100.0*[NICB Overtime]/[NICB Personnel] ELSE NULL END AS [NICB Percent Overtime], " & vbCrLf & _
	"	--[DPS Personnel], [DPS Overtime], CASE WHEN [DPS Personnel]>0 THEN 100.0*[DPS Overtime]/[DPS Personnel] ELSE NULL END AS [DPS Percent Overtime], " & vbCrLf & _
	"	[Subgrantee LEO Personnel], [Subgrantee LEO Overtime], CASE WHEN [Subgrantee LEO Personnel]>0 THEN 100.0*[Subgrantee LEO Overtime]/[Subgrantee LEO Personnel] ELSE NULL END AS [Subgrantee LEO Percent Overtime], " & vbCrLf & _
	"	[Subgrantee Support Personnel], [Subgrantee Support Overtime], CASE WHEN [Subgrantee Support Personnel]>0 THEN 100.0*[Subgrantee Support Overtime]/[Subgrantee Support Personnel] ELSE NULL END AS [Subgrantee Support Percent Overtime], " & vbCrLf & _
	"	ISNULL([Personnel],0.0) + ISNULL([Subgrantee LEO Personnel],0.0)+ISNULL([Subgrantee Support Personnel],0.0) AS [Total Personnel], " & vbCrLf & _
	"	ISNULL([Overtime],0.0) + ISNULL([Subgrantee LEO Overtime],0.0) + ISNULL([Subgrantee Support Overtime],0.0) AS [Total Overtime], " & vbCrLf & _
	"	CASE WHEN ISNULL([Personnel],0.0) + ISNULL([Subgrantee LEO Personnel],0.0)+ISNULL([Subgrantee Support Personnel],0.0) > 0 THEN " & vbCrLf & _
	"		100.0*(ISNULL([Overtime],0.0) + ISNULL([Subgrantee LEO Overtime],0.0) + ISNULL([Subgrantee Support Overtime],0.0)) / " & vbCrLF & _
	"		(ISNULL([Personnel],0.0) + ISNULL([Subgrantee LEO Personnel],0.0)+ISNULL([Subgrantee Support Personnel],0.0)) " & vbCrLf & _
	"	ELSE NULL END AS [Total Percent Overtime] " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"LEFT JOIN Application.IDs AS I ON I.GranteeID=A.GranteeID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".Main AS B ON B.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".vwPersonnelTotals C ON C.AppID=B.AppID " & vbCrLf & _
	"WHERE I.[FiscalYear]=" & prepIntegerSQL(FiscalYear) & vbCrLf
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
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Overtime Report</title>
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
%>&nbsp;&nbsp;&nbsp;
</select>&nbsp;&nbsp;&nbsp;<a href="Overtime.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a>
&nbsp;&nbsp;&nbsp;<span style="font-size: smaller; ">(Using <%=ApplicationSchema %> tables.)</span>
</form>
</div>
<table class="reporttable">
<%
End if

If rs.EOF = False Then
	Columns = rs.Fields.count
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
			ElseIf rs.Fields(i).Name = "Grantee ID" Then
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
			ElseIf rs.Fields(i).Name="Grantee Name" Then
				Response.Write("<td style=""text-align: left; white-space: nowrap"">" & rs.Fields(i).value & "</td>")
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal Year" Then
				Response.Write("<td style=""text-align: right"">" & rs.Fields(i).value & "</td>")
			ElseIf InStr(rs.Fields(i).Name, "Percent Overtime") Then
				If IsNull(rs.Fields(i).value) Then
					Response.Write("<td></td>" & vbCrLf)
				ElseIf CDbl(rs.Fields(i)) > 5.0 Then
					Response.Write("<td style=""text-align: right; color: red; "">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "%</td>")
				Else
					Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "%</td>")
				End If
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

If ShowExcel = False Then %>
<tr><th style="width: 100%; text-align: center" colspan="<%=columns %>"><input type="button" value="Close" onclick="window.close();" /></th></tr>
<%	
End If 
%>
</table>
<% If ShowExcel = False Then %>
</body>
</html>
<%	End If %>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/getApplicationSchema.asp"-->