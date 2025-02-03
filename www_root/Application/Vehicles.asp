﻿<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, OrderByDescription, OrderByField, ShowExcel, Columns, _
	ShowOnlySubmitted, ApplicationSchema
OrderByDescription = Array("App ID", "Grantee ID", "Grantee Name")
OrderByField = Array("I.AppID", "[Grantee ID]", "REPLACE([GranteeName],'City of ','')")
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

If FiscalYear = 2018 Then
	ShowOnlySubmitted = True
ElseIf FiscalYear = 2019 Then
	ShowOnlySubmitted = False
ElseIf FiscalYear = 2020 Then
	ShowOnlySubmitted = False
ElseIf FiscalYear = 2021 Then
	ShowOnlySubmitted = False
Else
	ShowOnlySubmitted = True
End If

sql = "SELECT I.AppID, I.FiscalYear AS [Fiscal Year], I.GranteeID AS [Grantee ID], B.GranteeName AS [Grantee Name], " & vbCrLf & _
	"	ROUND(AVG(ModelYear),1) AS [Average Of Model Year], COUNT(*) AS Vehicles, " & vbCrLF & _
	"	MAX(D.LEOs) AS LEOs, COUNT(*)/MAX(D.LEOs) AS [Inventory Vehicles / LEO], " & vbCrLf & _
	"	MAX(E.[Leased Vehicles]) AS [Leased Vehicles], MAX(CASE WHEN E.[Unspecified Leases]=1 THEN 'yes' ELSE '' END) AS [Unspecified Leases], " & vbCrLf & _
	"	(COUNT(*)+ MAX(E.[Leased Vehicles])) / MAX(D.LEOs) AS [Vehicles / LEO]" & vbCrLf & _
	"FROM Application.IDs AS I " & vbCrLf & _
	"LEFT JOIN Application.Main AS A ON A.AppID=I.AppID " & vbCrLf & _
	"JOIN Grantees AS B ON B.GranteeID=I.GranteeID " & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT InventoryID, GranteeID, AssetClassID, ItemDescription, ModelYear, Model, MakeManufacturer, UseID, ConditionID " & vbCrLf & _
	"	FROM Inventory " & vbCrLf & _
	"	WHERE AssetClassID='01-01' " & vbCrLf & _
	"		AND DateOfDisposal IS NULL " & vbCrLf & _
	"		AND ISNULL(ConditionID,2) IN (1,2) " & vbCrLf & _
	"		AND ISNULL(UseID,2)<> 1 " & vbCrLf & _
	"		AND MakeManufacturer NOT LIKE '%Honda%' AND MakeManufacturer NOT LIKE '%Bobcat%' " & vbCrLf & _
	"		AND MakeManufacturer NOT LIKE '%John Deere%' AND MakeManufacturer NOT LIKE '%Polaris%' " & vbCrLf & _
	") AS C ON C.GranteeID=I.GranteeID " & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT AppID, 1.0*SUM(PctTime)/100 AS LEOs " & vbCrLf & _
	"	FROM Application.BudgetDetails " & vbCrLf & _
	"	WHERE BudgetCategoryID=1 AND SubCategoryID=1 " & vbCrLf & _
	"	GROUP BY AppID " & vbCrLf & _
	") AS D ON D.AppID=A.AppID " & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT AppID, SUM(ISNULL(NoOfItems,1)) AS [Leased Vehicles], MAX(CASE WHEN NoOfItems IS NULL THEN 1 ELSE 0 END) AS [Unspecified Leases] " & vbCrLf & _
	"	FROM Application.BudgetDetails " & vbCrLf & _
	"	WHERE BudgetCategoryID=7 AND Description LIKE '%vehicle%' AND Description LIKE '%lease%' AND Description NOT LIKE '%maintenance%' " & vbCrLf & _
	"	GROUP BY AppID " & vbCrLf & _
	") AS E ON E.AppID=A.AppID " & vbCrLf & _
	"WHERE I.FiscalYear=" & FiscalYear & " " & vbCrLf & _
	"GROUP BY I.AppID, I.FiscalYear, I.GranteeID, B.GranteeName " & vbCrLf
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
<title>Vehicle Report</title>
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
</select>&nbsp;&nbsp;&nbsp;<a href="Vehicles.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a>
</form>

<table class="reporttable" style="width: 760px; ">
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
					Response.Write("<td style=""text-align: right; "">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "%</td>")
				End If
			ElseIf rs.Fields(i).Name = "Average Of Model Year" Then
				Response.Write("<td style=""text-align: center; "">" & formatnumber(rs.Fields(i).value, 1, true, true, false) & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right; "">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adNumeric OR rs.Fields(i).Type=adDouble Then
				Response.Write("<td style=""text-align: right; "">" & formatnumber(rs.Fields(i).value,2, true, false, true) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		'Response.Write("<td>" & rs.Fields("Average Of Model Year").Type & "</td>" & vbCrLf)
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