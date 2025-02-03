<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, OrderByDescription, OrderByField, _
	Show, ShowDescription, ShowClause, ShowExcel,LastAppID, Total, ApplicationSchema
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

ShowDescription = Array ("All", "Border", "Port", "Port 2", "Border and Port", "Border, Port, and Port 2")
ShowClause = Array ("1=1", "C.BorderCounty=1", "C.PortCounty=1", "C.Port2County=1", "(C.BorderCounty=1 OR C.PortCounty=1)", "(C.BorderCounty=1 OR C.PortCounty=1 OR C.Port2County=1)")

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

If Len(Request.Form("Show")) > 0 Then
	Show = CInt(Request.Form("Show"))
ElseIf Len(Request.QueryString("Show")) > 0 Then
	Show = CInt(Request.QueryString("Show"))
Else
	Show = 0
End If
If Request.Form("ShowExcel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

ApplicationSchema = getApplicationSchema(FiscalYear)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=GrantReport" & FiscalYear & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Grant Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2018 to Year(Now())+1
		Response.Write("<option value=""" & i & """ " & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="Show">Show:</label> <select name="Show" id="Show" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowDescription)
	Response.Write("<option value=""" & i & """ " & Selected(Show, i) & ">" & ShowDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;<a href="PersonnelReport.asp?FiscalYear=<%=FiscalYear %>&Show=<%=Show %>&ShowExcel=1">Show Excel</a>
&nbsp;&nbsp;&nbsp;<span style="font-size: smaller; ">(Using <%=ApplicationSchema %> tables.)</span>
</form>
<br />

<%
End If
sql = "SELECT ISNULL(REPLACE(C.GranteeName,'City of ',''),'" & ShowDescription(Show) & " Grants') AS [Grantee Name], " & vbCrLf & _
	"	E.BudgetCategory AS [Budget Category], " & vbCrLf & _
	"	F.SubCategory AS [Subcategory], COUNT(*) AS Count, SUM(PctTIme) AS [Pct Time], " & vbCrLf & _
	"	SUM((ISNULL(LineTotal,0.0)+ISNULL(InKindMatch,0.0))/PctTime*100.0)/COUNT(*) AS [Average] " & vbCrLf & _
	"FROM Application.Admin AS A " & vbCrLf & _
	"LEFT JOIN Application.IDs AS I ON I.AppID=A.AppID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".Main AS B ON B.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN Grantees AS C ON C.GranteeID=I.GranteeID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".BudgetDetails AS D ON D.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN Lookup.BudgetCategories AS E ON E.BudgetCategoryID=D.BudgetCategoryID " & vbCrLf & _
	"LEFT JOIN Lookup.BudgetSubcategories AS F ON F.BudgetCategoryID=E.BudgetCategoryID AND F.SubCategoryID=D.SubCategoryID " & vbCrLf & _
	"WHERE I.FiscalYear=" & FiscalYear & " AND (D.BudgetCategoryID=1 OR (D.BudgetCategoryID=4 AND D.SubCategoryID IN (1,4,7,13))) AND " & ShowClause(Show) & " " & vbCrLf & _
	"GROUP BY GROUPING SETS ( " & vbCrLf & _
	"	(REPLACE(C.GranteeName,'City of ',''), D.BudgetCategoryID, D.SubCategoryID, E.BudgetCategory, F.SubCategory), " & vbCrLf & _
	"	(D.BudgetCategoryID, D.SubCategoryID, E.BudgetCategory, F.SubCategory) ) " & vbCrLf & _
	"ORDER BY CASE WHEN REPLACE(C.GranteeName,'City of ','') IS NULL THEN 2 ELSE 1 END,REPLACE(C.GranteeName,'City of ',''), D.BudgetCategoryID, D.SubCategoryID"

	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

Set rs=Con.Execute(sql)

Total = 0
If rs.EOF = False Then
	Response.Write("<table class=""reporttable"">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr>" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)

	While rs.EOF = False
		Total = Total + rs.Fields("Count")
		Response.Write("<tr>" & vbCrLf)
		Response.Write("<td>" & rs.Fields("Grantee Name").value & "</td>")
		Response.Write("<td>" & rs.Fields("Budget Category").value & "</td>")
		Response.Write("<td>" & rs.Fields("Subcategory").value & "</td>")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("Count").value & "</td>")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("Pct Time").value & "%</td>")
		Response.Write("<td style=""text-align: right; "">$" & formatnumber(rs.Fields("Average").value, 2, true, true, true) & "</td>")
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.WRite("<tr><td>Overall Total</td><td></td><td></td><td style=""text-align: right; "">" & (Total/2) & "</td></tr>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If

Response.Write("<br />" & vbCrLf)

sql = "SELECT Position_No = ROW_NUMBER() OVER (PARTITION BY I.GranteeID ORDER BY D.BudgetCategoryID, D.SubCategoryID), " & vbCrLf & _
	"	A.AppID, I.GranteeID, D.BudgetCategoryID, D.SubCategoryID, " & vbCrLF & _
	"	REPLACE(C.GranteeName,'City of ','') AS GranteeName, B.ProgramName, E.BudgetCategory, F.SubCategory, " & vbCrLf & _
	"	D.Description, D.PctTime, ISNULL(D.LineTotal,0.0)+ISNULL(D.InKindMatch,0.0)/ISNULL(PctTime,100.0)*100.0 AS Salary" & vbCrLf & _
	"FROM Application.Admin AS A " & vbCrLf & _
	"LEFT JOIN Application.IDs AS I ON I.AppID=A.AppID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".Main AS B ON B.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN Grantees AS C ON C.GranteeID=I.GranteeID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".BudgetDetails AS D ON D.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN Lookup.BudgetCategories AS E ON E.BudgetCategoryID=D.BudgetCategoryID " & vbCrLf & _
	"LEFT JOIN Lookup.BudgetSubcategories AS F ON F.BudgetCategoryID=E.BudgetCategoryID AND F.SubCategoryID=D.SubCategoryID " & vbCrLf & _
	"WHERE I.FiscalYear=" & FiscalYear & " AND (D.BudgetCategoryID=1 OR (D.BudgetCategoryID=4 AND D.SubCategoryID IN (1,4,7,13))) AND " & ShowClause(Show) & " " & vbCrLf & _
	"ORDER BY REPLACE(C.GranteeName,'City of ',''), D.BudgetCategoryID, D.SubCategoryID"
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

Set rs=Con.Execute(sql)

Response.Write("<table class=""reporttable"">" & vbCrLf)
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr>" & vbCrLF)
	Response.Write(vbTab & "<th>#</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Budget Category</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Sub Category</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Description</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Pct Time</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Calc. Salary</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)

	LastAppID = 0
	While rs.EOF = False
		If LastAppID <> rs.Fields("AppID") Then
			LastAppID = rs.Fields("AppID")
			Response.Write("<tr><td colspan=""5"" style=""text-align: center; font-weight: bold; "">" & rs.Fields("GranteeName") & ", " & rs.Fields("ProgramName") & "</td></tr>" & vbCrLf)
		End If
		Response.Write("<tr>" & vbCrLF)
		Response.Write(vbTab & "<td>" & rs.Fields("Position_No") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("SubCategory") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("Description") & "</td>" & vbCrLf)
		If IsNull(rs.Fields("PctTime")) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right;"">" & rs.Fields("PctTime") & "%</td>" & vbCrLf)
		End If
		Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("Salary"),2, true, true, true) & "</td>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
Response.Write("</table>" & vbCrLf)

If ShowExcel = False Then
%>


<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/getApplicationSchema.asp"-->