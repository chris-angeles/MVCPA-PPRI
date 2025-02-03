<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><%
Dim debug, i, j, ShowExcel, RoundCurrency, FiscalYear, AppID, Grant_Number, Grant_Award_Amount, Grantee_Name, _
	TotalMVCPAFunds, GrandTotal, TotalCashMatch, InKind_Match, PctMVCPA, PctCashMatch, _
	Reimbursement_Rate, Reimbursement_Description, In_Lieu_Of_NICB, Submission, ApplicationSchema
Debug = False

RoundCurrency = True
ShowExcel = True
If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	Response.Write("No Fiscal Year provided")
	Response.End
End If

ApplicationSchema = getApplicationSchema(FiscalYear)

If ShowExcel = True and Debug = False Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=SGADetail" & FiscalYear & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%>
<!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Taskforce Grant Details for Statement of Grant Award</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style type="text/css">	th {
		text-align: center;
	}
</style>
</head>
<body>
<%
End If
%>
<table>
<% 
sql = "SELECT * " & vbCrLf & _
	"FROM " & ApplicationSchema & ".vwReimbursementRates " & vbCrLf & _
	"WHERE Fiscal_Year=" & FiscalYear & " " & vbCrLf & _
	"ORDER BY Grantee_Sort "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
While rs.EOF = False
	AppId = rs.Fields("AppID")
	Grant_Number = rs.Fields("Grant_Number")
	Grant_Award_Amount = rs.Fields("Grant_Award_Amount")
	Grantee_Name = rs.Fields("Grantee_Name")
	TotalCashMatch = rs.Fields("Cash_Match")
	InKind_Match = rs.Fields("InKind_Match")
	TotalMVCPAFunds = rs.Fields("MVCPA_Funds")
	GrandTotal = rs.Fields("Cash_Match") + rs.Fields("MVCPA_Funds")
	Reimbursement_Rate = rs.Fields("Reimbursement_Rate")
	In_Lieu_Of_NICB = rs.Fields("In_Lieu_Of_NICB")
	Submission = rs.Fields("Submission")

	PctMVCPA = 100.0*TotalMVCPAFunds / GrandTotal
	PctCashMatch = 100.0*TotalCashMatch / TotalMVCPAFunds

	Reimbursement_Description = "*Reimbursement Percent: " & prepNumberWeb(Reimbursement_Rate, 2) & "% - " & _
		prepCurrencyWebRound(Grant_Award_Amount, RoundCurrency) & "-MVCPA Amt / (" & _
		prepCurrencyWebRound(GrandTotal, RoundCurrency) & "-MVCPA Amt. plus " & _
		prepCurrencyWebRound(TotalCashMatch, RoundCurrency) & "-Cash Match"
	If IsNull(In_Lieu_Of_NICB) = False Then
		Reimbursement_Description = Reimbursement_Description & " minus " & _
			prepCurrencyWebRound(GrandTotal, RoundCurrency) & "-amt. in Lieu of"
	End If
	Reimbursement_Description = Reimbursement_Description & ")"

	Response.Write("<tr>")
	Response.Write(vbTab & "<td colspan=""5"" style=""text-align: center; font-weight: bold; color: red; "">" & Grantee_Name & ", " & Submission & "</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr><td>&nbsp;</td></tr>" & vbCrLf)

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>App ID:</td>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">" & AppID & "</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>Grant Number:</td>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">" & Grant_Number & "</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>Grantee:</td>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">" & Grantee_Name & "</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>Program Title:</td>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">" & rs.Fields("Program_Name") & "</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>Grant Award Amount:</td>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">" & prepCurrencyWebRound(Grant_Award_Amount, RoundCurrency) & "</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>Total Cash Match Amount:</td>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">" & prepCurrencyWebRound(TotalCashMatch, RoundCurrency) & "</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>In-Kind Match Amount:</td>" & vbCrLf)
	If IsNull(rs.Fields("InKind_Match")) = True Then
		Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">-0-</td>" & vbCrLf)
	Else
		Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">" & prepCurrencyWebRound(InKind_Match, RoundCurrency) & "</td>" & vbCrLf)
	End If
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>Reimbursement Percent*:</td>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">" & prepNumberWeb(Reimbursement_Rate, 2) & "%</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write(vbTab & "<td>Grant Term:</td>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""4"" style=""text-align: left; font-weight: bold; "">September 1, " & (FiscalYear-1) & " to August 31, " & FiscalYear & "</td>" & vbCrLf)
	Response.Write("</tr>")

	Response.Write("<tr><td>&nbsp;</td></tr>" & vbCrLf)

	DisplayBudget(AppID)

	Response.Write("<tr><td>&nbsp;</td></tr>" & vbCrLf)
	If ShowExcel = True Then
		Response.Write("<tr><td colspan=""5"">------------------------------------------------------------------------------------------------------------------------</td></tr>" & vbCrLf)
	Else
		Response.Write("<tr><td colspan=""5""><hr></td></tr>" & vbCrLf)
	End If
	Response.Write("<tr><td>&nbsp;</td></tr>" & vbCrLf)
	rs.MoveNext
Wend
%>
</table>
<%
If ShowExcel = False Or Debug = True Then
%>
</body>
</html>
<%
End If
%>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/getApplicationSchema.asp"-->
<%
Sub DisplayBudgetDetails(vAppID)
	Dim vsql, vrs, lastcategory 
	vsql = "SELECT B.BudgetItemID, A.BudgetCategoryID, A.BudgetCategory, " & vbCrLf & _
		"	CASE WHEN B.NoOfItems>0 THEN ISNULL(B.Description,'') + ' (' + CAST(B.NoOfItems AS VARCHAR) + ')' ELSE B.Description END AS Description, " & vbCrLf & _
		"	SubCategory, PctTime, LineTotal, MVCPAFunds, CashMatch, InKindMatch " & vbCrLf & _
		"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".BudgetDetails AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND AppID=" & prepIntegerSQL(vAppID) & " " & vbCrLf & _
		"LEFT JOIN Lookup.BudgetSubcategories AS C ON C.BudgetCategoryID=B.BudgetCategoryID AND C.SubCategoryID=B.SubCategoryID " & vbCrLf & _
		"UNION " & vbCrLf & _
		"SELECT 2147483647 AS BudgetItemID, A.BudgetCategoryID, A.BudgetCategory, 'Total ' + A.BudgetCategory AS Description, null, SUM(PctTime) AS PctTime, SUM(LineTotal) AS LineTotal, SUM(MVCPAFunds) AS MVCPAFunds, SUM(CashMatch) AS CashMatch, Sum(InKindMatch) AS InKindMatch " & vbCrLf & _
		"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".BudgetDetails AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND AppID=" & prepIntegerSQL(vAppID) & " " & vbCrLf & _
		"GROUP BY A.BudgetCategoryID, A.BudgetCategory" & vbCrLf & _
		"ORDER BY 2, 1 "
	LastCategory=0
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
		Response.Write("<th>Description</th>" & vbCrLf)
		Response.Write("<th>Subcategory</th>" & vbCrLf)
		Response.Write("<th>Pct Time</th>" & vbCrLf)
		Response.Write("<th style=""width: 100px; "">MVCPA Funds</th>" & vbCrLf)
		Response.Write("<th style=""width: 100px; "">Cash Match</th>" & vbCrLf)
		Response.Write("<th style=""width: 100px; "">Total</th>" & vbCrLf)
		Response.Write("<th style=""width: 100px; "">In-Kind Match</th>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)

		While vrs.EOF = False
			If LastCategory<>vrs.Fields("BudgetCategoryID") Then
				LastCategory=vrs.Fields("BudgetCategoryID")
				Response.Write("<tr><td colspan=""6"">&nbsp;</td></tr>" & vbCrLf)
				Response.Write("<tr><th colspan=""6"">" & vrs.Fields("BudgetCategory") & "</th></tr>" & vbCrLf)
			End If
			Response.Write("<tr>" & vbCrLf)
			Response.Write("<td>" & vrs.Fields("Description") & "</td>")
			Response.Write("<td>" & vrs.Fields("SubCategory") & "</td>")
			Response.Write("<td style=""text-align: right; "">" & vrs.Fields("PctTime") & "</td>")
			Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("MVCPAFunds"), RoundCurrency) & "</td>")
			Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("CashMatch"), RoundCurrency) & "</td>")
			Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("LineTotal"), RoundCurrency) & "</td>")
			Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(vrs.Fields("InKindMatch"), RoundCurrency) & "</td>")
			Response.Write("</tr>" & vbCrLf)
			vrs.MoveNext
		Wend
	End If
End Sub

Sub DisplayBudget(vAppID)
	Dim vsql, vrs, lastcategory 

	Response.Write("<tr style=""vertical-align: bottom"">" & vbCrLf)
	Response.Write("<th colspan=""5"">Grant Budget Summary: " & Grantee_Name & " (App ID: " & AppID & ")</th>" & vbCrLf)

	Response.Write("</tr>" & vbCrLf)

	Response.Write("<tr style=""vertical-align: bottom"">" & vbCrLf)
	Response.Write("<th>Budget Category</th>" & vbCrLf)
	Response.Write("<th>MVCPA<br />Expenditures</th>" & vbCrLf)
	Response.Write("<th>Cash<br />Match<br />Expenditures</th>" & vbCrLf)
	Response.Write("<th>Total<br />Expenditures</th>" & vbCrLf)
	Response.Write("<th>In-Kind<br />Match</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)

	vsql = "SELECT ISNULL(A.BudgetCategoryID,99) AS BudgetCategoryID, ISNULL(A.BudgetCategory, 'Total') As BudgetCategory, " & vbCrLf & _
		"	SUM(LineTotal) AS LineTotal, SUM(MVCPAFunds) AS [MVCPAFunds], " & vbCrLf & _
		"	SUM(CashMatch) AS [CashMatch], SUM(InKindMatch) AS [InKindMatch] " & vbCrLf & _
		"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".BudgetDetails AS B ON A.BudgetCategoryID=B.BudgetCategoryID AND B.AppID=" & _
			prepIntegerSQL(vAppID) & " " & vbCrLf & _
		"GROUP BY GROUPING SETS ((A.BudgetCategoryID,A.BudgetCategory),()) " & vbCrLf & _
		"ORDER BY ISNULL(A.BudgetCategoryID,99) "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	While vrs.EOF = False
		Response.Write(vbTab & "<tr style=""vertical-align: top; "">" & vbCrLf)
		If vrs.Fields("BudgetCategoryID")=99 Then
			Response.Write(vbTab & "<td style=""text-align: left; "">Totals</td>" & vbCrLf)  
		Else
			Response.Write(vbTab & "<td style=""text-align: left; "">" & vrs.Fields("BudgetCategory") & "</td>" & vbCrLf)  

		End If
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(vrs.Fields("MVCPAFunds"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(vrs.Fields("CashMatch"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(vrs.Fields("LineTotal"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(vrs.Fields("InKindMatch"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>")
		vrs.MoveNext
	Wend
	If ISNull(Reimbursement_Rate) = False Then
		Response.Write("<tr><td style=""text-align: left;"" colspan=""5"">" & Reimbursement_Description & "</td></tr>" & vbCrLf)
		'Response.Write("<tr><td style=""text-align: center;"">Cash Match Percentage</td><td></td><td style=""text-align: right; "">" & prepNumberWeb(PctCashMatch, 2) & "%</td><td></td><td></td></tr>" & vbCrLf)
	End If
	'Response.Write("<tr><td style=""text-align: center;"">Reimbursement Percentage</td><td></td><td style=""text-align: right; "">" & prepNumberWeb(Reimbursement_Rate, 2) & "%</td><td></td><td></td></tr>" & vbCrLf)
End Sub
%>