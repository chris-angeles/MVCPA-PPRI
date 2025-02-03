<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, GrantID, Quarter, FiscalYear, _
	StatePayeeIDNo, GrantNumber, GranteeName, ProgramName, _
	VendorOrganizationalUnit, VendorAddress1, VendorAddress2, VendorCity, VendorState, VendorZip, _
	InLieuOfDPS, InLieuOfNICB, PriorInLieuOfDPS, PriorInLieuOfNICB, UnbudgetedPI, _
	ReimbursableExpenditures, ReimbursementRate, ReimbursementYTD, PriorAmountPaid, Reimbursement, _
	CurrentSaleOfAssets, PriorSaleOfAssets, CurrentMVCPADeduction, PriorMVCPADeduction, _
	PriorYearFunds, CurrentYearFunds, ReviewDate, Reviewer, ReviewerTitle, _
	AuditApprovalDate, Auditor, AuditorTitle, _
	DirectorApprovalDate, Director, DirectorTitle, SubmitName, SubmitDate
Dim CashExpenditure_Total, InKindExpenditure_Total, YTDExpenditure_Total, _
	BudgetExpenditure_Total, RemainingBudget_Total
debug = False
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("<pre>Session(""" & i & """)='" & Session(i) & "'</pre>" & vbCrLf)
	Next
End If

If Len(Request.Form("GrantID"))>0 Then
	GrantID = Request.Form("GrantID")
ElseIf Len(Request.QueryString("GrantID"))>0 Then
	GrantID = Request.QueryString("GrantID")
Else
	Response.Write("Error: No GrantID provided")
	SendMessage "Error: No GrantID provided"
	Response.End
End If

If Len(Request.Form("Quarter"))>0 Then
	Quarter = Request.Form("Quarter")
ElseIf Len(Request.Querystring("Quarter"))>0 Then
	Quarter = Request.Querystring("Quarter")
Else
	Response.Write("Error: No Quarter provided")
	SendMessage "Error: No Quarter provided"
	Response.End
End If

sql = "SELECT B.GrantID, B.FiscalYear, C.Quarter, A.GranteeName, B.ProgramName, A.StatePayeeIDNo, B.GrantNumber, " & vbCrLf & _
	"	A.VendorOrganizationalUnit, A.VendorAddress1, A.VendorAddress2, A.VendorCity, A.VendorState, A.VendorZip, " & vbCrLf & _
	"	ISNULL(C.PriorYearFunds,0.0) AS PriorYearFunds, " & vbCrLf & _
	"	ISNULL(C.CurrentYearFunds,0.0) AS CurrentYearFunds, " & vbCrLf & _
	"	ISNULL(C.InLieuOfDPS,0.0) AS InLieuOfDPS, " & vbCrLf & _
	"	ISNULL(C.InLieuOfNICB, 0.0) AS InLieuOfNICB, " & vbCrLf & _
	"	ISNULL(C.PriorInLieuOfDPS,0.0) AS PriorInLieuOfDPS, " & vbCrLF & _
	"	ISNULL(PriorInLieuOfNICB, 0.0) AS PriorInLieuOfNICB, " & vbCrLF & _
	"	C.UnbudgetedPI, C.ReimbursableExpenditures, C.ReimbursementRate, " & vbCrLF & _
	"	C.ReimbursementYTD, C.PriorAmountPaid, C.CurrentSaleOfAssets, C.Reimbursement, " & vbCrLf & _
	"	ISNULL(C.CurrentMVCPADeduction, 0.0) AS CurrentMVCPADeduction, " & vbCrLf & _
	"	ISNULL(C.PriorMVCPADeduction, 0.0) AS PriorMVCPADeduction, " & vbCrLf & _
	"	ISNULL(C.PriorSaleOfAssets, 0.0) AS PriorSaleOfAssets, " & vbCrLf & _
	"	C.ReviewDate, ISNULL(G.Name, 'Not Reviewed') AS Reviewer, ISNULL(G.Title, 'Reviewer') AS ReviewerTitle, " & vbCrLF & _
	"	C.AuditApprovalDate, ISNULL(D.Name, 'Not Yet Approved') AS Auditor, ISNULL(D.Title, 'Auditor') AS AuditorTitle, " & vbCrLF & _
	"	C.DirectorApprovalDate, ISNULL(E.Name, 'Not Yet Approved') AS Director, ISNULL(E.Title, 'Director') AS DirectorTitle, " & vbCrLF & _
	"	F.Name AS SubmitName, C.SubmitTimestamp AS SubmitDate " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS B ON B.GranteeID=A.GranteeID " & vbCrLf & _
	"LEFT JOIN ER.Main AS C ON C.GrantID=B.GrantID AND C.Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"LEFT JOIN [System].Users AS D ON D.SystemID=C.AuditApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS E ON E.SystemID=C.DirectorApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS F ON F.SystemID=C.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS G ON G.SystemID=C.ReviewID " & vbCrLf & _
	"WHERE B.GrantID=" & prepIntegerSQL(GrantID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Expenditure Report record retrieved for Grant " & GrantID & " and Quarter" & Quarter)
	SendMessage "Error: No Expenditure Report record retrieved for Grant " & GrantID & " and Quarter" & Quarter
	Response.End
End If

GrantID = rs.Fields("GrantID")
FiscalYear = rs.Fields("FiscalYear")
Quarter = rs.Fields("Quarter")
GranteeName = rs.Fields("GranteeName")
ProgramName = rs.Fields("ProgramName")
StatePayeeIDNo = rs.Fields("StatePayeeIDNo")
GrantNumber = rs.Fields("GrantNumber")
VendorOrganizationalUnit = rs.Fields("VendorOrganizationalUnit")
VendorAddress1 = rs.Fields("VendorAddress1")
VendorAddress2 = rs.Fields("VendorAddress2")
VendorCity = rs.Fields("VendorCity")
VendorState = rs.Fields("VendorState")
VendorZip = rs.Fields("VendorZip")
InLieuOfDPS = rs.Fields("InLieuOfDPS")
InLieuOfNICB = rs.Fields("InLieuOfNICB")
UnbudgetedPI = rs.Fields("UnbudgetedPI")
PriorInLieuOfDPS = rs.Fields("PriorInLieuOfDPS")
PriorInLieuOfNICB = rs.Fields("PriorInLieuOfNICB")
ReimbursableExpenditures = rs.Fields("ReimbursableExpenditures")
ReimbursementRate = rs.Fields("ReimbursementRate")
ReimbursementYTD = rs.Fields("ReimbursementYTD")
PriorAmountPaid = rs.Fields("PriorAmountPaid")
PriorSaleOfAssets = rs.Fields("PriorSaleOfAssets")
CurrentSaleOfAssets = rs.Fields("CurrentSaleOfAssets")
CurrentMVCPADeduction = rs.Fields("CurrentMVCPADeduction")
PriorMVCPADeduction = rs.Fields("PriorMVCPADeduction")
Reimbursement = rs.Fields("Reimbursement")
PriorYearFunds = rs.Fields("PriorYearFunds")
CurrentYearFunds = rs.Fields("CurrentYearFunds")
ReviewDate = rs.Fields("ReviewDate")
Reviewer = rs.Fields("Reviewer")
ReviewerTitle = rs.Fields("ReviewerTitle")
AuditApprovalDate = rs.Fields("AuditApprovalDate")
Auditor = rs.Fields("Auditor")
AuditorTitle = rs.Fields("AuditorTitle")
DirectorApprovalDate = rs.Fields("DirectorApprovalDate")
Director = rs.Fields("Director")
DirectorTitle = rs.Fields("DirectorTitle")
SubmitName = rs.Fields("SubmitName")
SubmitDate = rs.Fields("SubmitDate")


%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Voucher,GrantID=<%=GrantID %>,Q<%=Quarter %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<table style="margin: auto; font-size: 15px; ">
<thead>
	<tr><th colspan="5">Texas Motor Vehicle Crime Prevention Authority</th></tr>
	<tr><th colspan="5">Grant Payment Voucher</th></tr>
</thead>
<tbody>
	<tr><td colspan="5">&nbsp;</td></tr>
	<tr>
		<th>Vendor No.:</th>
		<td><%=StatePayeeIDNo %></td>
		<td></td>
		<th>Voucher #:</th>
		<td><%=(GrantNumber & " Q" & Quarter) %></td>
	</tr>

	<tr>
		<th colspan="5">Payee:</th>
	</tr>

	<tr>
		<td colspan="5"><%
Response.Write(GranteeName & "<br />" & vbCrLf)
If IsNull(VendorOrganizationalUnit) = False Then
	Response.Write(VendorOrganizationalUnit & "<br />" & vbCrLf)
End If
If IsNull(VendorAddress1) = False Then
	Response.Write(VendorAddress1 & "<br />" & vbCrLf)
End If
If IsNull(VendorAddress2) = False Then
	Response.Write(VendorAddress2 & "<br />" & vbCrLf)
End If
Response.Write(VendorCity & ", " & VendorState & " " & VendorZip & "<br />" & vbCrLf)
%></td>
	</tr>

	<tr><td colspan="5">&nbsp;</td></tr>

	<tr>
		<th>Grant / PO Number</th>
		<th>PO Line Number</th>
		<th>AY</th>
		<th></th>
		<th>Amount</th>
	</tr>

<%	If CurrentYearFunds > 0.0 Then %>
	<tr>
		<td><%=GrantNumber %></td>
		<td>1</td>
		<td><%=(FiscalYear MOD 1000) %></td>
		<td></td>
		<td style="text-align: right;"><%=prepCurrencyWeb(CurrentYearFunds) %></td>
	</tr>
<%	End If %>

<%	If PriorYearFunds > 0.0 Then %>
	<tr>
		<td><%=GrantNumber %></td>
		<td>2</td>
		<td><%=((FiscalYear-1)MOD 1000) %></td>
		<td></td>
		<td style="text-align: right;"><%=prepCurrencyWeb(PriorYearFunds) %></td>
	</tr>
<%	End If %>
	<tr><td colspan="5">&nbsp;</td></tr>

	<tr>
		<th>Approved by:</th>
		<th>Date</th>
		<td></td>
		<th>Position</th>
	</tr>
<%	If Director = Auditor Then%>
	<tr>
		<td><%=Reviewer %></td>
		<td><%=ReviewDate %></td>
		<td></td>
		<td><%=ReviewerTitle %></td>
	</tr>
<%	Else%>
	<tr>
		<td><%=Auditor %></td>
		<td><%=AuditApprovalDate %></td>
		<td></td>
		<td><%=AuditorTitle %></td>
	</tr>
<%	End If %>
	<tr>
		<td><%=Director %></td>
		<td><%=DirectorApprovalDate %></td>
		<td></td>
		<td><%=DirectorTitle %></td>
	</tr>
</tbody>
</table>

<hr style="width: 680px; margin: auto; "/>

<%
CashExpenditure_Total = 0.0
InKindExpenditure_Total = 0.0
YTDExpenditure_Total = 0.0
BudgetExpenditure_Total = 0.0
RemainingBudget_Total = 0.0
sql = "SELECT C.BudgetCategoryID, C.BudgetCategory, " & vbCrLf & _
	"	D.CashExpenditure-ISNULL(D.ExcludedAmount,0.0) AS CashExpenditure, " & vbCrLf & _
	"	D.InKindExpenditure, " & vbCrLf & _
	"	ISNULL(D.YTDExpenditure,0.0) AS YTDExpenditure, " & vbCrLf & _
	"	ISNULL(D.BudgetExpenditure,0.0) AS BudgetExpenditure, " & vbCrLf & _
	"	ISNULL(D.RemainingBudget,0.0) AS RemainingBudget " & vbCrLf & _
	"FROM [Grants].Main AS A " & vbCrLf & _
	"LEFT JOIN ER.Main AS B ON B.GrantID=A.GrantID AND B.Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"CROSS JOIN lookup.BudgetCategories AS C " & vbCrLf & _
	"LEFT JOIN ER.Detail AS D ON D.GrantID=A.GrantID AND D.BudgetCategoryID=C.BudgetCategoryID AND D.Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"WHERE A.GrantID=" & prepIntegerSQL(GrantID) 
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
%>
<table style="margin: auto; font-size: 11px; ">
<tr><th>FY <%=FiscalYear %> Q<%=Quarter %> Quarterly Expenditure Report</th></tr>
<tr><th><%=GranteeName %>, <%=ProgramName %></th></tr>
<tr><th><%=GrantNumber %></th></tr>
<tr><th><%=ReportingPeriodDates(FiscalYear, Quarter) %></th></tr>
<tr><td style="text-align: center">Submitted by <%=SubmitName %> on <%=SubmitDate %></td></tr>
</table>

<table style="margin: auto; font-size: 11px; margin-top: 8px; ">
	<caption>Expenditures by Category</caption>
	<thead>
	<tr>
		<th rowspan="2">Budget Category</th>
		<th colspan="2">Quarterly Expenditures</th>
		<th colspan="3">Year to Date</th>
	</tr>
	<tr style="vertical-align: bottom">
		<th>Total Cash<br />Expenses:<br />MVCPA & Match</th>
		<th>In-Kind<br />Expenditures</th>
		<th>YTD<br />Expenditures</th>
		<th>Total<br />Grant<br />Budget</th>
		<th>Grant<br />Budget<br />Remaining</th>
	</tr>
	</thead>
	<tbody>
<%
rs.MoveFirst
While rs.EOF = False
	Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("CashExpenditure")) & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("InKindExpenditure")) & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("YTDExpenditure")) & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("BudgetExpenditure")) & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("RemainingBudget")) & "</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	If IsNull(rs.Fields("CashExpenditure")) = False Then
		CashExpenditure_Total = CashExpenditure_Total + CDbl(rs.Fields("CashExpenditure"))
	End If
	If IsNull(rs.Fields("InKindExpenditure")) = False Then
		InKindExpenditure_Total = InKindExpenditure_Total + CDbl(rs.Fields("InKindExpenditure"))
	End If
	YTDExpenditure_Total = YTDExpenditure_Total + CDbl(rs.Fields("YTDExpenditure"))
	BudgetExpenditure_Total = BudgetExpenditure_Total + CDbl(rs.Fields("BudgetExpenditure"))
	RemainingBudget_Total = RemainingBudget_Total + CDbl(rs.Fields("RemainingBudget"))
	rs.MoveNext
Wend
%>
</tbody>
<tr>
	<td style="font-weight: bold; ">Totals</td>
	<td style="text-align: right; "> <%=prepCurrencyWeb(CashExpenditure_Total) %></td>
	<td style="text-align: right; "> <%=prepCurrencyWeb(InKindExpenditure_Total) %></td>
	<td style="text-align: right; "> <%=prepCurrencyWeb(YTDExpenditure_Total) %></td>
	<td style="text-align: right; "> <%=prepCurrencyWeb(BudgetExpenditure_Total) %></td>
	<td style="text-align: right; "> <%=prepCurrencyWeb(RemainingBudget_Total) %></td>
</tr>
</table>

<table style="margin: auto; font-size: 11px; margin-top: 8px;">
<caption>Reimbursement Calculation</caption>
<tbody>
<tr>
	<td>Year To Date Expenditures</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(YTDExpenditure_Total)%></td>
</tr>
<!--<tr>
	<td>less In lieu of DPS for prior quarters</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(PriorInLieuOfDPS)%></td>
</tr>--->
<tr>
	<td>less In lieu of NICB for prior quarters</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(PriorInLieuOfNICB)%></td>
</tr>
<!--<tr>
	<td>less In lieu of DPS for quarter</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(InLieuOfDPS)%></td>
</tr>-->
<tr>
	<td>less In lieu of NICB for quarter</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(InLieuOfNICB)%></td>
</tr>
<tr title="Unbudgeted program income used is not reimbursed and limited to $1000 per fiscal year.">
	<td>less unbudgeted program income used</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(UnbudgetedPI)%></td>
</tr>
<tr>
	<td>Reimbursable Expenditures</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(ReimbursableExpenditures) %></td>
</tr>
<tr>
	<td>Reimbursement Rate</td>
	<td style="text-align: right; white-space: nowrap; "><%=ReimbursementRate%>%</td>
</tr>
<tr>
	<td style="white-space: nowrap; ">Reimbursement on YTD Expenditures</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(ReimbursementYTD) %></td>
</tr>
<tr>
	<td>less Prior Quarter Payments</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(PriorAmountPaid)%></td>
</tr>
<tr>
	<td>less Prior Quarter Sale of Assets</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(PriorSaleOfAssets)%></td>
</tr>
<tr>
	<td>less Current Sale of Assets</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(CurrentSaleOfAssets) %></td>
</tr>

<tr>
	<td>less Prior Quarter MVCPA Deduction</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(PriorMVCPADeduction)%></td>
</tr>
<tr>
	<td>less Current MVCPA Deduction</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(CurrentMVCPADeduction) %></td>
</tr>

<tr>
	<td>Reimbursement for this quarter</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(Reimbursement) %></td>
</tr>
</tbody>
</table>

</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<%
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
	Else
		ReportingPeriodDates = "Error in Reporting Period"
	End If
End Function
%>