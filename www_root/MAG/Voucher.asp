<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, MAGID, FiscalYear, PermitEdit, _
	StatePayeeIDNo, GrantNumber, GranteeName, ProgramName, OptionID, GrantOption,  _
	GrantAwardAmount, CashExpenditure, CashMatch, BudgetedAmount, ExcludedAmount, _
	VendorOrganizationalUnit, VendorAddress1, VendorAddress2, VendorCity, VendorState, VendorZip, _
	InLieuOfDPS, InLieuOfNICB, PriorInLieuOfDPS, PriorInLieuOfNICB, UnbudgetedPI, _
	ReimbursableExpenditure, ReimbursementRate, Reimbursement, AmountPaid,  _
	ReviewDate, Reviewer, ReviewerTitle, _
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

If Len(Request.Form("MAGID"))>0 Then
	MAGID = Request.Form("MAGID")
ElseIf Len(Request.QueryString("MAGID"))>0 Then
	MAGID = Request.QueryString("MAGID")
Else
	Response.Write("Error: No MAGID provided")
	SendMessage "Error: No MAGID provided"
	Response.End
End If

sql = "SELECT B.MAGID, B.FiscalYear, A.GranteeName, B.OptionID, H.GrantAwardAmount, H.CashMatch,  " & vbCrLf & _
	"	CASE WHEN B.OptionID=1 THEN 'Purchase' WHEN B.OptionID=2 THEN 'Lease' ELSE '' END AS [GrantOption], " & vbCrLf & _
	"	CASE WHEN B.FiscalYear=2022 THEN 25000 ELSE ISNULL(H.GrantAwardAmount,0.0) + ISNULL(H.CashMatch,0.0) END AS BudgetedAmount, " & vbCrLf & _
	"	C.CashExpenditure, C.ExcludedAmount, 'MVCPA Auxiliary Grant for ' + A.GranteeName AS ProgramName, " & vbCrLf & _
	"	A.StatePayeeIDNo, H.GrantNumber, C.AmountPaid, " & vbCrLf & _
	"	A.VendorOrganizationalUnit, A.VendorAddress1, A.VendorAddress2, A.VendorCity, A.VendorState, A.VendorZip, " & vbCrLf & _
	"	C.ReimbursableExpenditure, C.ReimbursementRate, C.Reimbursement, " & vbCrLf & _
	"	C.ReviewDate, ISNULL(G.Name, 'Not Reviewed') AS Reviewer, ISNULL(G.Title, 'Reviewer') AS ReviewerTitle, " & vbCrLF & _
	"	C.AuditApprovalDate, ISNULL(D.Name, 'Not Yet Approved') AS Auditor, ISNULL(D.Title, 'Auditor') AS AuditorTitle, " & vbCrLF & _
	"	C.DirectorApprovalDate, ISNULL(E.Name, 'Not Yet Approved') AS Director, ISNULL(E.Title, 'Director') AS DirectorTitle, " & vbCrLF & _
	"	F.Name AS SubmitName, C.SubmitTimestamp AS SubmitDate " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"LEFT JOIN [MAG].Main AS B ON B.GranteeID=A.GranteeID " & vbCrLf & _
	"LEFT JOIN MAG.ExpenditureReport AS C ON C.MAGID=B.MAGID " & vbCrLf & _
	"LEFT JOIN [System].Users AS D ON D.SystemID=C.AuditApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS E ON E.SystemID=C.DirectorApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS F ON F.SystemID=C.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS G ON G.SystemID=C.ReviewID " & vbCrLf & _
	"LEFT JOIN [MAG].Admin AS H ON H.MAGID=B.MAGID " & vbCrLf & _
	"WHERE B.MAGID=" & prepIntegerSQL(MAGID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Expenditure Report record retrieved for MAG " & MAGID)
	SendMessage "Error: No Expenditure Report record retrieved for MAG " & MAGID
	Response.End
End If

MAGID = rs.Fields("MAGID")
FiscalYear = rs.Fields("FiscalYear")
GranteeName = rs.Fields("GranteeName")
ProgramName = rs.Fields("ProgramName")
StatePayeeIDNo = rs.Fields("StatePayeeIDNo")
GrantOption = rs.Fields("GrantOption")
OptionID = rs.Fields("OptionID")
GrantNumber = rs.Fields("GrantNumber")
GrantAwardAmount = rs.Fields("GrantAwardAmount")
CashMatch = rs.Fields("CashMatch")
BudgetedAmount = rs.Fields("BudgetedAmount")
CashExpenditure = rs.Fields("CashExpenditure")
ExcludedAmount = rs.Fields("ExcludedAmount")
AmountPaid = rs.Fields("AmountPaid")
VendorOrganizationalUnit = rs.Fields("VendorOrganizationalUnit")
VendorAddress1 = rs.Fields("VendorAddress1")
VendorAddress2 = rs.Fields("VendorAddress2")
VendorCity = rs.Fields("VendorCity")
VendorState = rs.Fields("VendorState")
VendorZip = rs.Fields("VendorZip")
ReimbursableExpenditure = rs.Fields("ReimbursableExpenditure")
ReimbursementRate = rs.Fields("ReimbursementRate")
Reimbursement = rs.Fields("Reimbursement")
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
PermitEdit=False

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Voucher,MAGID=<%=MAGID %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<table style="margin: auto; font-size: 15px; ">
<thead>
	<tr><th colspan="5">Texas Motor Vehicle Crime Prevention Authority</th></tr>
	<tr><th colspan="5">Auxiliary Grant Payment Voucher</th></tr>
</thead>
<tbody>
	<tr><td colspan="5">&nbsp;</td></tr>
	<tr>
		<th>Vendor No.:</th>
		<td><%=StatePayeeIDNo %></td>
		<td></td>
		<th>Voucher #:</th>
		<td><%=GrantNumber %></td>
	</tr>

	<tr><td colspan="5">&nbsp;</td></tr>

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

	<tr>
		<td><%=GrantNumber %></td>
		<td>1</td>
		<td><%=(FiscalYear MOD 1000) %></td>
		<td></td>
		<td style="text-align: right;"><%=prepCurrencyWeb(AmountPaid) %></td>
	</tr>

	<tr><td colspan="5">&nbsp;</td></tr>

	<tr>
		<th>Approved by:</th>
		<th>Date</th>
		<td></td>
		<th>Position</th>
	</tr>
	<tr>
		<td><%=Reviewer %></td>
		<td><%=ReviewDate %></td>
		<td></td>
		<td><%=ReviewerTitle %></td>
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

<br />
<hr style="width: 680px; margin: auto; "/>
<br />
<table style="margin: auto; font-size: 15px; ">
<tr><th>FY <%=FiscalYear %> Expenditure Report</th></tr>
<tr><th><%=GranteeName %>, <%=ProgramName %></th></tr>
<tr><th><%=GrantNumber %></th></tr>
<tr><td style="text-align: center">Submitted by <%=SubmitName %> on <%=SubmitDate %></td></tr>
</table>

<br />

<div class="singleborder">
<table style="margin: auto; font-size: 15px; ">
<thead>
	<tr style="vertical-align: bottom; ">
		<th style="width: 175px; ">Budget Category</th>
		<th style="width: 200px; ">Description</th>
		<th>Budgeted<br />Amount</th>
		<th>Total Cash<br />Expenditure</th>
		<th>Excluded<br />Amount</th>
		<th>Reimbursable<br />Expenditure</th>
	</tr>
</thead>
<tbody>
<%	If OptionID = 1 Then %>
	<tr style="vertical-align: top;">
		<th style="vertical-align: top; ">Equipment</th>
		<td>Purchase of one (1) Mobile or Stationary Automatic License Plate Reader (ALPR)</td>
		<td><%=CurrencyField("BudgetedAmount", BudgetedAmount, 10, 12, False, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("CashExpenditure", CashExpenditure, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("ExcludedAmount", ExcludedAmount, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("ReimbursableExpenditure", ReimbursableExpenditure, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
<%	ElseIf OptionID = 2 Then %>
	<tr style="vertical-align: top;">
		<th style="vertical-align: top; ">Professional and Contract Services</th>
		<td>Lease of a multiple unit stationary Automatic License Plate Reader (ALPR) system</td>
		<td style="text-align: right; "><%=prepCurrencyWeb(BudgetedAmount) %></td>
		<td style="text-align: right; "><%=prepCurrencyWeb(CashExpenditure) %></td>
		<td style="text-align: right; "><%=prepCurrencyWeb(ExcludedAmount) %></td>
		<td style="text-align: right; "><%=prepCurrencyWeb(ReimbursableExpenditure) %></td>
	</tr>
<%	End If %>
</tbody>
</table>

<br />

<table style="margin: auto; font-size: 15px; margin-top: 8px;">
<caption>Reimbursement Calculation</caption>
<tbody>

<tr>
	<td>Reimbursable Expenditure</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(ReimbursableExpenditure) %></td>
</tr>
<tr>
	<td>Reimbursement Rate</td>
	<td style="text-align: right; white-space: nowrap; "><%=ReimbursementRate%>%</td>
</tr>

<tr>
	<td>Reimbursement</td>
	<td style="text-align: right; "><%=prepCurrencyWeb(Reimbursement) %></td>
</tr>
</tbody>
</table>

</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
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