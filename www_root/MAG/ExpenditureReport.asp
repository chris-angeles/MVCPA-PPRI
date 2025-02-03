<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, MAGID, PermitEdit, ViewDocuments, _
	GranteeID, GranteeName, ORI, StatePayeeIDNo, FiscalYear, PurchaseLease, OptionID, GrantOption, _
	GrantAwardAmount, CashMatch, GrantNumber, SerialNumbers, _
	BudgetedAmount, CashExpenditure, ExcludedAmount, ReimbursableExpenditure, _
	ReimbursableTotal, ReimbursementRate, Reimbursement, _
	SupplementaryComments, Confirmed, SubmitID, SubmitTimestamp, SubmitName, SubmitEmail, _
	ReviewID, ReviewName, ReviewDate, _
	AuditApprovalID, AuditApprovalDate, AuditApprovalName, _
	DirectorApprovalID, DirectorApprovalDate, DirectorApprovalName,  _
	AdministrativeComments, AmountPaid, DatePaid, UpdateID, UpdateTimestamp, _
	CanSubmit, CanApprove, CanInvoice
Debug = False

If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLf)
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
	Response.Write("</pre>" & vbCrLf)
End If

If Len(Request.Form("MAGID"))>0 Then
	MAGID = CInt(Request.Form("MAGID"))
ElseIf Len(Request.QueryString("MAGID"))>0 Then
	MAGID = CInt(Request.QueryString("MAGID"))
Else
	MAGID = Session("MAGID")
End If

sql = "SELECT G.GranteeID, G.GranteeName, G.ORI, G.StatePayeeIDNo, M.MAGID, M.FiscalYear, " & vbCrLf & _
	"	ISNULL(E.PurchaseLease, M.OptionID) AS PurchaseLease, M.OptionID, " & vbCrLf & _
	"	CASE WHEN M.OptionID=1 THEN 'Purchase' WHEN M.OptionID=2 THEN 'Lease' ELSE '' END AS [GrantOption], " & vbCrLf & _
	"	A.GrantAwardAmount, " & vbCrLF & _
	"	CASE WHEN M.FiscalYear=2022 THEN 5000 ELSE A.CashMatch END AS CashMatch, A.GrantNumber, E.SerialNumbers, " & vbCrLf & _
	"	ISNULL(E.CashExpenditure,0.0) AS CashExpenditure, ISNULL(E.ExcludedAmount, 0.0) AS ExcludedAmount, " & vbCrLf & _
	"	CASE WHEN M.FiscalYear=2022 THEN 25000 " & vbCrLF & _
	"		ELSE ISNULL(A.GrantAwardAmount,0.0) + ISNULL(A.CashMatch,0.0) END AS BudgetedAmount, " & vbCrLf & _
	"	ISNULL(E.ReimbursableExpenditure,0.0) AS ReimbursableExpenditure, E.ReimbursableTotal, " & vbCrLf & _
	"	CASE WHEN M.FiscalYear=2022 Then 80.0 " & vbCrLf & _
	"		WHEN GrantAwardAmount+CashMatch>0 THEN 100.0*A.GrantAwardAmount/(GrantAwardAmount+CashMatch) " & vbCrLf & _
	"		ELSE 0.0 END AS ReimbursementRate, " & vbCrLf & _
	"	E.Reimbursement, E.SupplementaryComments, E.Confirmed, " & vbCrLf & _
	"	ISNULL(E.SubmitID,0) AS SubmitID, E.SubmitTimestamp, U1.Name AS SubmitName, U1.Email AS SubmitEMail, " & vbCrLf & _
	"	CAST(CASE WHEN " & UserSystemID & " IN (G.FinancialOfficerID, G.FinancialAdministrativeContactID) THEN 1 ELSE 0 END AS BIT) AS CanSubmit, " & vbCrLf & _
	"	E.ReviewID, E.ReviewDate, U5.Name As ReviewName, " & vbCrLf & _
	"	E.AuditApprovalID, E.AuditApprovalDate, U2.Name AS AuditApprovalName, " & vbCrLf & _
	"	E.DirectorApprovalID, E.DirectorApprovalDate, U4.Name AS DirectorApprovalName, " & vbCrLf & _
	"	CASE WHEN E.DatePaid IS NULL AND E.Reimbursement<0 THEN 0.0 " & vbCrLf & _
	"		WHEN E.SubmitID IS NOT NULL AND E.DatePaid IS NULL THEN E.Reimbursement " & vbCrLf & _
	"		WHEN DatePaid IS NOT NULL THEN E.AmountPaid " & vbCrLf & _
	"		WHEN E.SubmitID IS NOT NULL AND E.AmountPaid IS NULL THEN E.Reimbursement " & vbCrLf & _
	"		ELSE E.AmountPaid END AS AmountPaid, " & vbCrLf & _
	"	E.AdministrativeComments, E.DatePaid, " & vbCrLf & _
	"	E.UpdateID, E.UpdateTimestamp " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"JOIN MAG.Main AS M ON M.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN MAG.Admin AS A ON A.MAGID=M.MAGID " & vbCrLf & _
	"LEFT JOIN MAG.ExpenditureReport AS E on E.MAGID=M.MAGID " & vbCrLf & _
	"LEFT JOIN [System].Users AS U1 ON U1.SystemID=E.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS U2 ON U2.SystemID=E.AuditApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS U3 ON U3.SystemID=E.UpdateID " & vbCrLf & _
	"LEFT JOIN [System].Users AS U4 ON U4.SystemID=E.DirectorApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS U5 ON U5.SystemID=E.ReviewID " & vbCrLf & _
	"WHERE M.MAGID=" & MAGID
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	ORI = rs.Fields("ORI")
	StatePayeeIDNo = rs.Fields("StatePayeeIDNo")
	MAGID = rs.Fields("MAGID")
	FiscalYear = rs.Fields("FiscalYear")
	PurchaseLease = rs.Fields("PurchaseLease")
	OptionID = rs.Fields("OptionID")
	GrantOption = rs.Fields("GrantOption")
	GrantAwardAmount = rs.Fields("GrantAwardAmount")
	CashMatch = rs.Fields("CashMatch")
	GrantNumber = rs.Fields("GrantNumber")
	SerialNumbers = rs.Fields("SerialNumbers")
	BudgetedAmount = rs.Fields("BudgetedAmount")
	CashExpenditure = rs.Fields("CashExpenditure")
	ExcludedAmount = rs.Fields("ExcludedAmount")
	ReimbursableExpenditure = rs.Fields("ReimbursableExpenditure")
	ReimbursementRate = rs.Fields("ReimbursementRate")
	ReimbursableTotal = rs.Fields("ReimbursableTotal")
	Reimbursement = rs.Fields("Reimbursement")
	SupplementaryComments = rs.Fields("SupplementaryComments")
	Confirmed = rs.Fields("Confirmed")
	SubmitID = rs.Fields("SubmitID")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	CanSubmit = rs.Fields("CanSubmit")
	SubmitName = rs.Fields("SubmitName")
	SubmitEMail = rs.Fields("SubmitEMail")
	ReviewID = rs.Fields("ReviewID")
	ReviewDate = rs.Fields("ReviewDate")
	ReviewName = rs.Fields("ReviewName")
	AuditApprovalID = rs.Fields("AuditApprovalID")
	AuditApprovalDate = rs.Fields("AuditApprovalDate")
	AuditApprovalName = rs.Fields("AuditApprovalName")
	DirectorApprovalID = rs.Fields("DirectorApprovalID")
	DirectorApprovalDate = rs.Fields("DirectorApprovalDate")
	DirectorApprovalName = rs.Fields("DirectorApprovalName")
	AdministrativeComments = rs.Fields("AdministrativeComments")
	AmountPaid = rs.Fields("AmountPaid")
	DatePaid = rs.Fields("DatePaid")
	UpdateID = rs.Fields("UpdateID")
	UpdateTimestamp = rs.Fields("UpdateTimestamp")
Else
	Response.Write("No such MAG Grant Found.")
	Response.End
End If

If SubmitID = 0 Then
	PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
ElseIf SubmitID > 0 Then
	PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, True)
Else
	PermitEdit = False
End If
ViewDocuments = CheckPermissions(UserSystemID, GranteeID, True)

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA MAG Expenditure Report for <%=GranteeName %></title>
<meta http-equiv="cache-control" content="no-cache" />
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm(action) {
		if (validateForm() == false) {
			return false;
		}
		if (action == "submit") {
<%			if PurchaseLease = 1 Then %>
			if (document.ER["SerialNumbers"].value.length == 0) {
				alert("You must enter serial numbers for the purchase before submitting.")
				return false;
			}
<%		End If %>
			if (document.ER["Confirmed"].checked == false) {
				alert("You must read certification and check box to submit expenditure report.");
				return false;
			}
		}
		document.ER.Action.value = action;
		document.ER.submit();
	}

	var reimbursementrate=<%=ReimbursementRate%>, grantawardamount=<%=grantawardamount%>, cashmatch=<%=CashMatch%>, budgetedamount=<%=BudgetedAmount%>;
	function updateTotals()
	{
		if (getNumericValue(document.ER.CashExpenditure.value) - budgetedamount>0.0) {
			document.ER.ExcludedAmount.value = currency(getNumericValue(document.ER.CashExpenditure.value) - budgetedamount);
		}
		else {
		document.ER.ExcludedAmount.value = "$0.00";
		}
		document.ER.ReimbursableExpenditure.value = currency(getNumericValue(document.ER.CashExpenditure.value) - getNumericValue(document.ER.ExcludedAmount.value));
		document.ER.ReimbursableTotal.value = document.ER.ReimbursableExpenditure.value;
		if (reimbursementrate * getNumericValue(document.ER.ReimbursableTotal.value) / 100 > grantawardamount) {
			document.ER.Reimbursement.value = currency(grantawardamount);
		}
		else {
			document.ER.Reimbursement.value = currency(reimbursementrate * getNumericValue(document.ER.ReimbursableTotal.value) / 100);
		}
	}

	function validateForm() {
		return true;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body onload="updateTotals();">

<div class="sectiontitle">Motor Vehicle Crime Prevention Authority Auxiliary Grant</div>
<div class="sectiontitle">FY <%=FiscalYear %> Quarterly Expenditure Report</div>
<%
If SubmitID > 0 Then
	Response.Write("<div style=""text-align: center;"">Submitted by " & SubmitName & ", " & SubmitTimeStamp & "</div>" & vbCrLf)
End If
If IsNull(DatePaid) = False Then ' Display paid date if any.
	Response.Write("<div style=""text-align: center;"">Date Paid: " & formatDateTime(DatePaid,vbShortDate) & "</div>" & vbCrLf)
End If
If IsNull(UpdateTimestamp) = False And UpdateTimestamp<CDate("2023-03-13 11:53") Then
	Response.Write("<div style=""text-align: center; color: red;"">Display shows new formula values but these newer values have not been saved.</div>" & vbCrLf)
End If
%>
<br />
<form name="ER" id="ER" method="post" action="ExpenditureReportSubmit.asp">
<%=HiddenField("MAGID",MAGID) %><%=HiddenField("Action","save") %><%=HiddenField("GrantAwardAmount",GrantAwardAmount) %>
<table>
	<tr><td>Grantee Name</td><td><%=GranteeName %></td></tr>
	<tr><td>Grant Number</td><td><%=GrantNumber %></td></tr>
	<tr><td>Fiscal Year</td><td><%=FiscalYear %></td></tr>
	<tr><td>Grant Award Amount</td><td><%=prepCurrencyWeb(GrantAwardAmount) %></td></tr>
	<tr><td>Cash Match</td><td><%=prepCurrencyWeb(CashMatch) %></td></tr>
	<tr><td>Grant Number</td><td><%=GrantNumber %></td></tr>
	<tr><td>Grant Option from Application</td><td><%=GrantOption %></td>
	<tr>
			<td>Grant Option for reimbursement</td>
		<td><select name="PurchaseLease">
			<%=SelectOption(1, "Purchase", PurchaseLease) %>
			<%=SelectOption(2, "Lease", PurchaseLease) %>
		    </select></td>
	</tr>
<%	If SubmitID = 0 Then %>
	<tr><td colspan="2" style="text-align: center; color: red; font-weight: bold">If changing purchase /
		lease option from original application, please select <br />new option in line above 
		and then save record to update fields before continuing.</td></tr>
<%	End If %>
<%	If PurchaseLease=1 Then %>
	<tr style="vertical-align: top; "><td>Serial Numbers<br />
		<span style="font-size: small; ">List serial numbers of purchased equipment.</span></td>
		<td><%=TextArea("SerialNumbers", SerialNumbers, 5, 40, 990, PermitEdit, "") %></td></tr>
<%	End If %>
	<tr><td></td><td></td></tr>
</table>

<div style="margin: auto; font-weight: bold; text-align: center; ">Expenditures</div>
<br />

<div class="singleborder">
<table>
<thead>
	<tr style="vertical-align: bottom; ">
		<th>Budget Category</th>
		<th>Description</th>
		<th>Budgeted<br />Amount</th>
		<th>Total Cash<br />Expenditure</th>
		<th>Excluded Amount</th>
		<th>Reimbursable<br />Expenditure</th>
	</tr>
</thead>
<tbody>
<%	If PurchaseLease = 1 Then %>
	<tr style="vertical-align: top;">
		<th style="vertical-align: top; ">Equipment</th>
		<td>Purchase of one (1) Mobile or Stationary Automatic License Plate Reader (ALPR)</td>
		<td><%=CurrencyField("BudgetedAmount", BudgetedAmount, 10, 12, False, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("CashExpenditure", CashExpenditure, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("ExcludedAmount", ExcludedAmount, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("ReimbursableExpenditure", ReimbursableExpenditure, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
<%	ElseIf PurchaseLease = 2 Then %>
	<tr style="vertical-align: top;">
		<th style="vertical-align: top; ">Professional and Contract Services</th>
		<td>Lease of a multiple unit stationary Automatic License Plate Reader (ALPR) system</td>
		<td><%=CurrencyField("BudgetedAmount", BudgetedAmount, 10, 12, False, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("CashExpenditure", CashExpenditure, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("ExcludedAmount", ExcludedAmount, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
		<td><%=CurrencyField("ReimbursableExpenditure", ReimbursableExpenditure, 10, 12, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
<%	End If %>
</tbody>
</table>
<br />

<div style="margin: auto; font-weight: bold; text-align: center; ">Reimbursement Calculation</div>
<br />
<table style="margin: auto; text-align: center; ">
	<tr>
		<td style="vertical-align: middle; text-align: left; ">Reimbursable Total</td>
		<td><%=CurrencyField("ReimbursableTotal", ReimbursableTotal, 10, 12, False, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td style="vertical-align: middle; text-align: left ">Reimbursement Rate</td>
		<td><%=NumberField("ReimbursementRate", prepNumberWeb(ReimbursementRate,6), 8, 8, False, "") %>%</td>
	</tr>
	<tr>
		<td style="vertical-align: middle; text-align: left ">Reimbursement</td>
		<td><%=CurrencyField("Reimbursement", Reimbursement, 10, 12, False, "checkCurrency(this); updateTotals()") %></td>
	</tr>
</table>
</div>
<br />
<%
If ViewDocuments = True Then
	Dim Folder, file, files, DocumentFolder, fso, counter
	counter=0
	DocumentFolder = Application("DocumentRoot") & "\MAG\" & MAGID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemOBject")
	Response.Write("<table style=""margin: auto; "">" & vbCrLf)
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		If PErmitEdit = True Then
			Response.Write("<tr style=""vertical-align: top; ""><td>Current Documents in folder: <a href=""../Upload/Upload.asp?fid=15&MAGID=" & MAGID & """ class=""plainlink"" target=""_blank"">Upload</a></td>" & vbCrLf)
		End If
		If files.count>0 Then 
			Response.Write("<tr><td>")
			For Each file in files
				If Left(file.Name,3)="ER " Then
					Response.Write("<a href=""../Documents/MAG/" & MAGID & "/" & file.Name & _
						""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
					counter = counter + 1
				End If
			Next
			Response.Write("</td></tr>" & vbCrLf)
		End If
	End If
	If counter = 0 Then
		Response.Write("<tr style=""vertical-align: top; ""><td style=""text-align: center; "">No Documents in folder</td></tr>" & vbCrLf)
	End If
	Response.Write("</table>" & vbCrLf)
	Response.Write("<hr />" & vbCrLf)
End If
%>
<br />
<div style="margin: auto; ">
Supplementary Comments. Provide additional informaton that would be helpful in evaluating this expenditure report.<br />
<%=TextArea2("SupplementaryComments", SupplementaryComments, 8, 960, 8000, PermitEdit, "") %>
<p><%
Response.Write(CheckBoxField2("Confirmed", Confirmed, PermitEdit) & vbCrLf)

If FiscalYear >= 2022 Then%>
By signing this report, I certify to the best of my knowledge and belief that the report is true, 
complete, and accurate, and the expenditures, disbursements and cash receipts are for the purposes 
and objectives set forth in the terms and conditions of the state award. I am aware that any false, 
fictitious, or fraudulent information, or the omission of any material fact, may subject me to 
criminal, civil or administrative penalties for fraud, false statements, false claims or otherwise.
<%
Else
%>
I acknowledge that I have reviewed and confirmed the accuracy of the information in this report, 
and I attest that this report is  correct and complete and that the costs incurred as stated 
herein are for allowable purposes as set forth in the Statement of Grant Award and, (1) pursuant 
to 43 TAC &sect; 57.9, I hereby further certify that MVCPA funds have not been used to replace 
state or local funds and (2) any false, fictitious, or fraudulent information provided herein 
may subject me to criminal, civil, or administrative penalties or sanctions. 
<%
End If

If SubmitID > 0 Then
	Response.Write(" <i>Submitted by " & SubmitName & ", " & SubmitTimeStamp & "</i>")
End If
%>
</p>
</div>
<br />
<%

' settings for admin section.
If SubmitID = 0 Then
	CanApprove = False
	CanInvoice = False
Else
	CanApprove = True
	CanInvoice = True ' Keep true if all conditions are passed.
End If

If MVCPARights = True Or MVCPAViewer = True Then
	Response.Write("<div style=""margin: auto; font-weight: bold; text-align: center; "">For Administrative Use Only</div>" & vbCrLf)
	Response.Write(HiddenField("ReviewID", ReviewID))
	Response.Write(HiddenField("AuditApprovalID", AuditApprovalID))
	Response.Write(HiddenField("DirectorApprovalID", DirectorApprovalID))
	Response.Write("<table style=""margin: auto; "">" & vbCrLf)
	Response.Write("<tr><td colspan=""2"">Administrative Comments:<br />" & vbCrLf)
	Response.Write(TextArea2("AdministrativeComments", AdministrativeComments, 6, 920, 8000, MVCPARights, "") & "</td></tr>" & vbCrLf)
	If SubmitID > 0 Then
		Response.Write("<tr><td>MVCPA Grant Coordinator Review Date:</td>" & vbCrLf)
		If ReviewID > 0 And IsNull(DirectorApprovalDate) = False And CanApprove = True Then
			Response.Write("<td>" & DateField("ReviewDate", ReviewDate, False))
		ElseIf IsNull(DirectorApprovalDate) = False And CanApprove = True Then
			Response.Write("<td>" & DateField("ReviewDate", ReviewDate, True))
		Else
			Response.Write("<td>" & DateField("ReviewDate", ReviewDate, MVCPARights))
		End If
		If IsNull(ReviewName) = False Then
			Response.Write(" by " & ReviewName)
		End If
		Response.Write("<tr><td>MVCPA Audit Approval Date:</td>" & vbCrLf)
		If IsNull(DirectorApprovalDate) = False And CanApprove = True Then ' no edit once director approval
			Response.Write("<td>" & DateField("AuditApprovalDate", AuditApprovalDate, False))
		Else
			Response.Write("<td>" & DateField("AuditApprovalDate", AuditApprovalDate, MVCPAAuditor))
		End If
		If IsNull(AuditApprovalName) = False Then
			Response.Write(" by " & AuditApprovalName)
		End If
		Response.Write("<tr><td>MVCPA Director Approval Date:</td>" & vbCrLf)
		If IsNull(DirectorApprovalDate) = False And CanApprove = True Then ' no edit once director approval
			Response.Write("<td>" & DateField("DirectorApprovalDate", DirectorApprovalDate, False))
		Else
			Response.Write("<td>" & DateField("DirectorApprovalDate", DirectorApprovalDate, MVCPAAdministrator))
		End If
		If IsNull(DirectorApprovalName) = False Then
			Response.Write(" by " & DirectorApprovalName)
		End If
		Response.Write("</td>")
		If IsNull(DatePaid) = False Then
			Response.Write("<tr><td>Amount Paid</td><td>" & CurrencyField("AmountPaid", AmountPaid, 10, 12, False, "checkCurrency(this); updateTotals();") & "</td>" & vbCrLf)
		Else
			Response.Write("<tr><td>Amount To Pay</td><td>" & CurrencyField("AmountPaid", AmountPaid, 10, 12, False, "checkCurrency(this); updateTotals();") & "</td>" & vbCrLf)
		End If
		If IsNull(DirectorApprovalDate) = True and IsNull(DatePaid)=True Then
			' Do not show date paid.
			Response.Write(HiddenField("DatePaid", DatePaid))
		ElseIf IsNull(DatePaid) = False Then ' no edit once there is a value except for director and auditor.
			If MVCPAAdministrator = True Or MVCPAAuditor=True Or Developer=True Then
				Response.Write("<tr><td>Date Paid</td><td>" & DateField("DatePaid", DatePaid, True))
			Else ' no edit once there is a value.
				Response.Write("<tr><td>Date Paid</td><td>" & DateField("DatePaid", DatePaid, False))
			End If
		Else
			Response.Write("<tr><td>Date Paid</td><td>" & DateField("DatePaid", DatePaid, MVCPARights))
			If CanInvoice = True Then
				Response.Write(" <a href=""Voucher.asp?MAGID=" & MAGID & """ target=""_blank"">invoice</a>" & vbCrLf)
			End If
			Response.Write("</td></tr>" & vbCrLf)
		End If
		If IsNull(DirectorApprovalDate) = True Then ' no edit once director approval
			Response.Write("<tr><td colspan=""2"">" &  CheckBoxField("Unsubmit", False) & " Unsubmit Expenditure Report (Clears submission and approval.)</td></tr>" & vbCrLf)
		Else
			Response.Write("<tr><td colspan=""2"">This Expenditure Report may not be unsubmitted because it has already received director approval.</td></tr>" & vbCrLf)
		End If
		Response.Write("</table>" & vbCrLf)
	Else
		Response.Write("</table>" & vbCrLf)
		Response.Write(HiddenField("ReviewDate", ReviewDate))
		Response.Write(HiddenField("AuditApprovalDate", AuditApprovalDate))
		Response.Write(HiddenField("DirectorApprovalDate", DirectorApprovalDate))
		Response.Write(HiddenField("AmountPaid", AmountPaid))
		Response.Write(HiddenField("DatePaid", DatePaid))
	End If
End If
%>
</form>

<div style="text-align: center">
<%	If (PermitEdit = True And SubmitID = 0) Or MVCPARights = True Then %>
<input type="button" name="Save" value="Save" onclick="return submitForm('save');" />
<%	
	End If	
	If SubmitID = 0 And CanSubmit=True  Then 
%>
<input type="button" name="Submit" value="Submit" onclick="return submitForm('submit');" />
<%	ElseIf SubmitID = 0 And CanSubmit=True Then 
%>
<input type="button" name="Submit" value="Submit" onclick="alert('Too early to submit'); return false;" />
<%	 
	End If
	If (PermitEdit = True And SubmitID = 0) Or MVCPARights = True Then
%>
	<input type="reset" name="Reset" value="Reset" />
<%	End If %>

<input type="button" value="Close" onclick="window.close();" />
</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
