<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, ViewDocuments, UploadDirectory, _
	GrantID, AdjustmentID, FiscalYear, _
	GranteeID, ProgramName, GrantNumber, GranteeName, InLieuOfNICBBudget, _
	CurrentReimbursementRate, ReimbursementRate, CashMatchRate, OvertimeRate, ProgramChange, BudgetChange, _
	BudgetModificationExplanation, ProgramModificationExplanation, _
	NewProgramIncome, ProgramIncomeToBeAddedToBudget, Confirmed, _
	SubmitID, SubmitTimeStamp, SubmitName, SubmitterEMail, _
	MailTo, AuthorizedofficialEMail, FinancialOfficerEmail, _
	ProgramManagerEmail, ProgramDirectorEMail, DenialDate, AdministrativelyClosedDate, _
	ExternalComments, FirstApprovalDate, FirstApprovalID, FirstApprover, _
	SecondApprovalID, SecondApprovalDate, SecondApprover, ChangesApplied, CanSubmit
debug = False
'MVCPAGrantCoordinator=True
'MVCPAAdministrator = True
'MVCPARights = False

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
	Response.Write("Application(""CurrentFiscalYear"")=" & Application("CurrentFiscalYear") & vbCrLf)
	Response.Write("</pre>" & vbCrLF)
End If

If Len(Request.Form("GrantID"))>0 Then
	GrantID = Request.Form("GrantID")
ElseIf Len(Request.QueryString("GrantID"))>0 Then
	GrantID = Request.QueryString("GrantID")
Else
	GrantID = Session("GrantID")
End If

If Len(Request.Form("AdjustmentID"))>0 Then
	AdjustmentID = Request.Form("AdjustmentID")
ElseIf Len(Request.QueryString("AdjustmentID"))>0 Then
	AdjustmentID = Request.QueryString("AdjustmentID")
Else
	AdjustmentID = -1
End If

If IsNumeric(GrantID) Then
	GrantID = CInt(GrantID)
Else
	Response.Write("Error: Non-numeric Fiscal Year Specified")
	SendMessage "Error: Non-numeric Fiscal Year Specified"
	Response.End
End If
If IsNumeric(AdjustmentID) Then
	AdjustmentID = CInt(AdjustmentID)
Else
	Response.Write("Error: Non-numeric AdjustmentID Specified")
	SendMessage "Error: Non-numeric AdjustmentID Specified"
	Response.End
End If

sql = "SELECT A.GrantID, B.GranteeID, A.FiscalYear, ISNULL(C.AdjustmentID,0) AS AdjustmentID, " & vbCrLF & _
	"	A.ProgramName, A.GrantNumber, A.ReimbursementRate AS CurrentReimbursementRate, B.GranteeName, " & vbCrLf & _
	"	CASE WHEN C.InLieuOfNICBBudget IS NULL OR C.SubmitTimeStamp IS NULL THEN ISNULL(A.InLieuOfNICBBudget,0.00) ELSE ISNULL(C.InLieuOfNICBBudget,0.00) END AS InLieuOfNICBBudget, " & vbCrLf & _
	"	C.ProgramChange, C.BudgetChange, C.BudgetModificationExplanation, C.ProgramModificationExplanation, " & vbCrLf & _
	"	C.NewProgramIncome, C.ProgramIncomeToBeAddedToBudget, " & vbCrLf & _
	"	CAST(CASE WHEN SubmitID IS NOT NULL THEN ISNULL(C.Confirmed,0) ELSE 0 END AS BIT) AS Confirmed, " & vbCrLf & _
	"	C.SubmitID, C.SubmitTimeStamp, D.Name AS SubmitName, " & vbCrLf & _
	"	CASE WHEN ISNULL(D.AccountDisabled,0)=0 THEN ISNULL(D.Email,'') ELSE '' END AS SubmitterEMail, " & vbCrLf & _
	"	CASE WHEN ISNULL(I.AccountDisabled,0)=0 THEN ISNULL(I.EMail,'') ELSE '' END AS AuthorizedOfficialEmail, " & vbCrLf & _
	"	CASE WHEN ISNULL(J.AccountDisabled,0)=0 THEN ISNULL(J.EMail,'') ELSE '' END AS FinancialOfficerEmail, " & vbCrLf & _
	"	CASE WHEN ISNULL(G.AccountDisabled,0)=0 THEN ISNULL(G.EMail,'') ELSE '' END AS ProgramManagerEmail, " & vbCrLf & _
	"	CASE WHEN ISNULL(H.AccountDisabled,0)=0 THEN ISNULL(H.Email,'') ELSE '' END AS ProgramDirectorEmail, " & vbCrLf & _
	"	C.DenialDate, C.AdministrativelyClosedDate, C.ExternalComments, " & vbCrLf & _
	"	CONVERT(VARCHAR,C.FirstApprovalDate,101) AS FirstApprovalDate, C.FirstApprovalID, E.Name AS FirstApprover, " & vbCrLf & _
	"	CONVERT(VARCHAR,C.SecondApprovalDate,101) AS SecondApprovalDate, C.SecondApprovalID, F.Name AS SecondApprover, " & vbCrLf & _
	"	CAST(CASE WHEN " & UserSystemID & " IN (B.AuthorizedOfficialID, B.ProgramDirectorID, B.FinancialOfficerId, B.ProgramManagerID) THEN 1 ELSE 0 END AS BIT) AS CanSubmit ," & vbCrLf & _
	"	IsNull(C.ChangesApplied,0) AS ChangesApplied " & vbCrLf & _
	"FROM [Grants].Main AS A " & vbCrLF & _
	"JOIN Grantees AS B ON B.GranteeID = A.GranteeID " & vbCrLf
If AdjustmentID > -1 Then
	sql = sql &	"LEFT JOIN [Grants].Adjustments AS C ON C.GrantID=A.GrantID AND C.AdjustmentID=" & prepIntegerSQL(AdjustmentID) & " " & vbCrLf
Else
	sql = sql & "LEFT JOIN [Grants].Adjustments AS C ON C.AdjustmentID=(SELECT MAX(AdjustmentID) FROM [Grants].Adjustments WHERE GrantID=A.GrantID) " & vbCrLf
End If
sql = sql &	"LEFT JOIN [System].Users AS D ON D.SystemID=C.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS E ON E.SystemID=C.FirstApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS F ON F.SystemID=C.SecondApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS G ON G.SystemID=B.ProgramManagerID " & vbCrLf & _
	"LEFT JOIN [System].Users AS H ON H.SystemID=B.ProgramDirectorID " & vbCrLf & _
	"LEFT JOIN [System].Users AS I ON I.SystemID=B.AuthorizedOfficialID " & vbCrLf & _
	"LEFT JOIN [System].Users AS J ON J.SystemID=B.FinancialOfficerID " & vbCrLf & _
	"WHERE A.GrantID=" & prepIntegerSQL(GrantID)

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Grant record retrieved")
	SendMessage "Error: No Grant record retrieved"
	Response.End
Else
	AdjustmentID = rs.Fields("AdjustmentID")
	GrantID = rs.Fields("GrantID")
	FiscalYear = rs.Fields("FiscalYear")
	ProgramName = rs.Fields("ProgramName")
	GranteeName = rs.Fields("GranteeName")
	GrantNumber = rs.Fields("GrantNumber")
	GranteeID = rs.Fields("GranteeID")
	CurrentReimbursementRate = FormatNumber(rs.Fields("CurrentReimbursementRate"),14,True,True,True) & "%"
	InLieuOfNICBBudget = rs.Fields("InLieuOfNICBBudget")
	ProgramChange = rs.Fields("ProgramChange")
	BudgetChange = rs.Fields("BudgetChange")
	BudgetModificationExplanation = rs.Fields("BudgetModificationExplanation")
	ProgramModificationExplanation = rs.Fields("ProgramModificationExplanation")
	NewProgramIncome = rs.Fields("NewProgramIncome")
	ProgramIncomeToBeAddedToBudget = rs.Fields("ProgramIncomeToBeAddedToBudget")
	Confirmed = rs.Fields("Confirmed")
	SubmitID = rs.Fields("SubmitID")
	SubmitTimeStamp = rs.Fields("SubmitTimeStamp")
	SubmitName = rs.Fields("SubmitName")
	SubmitterEMail = rs.Fields("SubmitterEMail")
	AuthorizedofficialEMail = rs.Fields("AuthorizedofficialEMail")
	FinancialOfficerEmail = rs.Fields("FinancialOfficerEmail")
	ProgramManagerEmail = rs.Fields("ProgramManagerEmail")
	ProgramDirectorEMail = rs.Fields("ProgramDirectorEMail")
	DenialDate = rs.Fields("DenialDate")
	AdministrativelyClosedDate = rs.Fields("AdministrativelyClosedDate")
	ExternalComments = rs.Fields("ExternalComments")
	FirstApprovalDate = rs.Fields("FirstApprovalDate")
	FirstApprovalID = rs.Fields("FirstApprovalID")
	FirstApprover = rs.Fields("FirstApprover")
	SecondApprovalDate = rs.Fields("SecondApprovalDate")
	SecondApprovalID = rs.Fields("SecondApprovalID")
	SecondApprover = rs.Fields("SecondApprover")
	ChangesApplied = rs.Fields("ChangesApplied")
	CanSubmit = rs.Fields("CanSubmit")
End If

If Len(Trim(SubmitterEmail))>0 Then
	MailTo = Trim(SubmitterEmail)
Else
	MailTo = ""
End If
'Response.Write("<td><pre>MailTo=""" & MailTo & """</pre>")
If Len(AuthorizedOfficialEmail)>0 Then
	If InStr(MailTo, AuthorizedOfficialEmail) = 0 Then
		If Len(Trim(MailTo))=0 Then
			MailTo = Trim(AuthorizedOfficialEmail)
		Else
			MailTo = MailTo & ";" & AuthorizedOfficialEmail
		End If
	End If
End If
'Response.Write("<pre>MailTo=""" & MailTo & """</pre>")
If Len(FinancialOfficerEmail)>0 Then
	If InStr(MailTo, FinancialOfficerEmail) = 0 Then
		If Len(Trim(MailTo))=0 Then
			MailTo = Trim(FinancialOfficerEmail)
		Else
			MailTo = MailTo & ";" & FinancialOfficerEmail
		End If
	End If
End If
'Response.Write("<pre>MailTo=""" & MailTo & """</pre>")
If Len(ProgramManagerEmail)>0 Then
	If InStr(MailTo, ProgramManagerEmail) = 0 Then
		If Len(Trim(MailTo))=0 Then
			MailTo = Trim(ProgramManagerEmail)
		Else
			MailTo = MailTo & ";" & ProgramManagerEmail
		End If
	End If
End If
'Response.Write("<pre>MailTo=""" & MailTo & """</pre>")
If Len(ProgramDirectorEMail)>0 Then
	If InStr(MailTo, ProgramDirectorEMail) = 0 Then
		If Len(Trim(MailTo))=0 Then
			MailTo = ProgramDirectorEMail
		Else
			MailTo = MailTo & ";" & ProgramDirectorEMail
		End If
	End If
End If
'Response.Write("<pre>MailTo=""" & MailTo & """</pre></td>")
If Debug = True Then
	Response.Write("<pre>MailTo='" & MailTo & "'</pre>" & vbCrLf)
	Response.Flush
End If

If FiscalYear < Application("CurrentFiscalYear")-1 And MVCPARights = False Then
	PermitEdit = False
	If Debug = True Then
		Response.Write("<pre>PermitEdit in first condition=" & PermitEdit & "; FiscalYear=" & FiscalYear & "</pre>" & vbCrLf)
	End If
ElseIf IsNull(SubmitID) = True Then
	PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
ElseIf ISNull(SubmitID) = False Then
	PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, True)
Else
	PermitEdit = False
End If
ViewDocuments = CheckPermissions(UserSystemID, GranteeID, True)

If Debug = True Then
	Response.Write("<pre>PermitEdit=" & PermitEdit & "</pre>" & vbCrLf)
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Grant Adjustment for <%=GranteeName %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function updateTotals()
	{
		var grandtotal = 0.0, mvcpatotal = 0.0, matchtotal = 0.0;
		var changegrandtotal = 0.0, mvcpachangetotal = 0.0, matchchangetotal = 0.0;
		for (var i = 1; i < 8; i++) {
			document.Adjustment["TotalExpenditures_" + i].value =
				currency(getNumericValue(document.Adjustment["MVCPAExpenditures_" + i].value) +
				getNumericValue(document.Adjustment["MatchExpenditures_" + i].value));
			grandtotal = grandtotal + getNumericValue(document.Adjustment["TotalExpenditures_" + i].value);
			mvcpatotal = mvcpatotal + getNumericValue(document.Adjustment["MVCPAExpenditures_" + i].value);
			matchtotal = matchtotal + getNumericValue(document.Adjustment["MatchExpenditures_" + i].value);

			document.Adjustment["TotalExpendituresChange_" + i].value =
				currency(getNumericValue(document.Adjustment["MVCPAExpendituresChange_" + i].value) +
				getNumericValue(document.Adjustment["MatchExpendituresChange_" + i].value));
			changegrandtotal = changegrandtotal + getNumericValue(document.Adjustment["TotalExpendituresChange_" + i].value);
			mvcpachangetotal = mvcpachangetotal + getNumericValue(document.Adjustment["MVCPAExpendituresChange_" + i].value);
			matchchangetotal = matchchangetotal + getNumericValue(document.Adjustment["MatchExpendituresChange_" + i].value);
			document.Adjustment["TotalExpendituresNew_" + i].value = 
				currency(getNumericValue(document.Adjustment["TotalExpenditures_" + i].value) +
				getNumericValue(document.Adjustment["TotalExpendituresChange_" + i].value));
			document.Adjustment["MVCPAExpendituresNew_" + i].value = 
				currency(getNumericValue(document.Adjustment["MVCPAExpenditures_" + i].value) +
				getNumericValue(document.Adjustment["MVCPAExpendituresChange_" + i].value));
			document.Adjustment["MatchExpendituresNew_" + i].value = 
				currency(getNumericValue(document.Adjustment["MatchExpenditures_" + i].value) +
				getNumericValue(document.Adjustment["MatchExpendituresChange_" + i].value));
		}
		document.Adjustment["TotalExpenditures_Total"].value = currency(grandtotal);
		document.Adjustment["MVCPAExpenditures_Total"].value = currency(mvcpatotal);
		document.Adjustment["MatchExpenditures_Total"].value = currency(matchtotal);

		document.Adjustment["TotalExpendituresChange_Total"].value = currency(changegrandtotal);
		document.Adjustment["MVCPAExpendituresChange_Total"].value = currency(mvcpachangetotal);
		document.Adjustment["MatchExpendituresChange_Total"].value = currency(matchchangetotal);
		document.Adjustment["TotalExpendituresNew_Total"].value = 
			currency(getNumericValue(document.Adjustment["TotalExpenditures_Total"].value) +
			getNumericValue(document.Adjustment["TotalExpendituresChange_Total"].value));
		document.Adjustment["MVCPAExpendituresNew_Total"].value = 
			currency(getNumericValue(document.Adjustment["MVCPAExpenditures_Total"].value) +
			getNumericValue(document.Adjustment["MVCPAExpendituresChange_Total"].value));
		document.Adjustment["MatchExpendituresNew_Total"].value = 
			currency(getNumericValue(document.Adjustment["MatchExpenditures_Total"].value) +
			getNumericValue(document.Adjustment["MatchExpendituresChange_Total"].value));
		if (document.Adjustment.ProgramIncomeToBeAddedToBudget.value == "")
			document.Adjustment.ProgramIncomeToBeAddedToBudget.value = "$0.00";
		if (document.Adjustment.TotalExpendituresChange_Total.value == "($0.00)")
			document.Adjustment.TotalExpendituresChange_Total.value = "$0.00";
		if (document.Adjustment.MVCPAExpendituresChange_Total.value == "($0.00)")
			document.Adjustment.MVCPAExpendituresChange_Total.value = "$0.00";
		document.Adjustment["ReimbursementRate"].value = 100.0 * getNumericValue(document.Adjustment["MVCPAExpendituresNew_Total"].value) / (getNumericValue(document.Adjustment["TotalExpendituresNew_Total"].value) - getNumericValue(document.Adjustment["InLieuOfNICBBudget"].value)) + "%";
		document.Adjustment["CashMatchRate"].value = 100.0 * getNumericValue(document.Adjustment["MatchExpendituresNew_Total"].value) / getNumericValue(document.Adjustment["MVCPAExpendituresNew_Total"].value) + "%";
		document.Adjustment["OvertimeRate"].value = 100.0 * getNumericValue(document.Adjustment["TotalExpendituresNew_3"].value) / getNumericValue(document.Adjustment["TotalExpendituresNew_1"].value) + "%";
	}

	function validateForm(action)
	{
		if (action.length>0)
			document.Adjustment.Action.value = action;
		if (document.Adjustment.ProgramChange.checked == false && document.Adjustment.BudgetChange.checked == false) {
			alert("You must select either a Program Change or a Budget Change to submit the adjustment request.")
			return false;
		}
		if (document.Adjustment.ProgramChange.checked == true && document.Adjustment.ProgramModificationExplanation.value.length == 0) {
			alert("A program change must have a program modification explanation");
			return false;
		}
		if (document.Adjustment.ProgramChange.checked == false && document.Adjustment.ProgramModificationExplanation.value.length > 0) {
			alert("The program change box should be checked if there is a program modification explanation");
			return false;
		}
		if (document.Adjustment.BudgetChange.checked == true && document.Adjustment.BudgetModificationExplanation.value.length == 0) {
			alert("A budget change must have a budget modification explanation");
			return false;
		}
		if (document.Adjustment.BudgetChange.checked == false && document.Adjustment.BudgetModificationExplanation.value.length > 0) {
			alert("The budget change box should be checked if there is a budget modification explanation");
			return false;
		}
		for (var i = 1; i < 8; i++) {
			if (checkCurrency(document.Adjustment["MVCPAExpendituresChange_"+i]) == false)
				return false;
			if (checkCurrency(document.Adjustment["MatchExpendituresChange_"+i]) == false)
				return false;
		}
<%	If AdjustmentID > 0 And SubmitID > 0 Then%>
		if (document.Adjustment.FirstApprovalDate.value.length > 0 && document.Adjustment.FirstApprovalID.value.length == 0) {
			document.Adjustment.FirstApprovalID.value = <%=usersystemid%>;
		}
		if (document.Adjustment.SecondApprovalDate.value.length > 0 && document.Adjustment.SecondApprovalID.value.length == 0) {
			document.Adjustment.SecondApprovalID.value = <%=usersystemid%>;
		}
<%	End If %>
		if (action == 'Submit') {
			var sumchanges = 0.0;
			for (var i = 1; i < 8; i++) {
				sumchanges = sumchanges + Math.abs(getNumericValue(document.Adjustment["TotalExpendituresChange_" + i].value))
			}
			if (document.Adjustment.BudgetChange.checked == true && sumchanges == 0.0 && getNumericValue(document.Adjustment.ProgramIncomeToBeAddedToBudget.value)==0) {
				alert("A budget change must have a budget modification");
				return false;
			}
			if (document.Adjustment.BudgetChange.checked == false && sumchanges > 0.0) {
				alert("The budget change box must checked if there is a budget modification");
				return false;
			}
<%	If (UserSystemID = 156 Or UserSystemID = 402 Or USerSystemID = 1) And FiscalYear = 2023 And DATE() < CDATE("2023/06/01") Then %>
<%	Else %>
			if (getNumericValue(document.Adjustment.MVCPAExpendituresChange_Total.value) > 0) {
				alert("Adjustment is not allowed to increase MVCPA Expenditures!");
				return false;
			}
			if (document.Adjustment.TotalExpendituresChange_Total.value != document.Adjustment.ProgramIncomeToBeAddedToBudget.value)
			{
				alert("The change in total expenditures must be equal to the change in program income in order to submit this budget adjustment");
				return false;
			}
<% End If %>
			for (var i = 1; i < 8; i++) {
				if (getNumericValue(document.Adjustment["MVCPAExpendituresNew_" + i].value) < 0.0) {
					alert("The New MVCPA Expenditures cannot be negative for any budget category.");
					return false;
				}
				if (getNumericValue(document.Adjustment["MatchExpendituresNew_"+i].value) < 0.0) {
					alert("The New Match Expenditures cannot be negative for any budget category.");
					return false;
				}
				if (getNumericValue(document.Adjustment["TotalExpendituresNew_"+i].value) < 0.0) {
					alert("The New Total Expenditures cannot be negative for any budget category.");
					return false;
				}
			}
			if (document.Adjustment["Confirmed"].checked == false) {
				alert("You must read certification and check box to submit progress report.");
				return false;
			}
		}
		return true;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body onload="updateTotals();">

<div class="sectiontitle">Motor Vehicle Crime Prevention Authority</div>
<div class="sectiontitle">FY <%=FiscalYear %> Grant Adjustment Request</div>
<%
If SubmitID > 0 Then
	Response.Write("<div style=""text-align: center;"">Submitted by " & SubmitName & ", " & SubmitTimeStamp & "</div>" & vbCrLf)
End If
If IsNull(SecondApprovalDate) = False Then
	Response.Write("<div style=""text-align: center;"">Approved by the MVCPA on " & SecondApprovalDate & "</div>" & vbCrLf)
End If
%>
<br />
<form name="Adjustment" id="Adjustment" method="post" action="AdjustmentSubmit.asp" onsubmit="return validateForm('');">
<%=HiddenField("GrantID",GrantID) %><%=HiddenField("AdjustmentID", AdjustmentID) %><%=HiddenField("FiscalYear", FiscalYear) %><%=HiddenField("Action","Save") %><%=HiddenField("InLieuOfNICBBudget",formatcurrency(InLieuOfNICBBudget,2,True,False,True)) %>
<table style="margin: auto; width: 900px">
	<tr>
		<td>Grantee</td>
		<td><%=GranteeName %></td>
	</tr>
	<tr>
		<td>Program Name</td>
		<td><%=ProgramName %></td>
	</tr>
	<tr>
		<td>Fiscal Year</td>
		<td><%=FiscalYear %></td>
	</tr>
	<tr>
		<td>Grant Number</td>
		<td><%=GrantNumber %></td>
	</tr>
</table>
<br />
<%
sql = "SELECT AdjustmentID, ProgramChange, BudgetChange, B.Name AS SubmitName, " & vbCrLF & _
	"	CAST(SubmitTimeStamp AS DATE) AS SubmitDate, SecondApprovalDate AS ApprovalDate, AdministrativelyClosedDate " & vbCrLf & _
	"FROM [Grants].Adjustments AS A " & vbCrLf & _
	"LEFT JOIN [System].Users AS B ON B.SystemID=A.SubmitID " & vbCrLf & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID) & " " & vbCrLf & _
	"ORDER BY AdjustmentID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	AdjustmentID = 0
	Response.Write("<p style=""text-align: center; "">There are no current Grant Adjustment Requests for this Grant. Starting new Grant Adjustment Request.</p>" & vbCrLf)
Else
	Response.Write("<table style=""margin: auto; width: 800px"">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write(vbTab & "<tr style=""vertical-align: bottom; ""><th colspan=""7"">Current Grant Adjustment Requests</th></tr>" & vbCrLf)
	Response.Write(vbTab & "<tr style=""vertical-align: bottom; "">" & vbCrLf)
	Response.Write(vbTab & "<th>Adjustment<br/>ID</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Submitted<br />By</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Submit<br />Date</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Program<br />Change</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Budget<br />Change</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Approval<br />Date</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Closed<br />Date</th>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	While rs.EOF = False
		Response.Write(vbTab & "<tr>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; ""><a href=""Adjustment.asp?GrantID=" & GrantID & "&AdjustmentID=" & rs.Fields("AdjustmentID") & """>" & rs.Fields("AdjustmentID") & "</a></td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("SubmitName") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & formatDate(rs.Fields("SubmitDate")) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; "">" & rs.Fields("ProgramChange") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; "">" & rs.Fields("BudgetChange") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; "">" & formatDate(rs.Fields("ApprovalDate")) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; "">" & formatDate(rs.Fields("AdministrativelyClosedDate")) & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>" & vbCrLf)
		rs.MoveNext()
		Wend
	If AdjustmentID > 0 Then
		Response.Write(vbTab & "<tr><td colspan=""7"" style=""text-align: center;""><a href=""Adjustment.asp?GrantID=" & GrantID & "&AdjustmentID=0"">Create a new Grant Adjustment Request</a></td></tr>" & vbCrLf)
	End If
	Response.Write("</table>" & vbCrLf)
	Response.Write("<br />" & vbCrLf)
End If

%>
<table style="margin: auto; width: 900px">
	<tr>
		<td>Grant Adjustment ID: <% 
		If AdjustmentID > 0 Then 
			Response.Write(AdjustmentID) 
		Else 
			Response.Write("<i>New Grant Adjustment Request</i>")
		End If
		%></td></tr>
	<tr>
		<td>This is a  
		<%=CheckBoxField("ProgramChange", ProgramChange) %>Program Change&nbsp;&nbsp;&nbsp;
		<%=CheckBoxField("BudgetChange", BudgetChange) %>Budget Change&nbsp;&nbsp;&nbsp;
		(Check each that applies)</td>
	</tr>
	<tr>
		<td><div class="boldtext">Program Modification Explanation and Reason:</div>
		<%=TextArea2("ProgramModificationExplanation", ProgramModificationExplanation, 10, 900, 8000, PermitEdit, "") %></td>
	</tr>
	<tr>
		<td><div class="boldtext">Budget Modification Explanation and Reason:</div>
		<%=TextArea2("BudgetModificationExplanation", BudgetModificationExplanation, 10, 900, 8000, PermitEdit, "") %></td>
	</tr>
</table>

<hr>

<%
sql = "SELECT A.GrantID, C.BudgetCategoryID, C.BudgetCategory, " & vbCrLf & _
	"	CASE WHEN B.ChangesApplied=1 And E.BudgetCategoryID>0 THEN E.TotalExpenditures ELSE D.TotalExpenditures END AS TotalExpenditures, " & vbCrLf & _
	"	CASE WHEN B.ChangesApplied=1 And E.BudgetCategoryID>0 THEN E.MVCPAExpenditures ELSE D.MVCPAExpenditures END AS MVCPAExpenditures, " & vbCrLf & _
	"	CASE WHEN B.ChangesApplied=1 And E.BudgetCategoryID>0 THEN E.MatchExpenditures ELSE D.MatchExpenditures END AS MatchExpenditures, " & vbCrLf & _
	"	E.TotalExpendituresChange, 	E.MVCPAExpendituresChange, E.MatchExpendituresChange, " & vbCrLf & _
	"	ISNULL(CASE WHEN B.ChangesApplied=1 And E.BudgetCategoryID>0 THEN E.TotalExpenditures ELSE D.TotalExpenditures END,0.0) + ISNULL(TotalExpendituresChange,0.0) AS TotalExpendituresNew, " & vbCrLf & _
	"	ISNULL(CASE WHEN B.ChangesApplied=1 And E.BudgetCategoryID>0 THEN E.MVCPAExpenditures ELSE D.MVCPAExpenditures END,0.0) + ISNULL(MVCPAExpendituresChange,0.0) AS MVCPAExpendituresNew, " & vbCrLf & _
	"	ISNULL(CASE WHEN B.ChangesApplied=1 And E.BudgetCategoryID>0 THEN E.MatchExpenditures ELSE D.MatchExpenditures END,0.0) + ISNULL(MatchExpendituresChange,0.0) AS MatchExpendituresNew " & vbCrLf & _
	"FROM [Grants].Main AS A " & vbCrLf & _
	"LEFT JOIN [Grants].Adjustments AS B ON B.GrantID=A.GrantID AND B.AdjustmentID=" & prepIntegerSQL(AdjustmentID) & " " & vbCrLf & _
	"CROSS JOIN Lookup.BudgetCategories AS C " & vbCrLf & _
	"LEFT JOIN [Grants].Budget AS D ON D.GrantID=A.GrantID AND D.BudgetCategoryID=C.BudgetCategoryID " & vbCrLf & _
	"LEFT JOIN [Grants].AdjustmentDetails AS E ON E.AdjustmentID=B.AdjustmentID AND E.BudgetCategoryID=C.BudgetCategoryID " & vbCrLf & _
	"WHERE A.GrantID=" & prepIntegerSQL(GrantID) & " " & vbCrLf & _ 
	"ORDER BY C.BudgetCategoryID"
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

'List out current budget
If rs.EOF = True Then
	Response.Write("Error: No Grant Budget records retrieved")
	SendMessage "Error: No Grant Budget records retrieved"
	Response.End
Else
	Response.Write("<table style=""margin: auto;"">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write(vbTab & "<tr><th colspan=""5"">Current Budget</th></tr>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & "<th>Budget Category</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Total<br />Expenditures</th>" & vbCrLf)
	Response.Write(vbTab & "<th>MVCPA<br />Expenditures</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Match<br />Expenditures</th>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write(vbTab & "<tr>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; white-space: nowrap; "">" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & CurrencyField("TotalExpenditures_" & rs.Fields("BudgetCategoryID"), rs.Fields("TotalExpenditures"), 12, 14, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & CurrencyField("MVCPAExpenditures_" & rs.Fields("BudgetCategoryID"), rs.Fields("MVCPAExpenditures"), 12, 14, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & CurrencyField("MatchExpenditures_" & rs.Fields("BudgetCategoryID"), rs.Fields("MatchExpenditures"), 12, 14, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>" & vbCrLf)

		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("<tfoot>" & vbCrLf)

	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: left; white-space: nowrap; font-weight: bold; "">Total</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & CurrencyField("TotalExpenditures_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & CurrencyField("MVCPAExpenditures_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & CurrencyField("MatchExpenditures_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab &"</tr>" & vbCrLf)

	Response.Write("</tfoot>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
End If
Response.Write("<br />" & vbCrLf)

' List out changes
rs.MoveFirst
If rs.EOF = True Then
	Response.Write("Error: No Grant Budget records retrieved")
	SendMessage "Error: No Grant Budget records retrieved"
	Response.End
Else
	Response.Write("<table style=""margin: auto;"">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write(vbTab & "<tr style=""vertical-align: bottom; ""><th colspan=""5"">Proposed Changes: indicate amount to increase or decrease budget item.</th></tr>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>Budget Category</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>Total<br />Expenditure<br />Change</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>MVCPA<br />Expenditure<br />Change</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>Match<br />Expenditure<br />Change</th>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write(vbTab & "<tr>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: left; white-space: nowrap; "">" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("TotalExpendituresChange_" & rs.Fields("BudgetCategoryID"), rs.Fields("TotalExpendituresChange"), 12, 14, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("MVCPAExpendituresChange_" & rs.Fields("BudgetCategoryID"), rs.Fields("MVCPAExpendituresChange"), 12, 14, PermitEdit, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("MatchExpendituresChange_" & rs.Fields("BudgetCategoryID"), rs.Fields("MatchExpendituresChange"), 12, 14, PermitEdit, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>" & vbCrLf)

		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("<tfoot>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)

	Response.Write(vbTab & vbTab & "<td style=""text-align: left; white-space: nowrap; font-weight: bold; "">Total</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("TotalExpendituresChange_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("MVCPAExpendituresChange_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("MatchExpendituresChange_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLf)

	Response.Write("</tfoot>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
End If
Response.Write("<br />" & vbCrLf)

' List out adjusted budget
rs.MoveFirst
If rs.EOF = True Then
	Response.Write("Error: No Grant Budget records retrieved")
	SendMessage "Error: No Grant Budget records retrieved"
	Response.End
Else
	Response.Write("<table style=""margin: auto;"">" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write(vbTab & "<tr><th colspan=""5"">Proposed New Budget</th></tr>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>Budget Category</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>New<br />Total<br />Expenditures</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>New<br />MVCPA<br />Expenditures</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>New<br />Match<br />Expenditures</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write(vbTab & "<tr>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: left; white-space: nowrap; "">" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("TotalExpendituresNew_" & rs.Fields("BudgetCategoryID"), rs.Fields("TotalExpendituresNew"), 12, 14, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("MVCPAExpendituresNew_" & rs.Fields("BudgetCategoryID"), rs.Fields("MVCPAExpendituresNew"), 12, 14, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("MatchExpendituresNew_" & rs.Fields("BudgetCategoryID"), rs.Fields("MatchExpendituresNew"), 12, 14, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>" & vbCrLf)

		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("<tfoot>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)

	Response.Write(vbTab & vbTab & "<td style=""text-align: left; white-space: nowrap; font-weight: bold; "">Total</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("TotalExpendituresNew_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("MVCPAExpendituresNew_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & CurrencyField("MatchExpendituresNew_Total", 0, 12, 14, False, "updateTotals();") & "</td>" & vbCrLf)

	Response.Write(vbTab & "</tr>" & vbCrLf)
	Response.Write("</tfoot>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
End If

Response.Write("<br />" & vbCrLF)
Response.Write("<hr>" & vbCrLf)

Response.Write("<table>" & vbCrLF)
Response.Write("<thead>" & vbCrLf & vbTab & "<tr><th colspan=""2"">Program Income</th></tr>" & vbCrLf & "</thead>" & vbCrLf)

Response.Write(vbTab & "<tr>" & vbCrLf)
Response.Write(vbTab & vbTab & "<td>Enter the amount of program income earned since the last submitted quarterly report</td>" & vbCrLf)
Response.Write(vbTab & vbTab & "<td>" & CurrencyField("NewProgramIncome", NewProgramIncome, 12, 14, PermitEdit, "checkCurrency(this); updateTotals();") & "</td>" & vbCrLf)
Response.Write(vbTab & "</tr>" & vbCrLf)

Response.Write(vbTab & "<tr>" & vbCrLf)
Response.Write(vbTab & vbTab & "<td>Enter the amount of program income to be moved into the program budget under this adjustment request.</td>" & vbCrLf)
Response.Write(vbTab & vbTab & "<td>" & CurrencyField("ProgramIncomeToBeAddedToBudget", ProgramIncomeToBeAddedToBudget, 12, 14, PermitEdit, "checkCurrency(this); updateTotals();") & "</td>" & vbCrLf)
Response.Write(vbTab & "</tr>" & vbCrLf)

Response.Write(vbTab & "<tr>" & vbCrLf)
Response.Write(vbTab & vbTab & "<td colspan=""2"">The amount moved into the budget must equal the change in total expenditures from the table above. " & _
	"Any increase in program expenditures must be supported by an increase in program income.</td>" & vbCrLf)
Response.Write(vbTab & "</tr>" & vbCrLf)

Response.Write("</table>" & vbCrLf)

If FiscalYear >= 2023 Then
	Response.Write("<br />" & vbCrLF & "<hr>" & vbCrLf)
	Response.Write("<table style=""width: 100%; "">" & vbCrLf & "<thead><tr><th colspan=""2"">Rates</th></tr></thead>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>Reimbursement rate before changes are approved:</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>" & NumberField("CurrentReimbursementRate", CurrentReimbursementRate, 18, 20, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>Updated reimbursement rate if changes are approved:</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>" & NumberField("ReimbursementRate", ReimbursementRate, 18, 20, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>Updated cash match if changes are approved (Must be at least 20%):</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>" & NumberField("CashMatchRate", CashMatchRate, 18, 20, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>Updated overtime if changes are approved (Must be less than 5%):</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>" & NumberField("OvertimeRate", OvertimeRate, 18, 20, False, "updateTotals();") & "</td>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
Else
	Response.Write(HiddenField("ReimbursementRate", ""))
	Response.Write(HiddenField("CashMatchRate", ""))
	Response.Write(HiddenField("OvertimeRate", ""))
End If

%><br />
<%	If AdjustmentID > 0 and ViewDocuments = True Then %>
<hr />
<%
	Dim Folder, file, files, DocumentFolder, fso, counter
	counter=0
	DocumentFolder = Application("DocumentRoot") & "\Grant\" & GrantID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemOBject")
	Response.Write("<table style=""margin: auto; "">" & vbCrLf)
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		Response.Write(vbTab & "<tr style=""vertical-align: top; ""><td></td><td>Current Documents in folder: ")
		If PermitEdit = True Then
		Response.Write("<a href=""../Upload/Upload.asp?fid=4&GrantID=" & GrantID & "&AdjustmentID=" & AdjustmentID & """ class=""plainlink"" target=""_blank"">Upload</a>")
		End If
		Response.Write("</td></tr>" & vbCrLf)
		Response.Write(vbTab & "<tr><td></td><td>")
		If files.count>0 Then 
			For Each file in files
				If Left(file.Name,7)="GA" & Right("00000" & AdjustmentID,5) Then
					Response.Write("<a href=""../Documents/Grant/" & GrantID & "/" & file.Name & _
						""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
					counter = counter + 1
				End If
			Next
		End If
	End If
	If counter = 0 Then
		Response.Write(vbTab & "<tr style=""vertical-align: top; ""><td colspan=""2"" style=""text-align: center; "">No Documents in folder</td></tr>" & vbCrLf)
	End If
	Response.Write("</table>" & vbCrLf)
	Response.Write("<hr />" & vbCrLf)
End If
%>
<p><%=CheckBoxField2("Confirmed", Confirmed, PermitEdit) %>
I have the authorization from the governing body to request and accept this proposed modification 
to the Statement of Grant Award. </p>
<%
' MVCPA Section
If SubmitID > 0 Then
	Response.Write("<br />" & vbCrLf)
	Response.Write("<table style=""margin: auto;"">" & vbCrLF)
	Response.Write("<thead>" & vbTab)
	Response.Write(vbTab & "<tr><th colspan=""2"">For Administrative Use Only</th></tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)

	Response.Write("<tr><td title=""First Approvers must be in MVCPA Grant Coordinator Role"">First Approver</td>")
	If MVCPAGrantCoordinator = True And ISNULL(FirstApprovalID) = False Then
		Response.Write("<td>" & DateField("FirstApprovalDate", FirstApprovalDate, MVCPAGrantCoordinator) & " by " & FirstApprover & "</td></tr>"  & vbCrLf)
	ElseIf MVCPAGrantCoordinator = True Then
		Response.Write("<td>" & DateField("FirstApprovalDate", FirstApprovalDate, MVCPAGrantCoordinator) & "</td></tr>"  & vbCrLf)
	ElseIf IsNull(FirstApprovalDate) = False Then
		Response.Write("<td>" & FirstApprovalDate & HiddenField("FirstApprovalDate", FirstApprovalDate) & " by " & FirstApprover & "</td></tr>"  & vbCrLf)
	Else
		Response.Write("<td style=""font-style: italic; "">Approval Pending" & HiddenField("FirstApprovalDate", FirstApprovalDate) & "</td></tr>"  & vbCrLf)
	End If

	Response.Write("<tr><td title=""Second Approvers must be in MVCPA Grant Administrator role"">Second Approver</td>")
	If MVCPAAdministrator = True And IsNull(SecondApprovalID) = False Then
		Response.Write("<td>" & DateField("SecondApprovalDate", SecondApprovalDate, MVCPAAdministrator) & " by " & SecondApprover & "</td></tr>"  & vbCrLf)
	ElseIf MVCPAAdministrator = True Then
		Response.Write("<td>" & DateField("SecondApprovalDate", SecondApprovalDate, MVCPAAdministrator) & "</td></tr>"  & vbCrLf)
	ElseIf IsNull(SecondApprovalDate) = False Then
		Response.Write("<td>" & SecondApprovalDate & HiddenField("SecondApprovalDate", SecondApprovalDate) & " by " & SecondApprover & "</td></tr>"  & vbCrLf)
	Else
		Response.Write("<td style=""font-style: italic; "">Approval Pending" & HiddenField("SecondApprovalDate", SecondApprovalDate) & "</td></tr>"  & vbCrLf)
	End If
	Response.Write("<tr><td title=""If the adjustment is denied by MVCPA, enter the date here that the grantee is notified."">If denied, date that grantee is notified</td>")
	If IsNull(FirstApprovalDate) AND ISNull(SecondApprovalDate) Then
		Response.Write("<td>" & DateField("DenialDate", DenialDate, MVCPAAdministrator) & "</td></tr>"  & vbCrLf)
	Else
		Response.Write("<td>" & DateField("DenialDate", DenialDate, MVCPAAdministrator) & "</td></tr>"  & vbCrLf)
	End If
	' Apply Changes Section
	If IsNull(FirstApprovalID) = False And IsNull(SecondApprovalID) = False And BudgetChange = True Then
		If ChangesApplied = True Then
			Response.Write("<td colspan=""2"" style=""text-align: center; "">The budget changes have been applied to the budget.</td>" & vbCrLf)
		Else
			Response.Write("<td>The budget changes have not been applied to the budget.</td>" & vbCrLf)
			Response.Write("<td>Apply Changes upon save (Director only.) " & CheckBoxField2("ApplyChanges", 0, MVCPAAdministrator) & "</td>" & vbCrLf)
		End If
	End If

	If MVCPARights=True Then
		Response.Write(vbTab & "<tr><td>Comments for Grantee</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td style=""text-align: right""><a href=""mailto:" & Trim(MailTo) & "?CC=grantsMVCPA@txdmv.gov" & _
			"&subject=MVCPA Grant Adjustment " & GranteeName & " Grant " & GrantNumber & " Adjustment ID=" & AdjustmentID & """>Send Email</a></td></tr>" & vbCrLf)
		Response.Write(vbTab & "<tr><td colspan=""2"">" & TextArea2("ExternalComments",ExternalComments, 6, 900, 10000, MVCPARights, "") & "</td></tr>" & vbCrLf)
	Else
		Response.Write(vbTab & "<tr><td colspan=""2"">Comments for Grantee: " & ExternalComments & HiddenField("ExternalComments",ExternalComments) & "</td></tr>" & vbCrLf)
	End If
	If MVCPARights=True Then
		If AdjustmentID>0 Then
			Response.Write(vbTab & "<tr><td colspan=""2"" style=""font-weight: bold; text-align: center"">Internal Comments</td></tr>" & vbCrLf)
			sql = "SELECT C.*, U.Name "& vbCrLf & _
				"FROM [Grants].AdjustmentComments AS C " & vbCrLf & _
				"LEFT JOIN [System].Users AS U ON U.SystemId=C.UpdateID " & vbCrLf & _
				"WHERE AdjustmentID=" & prepStringSQL(AdjustmentID) & " ORDER BY UpdateTimestamp"
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = Con.Execute(sql)
			While rs.EOF = False
				Response.Write(vbTab & "<tr style=""vertical-align: top; ""><td>" & rs.Fields("InternalComments") & "</td><td>" & rs.Fields("Name") & ", " & rs.Fields("UpdateTimestamp") & "</td></tr>" & vbCrLf)
				rs.MoveNext()
			Wend
		End If
		Response.Write("<tr><td colspan=""2"">New: " & TextArea2("InternalComments","", 4, 900, 10000, MVCPARights, "") & "</td></tr>" & vbCrLf)
		If ChangesApplied = False Then
			Response.Write("<tr><td colspan=""2"">" & CheckBoxField("ClearSubmit",false) & " Clear submission and any approvals to allow further edits.</td></tr>" & vbCrLf)
		End If
	End If
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
	Response.Write(HiddenField("FirstApprovalID",FirstApprovalID) & vbCrLf)
	Response.Write(HiddenField("SecondApprovalID",SecondApprovalID) & vbCrLf)
ElseIf MVCPARights = True Then
	Response.Write(HiddenField("FirstApprovalDate", FirstApprovalDate) & vbCrLf)
	Response.Write(HiddenField("FirstApprovalID", FirstApprovalID) & vbCrLf)
	Response.Write(HiddenField("SecondApprovalID", SecondApprovalID) & vbCrLf)
	Response.Write(HiddenField("SecondApprovalDate", SecondApprovalDate) & vbCrLf)
	'Response.Write(HiddenField("ExternalComments", ExternalComments) & vbCrLf)
	Response.Write(HiddenField("DenialDate", DenialDate) & vbCrLf)
	If IsNull(SubmitTimestamp) = True Then
		Response.Write("<br />" & vbCrLF)
		Response.Write("<table style=""margin: auto;"">" & vbCrLF)
		Response.Write("<thead>" & vbCrLf)
		Response.Write(vbTab & "<tr><th colspan=""2"">For Administrative Use Only</th></tr>" & vbCrLf)
		Response.Write("</thead>" & vbCrLf)
		Response.WRite("<tbody>" & vbCrLf)
		Response.Write(vbTab & "<tr><td title=""Enter date that adjustment is closed because Grantee does not intent to proceed."">Date that Adjustment is Administratively Closed</td>")
		Response.Write(vbTab & "<td>" & DateField("AdministrativelyClosedDate", AdministrativelyClosedDate, MVCPARights) & "</td></tr>"  & vbCrLf)
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
	End If
	If Len(ExternalComments)>0 Then
		Response.Write("<table style=""margin: auto; "">" & vbCrLf)
		Response.Write(vbTab & "<tr><td>Comments for Grantee</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right""><a href=""mailto: " & MailTo & "?CC=grantsMVCPA@txdmv.gov" & _
			"&subject=MVCPA Grant Adjustment " & GranteeName & " Grant " & GrantNumber & " Adjustment ID=" & AdjustmentID & """>Send Email</a></td></tr>")
		Response.Write(vbTab & "<tr><td colspan=""2"">" & TextArea2("ExternalComments",ExternalComments, 6, 900, 10000, MVCPARights, "") & "</td></tr>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
	End If

	If AdjustmentID>0 Then
		Response.Write("<table style=""margin: auto; "">" & vbCrLf)
		Response.Write(vbTab & "<tr><td colspan=""2"" style=""font-weight: bold; text-align: center"">Internal Comments</td></tr>" & vbCrLf)
		sql = "SELECT C.*, U.Name "& vbCrLf & _
			"FROM [Grants].AdjustmentComments AS C " & vbCrLf & _
			"LEFT JOIN [System].Users AS U ON U.SystemId=C.UpdateID " & vbCrLf & _
			"WHERE AdjustmentID=" & prepStringSQL(AdjustmentID) & " ORDER BY UpdateTimestamp"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		While rs.EOF = False
			Response.Write(vbTab & "<tr style=""vertical-align: top; "">" & vbCrLF)
			Response.Write(vbTab & "<td>" & rs.Fields("InternalComments") & "</td><td>" & rs.Fields("Name") & ", " & rs.Fields("UpdateTimestamp") & "</td>" & vbCrLf)
			Response.Write(vbTab & "</tr>" & vbCrLf)
			rs.MoveNext()
		Wend
	End If
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & "<td colspan=""2"">New: <br />" & TextArea2("InternalComments","", 4, 900, 10000, MVCPARights, "") & "</td>" & vbCrLf)
	Response.Write(vbTab &"</tr>" & vbCrLf)
	If ChangesApplied = False Then
		Response.Write(vbTab & "<tr>" & vbCrLf)
		Response.Write(vbTab & "<td colspan=""2"">" & CheckBoxField("ClearSubmit",false) & " Clear submission and any approvals to allow further edits.</td></tr>" & vbCrLf)
	End If
	Response.Write("</table>" & vbCrLf)
End If

If MVCPARights = True Then
	sql = "SELECT A.SubmitTimestamp, B.Name AS SubmitName " & vbCrLf & _
		"FROM [Grants].Adjustments AS A " & vbCrLf & _
		"LEFT JOIN [System].Users AS B ON A.SubmitID=B.SystemID " & vbCrLf & _
		"WHERE AdjustmentID=" & prepIntegerSQL(AdjustmentID) & " AND SubmitTimestamp IS NOT NULL " & vbCrLf & _
		"UNION " & vbCrLf & _
		"SELECT A.SubmitTimestamp, B.Name AS SubmitName " & vbCrLf & _
		"FROM [Grants].Adjustments_Log AS A " & vbCrLf & _
		"LEFT JOIN [System].Users AS B ON A.SubmitID=B.SystemID " & vbCrLf & _
		"WHERE AdjustmentID=" & prepIntegerSQL(AdjustmentID) & " AND SubmitTimestamp IS NOT NULL " & vbCrLf & _
		"ORDER BY 1 DESC "
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		Response.Write("<br>" & vbCrLf)
		Response.Write("<table style=""margin: auto; "">" & vbCrLf)
		Response.WRite("<thead>" & vbCrLF)
		Response.WRite(vbTab & "<tr><th>Submission History</th></tr>" & vbCrLF)
		Response.Write("</thead>" & vbCrLf)
		Response.Write("<tbody>" & vbCrLf)
		While rs.EOF = False
			Response.Write(vbTab & "<tr><td style=""text-align: center"">Submitted By " & rs.Fields("SubmitName") & ", " & rs.Fields("SubmitTimestamp") & "</td></tr>" & vbCrLf)
			rs.MoveNext()
		Wend
		Response.Write("</tbody>" & vbCrLf & "</table><br />" & vbCrLf)
	End If
End If
'CanSubmit = True
If CanSubmit = False Then
	If (UserSystemID=156 Or UserSystemID=402 Or USerSystemID=1) And FiscalYear=2023 And DATE()<CDATE("2023/05/01") Then
		CanSubmit = True
	End If
End If
%>
<br />
<div style="text-align: center">
<%	If (IsNull(SubmitID)=True And PermitEdit=True) Or MVCPARights = True Then %>
<input type="submit" name="Save" value="Save" onclick="return validateForm('Save');" />
<%
	End If
	If IsNull(SubmitID)=True And CanSubmit = True And PermitEdit=True Then 
%>
<input type="submit" name="Submit" value="Submit" onclick="return validateForm('Submit');" />
<%		 
	End If %>
<input type="reset" name="Reset" value="Reset" />
<input type="button" value="Close" onclick="window.close();" />
</div>

</form>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->