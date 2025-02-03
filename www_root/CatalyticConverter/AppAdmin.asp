<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j
debug = False
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

If MVCPARights = False And MVCPAViewer = False Then
	Response.Write("Forbidden: You do not have permissions to access this page.")
	SendMessage "Forbidden: You do not have permissions to access this page."
	Response.End
End If

Dim AppID, GranteeName, ProgramName, ORI, Agency, GrantClassID, GrantClass, FiscalYear, GrantTypeID, GrantType, _
	SubmitTimestamp, SubmitName, ConfirmationNumber, StaffShortProgramDescription, _
	ResolutionConfirmedDate, ResolutionFundsProvided, ResolutionReturnFunds, _
	ResolutionDesignateOfficals, ResolutionGoverningBody, ResolutionDelegationSupported, _
	ApplicationCertifiedCompleteDate, ApplicationConsideredDate, GrantResultID, _
	GrantAwardAmount, GrantNumber, POIssueDate, _
	InterlocalAgreementConfirmedDate, InterlocalAgreementConfirmedBy, InterlocalAgreementConfirmedByName, _
	ProsecutorAgreementConfirmedDate, ProsecutorAgreementConfirmedID, ProsecutorAgreementConfirmedByName, _
	OperationalPlanApprovalDate, OperationalPlanApprovalName, OperationalPlanApprovalID, MultiAgencyGrant, _
	InitialAwardTransmissionDate, CreateNegotiationRecords, AwardAcceptanceDate, _
	RevisedAppID, RevisedSubmitTimestamp, RevisedSubmitName, RevisedConfirmationNumber, RevisionsAcceptedDate, _
	OfficialGrantAwardLetterDate, AwardLetterTransmissionMethodID, NegotiationLocked, _
	SignedGrantAwardLetterDate, AwardAcceptanceSignatureConfirmedDate, GrantAwardCertifiedComplete, _
	GrantAwardDeclineLetterReceived, Notes, ExcludeFromConsideration, ConsiderationNotes, _
	GrantRecordPresent, BudgetRecordsPresent, GrantID, ShowOnlySubmitted, ApplicationSchema
Dim ProgramCategory(5)

AppID = Request.QueryString("AppID")
If Len(AppID) = 0 Then
	AppID=0
Else
	If IsNumeric(AppID) = False Then
		Response.Write("Error: Non-numeric AppID")
		Response.End
	End If
	AppID=CInt(AppID)
End If
FiscalYear=Request.QueryString("FiscalYear")
If Len(FiscalYear)=0 then
	FiscalYear = Session("FiscalYear")
ElseIf IsNumeric(FiscalYear) Then
	FiscalYear = CInt(FiscalYear)
Else
	Response.Write("Invalid value for fiscal year")
	Response.End
End If

ApplicationSchema = getCCApplicationSchema(FiscalYear)
If FiscalYear = 2024 Then
	ShowOnlySubmitted = False
Else
	ShowOnlySubmitted = True
End If
If Debug = True Then
	Response.Write("<pre>ApplicationSchema=" & ApplicationSchema & "</pre>" & vbCrLf)
	Response.Flush
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Application Administration</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm()
	{
		setInterlocalReviewer();
		saving = true;
		SaveChanges();
		form.submit();
	}

	function goHome()
	{
		if (!saving) {
			var f = FormChanges(form);
			if (f.length > 0) {
				if (window.confirm("Your form updates have not be saved. Do you wish to continue without saving?")) {
					location.href = '../Home/default.asp';
				}
				else {
					submitForm();
					return false;
				}
			}
			else {
				location.href = '../Home/default.asp';
			}
		}
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body style="width: 100%; ">

<div>

<form name="AppAdmin" id="AppAdmin" method="post" action="AppAdminSubmit.asp" onsubmit="return submitForm();">
<input type="hidden" name="Changes" />
<input type="hidden" name="FiscalYear" value="<%=FiscalYear %>" />
<table style="margin: auto; padding: 4px;">
<tr>
	<th colspan="2">Grant Application Administration</th>
</tr>
<tr>
	<td colspan="2" style="text-align: center; ">Application to display: <select name="NewAppID" onchange="submitForm();">
<%
sql = "SELECT A.AppID, A.ProgramName, C.GranteeName, A.GrantTypeID, B.GrantType " & vbCrLf & _
	"FROM Application.IDs AS I " & vbCrLf & _
	"LEFT JOIN CC.Application AS A ON A.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN Lookup.GrantType AS B ON B.GrantTypeID=A.GrantTypeID And B.Version=1 " & vbCrLf & _
	"LEFT JOIN Grantees AS C ON C.GranteeID=I.GranteeID " & vbCrLf
If ShowOnlySubmitted = True Then
	sql = sql & "WHERE I.GrantClassID=4 AND (I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND SubmitTimestamp IS NOT NULL) OR A.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf
Else
	sql = sql & "WHERE I.GrantClassID=4 AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " OR I.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf
End If
sql = sql &	"ORDER BY REPLACE(C.GranteeName,'City of ',''), A.GrantTypeID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<option value=""0"">Select Application</option>" & vbCrLf)
	While rs.EOF = False
		If rs.Fields("AppID") = AppID Then
			Response.Write("<option value=""" & rs.Fields("AppID") & """ selected>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</option>" & vbCrLf)
		Else
			Response.Write("<option value=""" & rs.Fields("AppID") & """>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</option>" & vbCrLf)
		End If
		rs.MoveNext()
	Wend
End If
%></select></td>
</tr>
<tr><td colspan="2">&nbsp;</td></tr>
<%

sql = "SELECT I.AppID, G.GranteeName, G.ORI, O.Agency, A.ProgramName, I.GrantClassID, GC.GrantClass, I.FiscalYear, A.GrantTypeID, T.GrantType, " & vbCrLf & _
	"	A.SubmitTimestamp, U.Name AS SubmitName, A.ConfirmationNumber, B.StaffShortProgramDescription, " & vbCrLf & _
	"	B.ResolutionConfirmedDate, B.ResolutionFundsProvided, B.ResolutionReturnFunds, " & vbCrLf & _
	"	B.ResolutionDesignateOfficals, B.ResolutionGoverningBody, B.ResolutionDelegationSupported, " & vbCrLf & _
	"	B.ApplicationCertifiedCompleteDate, B.ApplicationConsideredDate, B.GrantResultID, " & vbCrLf & _
	"	B.GrantAwardAmount, B.GrantNumber, B.POIssueDate, B.InterlocalAgreementConfirmedDate, B.InterlocalAgreementConfirmedBy, " & vbCrLf & _
	"	U3.Name AS InterlocalAgreementConfirmedByName, B.ProsecutorAgreementConfirmedDate, B.ProsecutorAgreementConfirmedID, " & vbCrLf & _
	"	B.OperationalPlanApprovalID, B.OperationalPlanApprovalDate, U5.Name AS OperationalPlanApprovalName, " & vbCrLf & _
	"	CAST(CASE WHEN NAC.NegotiationAgencyCount>1 THEN 1 WHEN AAC.ApplicationAgencyCount>1 THEN 1 ELSE 0 END AS BIT) AS MultiAgencyGrant, " & vbCrLf & _
	"	U4.Name AS ProsecutorAgreementConfirmedByName, B.InitialAwardTransmissionDate, " & vbCrLf & _
	"	ISNULL(B.NegotiationLocked,0) AS NegotiationLocked, B.AwardAcceptanceDate, " & vbCrLf & _
	"	A2.AppID AS RevisedAppID, A2.SubmitTimestamp AS RevisedSubmitTimestamp, " & vbCrLF & _
	"	U2.Name AS RevisedSubmitName, A2.ConfirmationNumber AS RevisedConfirmationNumber, " & vbCrLF & _
	"	B.RevisionsAcceptedDate, B.OfficialGrantAwardLetterDate, B.AwardLetterTransmissionMethodID, " & vbCrLf & _
	"	B.SignedGrantAwardLetterDate, B.AwardAcceptanceSignatureConfirmedDate, " & vbCrLf & _
	"	B.GrantAwardCertifiedComplete, B.GrantAwardDeclineLetterReceived, B.Notes, " & vbCrLf & _
	"	B.ExcludeFromConsideration, B.ConsiderationNotes, B.UpdateID, B.UpdateTimestamp,  " & vbCrLf & _
	"	CAST(CASE WHEN Grants.AppID IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS GrantRecordPresent, " & vbCrLf & _
	"	BudgetRecordsPresent = (SELECT CASE WHEN COUNT(*)>0 THEN 1 ELSE 0 END FROM [Grants].Budget WHERE GrantID=Grants.GrantID), " & vbCrLf & _
	"	Grants.GrantID " & vbCrLF & _
	"FROM Application.IDs AS I " & vbCrLF & _
	"LEFT JOIN CC.Application AS A ON A.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN CC.Admin AS B ON B.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN Grantees AS G ON G.GranteeID=I.GranteeID " & vbCrLf & _
	"LEFT JOIN Lookup.GrantClass AS GC ON GC.GrantClassID=I.GrantClassID " & vbCrLf & _
	"LEFT JOIN Lookup.ORI AS O ON O.ORI=G.ORI " & vbCrLf & _
	"LEFT JOIN Lookup.GrantType AS T ON T.GrantTypeID=A.GrantTypeID And T.Version=1 " & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=A.SubmitID " & vbCrLf & _
	"LEFT JOIN CC.Negotiation AS A2 ON A2.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN System.Users AS U2 ON U2.SystemID=A2.SubmitID " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS Grants ON Grants.AppID=A.AppID " & vbCrLf & _
	"LEFT JOIN System.Users AS U3 ON U3.SystemID=B.InterlocalAgreementConfirmedBy " & vbCrLf & _
	"LEFT JOIN System.Users AS U4 ON U4.SystemID=B.ProsecutorAgreementConfirmedID " & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT AppID, COUNT(*) AS ApplicationAgencyCount " & vbCrLf & _
	"	FROM Application.ParticipatingAgencies " & vbCrLf & _
	"	GROUP BY AppID " & vbCrLf & _
	") AS AAC ON AAC.AppID = A.AppID " & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT AppID, COUNT(*) AS NegotiationAgencyCount " & vbCrLf & _
	"	FROM Negotiation.ParticipatingAgencies " & vbCrLf & _
	"	GROUP BY AppID " & vbCrLf & _
	") AS NAC ON NAC.AppID = A.AppID " & vbCrLf & _
	"LEFT JOIN [Grants].OperationalPlan AS OP ON OP.AppID=A.AppID " & vbCrLf & _
	"LEFT JOIN System.Users AS U5 ON U5.SystemID=B.OperationalPlanApprovalID " & vbCrLf & _
	"WHERE A.AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error: Unable to retrive application record.")
		Response.End
	Else
		GrantClassID = rs.Fields("GrantClassID")
		GrantClass = rs.Fields("GrantClass")
		GranteeName = rs.Fields("GranteeName")
		ORI = rs.Fields("ORI")
		Agency = rs.Fields("Agency")
		ProgramName = rs.Fields("ProgramName")
		FiscalYear = rs.Fields("FiscalYear")
		GrantTypeID = rs.Fields("GrantTypeID")
		GrantType = rs.Fields("GrantType")
		SubmitTimestamp = rs.Fields("SubmitTimestamp")
		SubmitName = rs.Fields("SubmitName")
		ConfirmationNumber = rs.Fields("ConfirmationNumber")
		StaffShortProgramDescription = rs.Fields("StaffShortProgramDescription")
		ResolutionConfirmedDate = rs.Fields("ResolutionConfirmedDate")
		ResolutionFundsProvided = rs.Fields("ResolutionFundsProvided")
		ResolutionReturnFunds = rs.Fields("ResolutionReturnFunds")
		ResolutionDesignateOfficals = rs.Fields("ResolutionDesignateOfficals")
		ResolutionGoverningBody = rs.Fields("ResolutionGoverningBody")
		ResolutionDelegationSupported = rs.Fields("ResolutionDelegationSupported")
		ApplicationCertifiedCompleteDate = rs.Fields("ApplicationCertifiedCompleteDate")
		ApplicationConsideredDate = rs.Fields("ApplicationConsideredDate")
		GrantResultID = rs.Fields("GrantResultID")
		GrantAwardAmount = rs.Fields("GrantAwardAmount")
		GrantNumber = rs.Fields("GrantNumber")
		POIssueDate = rs.Fields("POIssueDate")
		InterlocalAgreementConfirmedDate = rs.Fields("InterlocalAgreementConfirmedDate")
		InterlocalAgreementConfirmedBy = rs.Fields("InterlocalAgreementConfirmedBy")
		InterlocalAgreementConfirmedByName = rs.Fields("InterlocalAgreementConfirmedByName")
		ProsecutorAgreementConfirmedDate = rs.Fields("ProsecutorAgreementConfirmedDate")
		ProsecutorAgreementConfirmedID = rs.Fields("ProsecutorAgreementConfirmedID")
		ProsecutorAgreementConfirmedByName = rs.Fields("ProsecutorAgreementConfirmedByName")
		OperationalPlanApprovalID = rs.Fields("OperationalPlanApprovalID")
		OperationalPlanApprovalDate = rs.Fields("OperationalPlanApprovalDate")
		OperationalPlanApprovalName = rs.Fields("OperationalPlanApprovalName")
		MultiAgencyGrant = rs.Fields("MultiAgencyGrant")
		InitialAwardTransmissionDate = rs.Fields("InitialAwardTransmissionDate")
		AwardAcceptanceDate = rs.Fields("AwardAcceptanceDate")
		RevisedAppID = rs.Fields("RevisedAppId")
		RevisedSubmitTimestamp = rs.Fields("RevisedSubmitTimestamp")
		RevisedSubmitName = rs.Fields("RevisedSubmitName")
		RevisedConfirmationNumber = rs.Fields("RevisedConfirmationNumber")
		RevisionsAcceptedDate = rs.Fields("RevisionsAcceptedDate")
		OfficialGrantAwardLetterDate = rs.Fields("OfficialGrantAwardLetterDate")
		AwardLetterTransmissionMethodID = rs.Fields("AwardLetterTransmissionMethodID")
		NegotiationLocked = rs.Fields("NegotiationLocked")
		SignedGrantAwardLetterDate = rs.Fields("SignedGrantAwardLetterDate")
		AwardAcceptanceSignatureConfirmedDate = rs.Fields("AwardAcceptanceSignatureConfirmedDate")
		GrantAwardCertifiedComplete  = rs.Fields("GrantAwardCertifiedComplete")
		GrantAwardDeclineLetterReceived = rs.Fields("GrantAwardDeclineLetterReceived")
		Notes = rs.Fields("Notes")
		ExcludeFromConsideration = rs.Fields("ExcludeFromConsideration")
		ConsiderationNotes = rs.Fields("ConsiderationNotes")
		GrantRecordPresent = rs.Fields("GrantRecordPresent")
		BudgetRecordsPresent = rs.Fields("BudgetRecordsPresent")
		GrantID = rs.Fields("GrantID")
		CreateNegotiationRecords = False
	End If

%>
<tr>
	<td>Application ID:</td>
	<td><%=AppID %><%=HiddenField("AppID", AppID) %>
		<div style="float: right; text-align: right; "><a href="<%=AppMailTo(AppID) %>">email</a></div></td>
</tr>

<tr>
	<td>Grant Class:</td>
	<td><%=GrantClass %><%=HiddenField("GrantClassID", GrantClassID) %>&nbsp;&nbsp;(<%=GrantClassID %>)</td>
</tr>

<tr>
	<td>Grantee Name:</td>
	<td><%=GranteeName %></td>
</tr>

<tr>
	<td>ORI:</td>
	<td><%=ORI %>&nbsp;<%=Agency %></td>
</tr>

<tr>
	<td>Program Name:</td>
	<td><%=ProgramName %></td>
</tr>

<tr>
	<td>Fiscal Year:</td>
	<td><%=FiscalYear %></td>
</tr>

<tr>
	<td>Application Category:</td>
	<td><%=GrantType %></td>
</tr>

<tr>
	<td>Application Submitted By:</td>
	<td><%=SubmitName %></td>
</tr>

<tr>
	<td>Submission Date and Time:</td>
	<td><%=SubmitTimestamp %></td>
</tr>
<tr>
	<td>Confirmation Number:</td>
	<td><%=ConfirmationNumber %> 
<%	
	If (IsNull(OfficialGrantAwardLetterDate)=True Or MVCPAAdministrator = True) And IsNull(SubmitTimestamp)=False Then %>
	<div style="float: right; text-align: right; ">
	<%=CheckBoxField("ClearSubmit", false) %>Clear Submission and allow editing</div>
<%	End If %>
	</td>
</tr>
<tr style="vertical-align: top; ">
	<td>Staff Short Program Description</td>
	<td><%=TextArea("StaffShortProgramDescription", StaffShortProgramDescription, 3, 80, 1000, True, "") %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Resolution Checklist:</td>
	<td><%=CheckBoxField("ResolutionFundsProvided", ResolutionFundsProvided) %> Funds for the MVCPA purpose provided in statute<br />
	<%=CheckBoxField("ResolutionReturnFunds", ResolutionReturnFunds) %> Return funds for loss or misuse<br />
	<%=CheckBoxField("ResolutionDesignateOfficals", ResolutionDesignateOfficals) %> Designate Officials<br />
	<%=CheckBoxField("ResolutionGoverningBody", ResolutionGoverningBody) %> Governing Body <b>OR</b> 
	<%=CheckBoxField("ResolutionDelegationSupported", ResolutionDelegationSupported) %> Delegation supported by delegation order With Attorney Letter
	</td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date that Resolution Confirmed:</td>
	<td><%=DateField("ResolutionConfirmedDate", ResolutionConfirmedDate, True) %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date <b>Application Certified Complete</b>:</td>
	<td><%=DateField("ApplicationCertifiedCompleteDate", ApplicationCertifiedCompleteDate, True) %></td>
</tr>
<tr><td colspan="2"><hr /></td></tr>
<tr style="vertical-align: top; ">
	<td>Exclude From Consideration:</td>
	<td><%=CheckBoxField("ExcludeFromConsideration", ExcludeFromConsideration) %> This checkbox removes the application from the Award Allocation worksheet.</td>
</tr>
<tr style="vertical-align: top; ">
	<td>Notes regarding exclusion of application from consideration</td>
	<td><%=TextArea("ConsiderationNotes", ConsiderationNotes, 5, 80, 1000, True, "") %></td>
</tr>

<tr><td colspan="2"><hr /></td></tr>
<tr style="vertical-align: top; ">
	<td>Date that Application is considered by MVCPA Board:</td>
	<td><%=DateField("ApplicationConsideredDate", ApplicationConsideredDate, True) %></td>
</tr>
<tr>
	<td>Grant Result by MVCPA Board</td>
	<td><select name="GrantResultID" id="GrantResultID">
		<option value="0">Select grant result</option>
<%
	sql = "SELECT GrantResultID, GrantResult FROM Lookup.GrantResults ORDER BY GrantResultSort"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		If GrantResultID = rs.Fields("GrantResultID") Then
			Response.Write("<option value=""" & rs.Fields("GrantResultID") & """ selected>" & rs.Fields("GrantResult") & "</option>" & vbCrLf)
		Else
			Response.Write("<option value=""" & rs.Fields("GrantResultID") & """>" & rs.Fields("GrantResult") & "</option>" & vbCrLf)
		End If
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Grant Amount Awarded by MVCPA Board:</td>
	<td><%=CurrencyField("GrantAwardAmount", GrantAwardAmount, 15, 15, True, "") %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Grant Number for awarded grants:</td>
	<td><%=TextFieldDblClick("GrantNumber", GrantNumber, 14, 16, True, "", "this.value='608-" & Mid(CStr(FiscalYear),3,2) & "-" & Mid(ORI,3) & "';") %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date that Purchase Order is issued:</td>
	<td><%=DateField("POIssueDate", POIssueDate, True) %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date that initial award amount is transmitted with instructions:</td>
	<td><%=DateField("InitialAwardTransmissionDate", InitialAwardTransmissionDate, True) %></td>
</tr>
<%	If ApplicationSchema = "Negotiation" Then %>
<tr><td colspan="2"><hr /></td></tr>
<%		If IsNull(RevisedAppID) = True Then %>
<tr><th colspan="2">Negotiation Application has not been created</th></tr>
<%			If IsNull(GrantAwardAmount) = False Then %>
<tr style="vertical-align: top; ">
	<td>Create Negotiation Records from Application:</td>
	<td><%=CheckBoxField("CreateNegotiationRecords", CreateNegotiationRecords) %> (This will delete any current negotiation records related to this application.)</td>
</tr>
<%			End If %>
<%		Else %>
<tr><th colspan="2">Negotiation Application has been created</th></tr>
<%		End If %>
<tr>
	<td>Revised Application Submitted By:</td>
	<td><%=RevisedSubmitName %></td>
</tr>
<tr>
	<td>Revised Submission Date and Time:</td>
	<td><%=RevisedSubmitTimestamp %></td>
</tr>
<tr>
	<td>Revised Application Confirmation Number:</td>
	<td><%=RevisedConfirmationNumber %>  
	<div style="float: right; text-align: right; ">
	<%=CheckBoxField("RevisedClearSubmit", false) %>Clear Revised Submission and allow editing</div></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date that <b>Revisions from grantee are Accepted</b>:</td>
	<td><%=DateField("RevisionsAcceptedDate", RevisionsAcceptedDate, True) %></td>
</tr>
<%	End If %>
<tr><td colspan="2"><hr /></td></tr>
<tr style="vertical-align: top; ">
	<td>Date that Official Grant Award Letter is transmitted:</td>
	<td><%=DateField("OfficialGrantAwardLetterDate", OfficialGrantAwardLetterDate, True) %>&nbsp;&nbsp;
	<%=CheckBoxField("NegotiationLocked", NegotiationLocked) %>Lock Negotiation Records
	</td>
</tr>
<tr style="vertical-align: top; ">
	<td>Transmission Method:</td>
	<td><%
	sql = "SELECT TransmissionMethodID, TransmissionMethod FROM Lookup.TransmissionMethods ORDER BY TransmissionSort"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(RadioInputField("AwardLetterTransmissionMethodID", AwardLetterTransmissionMethodID, rs.Fields("TransmissionMethodID")) & rs.Fields("TransmissionMethod") & "&nbsp;&nbsp;&nbsp;")
		rs.MoveNext()
	Wend
%></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date that Interlocal Agreement(s) are confirmed:</td>
	<td><%=DateField("InterlocalAgreementConfirmedDate", InterlocalAgreementConfirmedDate, True) %><%=HiddenField("InterlocalAgreementConfirmedBy", InterlocalAgreementConfirmedBy) %> 
	by <%=InterlocalAgreementConfirmedByName %><%=HiddenField("InterlocalAgreementConfirmedBy",InterlocalAgreementConfirmedBy) %>
	</td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date that Prosecutor Agreement(s) are confirmed:</td>
	<td><%=DateField("ProsecutorAgreementConfirmedDate", ProsecutorAgreementConfirmedDate, True) %><%=HiddenField("InterlocalAgreementConfirmedBy", InterlocalAgreementConfirmedBy) %> 
	by <%=ProsecutorAgreementConfirmedByName %><%=HiddenField("ProsecutorAgreementConfirmedID",ProsecutorAgreementConfirmedID) %>
	</td>
</tr>

<tr style="vertical-align: top; ">
	<td>Date that Operational Plans(s) for Multi-agency Grants are confirmed:</td>
	<td><%=DateField("OperationalPlanApprovalDate", OperationalPlanApprovalDate, True) %><%=HiddenField("OperationalPlanApprovalID", OperationalPlanApprovalID) %> 
	by <%=OperationalPlanApprovalName %> 
	<div style="float: right; text-align: right; ">Multi-Agency Grant? <% If MultiAgencyGrant = True Then Response.Write("Yes") Else Response.Write("No") End If %></div></td>
</tr>

<tr style="vertical-align: top; ">
	<td>Date that Award is Accepted:</td>
	<td><%=DateField("AwardAcceptanceDate", AwardAcceptanceDate, True) %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date that acceptance and signed Grant Award Acceptance Letter is received:</td>
	<td><%=DateField("SignedGrantAwardLetterDate", SignedGrantAwardLetterDate, True) %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date that Resolution and Award Acceptance signatory is confirmed:</td>
	<td><%=DateField("AwardAcceptanceSignatureConfirmedDate", AwardAcceptanceSignatureConfirmedDate, True) %></td>
</tr>
<%	If (ApplicationSchema="Application" And ISNull(SubmitTimestamp) = True) Or (ApplicationSchema="Negotiation" And ISNull(RevisedSubmitTimestamp) = True) Then %>
<tr style="vertical-align: top; ">
	<td>Date that <b>Grant Award Certified Complete</b>:</td>
	<td><%=DateField("GrantAwardCertifiedComplete", GrantAwardCertifiedComplete, False) %> 
	<font color="red"><%=ApplicationSchema %> has not been submitted</font></td>
</tr>
<%	Else %>
<tr style="vertical-align: top; ">
	<td>Date that <b>Grant Award Certified Complete</b>:</td>
	<td><%=DateField("GrantAwardCertifiedComplete", GrantAwardCertifiedComplete, True) %></td>
</tr>
<%	End If%>
<%	If GrantResultID=7 And GrantRecordPresent=False Then %>
<tr>
	<td>Application Process Complete, but no grant record exists.</td>
	<td><%=CheckBoxField("CreateGrantRecord", false) %>Create Grant Records from Application</td>
</tr>
<%	ElseIf GrantResultID=7 And GrantRecordPresent=True And BudgetRecordsPresent=False Then %>
<tr>
	<td>Application Process Complete. Only a stub grant record exists. Other grant records need to be created.</td>
	<td><%=CheckBoxField("CreateGrantRecord", false) %>Create Grant Records from Application</td>
</tr>
<%	ElseIf GrantRecordPresent = True And BudgetRecordsPresent = False  Then %>
<tr style="vertical-align: top; ">
	<td><b>Grant Record</b>:</td>
	<td>A stub grant record has been created with GrantID=<%=GrantID %>. The other records have not been generated.<br /><a href="../Upload/Upload.asp?GrantID=<%=GrantID%>&FID=11" class="plainlink" target="_blank">Upload SGA</a></td>
</tr>
<%	ElseIf IsNull(GrantID) = False Then %>
<tr style="vertical-align: top; ">
	<td><b>Grant Record</b>:</td>
	<td>A grant record has been created with GrantID=<%=GrantID %>.<br /><a href="../Upload/Upload.asp?GrantID=<%=GrantID%>&FID=11" class="plainlink" target="_blank">Upload SGA</a></td>
</tr>
<%	ElseIf GrantRecordPresent = False And IsNull(GrantAwardAmount) = False And IsNull(GrantNumber) = False Then %>
<tr>
	<td>Application Process is not Complete and no Grant Record Exists </td>
	<td><%=CheckBoxField("CreateStubRecord", false) %>Create Stub Grant Record from Application to allow for adding allocations.</td>
</tr>
<%	End If %>
<tr style="vertical-align: top; ">
	<td colspan="2">Notes:
	<%=TextArea("Notes", Notes, 4, 120, 4000, True, "") %></td>
</tr>
<%
Dim Folder, file, files, DocumentFolder, fso
DocumentFolder = Application("DocumentRoot") & "\Application\" & AppID & "\"
set fso = Server.CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(DocumentFolder) = False Then
	fso.CreateFolder(DocumentFolder)
End If

Set folder = fso.GetFolder(DocumentFolder)
Set files = folder.Files
Response.Write("<tr style=""vertical-align: top; ""><td>Current Documents in application folder: <a href=""../Upload/Upload.asp?AppID=" & AppID & "&FID=2"" class=""plainlink"" target=""_blank"">Upload</a></td>" & vbCrLf & vbTab & "<td>")
If files.count>0 Then 
	For Each file in files
		Response.Write("<a href=""../Documents/Application/" & AppID & "/" & file.Name & _
			""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
	Next
Else
	Response.Write("No files in Application folder<br />")
End If
If IsNull(GrantID) = False Then
	DocumentFolder = Application("DocumentRoot") & "\Grant\" & GrantID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(DocumentFolder) = False Then
		fso.CreateFolder(DocumentFolder)
	End If

	Set folder = fso.GetFolder(DocumentFolder)
	Set files = folder.Files
	If files.count>0 Then 
		For Each file in files
			Response.Write("<a href=""../Documents/Grant/" & GrantID & "/" & file.Name & _
				""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
		Next
	End If
End If
Response.Write("</td></tr>" & vbCrLf)

%>
<tr>
	<td colspan="2" style="text-align: center; "><input type="button" value="Save" onclick="submitForm();" />
	<input type="button" value="Close" onclick="window.close();" /></td>
</tr>
</table>
</form>

</div>

<%	If IsNull(RevisedAppID) = False Then %>
<table style="margin: auto;">
<tr>
	<td><a href="/CatalyticConverter/PrintApplication.asp?AppID=<%=AppID%>" target="DisplayApp" class="plainlink">Original Application</a></td>
	<td style="width: 400px; "></td>
	<td><a href="/CatalyticConverter/PrintNegotiation.asp?AppID=<%=AppID%>" target="DisplayApp" class="plainlink">Revised Application</a></td>
</tr>
</table>
<%	End If %>
<iframe name="DisplayApp" id="DisplayApp" style="width: 99%; height: 400px; margin: auto;" src="<%
If FiscalYear > 2021 Then
	Response.Write("../CatalyticConverter/PrintApplication.asp?AppID=" & AppID)
Else
	Response.Write("../" & ApplicationSchema & "/PrintApplication.asp?AppID=" & AppID)
End If
 %>"></iframe>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>

<script src="../includes/formchanges.js"></script>
<script type="text/javascript">
	var saving = false;
	var form = document.getElementById("AppAdmin");

	// form being updated
	form.onsubmit = function () { saving = true; };

	// form not saved warning
	/*
	window.onunload = function() {
		if (!saving) {
			var f = FormChanges(form);
			if (f.length > 0) 
			{
				if (window.confirm("Your form updates have not be saved. Do you wish to continue without saving?")) {
					return true;
				}
				else {
					submitForm();
					return false;
				}
			}
		}
	};*/

	// show changed messages
	function DetectChanges()
	{
		var f = FormChanges(form), msg = "";
		for (var e = 0, el = f.length; e < el; e++) msg += "\n" + f[e].id;
		alert((msg ? "Elements changed:" : "No changes made.") + msg);
	}

	// Save changes
	function SaveChanges()
	{
		var f = FormChanges(form), msg = "";
		for (var e = 0, el = f.length; e < el; e++) msg += f[e].id + "\n";
		document.AppAdmin.Changes.value = msg;
	}

	function setInterlocalReviewer() {
<% if IsNull(InterlocalAgreementConfirmedDate)=True Then %>
		if (document.AppAdmin.InterlocalAgreementConfirmedDate.value.length > 0) {
			document.AppAdmin.InterlocalAgreementConfirmedBy.value = <%=UserSystemID %>;
	}
<%	Else %>
		return;
<%	End If %>
	}
</script>
</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/getApplicationSchema.asp"-->
<% 
Function appMailTo(vAppID)
	Dim vrs, vsql, vto
	vto= ""
	vsql = "SELECT * FROM Application.vwMailTo WHERE AppID=" & prepIntegerSQL(vAppID)
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		vto = AddToList(vto, vrs.Fields("AO"))
		vto = AddToList(vto, vrs.Fields("PD"))
		vto = AddToList(vto, vrs.Fields("PM"))
		vto = AddToList(vto, vrs.Fields("FO"))
		vto = AddToList(vto, vrs.Fields("PAC"))
		vto = AddToList(vto, vrs.Fields("FAC"))
		vto = AddToList(vto, "grantsMVCPA@txdmv.gov")
		vto = AddToList(vto, "Bryan.Wilson@txdmv.gov")
	End If
	appMailTo = "mailto:" & vto & "?subject=" & vrs.Fields("ProgramName") & " MVCPA Grant Application"
End Function 

Function AddToList(vList, vAdd)
	If Len(vAdd)>0 Then
		If Len(vList)>0 And InStr(vList, vAdd)>0 Then
			'AddToList = vList & "; " & vAdd
			AddToList = vList
		ElseIf Len(vList)>0 Then
			AddToList = vList & "; " & vAdd
		Else
			AddToList = vAdd
		End If
	Else
		AddToList = vList
	End If
End Function
%>