<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit
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

Dim MonitorID, FiscalYear, GranteeID, GranteeName, GrantID, YearsReviewedStart, YearsReviewedEnd, _
	DateOfNotice, InformationOrFilesRequested, RequestedInformationReceivedDate, _
	StartDate, EndDate, ExitInterview, DataCollectionCompleteDate, DraftReportToGranteeDate, _
	GranteeResponseToDraftDueDate, GranteeResponseToDraftReceivedDate, FinalReportCompleteDate, ReportReceivedDate, ManagementLetterReceivedDate, _
	MVCPAFundsTested, MVCPAFundsTestedFinding, MVCPAStaffReviewDate, _
	DeskReview, SiteVisit, MonitoringVisit, CAFR, ExternalAudit, OtherStateAgencyAudit, OtherAudit, OtherAuditDescription, _
	SubgranteeReview, ProgramReview, FiscalReview, SpecialOrTargetReview, SpecialOrTargetReviewText, _
	OtherAgenciesOnVisit, ActionPlanRequired, ActionPlanDueDate, ActionPlanFollowupDate, _
	ActionPlanCompleteDate, RiskLevelAssigned, CompletionClosedDate, _
	UpdateID, UpdateName, UpdateTimestamp

If Len(Request.Form("MonitorID"))>0 Then
	MonitorID = Request.Form("MonitorID")
ElseIf Len(Request.Querystring("MonitorID")) > 0 Then
	MonitorID = Request.QueryString("MonitorID")
Else
	MonitorID = -1
End If
If IsNumeric(MonitorID) Then
	MonitorID = CInt(MonitorID)
Else
	Response.Write("Error: Invalid Grant Monitoring ID.")
	SendMessage "Error: Invalid Grant Monitoring ID."
	Response.End
End If

If MonitorID>0 Then
	sql = "SELECT MonitorID, FiscalYear, GranteeID FROM Monitor.Main WHERE MonitorID=" & prepIntegerSQL(MonitorID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error: Grant Monitoring ID not found.")
		SendMessage "Error: Grant Monitoring ID not found."
		Response.End
	Else
		MonitorID = rs.Fields("MonitorID")
		FiscalYear = rs.Fields("FiscalYear")
		GranteeID= rs.Fields("GranteeID")
	End If
Else
	If Len(Request.Form("FiscalYear"))>0 Then
		FiscalYear = Request.Form("FiscalYear")
	ElseIf Len(Request.Querystring("FiscalYear")) > 0 Then
		FiscalYear = Request.QueryString("FiscalYear")
	Else
		FiscalYear = Session("FiscalYear")
	End If
	If IsNumeric(FiscalYear) Then
		FiscalYear = CInt(FiscalYear)
	Else
		Response.Write("Error: Invalid FiscalYear.")
		SendMessage "Error: Invalid FiscalYear."
		Response.End
	End If
	If Len(Request.Form("GranteeID"))>0 Then
		GranteeID = Request.Form("GranteeID")
	ElseIf Len(Request.Querystring("GranteeID")) > 0 Then
		GranteeID = Request.QueryString("GranteeID")
	Else
		GranteeID = Session("GranteeID")
	End If
	If IsNumeric(GranteeID) Then
		GranteeID = CInt(GranteeID)
	Else
		Response.Write("Error: Invalid GranteeID.")
		SendMessage "Error: Invalid GranteeID."
		Response.End
	End If
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Monitoring: Site visits and compliance monitoring.</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm()
	{
		if (validateForm() == true) {
			saving = true;
			SaveChanges();
			// Values in multi-selects must be selected to be submitted!
			for (i = 0; i < GrantMonitoring.Participants.length; i++) {
				GrantMonitoring.Participants.options[i].selected = true;
			}
			form.submit();
		}
	}

	function validateForm()
	{
		if (document.GrantMonitoring.DeskReview.checked == false &&
			document.GrantMonitoring.SiteVisit.checked == false &&
			document.GrantMonitoring.MonitoringVisit.checked == false &&
			document.GrantMonitoring.CAFR.checked == false &&
			document.GrantMonitoring.ExternalAudit.checked == false &&
			document.GrantMonitoring.OtherStateAgencyAudit.checked == false &&
			document.GrantMonitoring.OtherAudit.checked == false &&
			document.GrantMonitoring.SubgranteeReview.checked == false) {
			alert("You must choose one of the 'Type of site-visit / monitoring visit / desk review / or audit' options to save a record.");
			return false;
		}
		if ((document.GrantMonitoring.ExternalAudit.checked == true ||
			document.GrantMonitoring.OtherStateAgencyAudit.checked == true ||
			document.GrantMonitoring.OtherAudit.checked == true) &&
			document.GrantMonitoring.OtherAuditDescription.value.length == 0) {
			alert("The name of other agency or outside firm should be entered for the audit");
			document.GrantMonitoring.OtherAuditDescription.focus();
			return false;
		}
		return true;
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

	function AddParticipant()
	{
		GrantMonitoring.ParticipantsChanged.value = "1";
		if (GrantMonitoring.AddParticipants.selectedIndex > 0) {
			GrantMonitoring.Participants.options[GrantMonitoring.Participants.length] =
				new Option(GrantMonitoring.AddParticipants.options[GrantMonitoring.AddParticipants.selectedIndex].text, GrantMonitoring.AddParticipants[GrantMonitoring.AddParticipants.selectedIndex].value);
			GrantMonitoring.AddParticipants.remove(GrantMonitoring.AddParticipants.selectedIndex);
		}
	}


	function removeParticipant()
	{
		GrantMonitoring.ParticipantsChanged.value = "1";
		for (i = 0; i < GrantMonitoring.Participants.length; i++) {
			if (GrantMonitoring.Participants.options[i].selected) {
				GrantMonitoring.AddParticipants.options[GrantMonitoring.AddParticipants.length] =
					new Option(GrantMonitoring.Participants.options[i].text, GrantMonitoring.Participants.options[i].value);
				GrantMonitoring.Participants.remove(i);
				i--;
			}
		}
		GrantMonitoring.AddParticipants.selectedIndex = 0;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
<style>
	hr {
		border-top: 1px solid #bbbbbb;
	}
</style>
</head>
<body style="width: 840px; text-align: center">
<h1>Grant Monitoring</h1>
<form name="Selection" id="Selection" method="post" action="Monitor.asp">
<table style="margin: auto">
<tr><th colspan="2">Select Record</th></tr>
<tr>
	<td style="text-align: right; ">Fiscal Year: </td>
	<td style="text-align: left; "><select name="FiscalYear" onchange="document.Selection.submit();">
		<option value="0">Select</option>
<%
For i = 2017 to Application("CurrentFiscalYear")+1
	Response.Write(vbTab & vbTab & SelectOption(i, i, FiscalYear))
Next
%>
	</select></td>
</tr>
<tr>
	<td style="text-align: right; ">Grantee: </td>
	<td style="text-align: left; "><select name="GranteeID" onchange="<% 
	If FiscalYear>0 And GranteeID>0 Then
		Response.Write("document.Selection.MonitorID.selectedIndex=0; ")
	End If
	%>document.Selection.submit();">
		<option value="0">Select Grantee</option>
<%
sql = "SELECT GranteeID, REPLACE(GranteeName, 'City of ','') AS GranteeName " & vbCrLF & _
	"FROM Grantees " & vbCrLf & _
	"ORDER BY 2 "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>")
	Response.Flush
End If
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write(vbTab & vbTab & SelectOption(rs.Fields("GranteeID"), rs.Fields("GranteeName"), GranteeID))
	rs.MoveNext
Wend
%>	</select></td>
</tr>
<%
If FiscalYear>0 And GranteeID>0 Then
	sql = "SELECT MonitorID, A.GranteeID, REPLACE(A.GranteeName, 'City of ','') AS GranteeName, " & vbCrLf & _
		"	CASE WHEN StartDate IS NOT NULL AND EndDate IS NOT NULL THEN CONVERT(VARCHAR,StartDate, 101) + ' - ' + CONVERT(VARCHAR,EndDate,101) " & vbCrLf & _
		"		WHEN StartDate IS NOT NULL THEN CONVERT(VARCHAR,StartDate, 101) ELSE '' END + ' ' + " & vbCrLf & _
		"		CASE WHEN DeskReview=1 THEN 'Desk Audit ' " & vbCrLf & _
		"			WHEN SiteVisit=1 THEN 'Site Visit ' " & vbCrLF & _
		"			WHEN MonitoringVisit=1 THEN 'Monitoring Visit ' " & vbCrLF & _
		"			WHEN CAFR=1 THEN 'CAFR ' " & vbCrLf & _
		"			WHEN ExternalAudit=1 THEN 'ExternalAudit ' " & vbCrLf & _
		"			WHEN OtherStateAgencyAudit=1 THEN 'Other State Agency Audit ' " & vbCrLf & _
		"			WHEN OTherAudit=1 THEN 'Other Audit ' END + " & vbCrLf & _
		"			'(MonitorID=' + CAST(MonitorID AS VARCHAR) + ')' AS Description " & vbCrLf & _
		"FROM Grantees AS A " & vbCrLf & _
		"JOIN Monitor.Main AS B ON B.GranteeID=A.GranteeID AND FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"WHERE A.GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf & _
		"ORDER BY 4 "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
%>
<tr>
	<tr>
	<td style="text-align: right; ">Dates: </td>
	<td style="text-align: left; "><select name="MonitorID" onchange="document.Selection.submit();">
		<option value="-1">Select record</option>
<%
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(vbTab & vbTab & SelectOption(rs.Fields("MonitorID"), rs.Fields("Description"), MonitorID) & vbCrLf)
		rs.MoveNext
	Wend
	'If MonitorID > -1 Then
		Response.Write("<option value=""0"" " & Selected(MonitorID, 0) & ">New Record</option>" & vbCrLf)
	'End If
%>
		</select></td>
</tr>
<%
End If
%>
</table>
</form>
<%	
If MonitorID>-1 Then
	If MonitorID = 0 Then
		sql = "SELECT 0 AS MonitorID, " & prepIntegerSQL(FiscalYear) & " AS FiscalYear, GranteeID, GranteeName " & vbCrLf & _
			"FROM Grantees " & vbCrLf & _
			"WHERE GranteeID=" & prepIntegerSQL(GranteeID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>")
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		If rs.EOF = False Then	
			MonitorID = rs.Fields("MonitorID")
			If IsNull(rs.Fields("FiscalYear")) = False Then
				FiscalYear = rs.Fields("FiscalYear")
			End If
			GranteeID = rs.Fields("GranteeID")
			GranteeName = rs.Fields("GranteeName")
		Else
			Response.Write("Error: Error retrieving Grantee Record.")
			SendMessage "Error: Error retrieving Grantee Record."
			Response.End
		End If
	Else
		sql = "SELECT * " & vbCrLf & _
			"FROM Monitor.vwMain " & vbCrLf & _
			"WHERE MonitorID=" & prepIntegerSQL(MonitorID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>")
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		If rs.EOF = False Then	
			MonitorID = rs.Fields("MonitorID")
			If IsNull(rs.Fields("FiscalYear")) = False Then
				FiscalYear = rs.Fields("FiscalYear")
			End If
			GranteeID = rs.Fields("GranteeID")
			GranteeName = rs.Fields("GranteeName")
			YearsReviewedStart = rs.Fields("YearsReviewedStart")
			YearsReviewedEnd = rs.Fields("YearsReviewedEnd")
			DateOfNotice = rs.Fields("DateOfNotice")
			InformationOrFilesRequested = rs.Fields("InformationOrFilesRequested")
			RequestedInformationReceivedDate = rs.Fields("RequestedInformationReceivedDate")
			StartDate = rs.Fields("StartDate")
			EndDate = rs.Fields("EndDate") 
			ExitInterview = rs.Fields("ExitInterview")
			DataCollectionCompleteDate = rs.Fields("DataCollectionCompleteDate")
			DraftReportToGranteeDate = rs.Fields("DraftReportToGranteeDate")
			GranteeResponseToDraftDueDate = rs.Fields("GranteeResponseToDraftDueDate")
			GranteeResponseToDraftReceivedDate = rs.Fields("GranteeResponseToDraftReceivedDate")
			FinalReportCompleteDate = rs.Fields("FinalReportCompleteDate")
			ReportReceivedDate = rs.Fields("ReportReceivedDate")
			ManagementLetterReceivedDate = rs.Fields("ManagementLetterReceivedDate")
			MVCPAFundsTested = rs.Fields("MVCPAFundsTested")
			MVCPAFundsTestedFinding = rs.Fields("MVCPAFundsTestedFinding")
			MVCPAStaffReviewDate = rs.Fields("MVCPAStaffReviewDate")
			DeskReview = rs.Fields("DeskReview")
			SiteVisit = rs.Fields("SiteVisit")
			MonitoringVisit = rs.Fields("MonitoringVisit")
			CAFR = rs.Fields("CAFR")
			ExternalAudit = rs.Fields("ExternalAudit")
			OtherStateAgencyAudit = rs.Fields("OtherStateAgencyAudit")
			OtherAudit = rs.Fields("OtherAudit")
			OtherAuditDescription = rs.Fields("OtherAuditDescription")
			SubgranteeReview = rs.Fields("SubgranteeReview")
			ProgramReview = rs.Fields("ProgramReview")
			FiscalReview = rs.Fields("FiscalReview")
			SpecialOrTargetReview = rs.Fields("SpecialOrTargetReview")
			SpecialOrTargetReviewText = rs.Fields("SpecialOrTargetReviewText")
			OtherAgenciesOnVisit = rs.Fields("OtherAgenciesOnVisit")
			ActionPlanRequired = rs.Fields("ActionPlanRequired")
			ActionPlanDueDate = rs.Fields("ActionPlanDueDate")
			ActionPlanFollowupDate = rs.Fields("ActionPlanFollowupDate")
			ActionPlanCompleteDate = rs.Fields("ActionPlanCompleteDate")
			RiskLevelAssigned = rs.Fields("RiskLevelAssigned")
			CompletionClosedDate = rs.Fields("CompletionClosedDate")
			UpdateID = rs.Fields("UpdateID")
			UpdateName = rs.Fields("UpdateName")
			UpdateTimestamp = rs.Fields("UpdateTimestamp")
		End If
	End If
	If Debug = True Then
		Response.Write("<pre>Year:" & YearsReviewedStart & ", " & YearsReviewedEnd & "</pre>")
	End If
	If IsNull(YearsReviewedStart) Then
		YearsReviewedStart = FiscalYear
	End If
	If IsNull(YearsReviewedEnd) Then
		YearsReviewedEnd = FiscalYear
	End If
	If Debug = True Then
		Response.Write("<pre>Year:" & YearsReviewedStart & ", " & YearsReviewedEnd & "</pre>")
	End If

	IF MonitorID = 0 Then
		PermitEdit = True
	ElseIf IsNull(CompletionClosedDate) = True Then
		PermitEdit = True
	ElseIf MVCPAAdministrator = True Then
		PermitEdit = True
	Else
		PermitEdit = False
	End If
	'PermitEdit = True ' For testing.
%>
<br />
<form name="GrantMonitoring" id="GrantMonitoring" method="post" action="MonitorSubmit.asp" onsubmit="return submitForm();">
<%
Response.Write(HiddenField("FiscalYear", FiscalYear))
Response.Write(HiddenField("MonitorID", MonitorID))
Response.Write(HiddenField("GranteeID", GranteeID))
Response.Write(HiddenField("ParticipantsChanged", 0))
Response.Write(HiddenField("Changes",""))
%>
<table style="margin: auto; text-align: left;">
<tr><th colspan="2">Edit Record</th></tr>
<tr>
	<td>Grant Monitoring ID:</td>
	<td><%
	If MonitorID=0 Then 
		Response.Write("New")
	Else
		Response.Write(MonitorID)
	End If
	%></td>
</tr>
<tr>
	<td>Grantee ID:</td>
	<td><%=GranteeID %></td>
</tr>
<tr>
	<td>Grantee Name:</td>
	<td><%=GranteeName %></td>
</tr>
<tr>
	<td>Fiscal Year:</td>
	<td><%=FiscalYear %></td>
</tr>
<tr>
	<td>If multiple years reviewed, indicate first and last year</td>
	<td>First Year: <select name="YearsReviewedStart" id="YearsReviewedStart"> 
		<option value="0">Select</option>
<%
For i = 2017 to Application("CurrentFiscalYear")+1
	Response.Write(SelectOption(i, i, YearsReviewedStart))
Next
%>
	    </select>&nbsp;&nbsp;Last Year:
	<select name="YearsReviewedEnd" id="YearsReviewedEnd">
	<option value="0">Select</option>
<%
For i = 2017 to Application("CurrentFiscalYear")+1
	Response.Write(SelectOption(i, i, YearsReviewedEnd))
Next
%>
	</select></td>
</tr>
<tr style="vertical-align: top; ">
	<td style="white-space: nowrap; ">Type of site-visit / monitoring visit / desk review / or audit:</td>
	<td style="white-space: nowrap; "><%=CheckBoxField("SiteVisit", SiteVisit) %>Site Visit<br />
	<%=CheckBoxField("MonitoringVisit", MonitoringVisit) %>Monitoring Visit<br />
	<%=CheckBoxField("DeskReview", DeskReview) %>Desk Review<br />
	<%=CheckBoxField("CAFR", CAFR) %>Comprehensive Financial Audit (CAFR)<br />
	<%=CheckBoxField("ExternalAudit", ExternalAudit) %>External / Independent Audit<br />
	<%=CheckBoxField("OtherStateAgencyAudit", OtherStateAgencyAudit) %>Other State Agency Audit<br />
	<%=CheckBoxField("OtherAudit", OtherAudit) %>Other Audit<br />
	<%=CheckBoxField("SubgranteeReview", SubgranteeReview) %>Subgrantee Review<br />
	</td>
</tr>
<tr style="vertical-align: top;">
	<td></td>
	<td>Name of other agency or outside firm conducting audit:<br />
	<%=TextField("OtherAuditDescription", OtherAuditDescription, 60, 250, PermitEdit, "") %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>Areas of Review:</td>
	<td><%=CheckBoxField("ProgramReview", ProgramReview) %>Program Review<br />
	<%=CheckBoxField("FiscalReview", FiscalReview) %>Fiscal Review<br />
	<%=CheckBoxField("SpecialOrTargetReview", SpecialOrTargetReview) %>Special Or Target Review (Specify:)<br />
	<%=TextField("SpecialOrTargetReviewText", SpecialOrTargetReviewText, 60, 250, PermitEdit, "") %></td>
</tr>
<tr><td colspan="2"><hr /></td></tr>
<tr style="vertical-align: top; ">
	<td>Date of Notice Provided of on-site visit to Grantee:</td>
	<td><%= DateField2("DateOfNotice", DateOfNotice, "01/01/2017", date(), PermitEdit) %></td>
</tr>
<tr>
	<td>Information or files requested: </td>
	<td><input type="radio" name="InformationOrFilesRequested" id="InformationOrFilesRequested1" value="1" <%=checked(InformationOrFilesRequested,1) %> />Yes&nbsp;&nbsp;
	<input type="radio" name="InformationOrFilesRequested" id="InformationOrFilesRequested2" value="2" <%=checked(InformationOrFilesRequested,2) %> />No&nbsp;&nbsp;
	Date Received: <%= DateField2("RequestedInformationReceivedDate", RequestedInformationReceivedDate, "01/01/2017", date(), PermitEdit) %>
	</td>
</tr>
<tr style="vertical-align: top;">
	<td style="white-space: nowrap; ">Date(s) of site-visit / monitoring visit /desk review:</td>
	<td style="white-space: nowrap; ">Start: <%= DateField2("StartDate", StartDate, "01/01/2017", date(), PermitEdit) %>- End: <%=DateField2("EndDate", EndDate, "01/01/2017", date(), PermitEdit) %><br />
	Exit Interview? <input type="radio" name="ExitInterview" id="ExitInterview1" value="1" <%=checked(ExitInterview,1) %> />Yes&nbsp;&nbsp;
	<input type="radio" name="ExitInterview" id="ExitInterview2" value="2" <%=checked(ExitInterview,2) %> />No</td>
</tr>
<tr style="vertical-align: top; ">
	<td>Staff members on visit: </td>
	<td>
<%
	If PermitEdit = True Then
%>
<select name="AddParticipants" id="AddParticipants" style="width: 200px;">
		<option value="0">Select Participant</option>
<%
	sql = "SELECT A.SystemID, A.Name " & vbCrLf & _
		"FROM [System].Users AS A" & vbCrLf & _
		"LEFT JOIN Monitor.Participants AS B ON B.SystemID=A.SystemID AND MonitorID=" & prepIntegerSQL(MonitorID) & " " & vbCrLf & _
		"WHERE A.MVCPAStaff=1 AND B.MonitorID IS NULL " & vbCrLf & _
		"ORDER BY A.LastName, A.FirstName "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(vbTab & "<option value=""" & rs.Fields("SystemID") & """>" & rs.Fields("Name") & "</option>" & vbCrLf)
		rs.MoveNext
	Wend
%>
	           </select>
			   <input type="button" value="Add" onclick="AddParticipant();" style="width: 70px; " 
				title="Select a name from dropdown and click on this button to add name to list." />
				<input type="button" value="Remove" onclick="removeParticipant();" style="width: 70px; vertical-align: top; " 
				title="select a name on list and click on this button to remove from list." /><br />
<%
End If
%>
			<select name="Participants" id="Participants" multiple size="4" style="width: 200px; ">
<%
	sql = "SELECT A.SystemID, B.Name " & vbCrLf & _
		"FROM Monitor.Participants AS A" & vbCrLF & _
		"JOIN [System].Users AS B ON B.SystemID=A.SystemID " & vbCrLf & _
		"WHERE A.MonitorID=" & prepIntegerSQL(MonitorID) & " " & vbCrLf & _
		"ORDER BY B.LastName, B.FirstName "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(vbTab & vbTab & "<option value=""" & rs.Fields("SystemID") & """>" & rs.Fields("Name") & "</option>" & vbCrLf)
		rs.MoveNext
	Wend

%>
		</select>
	</td>
</tr>
<tr style="vertical-align: top;">
	<td>Other State Agencies on visit:</td>
	<td><%=TextField("OtherAgenciesOnVisit", OtherAgenciesOnVisit, 60, 250, PermitEdit, "") %></td>
</tr>
<tr>
	<td>Date Data Collection Complete:</td>
	<td><%=DateField2("DataCollectionCompleteDate", DataCollectionCompleteDate, "01/01/2017", date(), PermitEdit) %></td>
</tr>
<tr>
	<td>Date Draft Report Provided to Grantee:</td>
	<td><%=DateField2("DraftReportToGranteeDate", DraftReportToGranteeDate, "01/01/2017", date(), PermitEdit) %></td>
</tr>
<tr>
	<td>Due Date for Grantee Response to Draft:</td>
	<td><%=DateField2("GranteeResponseToDraftDueDate", GranteeResponseToDraftDueDate, "01/01/2017", DateAdd("d",31,date()), PermitEdit) %> 
	Date Received: <%=DateField2("GranteeResponseToDraftReceivedDate", GranteeResponseToDraftReceivedDate, "01/01/2017", DateAdd("d",31,date()), PermitEdit) %></td>
</tr>
<tr>
	<td>Date Final Report Complete:</td>
	<td><%=DateField2("FinalReportCompleteDate", FinalReportCompleteDate, "01/01/2017", date(), PermitEdit) %></td>
</tr>
<tr><td colspan="2"><hr /></td></tr>
<tr>
	<td>Action Plan Required:</td>
	<td><input type="radio" name="ActionPlanRequired" id="ActionPlanRequired1" value="1" <%=checked(ActionPlanRequired,1) %> />Yes&nbsp;&nbsp;
	<input type="radio" name="ActionPlanRequired" id="ActionPlanRequired2" value="2" <%=checked(ActionPlanRequired,2) %> />No</td>
</tr>
<tr>
	<td>Action Plan Due Date:</td>
	<td><%=DateField("ActionPlanDueDate", ActionPlanDueDate, PermitEdit) %></td>
</tr>
<tr>
	<td>Date of Action Plan Follow-up:</td>
	<td><%=DateField("ActionPlanFollowupDate", ActionPlanFollowupDate, PermitEdit) %></td>
</tr>
<tr>
	<td>Date Action Plan Complete:</td>
	<td><%=DateField("ActionPlanCompleteDate", ActionPlanCompleteDate, PermitEdit) %></td>
</tr>
<tr>
	<td>Based on review and responses, Risk Level Assigned:</td>
	<td><input type="radio" name="RiskLevelAssigned" id="RiskLevelAssigned1" value="1" <%=checked(RiskLevelAssigned,1) %> />1 - high&nbsp;&nbsp;
	<input type="radio" name="RiskLevelAssigned" id="RiskLevelAssigned2" value="2" <%=checked(RiskLevelAssigned,2) %> />2 - medium&nbsp;&nbsp;
	<input type="radio" name="RiskLevelAssigned" id="RiskLevelAssigned3" value="3" <%=checked(RiskLevelAssigned,3) %> />3 - low</td>
</tr>
<tr><td>Notes</td><td></td></tr><%
	sql = "SELECT A.NoteID, A.MonitorID, A.Note, A.UpdateID, B.Name AS UpdateName, A.UpdateTimestamp " & vbCrLF & _
		"FROM Monitor.Notes AS A " & vbCrLF & _
		"LEFT JOIN [System].Users AS B ON B.SystemID=A.UpdateID " & vbCrLf & _
		"WHERE A.MonitorID=" & prepIntegerSQL(MonitorID) & " " & vbCrLf & _
		"ORDER BY NoteID "
	Set rs = Con.Execute(sql)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	While rs.EOF = False
		Response.Write("<tr><td style=""text-align: right; vertical-align: top; font-style: italic; "">" & rs.Fields("UpdateName") & ", " & rs.Fields("UpdateTimestamp") & "</td><td>" & rs.Fields("Note") & "</td></tr>")
		rs.MoveNext
	Wend
%>
<tr>
	<td style="text-align: right; vertical-align: top; font-style: italic; ">New Note:</td>
	<td><%=TextArea2("Note", "", 3, 430, 2040, PermitEdit, "") %></td>
</tr>
<tr>
	<td>Date Related Tasks Completed and this Item is Closed:</td>
	<td><%=DateField2("CompletionClosedDate", CompletionClosedDate, "01/01/2017", date(), PermitEdit) %></td>
</tr>

<tr><td colspan="2"><hr /></td></tr>
<tr style="vertical-align: top; ">
	<td>Date Outside Report Received:</td>
	<td><%= DateField2("ReportReceivedDate", ReportReceivedDate, "01/01/2017", date(), PermitEdit) %></td>
</tr>
<tr style="vertical-align: top; ">
	<td>If Management Letter Received, indicate Date:</td>
	<td><%= DateField2("ManagementLetterReceivedDate", ManagementLetterReceivedDate, "01/01/2017", date(), PermitEdit) %></td>
</tr>
<tr>
	<td>MVCPA Grant Funds Tested:</td>
	<td><input type="radio" name="MVCPAFundsTested" id="MVCPAFundsTested1" value="1" <%=checked(MVCPAFundsTested,1) %> />Yes&nbsp;&nbsp;
	<input type="radio" name="MVCPAFundsTested" id="MVCPAFundsTested2" value="2" <%=checked(MVCPAFundsTested,2) %> />No</td>
</tr>
<tr style="vertical-align: top; ">
	<td>Was there a finding?:</td>
	<td><%=TextArea2("MVCPAFundsTestedFinding", MVCPAFundsTestedFinding, 4, 430, 2040, PermitEdit, "") %>
	</td>
</tr>
<tr style="vertical-align: top; ">
	<td>Date of MVCPA Staff Review:</td>
	<td><%= DateField2("MVCPAStaffReviewDate", MVCPAStaffReviewDate, "01/01/2017", date(), PermitEdit) %></td>
</tr>

<%	If MonitorID>0 Then %>
<tr><td colspan="2"><hr /></td></tr>
<tr style="vertical-align: top; ">
	<td>Documents <a href="../Upload/Upload.asp?fid=12&MonitorID=<%=MonitorID %>" target="_blank">Upload</a></td>
		<td><%
Dim folder, fso, rsFSO
folder = Application("DocumentRoot") & "Monitor\" & MonitorID & "\"
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(folder) Then
	Response.Write("<table>" & vbCrLf)
	Set rsFSO = kc_fsoFiles(folder, "_")

	While Not rsFSO.EOF
		Response.Write("<tr><td><a href=""../Documents/Monitor/" & MonitorID & "/" & rsFSO("Name") & """ target=""_blank"">" & rsFSO("Name").Value & "</a></td><td>" & rsFSO.Fields("DateLastModified") & "</td><td>" & rsFSO.Fields("Type").Value & "</td></tr>")
		rsFso.MoveNext()
	Wend
	Response.Write("</table>" & vbCrLf)
  
	'finally, close out the recordset
	rsFSO.close()
	Set rsFSO = Nothing
End If
%></td>
</tr>
<tr>
	<td colspan="2" style="text-align: center;"><%="Last Update by " & UpdateName & " at " & UpdateTimestamp %></td>
</tr>
<%	End If %>
<tr>
	<td colspan="2" style="text-align: center">
<%	If PermitEdit = True Then %>
	<input type="button" name="submitbutton" value="Save" onclick="return submitForm('submit');" />&nbsp;&nbsp;
<%	End If %>
	<input type="button" name="close" value="Close" onclick="window.close();" />
	<input type="button" name="search" value="Search" onclick="location.href = 'Search.asp';" />
	</td>
</tr>
</table>
</form>
<script src="../includes/formchanges.js"></script>
<script type="text/javascript">
	var saving = false;
	var form = document.getElementById("GrantMonitoring");
	// form being updated
	form.onsubmit = function () { saving = true; };

	// form not saved warning
	/*window.onunload = function() {
		if (!saving) {
			var f = FormChanges(form);
			if (f.length > 0) 
			{
				if (window.confirm("Your form updates have not be saved. Do you wish to continue without saving?"))
					return true;
				else
					return false;
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
		document.GrantMonitoring.Changes.value = msg;
	}

</script>
<%	End If %>
</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<%
'**********
'kc_fsoFiles
'Purpose:
' 1. To create a recordset using the FSO object and ADODB
' 2. Allows you to exclude files from the recordset if needed
'Use:
' 1. Call the function when you're ready to open the recordset
' and output it onto the page.
' example:
' Dim rsFSO, strPath
' strPath = Server.MapPath("\PlayGround\FSO\Stuff\")
' Set rsFSO = kc_fsoFiles(strPath, "_")
' The "_" will exclude all files beginning with 
' an underscore 
'**********
Function kc_fsoFiles(theFolder, Exclude)
Dim rsFSO, objFSO, objFolder, File
  Const adInteger = 3
  Const adDate = 7
  Const adVarChar = 200
  
  'create an ADODB.Recordset and call it rsFSO
  Set rsFSO = Server.CreateObject("ADODB.Recordset")
  
  'Open the FSO object
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  
  'go get the folder to output it's contents
  Set objFolder = objFSO.GetFolder(theFolder)
  
  'Now get rid of the objFSO since we're done with it.
  Set objFSO = Nothing
  
  'create the various rows of the recordset
  With rsFSO.Fields
    .Append "Name", adVarChar, 200
    .Append "Type", adVarChar, 200
    .Append "DateCreated", adDate
    .Append "DateLastAccessed", adDate
    .Append "DateLastModified", adDate
    .Append "Size", adInteger
    .Append "TotalFileCount", adInteger
  End With
  rsFSO.Open()
	
  'Now let's find all the files in the folder
  For Each File In objFolder.Files
	
    'hide any file that begins with the character to exclude
    If (Left(File.Name, 1)) <> Exclude Then 
      rsFSO.AddNew
      rsFSO("Name") = File.Name
      rsFSO("Type") = File.Type
      rsFSO("DateCreated") = File.DateCreated
      rsFSO("DateLastAccessed") = File.DateLastAccessed
      rsFSO("DateLastModified") = File.DateLastModified
      rsFSO("Size") = File.Size
      rsFSO.Update
    End If

  Next
	
  'And finally, let's declare how we want the files 
  'sorted on the page. In this example, we are sorting 
  'by File Type in descending order,
  'then by Name in an ascending order.
  rsFSO.Sort = "Name ASC, DateCreated ASC "

  'Now get out of the objFolder since we're done with it.
  Set objFolder = Nothing

  'now make sure we are at the beginning of the recordset
  'not necessarily needed, but let's do it just to be sure.
  If rsFSO.BOF = False Then
	rsFSO.MoveFirst()
  End If
  Set kc_fsoFiles = rsFSO
	
End Function
%>