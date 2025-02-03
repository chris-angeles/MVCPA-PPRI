<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, LastCategory, PermitEdit, Submitted, AllowUpload, DocumentFolder, ApplicationSchema, RoundCurrency,  _
	fso, files, folder, file, _
	AppID, GranteeID, GranteeName, FiscalYear, ProgramName, ORI, ORIAgency, MultiAgencyGrant, _
	AuthorizedOfficialID, AuthorizedOfficial, AuthorizedOfficialTitle, _
	ProgramDirectorID, ProgramManagerID, _
	CoverageAreaDescription, StatewideCoverage, OtherCoverage, OtherCoverageText, _
	CashMatch, InKindMatch, GrandTotal, TotalMVCPAFunds, TotalCashMatch, TotalInkindMatch, PctMVCPA, PctCashMatch, _
	Section, Subsection, TaskForceStructureQuestion, TaskForceStructureResponse, _
	Colocation, MeetingsGranteeMethod, MeetingsGranteeFrequency, _
	MeetingsSubGranteeMethod, MeetingsSubGranteeFrequency, _
	MeetingsAllTFMethod, MeetingsAllTFFrequency, MeetingsDescription, _
	CommunicationGranteeMethod, CommunicationGranteeFrequency, _
	CommunicationSubGranteeMethod, CommunicationSubGranteeFrequency, _
	CommunicationAllTFMethod, CommunicationAllTFFrequency, CommunicationDescription, _
	CoverageAgencyMeetings, CoverageAgencyContacts, IntelligenceSharing, _
	OperationalCoordination, DirectOperatations, _ 
	SubmitID, SubmitName, SubmitTimestamp, OperationalPlanApprovalID, OperationalPlanApprovalName, OperationalPlanApprovalDate, _
	UpdateID, UpdateName, UpdateTimestamp
	
' Note: NegotiationLocked is only used at negotiation stage. But Variable loaded and referenced on application so code can be transfered.
Dim ProgramCategory(5)

debug = False
'ApplicationSchema = "Negotiation"
ApplicationSchema = "Application"
RoundCurrency = False

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
	Response.Write("Now=" & Now() & vbCrLf)
	Response.Write("</pre>" & vbCrLf)
End If

AppID = Request.QueryString("AppID")

If AppID="" Then
	Response.Write("Error: No AppID or GranteeID Specified")
	SendMessage "Error: No AppID or GranteeID Specified"
	Response.End
Else
	AppID=Cint(AppID)
End If

If AppID>0 Then 
	sql = "SELECT A.AppID, I.FiscalYear, G.GranteeID, G.GranteeName, A.ProgramName, " & vbCrLf & _
		"	G.ORI, ORI.Agency AS ORIAgency, G.OrganizationTypeID, OT.OrganizationType, " & vbCrLf & _
		"	AuthorizedOfficialID, AO.Name AS AuthorizedOfficial, AO.Title AS AuthorizedOfficialTitle,  " & vbCrLf & _
		"	G.ProgramDirectorID, G.ProgramManagerID, " & vbCrLf & _
		"	A.CoverageAreaDescription, A.StatewideCoverage, A.OtherCoverage, A.OtherCoverageText, " & vbCrLf & _
		"	ISNULL(B.TotalMVCPAFunds,0.0) AS TotalMVCPAFunds, " & vbCrLf & _
		"	ISNULL(B.TotalCashMatch,0.0) AS TotalCashMatch, " & vbCrLf & _
		"	ISNULL(B.GrandTotal,0.0) AS GrandTotal, " & vbCrLf & _
		"	ISNULL(B.TotalInKindMatch,0.0) AS TotalInKindMatch, " & vbCrLf & _
		"	TS.Section, TS.Subsection, TS.Question AS TaskForceStructureQuestion, ST.SectionText AS TaskForceSTructureResponse,  " & vbCrLf & _
		"	CAST(CASE WHEN AC.AgencyCount>1 THEN 1 ELSE 0 END AS BIT) AS MultiAgencyGrant " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".Main AS A ON A.AppID=I.AppID " & vbCrLf & _
		"LEFT JOIN Application.Admin AS L ON L.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN System.Users AS U ON U.SystemID=A.SubmitID " & vbCrLf & _
		"LEFT JOIN ( " & vbCrLf & _
		"	SELECT AppID, SUM(MVCPAFunds) AS TotalMVCPAFunds, SUM(CashMatch) AS TotalCashMatch, SUM(LineTotal) AS GrandTotal, SUM(InKindMatch) AS TotalInKindMatch " & vbCrLf & _
		"FROM " & ApplicationSchema & ".BudgetDetails " & vbCrLf & _
		"GROUP BY AppID) AS B ON B.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN ( " & vbCrLf & _
		"	SELECT AppID, SUM(CASE WHEN MatchTypeID=1 Then Amount ELSE NULL END) AS DetailCashMatch, " & vbCrLf & _
		"		SUM(CASE WHEN MatchTypeID=2 Then Amount ELSE NULL END) AS DetailInKindMatch,  " & vbCrLf & _
		"		SUM(Amount) AS DetailTotalMatch " & vbCrLf & _
		"	FROM " & ApplicationSchema & ".Matches " & vbCrLf & _
		"	GROUP BY AppID) AS M ON M.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".TargetAmounts AS N ON N.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN Lookup.OrganizationType AS OT ON OT.OrganizationTypeID=G.OrganizationTypeID " & vbCrLf & _
		"LEFT JOIN Lookup.ORI AS ORI ON ORI.ORI=G.ORI " & vbCrLf & _
		"LEFT JOIN [System].Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf & _
		"LEFT JOIN Lookup.TextSections AS TS ON TS.Version=2 AND TS.Section=1 AND TS.SubSection=2 " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".SectionText AS ST ON ST.TextSectionID=TS.TextSectionID AND ST.AppID=A.AppID" & vbCrLf & _
		"LEFT JOIN ( " & vbCrLf & _
		"	SELECT AppID, COUNT(*) AS AgencyCount " & vbCrLf & _
		"	FROM " & ApplicationSchema & ".ParticipatingAgencies " & vbCrLf & _
		"	GROUP BY AppID " & vbCrLf & _
		") AS AC ON AC.AppID = A.AppID " & vbCrLf & _
		"WHERE A.AppID=" & PrepIntegerSQL(AppID)
End If

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: Grantee and " & ApplicationSchema & " record not retrieved")
	SendMessage "Error: Grantee and " & ApplicationSchema & " record not retrieved"
	Response.End
Else
	AppID = rs.Fields("AppID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	ORI = rs.Fields("ORI")
	ORIAgency = rs.Fields("ORIAgency")
	AuthorizedOfficialID = rs.Fields("AuthorizedOfficialID")
	AuthorizedOfficial = rs.Fields("AuthorizedOfficial")
	AuthorizedOfficialTitle = rs.Fields("AuthorizedOfficialTitle")
	ProgramDirectorID = rs.Fields("ProgramDirectorID")
	ProgramManagerID = rs.Fields("ProgramManagerID")
	ProgramName = rs.Fields("ProgramName")
	CoverageAreaDescription = rs.Fields("CoverageAreaDescription")
	StatewideCoverage = rs.Fields("StatewideCoverage")
	OtherCoverage = rs.Fields("OtherCoverage")
	OtherCoverageText = rs.Fields("OtherCoverageText")
	TotalMVCPAFunds = rs.Fields("TotalMVCPAFunds")
	TotalCashMatch = rs.Fields("TotalCashMatch")
	TotalInkindMatch = rs.Fields("TotalInkindMatch")
	GrandTotal = rs.Fields("GrandTotal")
	Section = rs.Fields("Section")
	Subsection = rs.Fields("Subsection")
	TaskForceStructureQuestion = rs.Fields("TaskForceStructureQuestion")
	TaskForceStructureResponse = rs.Fields("TaskForceStructureResponse")
	MultiAgencyGrant = rs.Fields("MultiAgencyGrant")
	rs.Close()
End If

sql = "SELECT B.*, D.Name AS UpdateName, E.Name AS SubmitName, " & vbCrLf & _
	"	CAST(CASE WHEN B.SubmitID>0 THEN 1 ELSE 0 END AS BIT) AS Submitted, " & vbCrLF & _
	"	C.OperationalPlanApprovalDate, C.OperationalPlanApprovalID, F.Name as OperationalPlanApprovalName " & vbCrLf & _
	"FROM [Application].Main AS A " & vbCrLf & _
	"LEFT JOIN [Grants].OperationalPlan AS B ON B.AppID = A.AppID " & vbCrLf & _
	"LEFT JOIN [Application].Admin AS C ON C.AppID=B.AppID " & vbCrLF & _
	"LEFT JOIN [System].Users AS D ON D.SystemID = B.UpdateID " & vbCrLf & _
	"LEFT JOIN [System].Users AS E ON E.SystemID = B.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS F ON F.SystemID = C.OperationalPlanApprovalID " & vbCrLf & _
	"WHERE A.AppID=" & prepIntegerSQL(AppID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	'GrantID = rs.Fields("GrantID")
	Colocation = rs.Fields("Colocation") 
	MeetingsGranteeMethod = rs.Fields("MeetingsGranteeMethod") 
	MeetingsGranteeFrequency = rs.Fields("MeetingsGranteeFrequency") 
	MeetingsSubGranteeMethod = rs.Fields("MeetingsSubGranteeMethod") 
	MeetingsSubGranteeFrequency = rs.Fields("MeetingsSubGranteeFrequency") 
	MeetingsAllTFMethod = rs.Fields("MeetingsAllTFMethod") 
	MeetingsAllTFFrequency = rs.Fields("MeetingsAllTFFrequency") 
	MeetingsDescription = rs.Fields("MeetingsDescription") 
	CommunicationGranteeMethod = rs.Fields("CommunicationGranteeMethod") 
	CommunicationGranteeFrequency = rs.Fields("CommunicationGranteeFrequency") 
	CommunicationSubGranteeMethod = rs.Fields("CommunicationSubGranteeMethod") 
	CommunicationSubGranteeFrequency = rs.Fields("CommunicationSubGranteeFrequency") 
	CommunicationAllTFMethod = rs.Fields("CommunicationAllTFMethod") 
	CommunicationAllTFFrequency = rs.Fields("CommunicationAllTFFrequency")
	CommunicationDescription = rs.Fields("CommunicationDescription") 
	CoverageAgencyMeetings = rs.Fields("CoverageAgencyMeetings")
	CoverageAgencyContacts = rs.Fields("CoverageAgencyContacts")
	IntelligenceSharing = rs.Fields("IntelligenceSharing")
	OperationalCoordination = rs.Fields("OperationalCoordination")
	DirectOperatations = rs.Fields("DirectOperatations")
	SubmitID = rs.Fields("SubmitID")
	SubmitName = rs.Fields("SubmitName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	OperationalPlanApprovalID = rs.Fields("OperationalPlanApprovalID")
	OperationalPlanApprovalName = rs.Fields("OperationalPlanApprovalName")
	OperationalPlanApprovalDate = rs.Fields("OperationalPlanApprovalDate")
	UpdateID = rs.Fields("UpdateID")
	UpdateName = rs.Fields("UpdateName")
	UpdateTimestamp = rs.Fields("UpdateTimestamp")
	Submitted = rs.Fields("Submitted")
	rs.Close()
End If
' Start rounding dollar amounts as of 2020.
If FiscalYear>=2020 Then
	RoundCurrency = True
Else
	RoundCurrency = False
End If

PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, Submitted)
If MVCPARights = True Then
	AllowUpload = True
Else
	AllowUpload = PermitEdit
End If

DocumentFolder = Application("DocumentRoot") & "\Application\" & AppID & "\"

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Taskforce Multi-Agency Grant Operational Plan</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm(action)
	{
		document.Application.Button.value = action;

		if (action == "submit") {
			if (validateForm() == false)
				return false;
		}
		document.Application.Button.value = action;
		document.Application.submit();
	}

	function radioChecked(field)
	{
		return field.checked;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag"><%=GranteeName %> Taskforce Grant <%=ApplicationSchema %> for Fiscal Year <%=FiscalYear %></div>

<div class="widecontent">

<form name="Application" id="Application" method="post" action="OperationalPlanSubmit.asp" onsubmit="return validateForm()">
<%
Response.Write(HiddenField("GranteeID", GranteeID))
Response.Write(HiddenField("AppID", AppID))
Response.Write(HiddenField("FiscalYear", FiscalYear))
Response.Write(HiddenField("Button","save"))
Response.Write(HiddenField("Changes",""))
%>
<table style="width: 956px; ">

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<th colspan="2"><b>Multi-Agency Operational Plan</b></th>
</tr>

<tr><td colspan="2" style="text-align: center; "><%
If IsNull(UpdateID) Then
	Response.Write("No operational plan has been saved.<br/>")
Else
	Response.Write("Last save by " & UpdateName & " at " & UpdateTimestamp & "<br/>")
End If

If Submitted = True Then
	Response.Write("This plan was submitted by " & SubmitName & " at " & SubmitTimestamp & vbCrLf)
End If
%></td></tr>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<th colspan="2"><b>Grant Related Data from Application</b></th>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2"><b>Primary Agency / Grantee Legal Name:</b> <i><%=GranteeName %></i></td>
</tr>

<tr>
	<td colspan="2"><b>Organization ORI:</b> <i><%=ORI %>: <%=ORIAgency %></i></td>
</tr>

<tr>
	<td colspan="2"><b>Program Title:</b> <%=ProgramName %></td>
</tr>
<%
If MultiAgencyGrant = False Then
%>
<tr>
	<td><b>Multi-Agency Grant:</b></td>
	<td>The application shows only one particpating agency. If correct, a Multi-Agency Operational Plan is not required.</td>
</tr>
<%
End If
%>
<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<th colspan="2"><b>Taskforce Grant Participation and Coverage Area</b></th>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2"><b>Provide a General Description of the Participating and 
	Coverage Area of this Grant Application</b></td>
</tr>
<tr>
	<td></td>
	<td><%=CoverageAreaDescription %></td>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2"><table style="margin: auto;  border: 1px solid #dddddd; ">

	<tr>
		<td style="vertical-align: top; text-align: center"><b>Participating Agencies</b>
		<td style="vertical-align: top; text-align: center "><b>Coverage Agencies</b><br />
	</tr>
	<tr>
		<td style="vertical-align: top">

<%
	sql = "SELECT A.ORI, REPLACE(B.Agency,'&','&amp;') AS Agency" & vbCrLF & _
		"FROM " & ApplicationSchema & ".ParticipatingAgencies AS A" & vbCrLF & _
		"LEFT JOIN Lookup.ORI AS B ON B.ORI=A.ORI " & vbCrLf & _
		"WHERE A.AppID = " & prepIntegerSQL(AppID) & vbCrLF & _
		"ORDER BY A.ORI "
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(rs.Fields("ORI") & " " & rs.Fields("Agency") & "<br />" & vbCrLf)
		rs.MoveNext
	Wend
%></td>
	<td style="vertical-align: top">
<%
	sql = "SELECT A.ORI, REPLACE(B.Agency,'&','&amp;') AS Agency" & vbCrLF & _
		"FROM " & ApplicationSchema & ".CoverageAgencies AS A" & vbCrLF & _
		"LEFT JOIN Lookup.ORI AS B ON B.ORI=A.ORI " & vbCrLf & _
		"WHERE A.AppID = " & prepIntegerSQL(AppID) & vbCrLF & _
		"ORDER BY A.ORI "
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(rs.Fields("ORI") & " " & rs.Fields("Agency") & "<br />" & vbCrLf)
		rs.MoveNext
	Wend
%></td>
	</tr>
</table>
</td></tr></table>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Taskforce Governing, Organization and Command Structures</div>

<br />

<div style="width: 976px; text-align: left; "><b><%=Section %>.<%=SubSection %>&nbsp;<%=TaskForceStructureQuestion %></b><br />
	<%=TaskForceStructureResponse %>
</div>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Grant Budget Summary</div>

<br />

<table style="width: 896px; margin: auto">
<thead>
<tr style="vertical-align: bottom">
	<th>Budget Category</th>
	<th>MVCPA<br />Expenditures</th>
	<th>Cash<br />Match<br />Expenditures</th>
	<th>Total<br />Expenditures</th>
	<th>In-Kind<br />Match</th>
</tr>
</thead>
<tbody>
<%
sql = "SELECT ISNULL(A.BudgetCategoryID,99) AS BudgetCategoryID, ISNULL(A.BudgetCategory, 'Total') As BudgetCategory, " & vbCrLf & _
	"	SUM(LineTotal) AS LineTotal, SUM(MVCPAFunds) AS [MVCPAFunds], " & vbCrLf & _
	"	SUM(CashMatch) AS [CashMatch], SUM(InKindMatch) AS [InKindMatch] " & vbCrLf & _
	"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".BudgetDetails AS B ON A.BudgetCategoryID=B.BudgetCategoryID AND B.AppID=" & _
		prepIntegerSQL(AppID) & " " & vbCrLf & _
	"GROUP BY GROUPING SETS ((A.BudgetCategoryID,A.BudgetCategory),()) " & vbCrLf & _
	"ORDER BY ISNULL(A.BudgetCategoryID,99) "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(vbTab & "<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)  
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(rs.Fields("MVCPAFunds"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(rs.Fields("CashMatch"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(rs.Fields("LineTotal"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(rs.Fields("InKindMatch"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>")
		rs.MoveNext
	Wend
	If TotalMVCPAFunds>0 Then
		PctMVCPA = 100*TotalMVCPAFunds / GrandTotal
		PctCashMatch = 100*TotalCashMatch / TotalMVCPAFunds
		Response.Write("<tr><td style=""text-align: center;"">Cash Match Percentage</td><td style=""text-align: right; ""><!--" & prepNumberWeb(PctMVCPA, 2) & _
			"%--></td><td style=""text-align: right; "">" & prepNumberWeb(PctCashMatch, 2) & "%</td><td></td><td></td></tr>" & vbCrLf)
	End If
%>
</tbody>
</table>
<br />
<div style="width: 976px; text-align: center; "><a href="\<%=ApplicationSchema %>\TFGPrintApplication.asp?AppID=<%=AppID %>#BudgetDetail" target="_blank">Budget Detail</a></div>
<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Operational Plan New Data</div>
<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 1 Co-location</div>

<p><b>Are members of the taskforce co-located?</b><br />
	<%=RadioInputField("CoLocation", Colocation, 1) %>All of the time
	<%=RadioInputField("CoLocation", Colocation, 2) %>Occasionally
	<%=RadioInputField("CoLocation", Colocation, 3) %>Never
</p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 2 Grantee and Subgrantee Meetings</div>

<p><b>By what primary method do scheduled meetings occur and how often are they held for those in 
	the GRANTEE agency only?</b><br />
	<b>Method</b>
	<%=RadioInputField("MeetingsGranteeMethod", MeetingsGranteeMethod, 1) %>In-Person
	<%=RadioInputField("MeetingsGranteeMethod", MeetingsGranteeMethod, 2) %>Virtual
	<%=RadioInputField("MeetingsGranteeMethod", MeetingsGranteeMethod, 3) %>EMail
	<%=RadioInputField("MeetingsGranteeMethod", MeetingsGranteeMethod, 4) %>Phone
	<%=RadioInputField("MeetingsGranteeMethod", MeetingsGranteeMethod, 5) %>Other
	<br />
	<b>Frequency</b>
	<%=RadioInputField("MeetingsGranteeFrequency", MeetingsGranteeFrequency, 1) %>Daily
	<%=RadioInputField("MeetingsGranteeFrequency", MeetingsGranteeFrequency, 2) %>Weekly
	<%=RadioInputField("MeetingsGranteeFrequency", MeetingsGranteeFrequency, 3) %>Every two weeks
	<%=RadioInputField("MeetingsGranteeFrequency", MeetingsGranteeFrequency, 4) %>Monthly
	<%=RadioInputField("MeetingsGranteeFrequency", MeetingsGranteeFrequency, 5) %>Quarterly
	<%=RadioInputField("MeetingsGranteeFrequency", MeetingsGranteeFrequency, 6) %>Yearly
</p>

<p><b>By what primary method do scheduled meetings occur and how often are they held that 
	include the GRANTEE agency and INDIVIDUAL SUBGRANTEE agencies?</b><br />
	<b>Method</b>
	<%=RadioInputField("MeetingsSubGranteeMethod", MeetingsSubGranteeMethod, 1) %>In-Person
	<%=RadioInputField("MeetingsSubGranteeMethod", MeetingsSubGranteeMethod, 2) %>Virtual
	<%=RadioInputField("MeetingsSubGranteeMethod", MeetingsSubGranteeMethod, 3) %>EMail
	<%=RadioInputField("MeetingsSubGranteeMethod", MeetingsSubGranteeMethod, 4) %>Phone
	<%=RadioInputField("MeetingsSubGranteeMethod", MeetingsSubGranteeMethod, 5) %>Other
	<br />
	<b>Frequency</b>
	<%=RadioInputField("MeetingsSubGranteeFrequency", MeetingsSubGranteeFrequency, 1) %>Daily
	<%=RadioInputField("MeetingsSubGranteeFrequency", MeetingsSubGranteeFrequency, 2) %>Weekly
	<%=RadioInputField("MeetingsSubGranteeFrequency", MeetingsSubGranteeFrequency, 3) %>Every two weeks
	<%=RadioInputField("MeetingsSubGranteeFrequency", MeetingsSubGranteeFrequency, 4) %>Monthly
	<%=RadioInputField("MeetingsSubGranteeFrequency", MeetingsSubGranteeFrequency, 5) %>Quarterly
	<%=RadioInputField("MeetingsSubGranteeFrequency", MeetingsSubGranteeFrequency, 6) %>Yearly
</p>

<p><b>By what primary method do scheduled meetings occur and how often are they held that 
	include the ENTIRE TASKFORCE?</b><br />
	<b>Method</b>
	<%=RadioInputField("MeetingsAllTFMethod", MeetingsAllTFMethod, 1) %>In-Person
	<%=RadioInputField("MeetingsAllTFMethod", MeetingsAllTFMethod, 2) %>Virtual
	<%=RadioInputField("MeetingsAllTFMethod", MeetingsAllTFMethod, 3) %>EMail
	<%=RadioInputField("MeetingsAllTFMethod", MeetingsAllTFMethod, 4) %>Phone
	<%=RadioInputField("MeetingsAllTFMethod", MeetingsAllTFMethod, 5) %>Other
	<br />
	<b>Frequency</b>
	<%=RadioInputField("MeetingsAllTFFrequency", MeetingsAllTFFrequency, 1) %>Daily
	<%=RadioInputField("MeetingsAllTFFrequency", MeetingsAllTFFrequency, 2) %>Weekly
	<%=RadioInputField("MeetingsAllTFFrequency", MeetingsAllTFFrequency, 3) %>Every two weeks
	<%=RadioInputField("MeetingsAllTFFrequency", MeetingsAllTFFrequency, 4) %>Monthly
	<%=RadioInputField("MeetingsAllTFFrequency", MeetingsAllTFFrequency, 5) %>Quarterly
	<%=RadioInputField("MeetingsAllTFFrequency", MeetingsAllTFFrequency, 6) %>Yearly
</p>

<p><b>Describe the taskforce meetings with grantee and subgrantee agencies. Include meeting 
	organization, attendees, information, operational issues and progress report and performance 
	data collection issue.</b><br />
	<%=TextArea("MeetingsDescription", MeetingsDescription, 8, 120, 4000, PermitEdit, "") %>
</p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 3 Grantee and Subgrantee Contacts and Communication</div>

<p><b>By what primary method and how often does communication occur among those 
		in the GRANTEE agency only?</b><br />
	<b>Method</b>
	<%=RadioInputField("CommunicationGranteeMethod", CommunicationGranteeMethod, 1) %>In-Person
	<%=RadioInputField("CommunicationGranteeMethod", CommunicationGranteeMethod, 2) %>Virtual
	<%=RadioInputField("CommunicationGranteeMethod", CommunicationGranteeMethod, 3) %>EMail
	<%=RadioInputField("CommunicationGranteeMethod", CommunicationGranteeMethod, 4) %>Phone
	<%=RadioInputField("CommunicationGranteeMethod", CommunicationGranteeMethod, 5) %>Other
	<br />
	<b>Frequency</b>
	<%=RadioInputField("CommunicationGranteeFrequency", CommunicationGranteeFrequency, 1) %>Daily
	<%=RadioInputField("CommunicationGranteeFrequency", CommunicationGranteeFrequency, 2) %>Weekly
	<%=RadioInputField("CommunicationGranteeFrequency", CommunicationGranteeFrequency, 3) %>Every two weeks
	<%=RadioInputField("CommunicationGranteeFrequency", CommunicationGranteeFrequency, 4) %>Monthly
	<%=RadioInputField("CommunicationGranteeFrequency", CommunicationGranteeFrequency, 5) %>Quarterly
	<%=RadioInputField("CommunicationGranteeFrequency", CommunicationGranteeFrequency, 6) %>Yearly
</p>

<p><b>By what primary method and how often does communication occur that include 
	the GRANTEE agency and INDIVIDUAL SUBGRANTEES?</b><br />
	<b>Method</b>
	<%=RadioInputField("CommunicationSubGranteeMethod", CommunicationSubGranteeMethod, 1) %>In-Person
	<%=RadioInputField("CommunicationSubGranteeMethod", CommunicationSubGranteeMethod, 2) %>Virtual
	<%=RadioInputField("CommunicationSubGranteeMethod", CommunicationSubGranteeMethod, 3) %>EMail
	<%=RadioInputField("CommunicationSubGranteeMethod", CommunicationSubGranteeMethod, 4) %>Phone
	<%=RadioInputField("CommunicationSubGranteeMethod", CommunicationSubGranteeMethod, 5) %>Other
	<br />
	<b>Frequency</b>
	<%=RadioInputField("CommunicationSubGranteeFrequency", CommunicationSubGranteeFrequency, 1) %>Daily
	<%=RadioInputField("CommunicationSubGranteeFrequency", CommunicationSubGranteeFrequency, 2) %>Weekly
	<%=RadioInputField("CommunicationSubGranteeFrequency", CommunicationSubGranteeFrequency, 3) %>Every two weeks
	<%=RadioInputField("CommunicationSubGranteeFrequency", CommunicationSubGranteeFrequency, 4) %>Monthly
	<%=RadioInputField("CommunicationSubGranteeFrequency", CommunicationSubGranteeFrequency, 5) %>Quarterly
	<%=RadioInputField("CommunicationSubGranteeFrequency", CommunicationSubGranteeFrequency, 6) %>Yearly
</p>

<p><b>By what primary method and how often does communication occur that include 
	the ENTIRE TASKFORCE?</b><br />
	<b>Method</b>
	<%=RadioInputField("CommunicationAllTFMethod", CommunicationAllTFMethod, 1) %>In-Person
	<%=RadioInputField("CommunicationAllTFMethod", CommunicationAllTFMethod, 2) %>Virtual
	<%=RadioInputField("CommunicationAllTFMethod", CommunicationAllTFMethod, 3) %>EMail
	<%=RadioInputField("CommunicationAllTFMethod", CommunicationAllTFMethod, 4) %>Phone
	<%=RadioInputField("CommunicationAllTFMethod", CommunicationAllTFMethod, 5) %>Other
	<br />
	<b>Frequency</b>
	<%=RadioInputField("CommunicationAllTFFrequency", CommunicationAllTFFrequency, 1) %>Daily
	<%=RadioInputField("CommunicationAllTFFrequency", CommunicationAllTFFrequency, 2) %>Weekly
	<%=RadioInputField("CommunicationAllTFFrequency", CommunicationAllTFFrequency, 3) %>Every two weeks
	<%=RadioInputField("CommunicationAllTFFrequency", CommunicationAllTFFrequency, 4) %>Monthly
	<%=RadioInputField("CommunicationAllTFFrequency", CommunicationAllTFFrequency, 5) %>Quarterly
	<%=RadioInputField("CommunicationAllTFFrequency", CommunicationAllTFFrequency, 6) %>Yearly
</p>

<p><b>Describe the taskforce communication with grantee and subgrantee agencies. 
	Include regular, occasional and ad hoc communication about cases, reporting, and trends.</b><br />
	<%=TextArea("CommunicationDescription", CommunicationDescription, 8, 120, 4000, PermitEdit, "") %>
</p>
<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 4 Coverage Agency Meetings</div>

<br />

<p><b>Describe meetings that grantee and subgrantee agencies perform with or for coverage agencies. 
	Include purpose, method and frequency of meetings.</b><br />
	<%=TextArea("CoverageAgencyMeetings", CoverageAgencyMeetings, 8, 120, 4000, PermitEdit, "") %>
</p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 5 Coverage Agency Contacts</div>

<p><b>Describe contact that grantee and subgrantee have with coverage agencies. 
Include purpose, method and frequency of contact. </b><br />
	<%=TextArea("CoverageAgencyContacts", CoverageAgencyContacts, 8, 120, 4000, PermitEdit, "") %>
</p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 6 Intelligence Sharing</div>

<br />

<p><b>Describe a plan to develop, collect, process, disseminate, and receive feedback, 
	intelligence information.  Describe who (sub grantee, coverage agencies, and or other) and 
	how the intelligence is disseminated. Is the information posted to the Virtual Command Center?</b><br />
	<%=TextArea("IntelligenceSharing", IntelligenceSharing, 8, 120, 4000, PermitEdit, "") %>
</p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 7 Operational and Investigative Coordination</div>

<br />

<p><b>Describe how cases are assigned to taskforce personnel.  Include if subgrantees are assigned cases from 
	taskforce commander, the sub grantee agency, or both.</b><br />
	<%=TextArea("OperationalCoordination", OperationalCoordination, 8, 120, 4000, PermitEdit, "") %>
</p>

<br />


<div style="width: 976px; text-align: center; font-weight: bold; ">Section 8 Direct Operations</div>

<br />

<p><b>Describe how taskforce personnel conduct operations/activities as a group.  
	Include how and what types of operations occur in participating agency jurisdictions.  
	Include how border/bridge  and port operations are coordinated and occur (planned and unplanned) if applicable.</b><br />
	<%=TextArea("DirectOperatations", DirectOperatations, 8, 120, 4000, PermitEdit, "") %>
</p>

<br />

<%
	If AppID>0 And (AllowUpload = True) Then
		Response.Write("<div style=""text-align: center; ""><a href=""../Upload/Upload.asp?FID=13&AppID=" & AppID & """ target=""_blank"">File Upload</a></div><br />" & vbCrLf)
	End If

	DocumentFolder = Application("DocumentRoot") & "\Application\" & AppID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemOBject")
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		If files.count>0 Then 
			Response.Write("<div style=""width: 600px; margin: auto; ""><h2>Current Documents in folder</h2>" & vbCrLf)
		For Each file in files
				Response.Write("<a href=""../Documents/Application/" & AppID & "/" & file.Name & _
					""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
		Next
			Response.Write("<br /></div>" & vbCrLf)
		End If
	End If
%>
<br />
<div style="text-align: center; ">
<%
If MVCPARights = True And IsNull(SubmitID) = False Then
	Response.Write("<b>Administrative Use Only</b><br />")
	If IsNull(OperationalPlanApprovalDate) = True Then
		Response.Write("Multi-Agency Operations Plan Approved: " & DateField("OperationalPlanApprovalDate", OperationalPlanApprovalDate, MVCPARights) & "<br />")
	Else
		Response.Write("Multi-Agency Operations Plan Approved: " & DateField("OperationalPlanApprovalDate", OperationalPlanApprovalDate, MVCPARights) & " by " & OperationalPlanApprovalName & "<br />")
	End If
	Response.Write("<br />")
End If
If Debug = True Then
	Response.Write("<pre>")
	Response.Write("AuthorizedOfficialID=" & AuthorizedOfficialID & "; User=" & UserSystemID & vbCrLf)
	Response.Write("PermitEdit=" & PermitEdit & vbCrLf)
	Response.Write("MVCPA Rights=" & MVCPARights & vbCrLf)
	Response.Write("Allow Upload=" & AllowUpload & vbCrLf)
	'Response.Write("ReadyToSubmit=" & ReadyToSubmit & vbCrLf)
	Response.Write("</pre>")
End If
%>
		<input type="button" value="Save" onclick="return submitForm('save');" 
			title="Save what you have currently and remain on the page."/>
<%	IF Submitted = False And (UserSystemID = ProgramDirectorID Or UserSystemID = ProgramManagerID) Then%>
		<input type="button" value="Submit" onclick="return submitForm('submit');" 
			title="Only the authorized official may submit the application. After submitting, you will be returned to the home page."/>
<%	End if %>
		<input type="button" value="Close" onclick="window.close();" 
			title="Do to browser security changes, you will likerly have to click on the 'X' at the upper right corner of the window to close."/>
<%
If SubmitID>0 And MVCPARights = False Then
	Response.Write("<br /><br />Click on ""X"" at upper right of window to close.<br />")
End If
%>
</div>
</form>
<br />
</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
<script type="text/javascript">

	function validateForm()
	{
		return true;
	}

	function checkTypes()
	{
		// Add validation for things that are required to save and avoid an error.
		document.Application.ProgramName.value = replaceWordChars(document.Application.ProgramName.value);
		document.Application.OtherCoverageText.value = replaceWordChars(document.Application.OtherCoverageText.value);
		return true;
	}
</script>
<script src="../includes/formchanges.js"></script>
<script type="text/javascript">
	var saving = false;
	var form = document.getElementById("Application");

	// form being updated
	form.onsubmit = function() { saving = true; };

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
	function DetectChanges() {
		var f = FormChanges(form), msg = "";
		for (var e = 0, el = f.length; e < el; e++) msg += "\n" + f[e].id;
		alert((msg ? "Elements changed:" : "No changes made.") + msg);
	}

	// Save changes
	function SaveChanges() {
		var f = FormChanges(form), msg = "";
		for (var e = 0, el = f.length; e < el; e++) msg += f[e].id + "\n";
		document.Application.Changes.value=msg;
	}

</script>
</body>
</html>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
