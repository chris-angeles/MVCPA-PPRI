<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, LastCategory, PermitEdit, AllowUpload, ApplicationSchema, AppID, FiscalYear, GranteeID, GranteeName, _
	ORI, ORIAgency, OrganizationTypeID, OrganizationType, StatePayeeIDNo, _
	AuthorizedOfficialID, AuthorizedOfficial, AuthorizedOfficialTitle, ProgramName, GrantTypeID, _
	CoverageAreaDescription, StatewideCoverage, OtherCoverage, OtherCoverageText, LawEnforcementGrant, _
	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, _
	HistoricalDataYear,  LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, _
	LarcenyFromMVParts1, LarcenyFromMVParts2, LarcenyFromMVParts3, LarcenyJurisdiction, _
	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, DataProblems, _
	Certification, SubmitID, SubmitByName, SubmitTimestamp, ConfirmationNumber, ReadyToSubmit, _
	CashMatch, InKindMatch, GrandTotal, TotalMVCPAFunds, TotalCashMatch, TotalInkindMatch, _
	DetailCashMatch, DetailInKIndMatch, DetailTotalMatch, PctMVCPA, PctCashMatch, _
	BudgetEntryOption, BudgetCashMatch, RoundCurrency, _
	DocumentFolder, fso, folder, file, files, _
	TargetAwardAmount, TargetMatchAmount, TargetReimbursementRate, GSATargets
Dim ProgramCategory(5)

debug = False
ApplicationSchema = "Negotiation"

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

AppID = Request.QueryString("AppID")
GranteeID = Request.QueryString("GranteeID")
FiscalYear = Request.QueryString("FiscalYear")
GrantTypeID = Request.QueryString("GrantTypeID")

If AppID="" Then
	AppID=0
	If GranteeID="" Then
		GranteeID = Session("GranteeID")
	End If
	If GranteeID="" Or GranteeID=0 Then
		Response.Write("Error: No AppID or GranteeID Specified")
		SendMessage "Error: No AppID or GranteeID Specified"
		Response.End
	End If
Else
	AppID=Cint(AppID)
End If

' Disable Application
'If UserSystemID<>1 and UserSystemID<>2 Then
'	Response.Redirect("PrintApplication.asp?AppID=" & AppID)
'End If


If AppID>0 Then 
	sql = "SELECT G.GranteeID, G.GranteeName, G.ORI, ORI.Agency AS ORIAgency, G.OrganizationTypeID, OT.OrganizationType, G.StatePayeeIDNo, " & vbCrLf & _
		"	ISNULL(I.FiscalYear, " & prepIntegerSQL(FiscalYear) & ") AS FiscalYear, " & vbCrLf & _
		"	AuthorizedOfficialID, AO.Name AS AuthorizedOfficial, AO.Title AS AuthorizedOfficialTitle, ISNULL(A.AppID,0) AS AppID, ProgramName, " & vbCrLf & _
		"	A.GrantTypeID, A.CoverageAreaDescription, A.StatewideCoverage, A.OtherCoverage, A.OtherCoverageText, A.LawEnforcementGrant, " & vbCrLf & _
		"	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, " & vbCrLf & _
		"	ProgramCategory1, ProgramCategory2, ProgramCategory3, ProgramCategory4, ProgramCategory5, " & vbCrLf & _
		"	HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, " & vbCrLf & _
		"	LarcenyFromMVParts1, LarcenyFromMVParts2, LarcenyFromMVParts3, " & vbCrLf & _
		"	LarcenyJurisdiction, DataProblems, " & vbCrLf & _
		"	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, " & vbCrLf & _
		"	A.Certification, A.SubmitID, U.Name AS SubmitByName, A.SubmitTimestamp, A.ConfirmationNumber, " & vbCrLf & _
		"	CASE WHEN A.BudgetCashMatch IS NOT NULL THEN 2 ELSE 1 END AS BudgetEntryOption, " & vbCrLf & _
		"	A.BudgetCashMatch, " & vbCrLf & _
		"	ISNULL(B.TotalMVCPAFunds,0.0) AS TotalMVCPAFunds, " & vbCrLf & _
		"	ISNULL(B.TotalCashMatch,0.0) AS TotalCashMatch, " & vbCrLf & _
		"	ISNULL(B.GrandTotal,0.0) AS GrandTotal, " & vbCrLf & _
		"	ISNULL(B.TotalInKindMatch,0.0) AS TotalInKindMatch, " & vbCrLf & _
		"	ISNULL(M.DetailCashMatch,0.0) AS DetailCashMatch, " & vbCrLf & _
		"	ISNULL(M.DetailInKindMatch,0.0) AS DetailInKindMatch, " & vbCrLf & _
		"	ISNULL(DetailTotalMatch,0.0) AS DetailTotalMatch, " & vbCrLf & _
		"	N.AwardAmount AS TargetAwardAmount, N.MatchAmount AS TargetMatchAmount, N.ReimbursementRate AS TargetReimbursementRate, " & vbCrLf & _
		"	GSATargets = (SELECT COUNT(*) AS GSATargets FROM " & ApplicationSchema & ".GSATargets WHERE AppID=A.AppID) " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".Main AS A ON A.AppID=I.AppID " & vbCrLf & _
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
		"WHERE A.AppID=" & PrepIntegerSQL(AppID)
Else
	sql = "SELECT G.GranteeID, G.GranteeName, G.ORI, ORI.Agency AS ORIAgency, G.OrganizationTypeID, OT.OrganizationType, G.StatePayeeIDNo, " & vbCrLf & _
		"	ISNULL(I.FiscalYear, " & prepIntegerSQL(FiscalYear) & ") AS FiscalYear, " & vbCrLf & _
		"	AuthorizedOfficialID, AO.Name AS AuthorizedOfficial, AO.Title AS AuthorizedOfficialTitle, ISNULL(A.AppID,0) AS AppID, ProgramName, " & vbCrLf & _
		"	ISNULL(A.GrantTypeID," & prepIntegerSQL(GrantTypeID) & ") AS GrantTypeID, " & vbCrLf & _
		"	A.CoverageAreaDescription, A.StatewideCoverage, A.OtherCoverage, A.OtherCoverageText, ISNULL(LawEnforcementGrant,1) AS LawEnforcementGrant, " & vbCrLf & _
		"	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, " & vbCrLf & _
		"	ProgramCategory1, ProgramCategory2, ProgramCategory3, ProgramCategory4, ProgramCategory5, " & vbCrLf & _
		"	HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, " & vbCrLf & _
		"	LarcenyFromMVParts1, LarcenyFRomMVParts2, LarcenyFromMVParts3, " & vbCrLf & _
		"	LarcenyJurisdiction, DataProblems, " & vbCrLf & _
		"	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, " & vbCrLf & _
		"	A.Certification, A.SubmitID, U.Name AS SubmitByName, A.SubmitTimestamp, A.ConfirmationNumber, " & vbCrLf & _
		"	CASE WHEN A.AppID IS NULL THEN 2 WHEN A.BudgetCashMatch IS NOT NULL THEN 2 ELSE 1 END AS BudgetEntryOption, " & vbCrLF & _
		"	CASE WHEN A.AppID IS NULL THEN 20 ELSE A.BudgetCashMatch END AS BudgetCashMatch, " & vbCrLf & _
		"	ISNULL(B.TotalMVCPAFunds,0.0) AS TotalMVCPAFunds, " & vbCrLf & _
		"	ISNULL(B.TotalCashMatch,0.0) AS TotalCashMatch, " & vbCrLf & _
		"	ISNULL(B.GrandTotal,0.0) AS GrandTotal, " & vbCrLf & _
		"	ISNULL(B.TotalInKindMatch,0.0) AS TotalInKindMatch, " & vbCrLf & _
		"	ISNULL(M.DetailCashMatch,0.0) AS DetailCashMatch, " & vbCrLf & _
		"	ISNULL(M.DetailInKindMatch,0.0) AS DetailInKindMatch, " & vbCrLf & _
		"	ISNULL(DetailTotalMatch,0.0) AS DetailTotalMatch, " & vbCrLf & _
		"	NULL AS TargetAwardAmount, NULL AS TargetMatchAmount, NULL AS TargetReimbursementRate, " & vbCrLf & _
		"	0 AS GSATArgets " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf &_
		"LEFT JOIN " & ApplicationSchema & ".Main AS A ON A.AppID=G.AppID AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"LEFT JOIN Lookup.OrganizationType AS OT ON OT.OrganizationTypeID=G.OrganizationTypeID " & vbCrLf & _
		"LEFT JOIN Lookup.ORI AS ORI ON ORI.ORI=G.ORI " & vbCrLf & _
		"LEFT JOIN [System].Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf
	If Len(GrantTypeID)>0 Then
		sql = sql &  vbTab & "AND ISNULL(GrantTypeID," & prepIntegerSQL(GrantTypeID) & ")=" & prepIntegerSQL(GrantTypeID)
	End If
	sql = sql & "LEFT JOIN System.Users AS U ON U.SystemID=A.SubmitID " & vbCrLf & _
		"LEFT JOIN ( " & vbCrLf & _
		"	SELECT AppID, SUM(MVCPAFunds) AS TotalMVCPAFunds, SUM(CashMatch) AS TotalCashMatch, SUM(LineTotal) AS GrandTotal, SUM(InKindMatch) AS TotalInKindMatch " & vbCrLf & _
		"	FROM " & ApplicationSchema & ".BudgetDetails " & vbCrLf & _
		"	GROUP BY AppID) AS B ON B.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN ( " & vbCrLf & _
		"	SELECT AppID, SUM(CASE WHEN MatchTypeID=1 Then Amount ELSE NULL END) AS DetailCashMatch, " & vbCrLf & _
		"		SUM(CASE WHEN MatchTypeID=2 Then Amount ELSE NULL END) AS DetailInKindMatch,  " & vbCrLf & _
		"		SUM(Amount) AS DetailTotalMatch " & vbCrLf & _
		"	FROM " & ApplicationSchema & ".Matches " & vbCrLf & _
		"	GROUP BY AppID) AS M ON M.AppID=A.AppID " & vbCrLf & _
	"WHERE G.GranteeID=" & PrepIntegerSQL(GranteeID)
End If

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Grantee and " & ApplicationSchema & " record retrieved")
	SendMessage "Error: No Grantee and " & ApplicationSchema & " record retrieved"
	Response.End
Else
	AppID = rs.Fields("AppID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	ORI = rs.Fields("ORI")
	ORIAgency = rs.Fields("ORIAgency")
	OrganizationTypeID = rs.Fields("OrganizationTypeID")
	OrganizationType = rs.Fields("OrganizationType")
	StatePayeeIDNo = rs.Fields("StatePayeeIDNo")
	AuthorizedOfficialID = rs.Fields("AuthorizedOfficialID")
	AuthorizedOfficial = rs.Fields("AuthorizedOfficial")
	AuthorizedOfficialTitle = rs.Fields("AuthorizedOfficialTitle")
	ProgramName = rs.Fields("ProgramName")
	GrantTypeID = rs.Fields("GrantTypeID")
	For i = 1 to 5
		ProgramCategory(i) = rs.Fields("ProgramCategory" & i)
	Next
	CoverageAreaDescription = rs.Fields("CoverageAreaDescription")
	StatewideCoverage = rs.Fields("StatewideCoverage")
	OtherCoverage = rs.Fields("OtherCoverage")
	OtherCoverageText = rs.Fields("OtherCoverageText")
	LawEnforcementGrant = rs.Fields("LawEnforcementGrant")
	NationalInsuranceCrimeBureau = rs.Fields("NationalInsuranceCrimeBureau")
	TexasDepartmentOfPublicSafety = rs.Fields("TexasDepartmentOfPublicSafety")
	OtherAgency = rs.Fields("OtherAgency")
	OtherAgencySpecify = rs.Fields("OtherAgencySpecify")
	HistoricalDataYear = rs.Fields("HistoricalDataYear")
	LarcenyFromMV1 = rs.Fields("LarcenyFromMV1")
	LarcenyFromMV2 = rs.Fields("LarcenyFromMV2")
	LarcenyFromMV3 = rs.Fields("LarcenyFromMV3")
	LarcenyFromMVParts1 = rs.Fields("LarcenyFromMVParts1")
	LarcenyFromMVParts2 = rs.Fields("LarcenyFromMVParts2")
	LarcenyFromMVParts3 = rs.Fields("LarcenyFromMVParts3")
	LarcenyJurisdiction = rs.Fields("LarcenyJurisdiction")
	MVT1 = rs.Fields("MVT1")
	MVT2 = rs.Fields("MVT2")
	MVT3 = rs.Fields("MVT3")
	RecoveryMVT1 = rs.Fields("RecoveryMVT1")
	RecoveryMVT2 = rs.Fields("RecoveryMVT2")
	RecoveryMVT3 = rs.Fields("RecoveryMVT3")
	MVTJurisdiction = rs.Fields("MVTJurisdiction")
	DataProblems = rs.Fields("DataProblems")
	BudgetEntryOption = rs.Fields("BudgetEntryOption")
	BudgetCashMatch = rs.Fields("BudgetCashMatch")
	Certification = rs.Fields("Certification")
	SubmitID = rs.Fields("SubmitID")
	SubmitByName = rs.Fields("SubmitByName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	ConfirmationNumber = rs.Fields("ConfirmationNumber")
	TotalMVCPAFunds = rs.Fields("TotalMVCPAFunds")
	TotalCashMatch = rs.Fields("TotalCashMatch")
	TotalInkindMatch = rs.Fields("TotalInkindMatch")
	GrandTotal = rs.Fields("GrandTotal")
	DetailCashMatch = rs.Fields("DetailCashMatch")
	DetailInKIndMatch = rs.Fields("DetailInKIndMatch")
	DetailTotalMatch = rs.Fields("DetailTotalMatch")
	TargetAwardAmount = rs.Fields("TargetAwardAmount")
	TargetMatchAmount = rs.Fields("TargetMatchAmount")
	TargetReimbursementRate = rs.Fields("TargetReimbursementRate")
	GSATargets = rs.Fields("GSATargets")
End If

' Start rounding dollar amounts as of 2020.
If FiscalYear>=2020 Then
	RoundCurrency = True
Else
	RoundCurrency = False
End If

If FiscalYear=2022 Then
	GrantTypeID=3
End If
If AppID=0 Then 
	ReadyToSubmit = False
ElseIf IsNull(OrganizationTypeID) Or IsNull(StatePayeeIDNo) Or IsNull(AuthorizedOfficialID) Then
	ReadyToSubmit = False
ElseIf UserSystemID <> AuthorizedOfficialID Then
	ReadyToSubmit = False
ElseIf TotalCashMatch <> DetailCashMatch Then
	ReadyToSubmit = False
ElseIf TotalInKindMatch <> DetailInKindMatch Then
	ReadyToSubmit = False
ElseIf IsNull(TargetAwardAmount) = False And IsNull(TotalMVCPAFunds) = False Then
	If TotalMVCPAFunds > TargetAwardAmount Then
		ReadyToSubmit = False
	Else
		ReadyToSubmit = True
	End If
ElseIf IsNull(TargetMatchAmount) = False And IsNull(TotalCashMatch) = False Then
	If TargetMatchAmount < TotalCashMatch Then
		ReadyToSubmit = False
	Else
		ReadyToSubmit = True
	End If
Else
	ReadyToSubmit = True
End If

If Debug = True Then
	Response.Write("<pre>ReadyToSubmit=" & ReadyToSubmit & "</pre>")
End If
DocumentFolder = Application("DocumentRoot") & "\Application\" & AppID & "\"

If GranteeID>0 Then
	If IsNull(SubmitID) = True Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
		'PermitEdit = False
	ElseIf ISNull(SubmitID) = False Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, True)
		PermitEdit = False
	Else
		PermitEdit = False
	End If
	AllowUpload = CheckPermissions(UserSystemID, GranteeID, True)
Else
	PermitEdit = False
	AllowUpoad = False
End If

If IsNull(HistoricalDataYear) Then
	HistoricalDataYear = FiscalYear - 2
End If

sql = "SELECT ISNULL(SUM(CASE WHEN MatchTypeID=1 THEN Amount ELSE NULL END),0) AS CashMatch, " & vbCrLf & _
	"	ISNULL(SUM(CASE WHEN MatchTypeID=2 THEN Amount ELSE NULL END),0) AS InKindMatch " & vbCrLf & _
	"FROM " & ApplicationSchema & ".Matches " & vbCrLf & _
	"WHERE AppID=" & prepIntegerSQL(AppID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	CashMatch = 0
	InKIndMatch = 0
Else
	CashMatch = rs.Fields("CashMatch")
	InKIndMatch = rs.Fields("InKindMatch")
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Taskforce Grant <%=ApplicationSchema %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style type="text/css">	th {
		text-align: center;
	}
</style>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag"><%=GranteeName %> Taskforce Grant <%=ApplicationSchema%> for Fiscal Year <%=FiscalYear %></div>

<div class="widecontent">

<form name="Application" id="Application" method="post" action="TFGApplicationSubmit.asp" onsubmit="return validateForm()">
<%
Response.Write(HiddenField("GranteeID", GranteeID))
Response.Write(HiddenField("AppID", AppID))
Response.Write(HiddenField("FiscalYear", FiscalYear))
Response.Write(HiddenField("HistoricalDataYear", HistoricalDataYear))
Response.Write(HiddenField("Button","save"))
Response.Write(HiddenField("ParticipatingAgenciesChanged",""))
Response.Write(HiddenField("CoverageAgenciesChanged",""))
Response.Write(HiddenField("Changes",""))
Response.Write(HiddenField("LawEnforcementGrant", LawEnforcementGrant))
%>
<table style="width: 956px; ">
<%	
If SubmitID>0 Then 
%>
<tr><td colspan="2" style="text-align: center; font-weight: bold; ">The Application was submitted by <%=SubmitByName%> at <%=SubmitTimestamp %> and is now locked.<br />
	The confirmation Number is <%=ConfirmationNumber %>.</td></tr>
<tr><td colspan="2">&nbsp;</td></tr>
<%	
Else
	If FiscalYear=2024 Then
		Response.Write("<tr><td colspan=""2"" style=""text-align: center; ""><a href=""https://www.txdmv.gov/sites/default/files/body-files/FY%202024%20SB%20224%20Catalytic%20Converter%20Grant%20RFA.pdf"" target=""_blank"">Request for Application (RFA)</a></td></tr>" & vbCrLf)
	Else
		Response.Write("<tr><td colspan=""2"" style=""text-align: center; ""><a href=""../RFA/RFA2022-23.pdf"" target=""_blank"">Request for Application (RFA)</a> <i>(need link to rfa)</i></td></tr>" & vbCrLf)
	End If
End If 

If IsNull(OrganizationTypeID) Or IsNull(StatePayeeIDNo) Or IsNull(AuthorizedOfficialID) Then
	Response.Write("<tr><td colspan=""2"" style=""text-align: center; font-weight: bold; color: red; ""><br />" & vbCrLf & _
		"The grantee information on the grantee screen has not been completed. " & vbCrLf & _
		"This must be done before the application can be submitted.<br /><br /></td></tr>" & vbCrLf)
End If
If TotalCashMatch <> DetailCashMatch Then
	Response.Write("<tr><td colspan=""2"" style=""text-align: center; font-weight: bold; color: red; ""><br />" & vbCrLf & _
		"The total cash match from the budget does not match the total cash match from the source of the matches detail. " & vbCrLf & _
		"These must match before the application can be submitted.<br /><br /></td></tr>" & vbCrLf)
End If
If TotalInKindMatch <> DetailInKindMatch Then
	Response.Write("<tr><td colspan=""2"" style=""text-align: center; font-weight: bold; color: red; ""><br />" & vbCrLf & _
		"The total in-kind match from the budget does not match the total in-kind match from the source of the matches detail. " & vbCrLf & _
		"These must match before the application can be submitted.<br /><br /></td></tr>" & vbCrLf)
End If
%>

<tr>
	<td colspan="2">Primary Agency / Grantee Legal Name: <i><%=GranteeName %></i></td>
</tr>

<tr>
	<td colspan="2">Organization Type: <i><%=OrganizationType %></i></td>
</tr>

<tr>
	<td colspan="2">Organization ORI (if applicable): <i><%=ORI %>: <%=ORIAgency %></i></td>
</tr>
<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2"><b>Program Title</b> Please enter a short description of the proposed program that can be used as the title.<br />
	<span class="usertext"><%=ProgramName%></span>
	</td>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2"><b>Application Category</b> (See <b>Request for Applications</b> [RFA] for category details and descriptions RFA Priority Funding Section):</td>
</tr>
<%
	If FiscalYear=2023 Then
		sql = "SELECT GrantTypeID, GrantType, GrantTypeDescription FROM Lookup.GrantType WHERE GrantTypeID=1 AND Version=1 "
	ElseIf FiscalYear=2022 Or FiscalYear=2024 Then
		sql = "SELECT GrantTypeID, GrantType, GrantTypeDescription FROM Lookup.GrantType WHERE GrantTypeID=1 AND Version=2 "
	ElseIf IsNull(TargetAwardAmount) = False Then
		sql = "SELECT GrantTypeID, GrantType, GrantTypeDescription FROM Lookup.GrantType WHERE GrantTypeID=1 AND Version=2"
	Else
		sql = "SELECT GrantTypeID, GrantType, GrantTypeDescription FROM Lookup.GrantType WHERE Version=2 ORDER BY GrantTypeID"
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		If GrantTypeID=rs.Fields("GrantTypeID") Or IsNull(GrantTypeID) Then
		Response.Write(vbTab & "<tr style=""vertical-align: top""><td></td><td><b>" & _
			rs.Fields("GrantType") & "</b> - " & replace(replace(rs.Fields("GrantTypeDescription"),"{PreviousYear}",(FiscalYear-1)),"{CurrentYear}",FiscalYear) & "</td></tr>" & vbCrLf)
		End If
		rs.MoveNext
		If Debug = True Then
			Response.Write("<!--GrantTypeID=" & GrantTypeID & "; rs.Fields(""GrantTypeID"")=" & rs.Fields("GrantTypeID") & "-->")
		End If
	Wend
	If IsNull(GrantTypeID) = True then
		Response.Write("<tr><td></td><td style=""font-weight: bold; "">Note: Grant Type not selected on application!</td></tr>" & vbCrLf)
	End If
%>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2"><b>MVCPA Program Category</b> (see <b>RFA</b> and TAC 43, 3 &sect;57.14). Check all that apply.</td>
</tr>
<%
	sql = "SELECT ProgramCategoryID, ProgramCategory FROM Lookup.ProgramCategory WHERE Version=1 ORDER BY ProgramCategoryID "
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		If ProgramCategory(rs.Fields("ProgramCategoryID")) = True Then
			Response.Write(vbTab & "<tr><td></td><td>&bullet; " & rs.Fields("ProgramCategory") & "</td></tr>" & vbCrLf)
		End If
		rs.MoveNext
	Wend
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
	<td><%=CoverageAreaDescription %></span></td></tr>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2"><b>Define in the tables below the grant relationships and geographic area 
	of the taskforce:</b></br> 
	Applicant will add the participating and coverage agencies from the ORI list below. 
	If an agency is not in the ORI list, please include the agency and role in the general 
	description above. Make sure to follow the definitions below and select an agency in the 
	dropdown. Use the <i>Add as Participating Agency</i> or <i>Add as Coverage Agency</i> button to 
	populate the list.</td>
</tr>

<tr style="vertical-align: top; ">
	<td></td>
	<td><p><b>Participating Agencies</b> are agencies that materially participate in the grant 
	application through the exchange of funds for reimbursement and cash match. Participating 
	agencies are defined after the grant award by interlocal/interagency agreements. Each 
	applicant must select their own agency first. Then select agencies that will receive or 
	provide funding and/or resources. [Note: Interlocal/interagency agreements do not need to 
	be submitted with the application. Interlocal agreements will need to be executed prior to 
	the first payment being made if selected for a grant. Letters of support with the application 
	from the participating agencies are strongly recommended.]</p>
	<p><b>Coverage Agencies</b> are agencies that provided some level of coverage, assistance 
	or support by this grant application but will not materially exchange funds as cash match 
	or reimbursement. The coverage is not supported by an after the award with interlocal/interagency 
	agreements. Coverage agencies as law enforcement agencies may have jurisdictional coverage 
	agreements unrelated to the grant (Ex. City Y is within County X or vice versa). Agencies 
	selected in this list include agencies that will be covered or where the agency indicates 
	that their agency will coordinate or call upon the taskforce. Letters of support with the 
	application from the participating agencies are strongly recommended.</p></td>
</tr>

<tr>
	<td></td>
	<td>
<table style="margin: auto;  border: 1px solid #dddddd; ">

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
</table></td>
</tr>
<tr><td colspan="2">&nbsp;</td></tr>
<%	If OtherCoverage = True Or Len(OtherCoverageText)>0 Then %>
<tr style="vertical-align: top; ">
	<td>&bullet;</td>
	<td><b>Other Coverage</b> (Use if ORI not listed or explanation is necessary.): <br />
	<span class="usertext"><%=OtherCoverageText %></span></td>
</tr>

<%
	End If
	If NationalInsuranceCrimeBureau = True Then%>
<tr style="vertical-align: top; ">
	<td>&bullet;</td>
	<td><b>National Insurance Crime Bureau (NICB)</b> Used as Match (Documentation and time certification required.)</td>
</tr>
<%	End If 
	If TexasDepartmentOfPublicSafety = True Then %>
<tr style="vertical-align: top; ">
	<td>&bullet;</td>
	<td><b>Texas Department of Public Safety (DPS)</b></td>
</tr>
<%	End If 
	If OtherAgency = True Or IsNull(OtherAgencySpecify) = False Then %>
<tr style="vertical-align: top; ">
	<td>&bullet;</td>
	<td><b>Other State or Federal Agency</b> (specify:) <span class="usertext"><%=OtherAgencySpecify %></span></td>
</tr>
<%	End If %>
</table>

<br />

<div style="width: 980; margin: auto"><b>Resolution</b>: Complete a Resolution and submit to local governing body 
	for approval. <a href="Resolution.asp?AppID=<%=AppID %>&GranteeID=<%=GranteeID %>&FiscalYear=<%=FiscalYear%>" target="_blank" class="plainlink">Sample Resolution</a> 
	is found in the Request for Application or send a request for an electronic copy to 
	<a href="mailto:grantsMVCPA@txdmv.gov?subject=Resolution Request" class="plainlink">grantsMVCPA@txdmv.gov</a>. 
	The completed and executed Resolution must be attached to this on-line application. 
</div>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Grant Budget Form</div>

<div style="width: 976px; text-align: left;">MVCPA recommends that the applicant complete the total costs 
(MVCPA and Cash Match combined) for this program. The applicant can then enter the desired amount of 
Cash Match (not less than 20% per TAC Title 43, §57.36). The system will then calculate the correct 
grant and match amounts.</div>
<br />
<div style="width: 976px; text-align: left; "><b>Budget Entry Option:</b><br /> 
<%	If BudgetEntryOption = 2 Then %>
Enter Total and let system calculate MVCPA Funds and Cash Match, Match Percentage: <%=BudgetCashMatch %>%<br />
<%	ElseIf BudgetEntryOption = 1 then %>
Enter MVCPA and Cash Match Amounts</div>
<%	End If %>
<br />
<%	
If AppID>0 Then
%>
<a name="budget" />
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
		If rs.Fields("BudgetCategoryID")=99 Then
			Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)  
		ElseIf PermitEdit = False Then
			If BudgetEntryOption = 2 Then
				Response.Write(vbTab & "<td><a href=""BudgetDetail2.asp?AppID=" & AppID & _
					"&BudgetCategoryID=" & rs.Fields("BudgetCategoryID") & """ class=""plainlink"">" & rs.Fields("BudgetCategory") & "</a></td>" & vbCrLf)  
			Else
				Response.Write(vbTab & "<td><a href=""BudgetDetail.asp?AppID=" & AppID & _
					"&BudgetCategoryID=" & rs.Fields("BudgetCategoryID") & """ class=""plainlink"">" & rs.Fields("BudgetCategory") & "</a></td>" & vbCrLf)  
			End If
		Else
			Response.Write(vbTab & "<td><a onclick=""submitForm('" & rs.Fields("BudgetCategoryID") & "');"" class=""plainlink"">" & rs.Fields("BudgetCategory") & "</a></td>" & vbCrLf)  
		End If
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
	If IsNull(TargetAwardAmount) = False And IsNull(TotalMVCPAFunds) = False Then
		If TotalMVCPAFunds > TargetAwardAmount Then
			ReadyToSubmit = False
			Response.Write(vbTab & "<tr><td colspan=""5"">&nbsp;</td></tr>" & vbCrLf)
			Response.Write(vbTab & "<tr><td colspan=""5"" style=""text-align: center; color: red; "">The maximum award for this grant is " & formatcurrencyRound(TargetAwardAmount, RoundCurrency) & ".</td></tr>" & vbCrLf)
			'Response.Write(vbTab & "<tr><td colspan=""5"" style=""text-align: center; font-style: italic; "">The maximum award amount for this grant is " & formatCurrencyRound(TargetAwardAmount, RoundCurrency) & ", the minimum cash match amount is " & formatCurrencyRound(TargetMatchAmount, RoundCurrency) & ".</td></tr>" & vbCrLf)
		End If
	End If
	If IsNull(TargetMatchAmount) = False And IsNull(TotalCashMatch) = False Then
		If TotalCashMatch < TargetMatchAmount Then
			ReadyToSubmit = False
			Response.Write(vbTab & "<tr><td colspan=""5"">&nbsp;</td></tr>" & vbCrLf)
			Response.Write(vbTab & "<tr><td colspan=""5"" style=""text-align: center; color: red; "">The cash match specified is below the minimum cash match allowed for this grant.</td></tr>" & vbCrLf)
			'Response.Write(vbTab & "<tr><td colspan=""5"" style=""text-align: center; color: red; "">The minimum cash match for this grant is " & formatcurrencyRound(TargetMatchAmount, RoundCurrency) & ".</td></tr>" & vbCrLf)
			'Response.Write(vbTab & "<tr><td colspan=""5"" style=""text-align: center; font-style: italic; "">The maximum award amount for this grant is " & formatCurrencyRound(TargetAwardAmount, RoundCurrency) & ", the minimum cash match amount is " & formatCurrencyRound(TargetMatchAmount, RoundCurrency) & ".</td></tr>" & vbCrLf)
		End If
	End If
%>

</tbody>
</table>
<br />
<a name="BudgetDetail" />
<%
Else
	Response.Write("<div style=""width: 100%; text-align: center; font-style: italic; font-weight: bold; "">Grant Budget Form Will Be Displayed After First Save of Application.</div>")
	Response.Write("<br />" & vbCrLf)
End If

If AppID>0 And GrandTotal>0 Then 
sql = "SELECT B.BudgetItemID, A.BudgetCategoryID, A.BudgetCategory, " & vbCrLf & _
	"	CASE WHEN B.NoOfItems>0 THEN ISNULL(B.Description,'') + ' (' + CAST(B.NoOfItems AS VARCHAR) + ')' ELSE B.Description END AS Description, " & vbCrLf & _
	"	SubCategory, PctTime, LineTotal, MVCPAFunds, CashMatch, InKindMatch " & vbCrLf & _
	"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".BudgetDetails AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"LEFT JOIN Lookup.BudgetSubcategories AS C ON C.BudgetCategoryID=B.BudgetCategoryID AND C.SubCategoryID=B.SubCategoryID " & vbCrLf & _
	"UNION " & vbCrLf & _
	"SELECT 2147483647 AS BudgetItemID, A.BudgetCategoryID, A.BudgetCategory, 'Total ' + A.BudgetCategory AS Description, null, SUM(PctTime) AS PctTime, SUM(LineTotal) AS LineTotal, SUM(MVCPAFunds) AS MVCPAFunds, SUM(CashMatch) AS CashMatch, Sum(InKindMatch) AS InKindMatch " & vbCrLf & _
	"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".BudgetDetails AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"GROUP BY A.BudgetCategoryID, A.BudgetCategory" & vbCrLf & _
	"ORDER BY 2, 1 "
LastCategory=0
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<table style=""margin: auto; border: "">" & vbCrLf)
	Response.Write("<thead><tr style=""vertical-align: bottom; "">" & vbCrLf)
	Response.Write("<th>Description</th>" & vbCrLf)
	Response.Write("<th>Subcategory</th>" & vbCrLf)
	Response.Write("<th>Pct Time</th>" & vbCrLf)
	Response.Write("<th style=""width: 100px; "">MVCPA Funds</th>" & vbCrLf)
	Response.Write("<th style=""width: 100px; "">Cash Match</th>" & vbCrLf)
	Response.Write("<th style=""width: 100px; "">Total</th>" & vbCrLf)
	Response.Write("<th style=""width: 100px; "">In-Kind Match</th>" & vbCrLf)
	Response.Write("</tr></thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		If LastCategory<>rs.Fields("BudgetCategoryID") Then
			LastCategory=rs.Fields("BudgetCategoryID")
			Response.Write("<tr><td colspan=""6"">&nbsp;</td></tr>" & vbCrLf)
			Response.Write("<tr><th colspan=""6"">" & rs.Fields("BudgetCategory") & "</th></tr>" & vbCrLf)
		End If
		Response.Write("<tr>" & vbCrLf)
		Response.Write("<td>" & rs.Fields("Description") & "</td>")
		Response.Write("<td>" & rs.Fields("SubCategory") & "</td>")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("PctTime") & "</td>")
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("MVCPAFunds"), RoundCurrency) & "</td>")
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("CashMatch"), RoundCurrency) & "</td>")
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("LineTotal"), RoundCurrency) & "</td>")
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("InKindMatch"), RoundCurrency) & "</td>")
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)

	Response.Write("<br />" & vbCrLf)
	sql = "SELECT A.BudgetCategoryID, A.BudgetCategory, A.BudgetCategoryLetter, B.Narrative " & vbCrLf & _
	"FROM [Lookup].BudgetCategories AS A " & vbCrLf & _
	"LEFT JOIN [" & ApplicationSchema & "].[BudgetDetailNarrative] AS B ON B.BudgetCategoryID=A.BudgetCategoryID " & vbCrLf & _
	"WHERE B.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"ORDER BY A.BudgetCategoryID "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		Response.Write("<table style=""margin: auto; width: 100%; "">" & vbCrLf)
		Response.Write("<thead><tr><th colspan=""2"">Budget Narrative</th></tr></thead>" & vbCrLf)
		Response.Write("<tbody>" & vbCrLf)
		While rs.EOF = False
			Response.Write("<tr><td style=""width: 10px; font-weight: bold; "">" & rs.Fields("BudgetCategoryLetter") & ".</td><td style=""font-weight: bold; text-align: left;"" >" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
			Response.Write("<tr><td></td><td>" & rs.Fields("Narrative") & "</td></tr>" & vbCrLf)
			rs.MoveNext
		Wend
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
	End If

End If
%>
<br />
<div style="text-align: center; "><b>Revenue</b></div> 
<p>Indicate Source of Cash and In-Kind Matches for the proposed program. Click on links to go to 
	match detail pages for entry of data.</p>
<%

sql = "SELECT A.Source, B.MatchSource, A.Amount " & vbCrLf & _
	"FROM " & ApplicationSchema & ".Matches AS A " & vbCrLf & _
	"LEFT JOIN Lookup.MatchSources AS B ON B.MatchSourceID=A.MatchSourceID " & vbCrLf & _
	"WHERE A.MatchTypeID=1 AND A.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"ORDER BY A.MatchID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<table style=""margin: auto; width: 500px; "">" & vbCrLf)
	Response.Write("<thead><tr><th colspan=""3"">Source of Cash Match</th></tr></thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr><td>" & rs.Fields("Source") & "</td>" & vbCrLf)
		Response.Write("<td>" & rs.Fields("MatchSource") & "</td>" & vbCrLf)
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("Amount"), RoundCurrency) & "</td></tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("<tfoot><tr><td style=""font-weight: bold; "">Total Cash Match</td><td></td><td style=""text-align: right; "">" & prepCurrencyWebRound(CashMatch, RoundCurrency) & "</td></tr></tfoot>")
	Response.Write("</table>" & vbCrLf)
End If

Response.Write("<br />" & vbCrLf)

sql = "SELECT A.Source, B.MatchSource, A.Amount " & vbCrLf & _
	"FROM " & ApplicationSchema & ".Matches AS A " & vbCrLf & _
	"LEFT JOIN Lookup.MatchSources AS B ON B.MatchSourceID=A.MatchSourceID " & vbCrLf & _
	"WHERE A.MatchTypeID=2 AND A.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"ORDER BY A.MatchID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<table style=""margin: auto; width: 500px; "">" & vbCrLf)
	Response.Write("<thead><tr><th colspan=""3"">Source of In-Kind Match</th></tr></thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr><td>" & rs.Fields("Source") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("MatchSource") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("Amount")) & "</td></tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("<tfoot><tr><td style=""font-weight: bold; "">Total In-Kind Match</td><td></td><td style=""text-align: right; "">" & prepCurrencyWeb(InKindMatch) & "</td></tr></tfoot>")
	Response.Write("</table>" & vbCrLf)
End If
Response.Write("<br />" & vbCrLf)
%>

<br />
<%	End If %>
<div style="text-align: center; font-weight: bold; ">Statistics to Support Grant Problem Statement</div>
<table class="bordertable">
<thead>
	<tr>
		<th>Reported Cases</th>
		<th colspan="3" style="border: solid black thin; "><%=(HistoricalDataYear-1) %></th>
		<th colspan="3" style="border: solid black thin; "><%=(HistoricalDataYear) %></th>
	</tr>
	<tr style="vertical-align: bottom; ">
		<th style="width: 175px; ">Jurisdiction</th>
		<th style="width: 115px; ">Motor Vehicle Theft<br />(MVT)</th>
		<th style="width: 115px; " title="Burglary from Motor Vehicle including theft of parts">Burglary from Motor Vehicle<br />(BMV)</th>
		<th style="width: 115px; ">Fraud-Related Motor Vehicle Crime<br />(FRMVC)</th>
		<th style="width: 115px; ">Motor Vehicle Theft<br />(MVT)</th>
		<th style="width: 115px; " title="Burglary from Motor Vehicle including theft of parts">Burglary from Motor Vehicle<br />(BMV)</th>
		<th style="width: 115px; ">Fraud-Related Motor Vehicle Crime<br />(FRMVC)</th>
	</tr>
</thead>
<tbody>
<%
sql = "SELECT StatisticsID, AppID, Jurisdiction, MVT1, BMV1, FRMVC1, MVT2, BMV2, FRMVC2 " & vbCrLf & _
	"FROM " & ApplicationSchema & ".[Statistics] " & vbCrLF & _
	"WHERE AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"ORDER BY AppID, StatisticsID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write(vbTab & "<tr><td colspan=""7"">&nbsp;</td></tr>" & vbCrLf)
	Response.Write(vbTab & "<tr><th colspan=""7""><i>No Statistical Data has been entered yet.</i></th></tr>" & vbCrLf)
Else
While rs.EOF = False
	Response.Write(vbtab & "<tr style=""vertical-align: top; "">" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td>" & rs.Fields("Jurisdiction") & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("MVT1")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("BMV1")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("FRMVC1")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("MVT2")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("BMV2")) & "</td>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<td style=""text-align: right; "">" & formatInteger(rs.Fields("FRMVC2")) & "</td>" & vbCrLf)
	Response.Write(vbtab & "</tr>" & vbCrLf)
	rs.MoveNext
Wend
End If
%>
</tbody>
</table>

<br />
<div style="width: 976px; text-align: center; font-weight: bold; ">Application Narrative</div>

<table style="width: 100%">
<%
sql = "SELECT A.TextSectionID, A.Section, A.SubSection, A.SectionTitle, A.QuestionPreText, " & vbCrLf & _
	"	A.Question, B.SectionText, A.SpecialTreatments " & vbCrLf & _
	"FROM Lookup.TextSections AS A " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".SectionText AS B ON A.TextSectionID=B.TextSectionID AND B.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"WHERE A.Version=2 " & vbCrLf & _
	"ORDER BY Section, SubSection "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
j = 0
While rs.EOF = False
	If rs.Fields("SpecialTreatments") = "GSA" Then 
		Response.Write("<tr><td colspan=""2"">" & vbCrLF)
		Response.Write("<div style=""font-weight: bold"">Part II</div>" & vbCrLf)
		Response.Write("<div style=""text-align: center; font-weight: bold;"">Goals, Strategies, and Activities</div>")
		Response.Write("<p style=""text-align: left"">Select Goals, Strategies, and Activity Targets for the proposed program.</p>")
		Response.Write("<p>Click on the link above and select the method by which statutory measures will be collected. Law Enforcement programs must also estimate targets for the MVCPA predetermined activities. The MVCPA board has determined that grants programs must document specific activities that are appropriate under each of the three goals. Applicants are allowed to write a limited number of user defined activities.</p></td></tr>")
		Response.Write("<tr><td colspan=""2"">")
		outputGSA()
		Response.Write("</td></tr>")
	End If
	if j <> rs.Fields("Section") Then
		j= rs.Fields("Section")
		Response.Write("<tr><th colspan=""2"">" & rs.Fields("SectionTitle") & "</th></tr>" & vbCrLf)
	End If
	If IsNull(rs.Fields("QuestionPreText")) = False Then
		Response.Write("<tr style=""vertical-align: top""><td colspan=""2"">" & rs.Fields("QuestionPreText") & "</td></tr>" & vbCrLf)
	End If
	Response.Write("<tr style=""vertical-align: top""><td>" & rs.Fields("Section") & "." & rs.Fields("SubSection")  & "</td>" & vbCrLf)
	Response.Write("<td>" & rs.Fields("Question") & "</td></tr>" & vbCrLf)
	Response.Write("<tr><td></td><td><span class=""usertext"">" & textarea2html(rs.Fields("SectionText")) & "</span></td></tr>" & vbCrLf)
	Response.Write("<tr><td colspan=""2"">&nbsp;</td></tr>" & vbCrLf)
	rs.MoveNext()
Wend
%>
</table>

<div style="width: 976px; text-align: center; font-weight: bold; ">TxGMS Standard Assurances by Local Governments</div>
<div style="width: 976px; font-weight: normal; "><%=CheckBoxField("Certification", Certification) %> We acknowledge reviewing the 
<a href="../RFA/UniformAssurances.pdf" target="_blank" class="plainlink">TxGMS Standard Assurances by Local Governments</a> as 
promulgated by the Texas Comptroller of Public Accounts and agree to abide by the terms stated therein.</div>
<br />
<%
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

<div style="text-align: center; font-weight: bold; ">Certifications</div>

<p>The certifying official is the authorized official, <%=AuthorizedOfficial %>, <%=AuthorizedOfficialTitle %>.</p>

<p>By submitting this application I certify that I have been designated by my jurisdiction as the authorized 
official to accept the terms and conditions of the grant. 
The statements herein are true, complete, and accurate to the best of my knowledge. I am aware that any false, 
fictitious, or fraudulent statements or claims may subject me to criminal, civil, or administrative penalties.</p>

<p>By submitting this application I certify that my jurisdiction agrees to comply with all terms and conditions if the grant 
is awarded and accepted. I further certify that my jurisdiction will comply with all applicable state and federal laws, rules 
and regulations in the application, acceptance, administration and operation of this grant.</p>

</form>

</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>

</body>
</html>
<%
Function outputGSA()
	Response.Write("<table style=""margin: auto""><thead><tr><th>ID</th><th>Activity</th><th>Measure</th><th>Target</th></tr></thead>" & vbCrLf)
	Dim vrs, vsql, LastMandatory, LastGoal, LastStrategy, vVersion
	If FiscalYear>= 2022 Then
		vVersion = 5
	End If
	vsql = "SELECT G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
		"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
		"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
		"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.ResponseTypeID, " & vbCrLf & _
		"	T.IntegerResponse, T.DecimalResponse " & vbCrLf & _
		"FROM Lookup.Goals AS G " & vbCrLf & _
		"LEFT JOIN Lookup.Strategies AS S ON S.Version=G.Version AND S.GoalID=G.GoalID " & vbCrLf & _
		"LEFT JOIN Lookup.Activities AS A ON A.Version=G.Version AND A.GoalID=S.GoalID AND A.StrategyID=S.StrategyID " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".GSATargets AS T ON T.AppID=" & prepIntegerSQL(AppID) & " AND T.Version=G.Version AND T.GoalID=G.GoalID AND T.StrategyID=S.StrategyID AND T.ActivityID=A.ActivityID AND T.MeasureID=A.MeasureID " & vbCrLF & _
		"WHERE G.Version=" & prepIntegerSQL(vVersion) & " AND G.GoalID NOT IN (4,5, 6, 7) AND (Mandatory=1 OR NoTarget=0) " & vbCrLf & _
		"ORDER BY A.Mandatory DESC, G.Version, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	LastMandatory = True
	LastGoal=0
	LastStrategy=0
	Set vrs=Con.Execute(vsql)
	While vrs.EOF = False
		If LastMandatory <> vrs.Fields("Mandatory") Then
			LastMandatory = vrs.Fields("Mandatory")
			If LastMandatory = False Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: YellowGreen; "">Measures for Grantees. Add Target values for those that you will measure.</th></tr>" & vbCrLF)
			End If
		End If
		If LastGoal <> vrs.Fields("GoalID") And vrs.Fields("Mandatory") = False Then
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			LastGoal = vrs.Fields("GoalID")
			Response.Write("<td style=""text-align: right; "">" & vrs.Fields("GoalID") & "</td>" & vbCrLf)
			Response.Write("<th colspan=""3"" style=""background-color: PowderBlue;"">Goal " & vrs.Fields("GoalID") & ": " & vrs.Fields("Goal") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		ElseIf LastGoal <> vrs.Fields("GoalID") And vrs.Fields("Mandatory") = True Then
			LastGoal = vrs.Fields("GoalID")
			If vrs.Fields("GoalID") = 1 Then
				Response.Write("<tr style=""vertical-align: top; ""><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Statutory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLF)
			ElseIf vrs.Fields("GoalID")=2 Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Statutory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLF)
			ElseIf vrs.Fields("GoalID")=8 Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Fraud must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Statutory Fraud-Related Motor Vehicle Crime Measures Required for all Grantees</th></tr>" & vbCrLF)
			End If
		End If
		If LastStrategy <> vrs.Fields("StrategyID") And vrs.Fields("Mandatory") = False  Then
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			LastStrategy = vrs.Fields("StrategyID")
			Response.Write("<td style=""text-align: right; "">" & vrs.Fields("GoalID") & "." & vrs.Fields("StrategyID") & "</td>" & vbCrLf)
			Response.Write("<th colspan=""3"" style=""background-color: PeachPuff; "">Strategy " & vrs.Fields("StrategyID") & ": " & vrs.Fields("Strategy") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & vrs.Fields("MeasureNumber") & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td>" & vrs.Fields("Activity") & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td>" & vrs.Fields("Measure") & "</td>" & vbCrLF)
		If vrs.Fields("Mandatory") Then
			Response.Write(vbTab & "<td class=""usertext""></td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=1 Then
			Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & vrs.Fields("IntegerResponse") & "</td>" & vbCrLF)
		ElseIf vrs.Fields("ResponseTypeID")=2 Then
			Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & formatnumber(vrs.Fields("DecimalResponse")) & "</td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=3 Then
				Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & formatnumber(vrs.Fields("DecimalResponse")) & "</td>" & vbCrLF)
		End If
		Response.Write("</tr>" & vbCrLf)
		vrs.MoveNext()
	Wend
	Response.Write("</table>" & vbCrLf)
	Response.Write("<br />")
End Function

function textarea2html(vText)
	if IsNull(vText) = true Then
		textarea2html = null
	ElseIf Len(vText)=0 Then
		textarea2html = ""
	Else
		textarea2html = Replace(vText, vbCrLf, "<br />")
	End If
end function
%>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->