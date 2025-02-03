﻿<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, LastCategory, PermitEdit, AppID, FiscalYear, Version, GranteeID, GranteeName, _
	AuthorizedOfficialID, AuthorizedOfficial, AuthorizedOfficialTitle, _
	ProgramName, GrantTypeID, _
	StatewideCoverage, OtherCoverage, OtherCoverageText, LawEnforcementGrant, _
	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, _
	HistoricalDataYear,  LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, _
	LarcenyFromMVParts1, LarcenyFRomMVParts2, LarcenyFromMVParts3, LarcenyJurisdiction, _
	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, DataProblems, _
	SubmitID, SubmitByName, SubmitTimestamp, ConfirmationNumber, _
	CashMatch, InKindMatch, GrandTotal, TotalMVCPAFunds, TotalCashMatch, TotalInkindMatch, _
	PctMVCPA, PctCashMatch, TargetAwardAmount, TargetMatchAmount, TargetReimbursementRate, _
	BudgetEntryOption, BudgetCashMatch, RoundCurrency
Dim ProgramCategory(5)

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

AppID = Request.QueryString("AppID")
GranteeID = Request.QueryString("GranteeID")
FiscalYear = Request.QueryString("FiscalYear")

If AppID="" Then
	AppID=0
	If GranteeID="" Then
		GranteeID = Session("GranteeID")
	End If
	If GranteeID="" Or GranteeID=0 Then
		Response.Write("Error: No AppID or GranteeID Specified")
		Response.End
	End If
Else
	AppID=Cint(AppID)
End If

If AppID>0 Then 
	sql = "SELECT G.GranteeID, G.GranteeName, ISNULL(A.FiscalYear, " & prepIntegerSQL(FiscalYear) & _
		") AS FiscalYear, G.AuthorizedOfficialID, " & vbCrLf & _
		"	ISNULL(AO.Name, 'Authorized Official') AS AuthorizedOfficial, " & vbCrLf & _
		"	ISNULL(AO.Title, 'Authorized Official Title') AS AuthorizedOfficialTitle, " & vbCrLf & _
		"	ISNULL(A.AppID,0) AS AppID, ProgramName, " & vbCrLf & _
		"	A.GrantTypeID, A.StatewideCoverage, A.OtherCoverage, A.OtherCoverageText, A.LawEnforcementGrant, " & vbCrLf & _
		"	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, " & vbCrLf & _
		"	ProgramCategory1, ProgramCategory2, ProgramCategory3, ProgramCategory4, ProgramCategory5, " & vbCrLf & _
		"	HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, " & vbCrLf & _
		"	LarcenyFromMVParts1, LarcenyFRomMVParts2, LarcenyFromMVParts3, " & vbCrLf & _
		"	LarcenyJurisdiction, DataProblems, " & vbCrLf & _
		"	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, " & vbCrLf & _
		"	A.SubmitID, U.Name AS SubmitByName, A.SubmitTimestamp, A.ConfirmationNumber, " & vbCrLf & _
		"	N.AwardAmount AS TargetAwardAmount, N.MatchAmount AS TargetMatchAmount, N.ReimbursementRate AS TargetReimbursementRate " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"LEFT JOIN Negotiation.Main AS A ON A.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN System.Users AS U ON U.SystemID=A.SubmitID " & vbCrLf & _
		"LEFT JOIN System.Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf & _
		"LEFT JOIN Negotiation.TargetAmounts AS N ON N.AppID=A.AppID " & vbCrLf & _
		"WHERE A.AppID=" & PrepIntegerSQL(AppID)
Else
	sql = "SELECT G.GranteeID, G.GranteeName, ISNULL(A.FiscalYear, " & prepIntegerSQL(FiscalYear) & _
		") AS FiscalYear, G.AuthorizedOfficialID, " & vbCrLf & _
		"	ISNULL(AO.Name, 'Authorized Official') AS AuthorizedOfficial, " & vbCrLf & _
		"	ISNULL(AO.Title, 'Authorized Official Title') AS AuthorizedOfficialTitle, " & vbCrLf & _
		"	ISNULL(A.AppID,0) AS AppID, ProgramName, " & vbCrLf & _
		"	A.GrantTypeID, A.StatewideCoverage, A.OtherCoverage, A.OtherCoverageText, A.LawEnforcementGrant, " & vbCrLf & _
		"	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, " & vbCrLf & _
		"	ProgramCategory1, ProgramCategory2, ProgramCategory3, ProgramCategory4, ProgramCategory5, " & vbCrLf & _
		"	HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, " & vbCrLf & _
		"	LarcenyFromMVParts1, LarcenyFRomMVParts2, LarcenyFromMVParts3, " & vbCrLf & _
		"	LarcenyJurisdiction, DataProblems, " & vbCrLf & _
		"	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, " & vbCrLf & _
		"	A.SubmitID, U.Name AS SubmitByName, A.SubmitTimestamp, A.ConfirmationNumber " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"LEFT JOIN Negotiation.Main AS A ON A.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN System.Users AS U ON U.SystemID=A.SubmitID " & vbCrLf & _
		"LEFT JOIN System.Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf & _
		"LEFT JOIN Negotiation.TargetAmounts AS N ON N.AppID=A.AppID " & vbCrLf & _
		"WHERE G.GranteeID=" & PrepIntegerSQL(GranteeID)
End If

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Grantee and Application record retrieved")
	Response.End
Else
	AppID = rs.Fields("AppID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	AuthorizedOfficialID = rs.Fields("AuthorizedOfficialID")
	AuthorizedOfficial = rs.Fields("AuthorizedOfficial")
	AuthorizedOfficialTitle = rs.Fields("AuthorizedOfficialTitle")
	ProgramName = rs.Fields("ProgramName")
	GrantTypeID = rs.Fields("GrantTypeID")
	For i = 1 to 5
		ProgramCategory(i) = rs.Fields("ProgramCategory" & i)
	Next
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
	SubmitID = rs.Fields("SubmitID")
	SubmitByName = rs.Fields("SubmitByName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	ConfirmationNumber = rs.Fields("ConfirmationNumber")
	TargetAwardAmount = rs.Fields("TargetAwardAmount")
	TargetMatchAmount = rs.Fields("TargetMatchAmount")
	TargetReimbursementRate = rs.Fields("TargetReimbursementRate")
End If

' Start rounding dollar amounts as of 2020.
If FiscalYear>=2020 Then
	RoundCurrency = True
Else
	RoundCurrency = False
End If

If FiscalYear>= 2021 Then
	Version = 4
ElseIf FiscalYear>= 2020 Then
	Version = 2
ElseIf FiscalYear>= 2018 Then
	Version = 2
Else
	Version = 1
End If

If GranteeID>0 Then
	If IsNull(SubmitID) = True Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
	ElseIf ISNull(SubmitID) = False Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, True)
	Else
		PermitEdit = False
	End If
Else
		PermitEdit = False
End If

If IsNull(HistoricalDataYear) Then
	HistoricalDataYear = FiscalYear - 2
End If

sql = "SELECT ISNULL(SUM(CASE WHEN MatchTypeID=1 THEN Amount ELSE NULL END),0) AS CashMatch, " & vbCrLf & _
	"	ISNULL(SUM(CASE WHEN MatchTypeID=2 THEN Amount ELSE NULL END),0) AS InKindMatch " & vbCrLf & _
	"FROM Negotiation.Matches " & vbCrLf & _
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
<html lang="en-us" moznomarginboxes mozdisallowselectionprint>
<head>
<title><%=GranteeName %> Grant Application Negotiation for Fiscal Year <%=FiscalYear %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 96%; margin: auto">

<div class="sectiontitle"><%=GranteeName %> Grant Application Negotiation for Fiscal Year <%=FiscalYear %></div>


<form name="Application">

<table>
<%	If SubmitID>0 Then %>
<tr><td colspan="2" style="text-align: center; font-weight: bold; ">The Application was submitted by <%=SubmitByName%> at <%=SubmitTimestamp %> and is now locked.<br />
	The confirmation Number is <%=ConfirmationNumber %>.</td></tr>
<tr><td colspan="2">&nbsp;</td></tr>
<%	End If %>

<tr>
	<td colspan="2"><b>Program Title</b> Please enter a short description of the proposed program that can be used as the title.
	<span class="usertext"><%=ProgramName%></span>
	</td>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2">Which type of grant are you applying for?</td>
</tr>
<%
	sql = "SELECT GrantTypeID, GrantType, GrantTypeDescription FROM Lookup.GrantType WHERE Version=1 ORDER BY GrantTypeID "
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		If GrantTypeID=rs.Fields("GrantTypeID") Or IsNull(GrantTypeID) Then
		Response.Write(vbTab & "<tr style=""vertical-align: top""><td></td><td><b>" & _
			rs.Fields("GrantType") & "</b> - " & replace(rs.Fields("GrantTypeDescription"),"{PreviousYear}",(FiscalYear-1)) & "</td></tr>" & vbCrLf)
		End If
		rs.MoveNext
	Wend
	If IsNull(GrantTypeID) = True then
		Response.Write("<tr><td></td><td style=""font-weight: bold; "">Note: Grant Type not selected on application!</td></tr>" & vbCrLf)
	End If
%>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
	<td colspan="2">To be eligible for consideration for funding, a program must be designed to 
	support one or more of the following <b>MVCPA program categories</b>.</td>
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
	<td colspan="2"><b>Grant Participation and Coverage Area</b></td>
</tr>
<%	If StatewideCoverage=True Then %>
<tr>
	<td>&bullet;</td>
	<td><b>Statewide Coverage</b></td>
</tr>
<%	End If 
	If OtherCoverage = True Then%>
<tr style="vertical-align: top; ">
	<td>&bullet;</td>
	<td><b>Other Coverage</b> (Describe): 
	<span class="usertext"><%=OtherCoverageText %></span></td>
</tr>
<%	End If 
	If LawEnforcementGrant = True Then %>
<tr style="vertical-align: top; ">
	<td>&bullet;</td>
	<td><b>Law Enforcement Grant</b><br />
	Participating and coverage agencies below.
	</td>
</tr>

<tr style="vertical-align: top; ">
	<td></td>
	<td><p><b>Participating Agencies</b>: agencies that will materially 
	participate in the grant application through the use of interlocal agreements. The agencies 
	selected in this list only includes agencies that will receive or provide funding and/or 
	resources. The interlocal agreements do not need to be submitted with the application. 
	Interlocal agreements will need to be executed prior to the first payment being made if 
	selected for a grant.   Letters of support with the application from the participating 
	agencies are strongly recommended.</p>
	<p><b>Coverage Agencies</b>: agencies that will be covered by the grant but not 
	materially participating in the grant application. These agencies will not be covered by 
	a grant interlocal agreement but as law enforcement agencies may have jurisdictional 
	coverage agreements unrelated to the grant. The agencies selected in this list only 
	includes agencies that will be covered or where the chief of police or county sheriff 
	indicates that their agency will coordinate or call upon the taskforce. These will not 
	directly receive or provide funding and/or resources. Letters of support with the 
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
	sql = "SELECT A.ORI, REPLACE(B.Agency,'&','&amp;') AS Agency" & vbCrLf & _
		"FROM Negotiation.ParticipatingAgencies AS A" & vbCrLf & _
		"LEFT JOIN Lookup.ORI AS B ON B.ORI=A.ORI " & vbCrLf & _
		"WHERE A.AppID = " & prepIntegerSQL(AppID) & vbCrLf & _
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
	sql = "SELECT A.ORI, REPLACE(B.Agency,'&','&amp;') AS Agency" & vbCrLf & _
		"FROM Negotiation.CoverageAgencies AS A" & vbCrLf & _
		"LEFT JOIN Lookup.ORI AS B ON B.ORI=A.ORI " & vbCrLf & _
		"WHERE A.AppID = " & prepIntegerSQL(AppID) & vbCrLf & _
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
<%	End If 
	If NationalInsuranceCrimeBureau = True Then%>
<tr style="vertical-align: top; ">
	<td>&bullet;</td>
	<td><b>National Insurance Crime Bureau (NICB)</b> Used as Match (Documentation and time certification required.)</td>
</tr>
<%	End If 
	If TexasDepartmentOfPublicSafety = True Then %>
<tr style="vertical-align: top; ">
	<td>&bullet;</td>
	<td><b>Texas Department of Public Safety (DPS)</b> Used as Match (Documentation and time certification required.)</td>
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

<div style="margin: auto"><b>Resolution</b>: Complete a Resolution and submit to local governing body 
	for approval. <a href="../Application/Resolution.asp?AppID=<%=AppID %>&GranteeID=<%=GranteeID %>&FiscalYear=<%=FiscalYear%>" target="_blank" class="plainlink">Sample Resolution</a> 
	is found in the Request for Application or send a request for an electronic copy to 
	<a href="mailto:grantsMVCPA@txdmv.gov?subject=Resolution Request" class="plainlink">grantsMVCPA@txdmv.gov</a>.
</div>

<br />

<div style="margin: auto; text-align: center; font-weight: bold; ">Grant Budget Summary</div>

<table style="margin: auto; ">
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
	"LEFT JOIN Negotiation.BudgetDetails AS B ON A.BudgetCategoryID=B.BudgetCategoryID AND B.AppID=" & _
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
			TotalMVCPAFunds = rs.Fields("MVCPAFunds")
			TotalCashMatch = rs.Fields("CashMatch")
			TotalInkindMatch = rs.Fields("InkindMatch")
			GrandTotal = rs.Fields("LineTotal")
		Else
			Response.Write(vbTab & "<td>" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)  
		End If
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(rs.Fields("MVCPAFunds"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(rs.Fields("CashMatch"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(rs.Fields("LineTotal"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWebRound(rs.Fields("InKindMatch"), RoundCurrency) & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>")
		rs.MoveNext
	Wend
	If TotalMVCPAFunds>0 Then
		PctMVCPA = 100* TotalMVCPAFunds / TotalMVCPAFunds
		PctCashMatch = 100 * TotalCashMatch / TotalMVCPAFunds
		Response.Write("<tr><td></td><td style=""text-align: right; ""><!--" & prepNumberWeb(PctMVCPA, 2) & _
			"%--></td><td style=""text-align: right; "">" & prepNumberWeb(PctCashMatch, 2) & "%</td><td></td><td></td></tr>" & vbCrLf)
	End If

	If IsNull(TargetAwardAmount) = False And IsNull(TotalMVCPAFunds) = False Then
		If TotalMVCPAFunds > TargetAwardAmount Then
			Response.Write(vbTab & "<tr><td colspan=""5"">&nbsp;</td></tr>" & vbCrLf)
			Response.Write(vbTab & "<tr><td colspan=""5"" style=""text-align: center; color: red; "">The maximum award for this grant is " & formatcurrencyRound(TargetAwardAmount, RoundCurrency) & ".</td></tr>" & vbCrLf)
			'Response.Write(vbTab & "<tr><td colspan=""5"" style=""text-align: center; font-style: italic; "">The maximum award amount for this grant is " & formatCurrencyRound(TargetAwardAmount, RoundCurrency) & ", the minimum cash match amount is " & formatCurrencyRound(TargetMatchAmount, RoundCurrency) & ".</td></tr>" & vbCrLf)
		Else
			Response.Write(vbTab & "<tr><td colspan=""5"" style=""text-align: center; "">The maximum award for this grant is " & formatcurrencyRound(TargetAwardAmount, RoundCurrency) & ".</td></tr>" & vbCrLf)
		End If
	End If
	If IsNull(TargetMatchAmount) = False And IsNull(TotalCashMatch) = False Then
		If TotalCashMatch < TargetMatchAmount Then
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
<%	If AppID>0 Then 
sql = "SELECT B.BudgetItemID, A.BudgetCategoryID, A.BudgetCategory, " & vbCrLf & _
	"	CASE WHEN B.NoOfItems>0 THEN ISNULL(B.Description,'') + ' (' + CAST(B.NoOfItems AS VARCHAR) + ')' ELSE B.Description END AS Description, " & vbCrLf & _
	"	SubCategory, PctTime, LineTotal, MVCPAFunds, CashMatch, InKindMatch, ISNULL(UnallowedItem,0) AS UnallowedItem, " & vbCrLf & _
	"	CASE WHEN AllowedAmount IS NULL THEN 0 WHEN AllowedAmount>LineTotal THEN 1 ELSE 0 END AS AllowedAmount, ISNULL(Issue, 0) AS Issue " & vbCrLf & _
	"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
	"LEFT JOIN Negotiation.BudgetDetails AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"LEFT JOIN Lookup.BudgetSubcategories AS C ON C.BudgetCategoryID=B.BudgetCategoryID AND C.SubCategoryID=B.SubCategoryID " & vbCrLf & _
	"UNION " & vbCrLf & _
	"SELECT 2147483647 AS BudgetItemID, A.BudgetCategoryID, A.BudgetCategory, null AS PctTime, 'Total ' + A.BudgetCategory AS Description, " & vbCrLf & _
	"	null, SUM(LineTotal) AS LineTotal, SUM(MVCPAFunds) AS MVCPAFunds, SUM(CashMatch) AS CashMatch, Sum(InKindMatch) AS InKindMatch, 0 AS UnallowedItem, " & vbCrLf & _
	"	null AS AllowedAmount, 0 As Issue " & vbCrLf & _
	"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
	"LEFT JOIN Negotiation.BudgetDetails AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"GROUP BY A.BudgetCategoryID, A.BudgetCategory" & vbCrLf & _
	"ORDER BY 2, 1 "
LastCategory=0
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<table style=""margin: auto; "">" & vbCrLf)
	Response.Write("<thead><tr>" & vbCrLf)
	Response.Write("<th>Description</th>" & vbCrLf)
	Response.Write("<th>Subcategory</th>" & vbCrLf)
	Response.Write("<th>Pct Time</th>" & vbCrLf)
	Response.Write("<th>MVCPA Funds</th>" & vbCrLf)
	Response.Write("<th>Cash Match</th>" & vbCrLf)
	Response.Write("<th>Total</th>" & vbCrLf)
	Response.Write("<th>In-Kind Match</th>" & vbCrLf)
	Response.Write("</tr></thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		If LastCategory<>rs.Fields("BudgetCategoryID") Then
			LastCategory=rs.Fields("BudgetCategoryID")
			Response.Write("<tr><td colspan=""6"">&nbsp;</td></tr>" & vbCrLf)
			Response.Write("<tr><th colspan=""6"">" & rs.Fields("BudgetCategory") & "</th></tr>" & vbCrLf)
		End If
		If rs.Fields("UnallowedItem") = 1 Then
			Response.Write("<tr style=""background-color: Yellow; "">" & vbCrLf)
		ElseIf rs.Fields("Issue") = 1 Then
			Response.Write("<tr style=""background-color: Yellow; "">" & vbCrLf)
		ElseIf rs.Fields("AllowedAmount") = 1 Then
			Response.Write("<tr style=""background-color: Yellow; "">" & vbCrLf)
		Else
			Response.Write("<tr>" & vbCrLf)
		End If
		Response.Write("<td>" & rs.Fields("Description") & "</td>")
		Response.Write("<td>" & rs.Fields("SubCategory") & "</td>")
		If (rs.Fields("BudgetCategoryID")=1 Or rs.Fields("BudgetCategoryID")=4) And IsNull(rs.Fields("PctTime")) = False Then
			Response.Write("<td style=""text-align: right; "">" & prepNumberWeb(rs.Fields("PctTime"),2) & "%</td>")
		Else
			Response.Write("<td></td>")
		End If
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("MVCPAFunds"), RoundCurrency) & "</td>")
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("CashMatch"), RoundCurrency) & "</td>")
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("LineTotal"), RoundCurrency) & "</td>")
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWebRound(rs.Fields("InKindMatch"), RoundCurrency) & "</td>")
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
End If

sql = "SELECT A.BudgetCategoryID, A.BudgetCategory, B.Narrative " & vbCrLf & _
	"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
	"JOIN Negotiation.BudgetDetailNarrative AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND B.AppID=" & prepIntegerSQL(AppID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<br /><table>" & vbCrLf)
	Response.Write("<thead><tr><th>Budget Narrative</td></tr></thead>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr><td><b>" & rs.Fields("BudgetCategory") & "</b>: " & textarea2html(rs.Fields("Narrative")) & "</td></tr>" & vbCrLf)
		rs.MoveNext()
	Wend
	Response.Write("</table>" & vbCrLf)
End If
%>
<br />
<div style="text-align: center; "><b>Revenue</b></div> 
<p>Indicate Source of Cash and In-Kind Matches for the proposed program.</p>
<%
Response.Write(vbTab & "<div style=""text-align: center"">Cash Match</div>" & vbCrLf)  

sql = "SELECT A.Source, B.MatchSource, A.Amount " & vbCrLf & _
	"FROM Negotiation.Matches AS A " & vbCrLf & _
	"LEFT JOIN Lookup.MatchSources AS B ON B.MatchSourceID=A.MatchSourceID " & vbCrLf & _
	"WHERE A.MatchTypeID=1 AND A.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"ORDER BY A.MatchID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<table style=""margin: auto; width: 650px; "">" & vbCrLf)
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
	Response.Write(vbTab & "<div style=""text-align: center"">In-Kind Match</div>" & vbCrLf)  

sql = "SELECT A.Source, B.MatchSource, A.Amount " & vbCrLf & _
	"FROM Negotiation.Matches AS A " & vbCrLf & _
	"LEFT JOIN Lookup.MatchSources AS B ON B.MatchSourceID=A.MatchSourceID " & vbCrLf & _
	"WHERE A.MatchTypeID=2 AND A.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"ORDER BY A.MatchID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<table style=""margin: auto; width: 650px; "">" & vbCrLf)
	Response.Write("<thead><tr><th colspan=""3"">Source of In-Kind Match</th></tr></thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr><td>" & rs.Fields("Source") & "</td>" & vbCrLf)
		Response.Write("<td>" & rs.Fields("MatchSource") & "</td>" & vbCrLf)
		Response.Write("<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("Amount")) & "</td></tr>" & vbCrLf)
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

<table style="margin: auto; ">
	<thead>
	<tr>
		<td style="text-align: center; ">Use UCR data</td>
		<th><%=(HistoricalDataYear-2) %></th>
		<th><%=(HistoricalDataYear-1) %></th>
		<th><%=(HistoricalDataYear) %></th>
	</tr>
	</thead>
	<tbody>
	<tr>
		<td style="font-weight: bold; ">Larceny from a motor vehicle</td>
		<td style="text-align: right; "><%=prepIntegerWeb(LarcenyFromMV1) %></td>
		<td style="text-align: right; "><%=prepIntegerWeb(LarcenyFromMV2) %></td>
		<td style="text-align: right; "><%=prepIntegerWeb(LarcenyFromMV3) %></td>
	</tr>
	<tr>
		<td style="font-weight: bold; ">Larceny from a motor vehicle - Parts</td>
		<td style="text-align: right; "><%=prepIntegerWeb(LarcenyFromMVParts1) %></td>
		<td style="text-align: right; "><%=prepIntegerWeb(LarcenyFromMVParts2) %></td>
		<td style="text-align: right; "><%=prepIntegerWeb(LarcenyFromMVParts3) %></td>
	</tr>
	<tr>
		<td style="font-weight: bold; ">Jurisdictions included in totals</td>
		<td colspan="3" style="text-align: center; ">
<%
	If LarcenyJurisdiction= 0 Then
		Response.Write("Select Jurisdiction")
	ElseIf LarcenyJurisdiction=1 Then
		Response.Write("Statistics for Taskforce Only")
	ElseIf LarcenyJurisdiction=2 Then
		Response.Write("Statistics for Area of Jurisdiction")
	ElseIf LarcenyJurisdiction=3 Then
		Response.Write("Statistics a combination of Taskforce and Jurisdiction")
	End If
%></td>
	</tr>
	<tr>
		<td style="font-weight: bold; ">Theft of a motor vehicle</td>
		<td style="text-align: right; "><%=prepIntegerWeb(MVT1) %></td>
		<td style="text-align: right; "><%=prepIntegerWeb(MVT2) %></td>
		<td style="text-align: right; "><%=prepIntegerWeb(MVT3) %></td>
	</tr>
	<tr>
		<td style="font-weight: bold; ">Recoveries of Motor Vehicles</td>
		<td style="text-align: right; "><%=prepIntegerWeb(RecoveryMVT1) %></td>
		<td style="text-align: right; "><%=prepIntegerWeb(RecoveryMVT2) %></td>
		<td style="text-align: right; "><%=prepIntegerWeb(RecoveryMVT3) %></td>
	</tr>
	<tr>
		<td style="font-weight: bold; ">Jurisdictions included in totals</td>
		<td colspan="3" style="text-align: center; "><%
	If MVTJurisdiction= 0 Then
		Response.Write("Select Jurisdiction")
	ElseIf MVTJurisdiction=1 Then
		Response.Write("Statistics for Taskforce Only")
	ElseIf MVTJurisdiction=2 Then
		Response.Write("Statistics for Area of Jurisdiction")
	ElseIf MVTJurisdiction=3 Then
		Response.Write("Statistics a combination of Taskforce and Jurisdiction")
	End If
%></td>
	</tr>
	<tr>
		<td colspan="4">Provide any additional information or limitations about the data provide above<br /><span class="usertext"><%=DataProblems%></span></td>
	</tr>
	</tbody>
</table>
<br />
<div style="text-align: center; font-weight: bold; ">Application Narrative</div>

<table>
<%
sql = "SELECT A.TextSectionID, A.Section, A.SubSection, A.SectionTitle, A.QuestionPreText, " & vbCrLf & _
	"	A.Question, B.SectionText " & vbCrLf & _
	"FROM Lookup.TextSections AS A " & vbCrLf & _
	"LEFT JOIN Negotiation.SectionText AS B ON A.TextSectionID=B.TextSectionID AND B.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"WHERE A.TextSectionID<13 AND A.Version=1 " & vbCrLf & _
	"ORDER BY Section, SubSection "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
j = 0
While rs.EOF = False
	If rs.Fields("TextSectionID") = 10 Then 
		Response.Write("<tr><td colspan=""2"">" & vbCrLf)
		Response.Write("<div style=""font-weight: bold"">Part II</div>" & vbCrLf)
		Response.Write("<div style=""text-align: center; font-weight: bold;"">Goals, Strategies, and Activities</div>")
		Response.Write("<p style=""text-align: left"">Select Goals, Strategies, and Activity Targets for the proposed program.</p>")
		Response.Write("<p>Click on the link above and select the method by which statutory measures will be collected. Law Enforcement programs must also estimate targets for the MVCPA predetermined activities. The MVCPA board has determined that grants programs must document specific activities that are appropriate under each of the three goals. Applicants are allowed to write a limited number of user defined activities.</p></td></tr>")
		Response.Write("<tr><td colspan=""2"">")
		outputGSA()
		Response.WRite("</td></tr>")
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

<%
If AppID > 0 Then
	Dim DocumentFolder, fso, folder, files, file
	DocumentFolder = Application("DocumentRoot") & "\Application\" & AppID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemOBject")
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		If files.count>0 Then 
			Response.Write("<div style=""width: 600px; margin: auto; ""><div style=""text-align: center; font-weight: bold; "">Current Documents in folder</div>" & vbCrLf)
		For Each file in files
				Response.Write("<a href=""../Documents/Application/" & AppID & "/" & file.Name & _
					""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
		Next
			Response.Write("<br /></div>" & vbCrLf)
		End If
	End If
End If
%>

<div style="text-align: center; font-weight: bold; ">Certifications</div>

<p>The certifying official is the authorized official, <%=AuthorizedOfficial %>, <%=AuthorizedOfficialTitle %>.</p>

<p>By submitting this application I certify that I have been designated by my jurisdiction as the authorized official to accept the terms and conditions of the grant. The statements herein are true, complete, and accurate to the best of my knowledge. I am aware that any false, fictitious, or fraudulent statements or claims may subject me to criminal, civil, or administrative penalties.</p>

<p>By submitting this application I certify that my jurisdiction agrees to comply with all terms and conditions if the grant is awarded and accepted. I further certify that my jurisdiction will comply with all applicable state and federal laws, rules and regulations in the application, acceptance, administration and operation of this grant.</p>

</form>


</body>
</html>
<%
Function outputGSA()
	Response.Write("<table style=""margin: auto""><thead><tr><th>ID</th><th>Activity</th><th>Measure</th><th>Target</th></tr></thead>" & vbCrLf)
	Dim vrs, vsql, LastMandatory, LastGoal, LastStrategy
	vsql = "SELECT G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
		"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
		"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
		"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.ResponseTypeID, " & vbCrLf & _
		"	T.IntegerResponse, T.DecimalResponse " & vbCrLf & _
		"FROM Lookup.Goals AS G " & vbCrLf & _
		"LEFT JOIN Lookup.Strategies AS S ON S.GoalID=G.GoalID " & vbCrLf & _
		"LEFT JOIN Lookup.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID " & vbCrLf & _
		"LEFT JOIN Negotiation.GSATargets AS T ON T.AppID=" & prepIntegerSQL(AppID) & " AND T.GoalID=G.GoalID AND T.StrategyID=S.StrategyID AND T.ActivityID=A.ActivityID AND T.MeasureID=A.MeasureID " & vbCrLf & _
		"WHERE T.IntegerResponse IS NOT NULL OR T.DecimalResponse IS NOT NULL " & vbCrLf & _
		"ORDER BY A.Mandatory DESC, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
vsql = "SELECT G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.ResponseTypeID, " & vbCrLf & _
	"	T.IntegerResponse, T.DecimalResponse " & vbCrLf & _
	"FROM Lookup.Goals AS G " & vbCrLf & _
	"LEFT JOIN Lookup.Strategies AS S ON S.Version=G.Version AND S.GoalID=G.GoalID " & vbCrLf & _
	"LEFT JOIN Lookup.Activities AS A ON A.Version=G.Version AND A.GoalID=S.GoalID AND A.StrategyID=S.StrategyID " & vbCrLf & _
	"LEFT JOIN Negotiation.GSATargets AS T ON T.AppID=" & prepIntegerSQL(AppID) & " AND T.Version=G.Version AND T.GoalID=G.GoalID AND T.StrategyID=S.StrategyID AND T.ActivityID=A.ActivityID AND T.MeasureID=A.MeasureID " & vbCrLf & _
	"WHERE G.Version=" & prepIntegerSQL(Version) & " " & vbCrLf & _
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
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: YellowGreen; "">Measures for Grantees. Add Target values for those that you will measure.</th></tr>" & vbCrLf)
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
				Response.Write("<tr style=""vertical-align: top; ""><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Mandatory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLf)
			ElseIf vrs.Fields("GoalID")=2 Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Mandatory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLf)
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
		Response.Write(vbTab & "<td style=""text-align: right; "">" & vrs.Fields("MeasureNumber") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & vrs.Fields("Activity") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & vrs.Fields("Measure") & "</td>" & vbCrLf)
		If vrs.Fields("Mandatory") Then
			Response.Write(vbTab & "<td class=""usertext"">Mandatory. Reporting for ")
			If vrs.Fields("IntegerResponse") = 0 Then
				Response.Write("Select Jurisdiction")
			ElseIf vrs.Fields("IntegerResponse") = 1 Then
				Response.Write("Taskforce Only")
			ElseIf vrs.Fields("IntegerResponse") = 2 Then
				Response.Write("Area of Jurisdiction")
			ElseIf vrs.Fields("IntegerResponse") = 3 Then
				Response.Write("Combination of TF and Jurisdiction")
			End If
			Response.Write("</td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=1 Then
			Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & vrs.Fields("IntegerResponse") & "</td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=2 Then
			Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & formatnumber(vrs.Fields("DecimalResponse")) & "</td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=3 Then
				Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & formatnumber(vrs.Fields("DecimalResponse")) & "</td>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
		vrs.MoveNext()
	Wend
	Response.Write("</table>" & vbCrLf)
End Function

function textarea2html(vText)
	if IsNull(vText) = true Then
		textarea2html = null
	ElseIf Len(vText)=0 Then
		textarea2html = ""
	Else
		textarea2html = Replace(vText, vbCrLf&vbCrLf, "<br /><br />")
	End If
end function
%>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->