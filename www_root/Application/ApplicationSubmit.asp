<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, TimeStamp, Button, SectionTextID, SectionText, _
	AppID, FiscalYear, GranteeID, ProgramName, _
	GrantTypeID, StatewideCoverage, OtherCoverage, OtherCoverageText, LawEnforcementGrant, _
	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, _
	ProgramCategory1, ProgramCategory2, ProgramCategory3, ProgramCategory4, ProgramCategory5, _
	ParticipatingAgencies, CoverageAgencies, List, _
	HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, _
	LarcenyFromMVParts1, LarcenyFRomMVParts2, LarcenyFromMVParts3, LarcenyJurisdiction, _
	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, DataProblems, _
	BudgetEntryOption, BudgetCashMatch, _
	ParticipatingAgenciesChanged, CoverageAgenciesChanged, Changes, ChangesArray, foundchange, ApplicationSchema
ReDim ProgramCategory(5)
TimeStamp = Now()

debug = False
ApplicationSchema = "Application"

'PermitEdit = False
If Debug = True Then
	Response.Write("<!DOCTYPE html><pre>Dubugging Information: " & vbCrLf)
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

Button = Request.Form("Button")
AppID = Request.Form("AppID")
GranteeID = Request.Form("GranteeID")
FiscalYear = Request.Form("FiscalYear")
ProgramName = Request.Form("ProgramName")
GrantTypeID = Request.Form("GrantTypeID")
ProgramCategory1 = Request.Form("ProgramCategory1")
ProgramCategory2 = Request.Form("ProgramCategory2")
ProgramCategory3 = Request.Form("ProgramCategory3")
ProgramCategory4 = Request.Form("ProgramCategory4")
ProgramCategory5 = Request.Form("ProgramCategory5")
StatewideCoverage = Request.Form("StatewideCoverage")
ParticipatingAgencies = Request.Form("ParticipatingAgencies")
CoverageAgencies = Request.Form("CoverageAgencies")
OtherCoverage = Request.Form("OtherCoverage")
OtherCoverageText = Request.Form("OtherCoverageText")
LawEnforcementGrant = Request.Form("LawEnforcementGrant")
NationalInsuranceCrimeBureau = Request.Form("NationalInsuranceCrimeBureau")
TexasDepartmentOfPublicSafety = Request.Form("TexasDepartmentOfPublicSafety")
OtherAgency = Request.Form("OtherAgency")
OtherAgencySpecify = Request.Form("OtherAgencySpecify")
HistoricalDataYear = Request.Form("HistoricalDataYear")
LarcenyFromMV1 = Request.Form("LarcenyFromMV1")
LarcenyFromMV2 = Request.Form("LarcenyFromMV2")
LarcenyFromMV3 = Request.Form("LarcenyFromMV3")
LarcenyFromMVParts1 = Request.Form("LarcenyFromMVParts1")
LarcenyFromMVParts2 = Request.Form("LarcenyFromMVParts2")
LarcenyFromMVParts3 = Request.Form("LarcenyFromMVParts3")
LarcenyJurisdiction = Request.Form("LarcenyJurisdiction")
MVT1 = Request.Form("MVT1")
MVT2 = Request.Form("MVT2")
MVT3 = Request.Form("MVT3")
RecoveryMVT1 = Request.Form("RecoveryMVT1")
RecoveryMVT2 = Request.Form("RecoveryMVT2")
RecoveryMVT3 = Request.Form("RecoveryMVT3")
MVTJurisdiction = Request.Form("MVTJurisdiction")
DataProblems = Request.Form("DataProblems")
BudgetEntryOption = Request.Form("BudgetEntryOption")
If BudgetEntryOption = "1" Then
	BudgetCashMatch = null
Else
	BudgetCashMatch = Request.Form("BudgetCashMatch")
End If
If Request.Form("ParticipatingAgenciesChanged") = "1" Then
	ParticipatingAgenciesChanged = True
ElseIf Request.Form("ParticipatingAgenciesChanged") = "0" Then
	ParticipatingAgenciesChanged = False
Else ' Just to be safe in transition. If variable not present, then do updates.
	ParticipatingAgenciesChanged = True 
End If
If Request.Form("CoverageAgenciesChanged") = "1" Then
	CoverageAgenciesChanged = True
ElseIf Request.Form("CoverageAgenciesChanged") = "0" Then
	CoverageAgenciesChanged = False
Else ' Just to be safe in transition. If variable not present, then do updates.
	CoverageAgenciesChanged = True 
End If
Changes = Request.Form("Changes")

' Are there changes other than the question text questions?
ChangesArray=Split(Changes, vbCrLf)
foundchange = false
For Each i in ChangesArray
	if Len(i)>0 And InStr(i,"Question_")=0 Then
		foundchange=true
	End If
	If Debug = True Then
		Response.Write("<pre>")
		Response.Write("value='" & i & "' " & " Changed (InStr(question_)=" & InStr(i,"question_") & ")" & vbCrLf)
		Response.Write("</pre>")
	End If
Next
If Debug = True Then
	Response.Write(vbCrLf & "<pre>foundchange=" & foundchange & "</pre>" & vbCrLf)
End If


If AppID=0 Or Button="submit" Or foundchange = true Then
	If AppID = 0 Then
		sql = "INSERT INTO Application.IDs (GrantClassID, FiscalYear, GranteeID) VALUES (1, " & _
			prepIntegerSQL(FiscalYear) & ", " & _
			prepIntegerSQL(GranteeID) & ")"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
		sql = "SELECT IDENT_CURRENT('Application.IDs') AS AppID"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		If rs.EOF = True Then
			Response.Write("Error: No Grantee and Application record retrieved")
			sendWarning("Error: No Grantee and Application record retrieved")
			Response.End
		Else
			AppID = CLng(rs.Fields("AppID"))
			If Debug = True Then
				Response.Write("<pre>AppID='" & AppID & "'</pre>" & vbCrLf)
				Response.Write("<pre>IsNumeric(AppID)=" & IsNumeric(AppID) & "</pre>" & vbCrLf)
				Response.Flush
			End If
		End If
		sql = "INSERT INTO Application.Main (AppID, ProgramName, GrantTypeID, " & vbCrLf & _
			"	StatewideCoverage, OtherCoverage, OtherCoverageText, LawEnforcementGrant, " & vbCrLf & _
			"	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, " & vbCrLf & _
			"	ProgramCategory1, ProgramCategory2, ProgramCategory3, ProgramCategory4, ProgramCategory5, " & vbCrLf & _
			"	HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, " & vbCrLf & _
			"	LarcenyFromMVParts1, LarcenyFRomMVParts2, LarcenyFromMVParts3, LarcenyJurisdiction, " & vbCrLf & _
			"	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, DataProblems, " & vbCrLf & _
			"	BudgetCashMatch, " & vbCrLf
		If Button="submit" Then
			sql = sql & "SubmitID, SubmitTimestamp, ConfirmationNumber, "
		End If
		sql = sql & "	UpdateID, UpdateTimestamp) " & vbCrLf & _
			"VALUES (" & prepIntegerSQL(AppID) & ", " & vbCrLf & _
			prepStringSQL(ProgramName) & ", " & _
			prepIntegerSQL(GrantTypeID) & ", " & _
			prepBitRequiredSQL(StatewideCoverage) & ", " & _
			prepBitRequiredSQL(OtherCoverage) & ", " & _
			prepStringSQL(OtherCoverageText) & ", " & _
			prepBitRequiredSQL(LawEnforcementGrant) & ", " & _
			prepBitRequiredSQL(NationalInsuranceCrimeBureau) & ", " & _
			prepBitRequiredSQL(TexasDepartmentOfPublicSafety) & ", " & _
			prepBitRequiredSQL(OtherAgency) & ", " & _
			prepStringSQL(OtherAgencySpecify) & ", " & _
			prepBitRequiredSQL(ProgramCategory1) & ", " & _
			prepBitRequiredSQL(ProgramCategory2) & ", " & _
			prepBitRequiredSQL(ProgramCategory3) & ", " & _
			prepBitRequiredSQL(ProgramCategory4) & ", " & _
			prepBitRequiredSQL(ProgramCategory5) & ", " & _
			prepIntegerSQL(HistoricalDataYear) & ", " & _
			prepIntegerSQL(LarcenyFromMV1) & ", " & _
			prepIntegerSQL(LarcenyFromMV2) & ", " & _
			prepIntegerSQL(LarcenyFromMV3) & ", " & _
			prepIntegerSQL(LarcenyFromMVParts1) & ", " & _
			prepIntegerSQL(LarcenyFRomMVParts2) & ", " & _
			prepIntegerSQL(LarcenyFromMVParts3) & ", " & _
			prepIntegerSQL(LarcenyJurisdiction) & ", " & _
			prepIntegerSQL(MVT1) & ", " & _
			prepIntegerSQL(MVT2) & ", " & _
			prepIntegerSQL(MVT3) & ", " & _
			prepIntegerSQL(RecoveryMVT1) & ", " & _
			prepIntegerSQL(RecoveryMVT2) & ", " & _
			prepIntegerSQL(RecoveryMVT3) & ", " & _
			prepIntegerSQL(MVTJurisdiction) & ", " & _
			prepStringSQL(DataProblems) & ", " & _
			prepNumberSQL(BudgetCashMatch) & ", "
		If Button="submit" Then
			sql = sql & prepIntegerSQL(UserSystemID) & ", " & _
			prepStringSQL(Timestamp) & ", " & _
			prepStringSQL(Year(Timestamp) & Right("00" & Month(Timestamp),2) & Right("00" & Day(Timestamp),2) & Right("00000" & AppID,5))
		End If
			sql = sql & prepIntegerSQL(UserSystemID) & ", " & _
			prepStringSQL(Timestamp) & ")"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	Else
		sql = "UPDATE Application.Main " & vbCrLf & "SET " & _
				"ProgramName=" & prepStringSQL(ProgramName) & ", " & _
				"GrantTypeID=" & prepIntegerSQL(GrantTypeID) & ", " & _
				"StatewideCoverage=" & prepBitRequiredSQL(StatewideCoverage) & ", " & _
				"OtherCoverage=" & prepBitRequiredSQL(OtherCoverage) & ", " & _
				"OtherCoverageText=" & prepStringSQL(OtherCoverageText) & ", " & _
				"LawEnforcementGrant=" & prepBitRequiredSQL(LawEnforcementGrant) & ", " & _
				"NationalInsuranceCrimeBureau=" & prepBitRequiredSQL(NationalInsuranceCrimeBureau) & ", " & _
				"TexasDepartmentOfPublicSafety=" & prepBitRequiredSQL(TexasDepartmentOfPublicSafety) & ", " & _
				"OtherAgency=" & prepBitRequiredSQL(OtherAgency) & ", " & _
				"OtherAgencySpecify=" & prepStringSQL(OtherAgencySpecify) & ", " & _
				"ProgramCategory1=" & prepBitRequiredSQL(ProgramCategory1) & ", " & _
				"ProgramCategory2=" & prepBitRequiredSQL(ProgramCategory2) & ", " & _
				"ProgramCategory3=" & prepBitRequiredSQL(ProgramCategory3) & ", " & _
				"ProgramCategory4=" & prepBitRequiredSQL(ProgramCategory4) & ", " & _
				"ProgramCategory5=" & prepBitRequiredSQL(ProgramCategory5) & ", " & _
				"HistoricalDataYear=" & prepIntegerSQL(HistoricalDataYear) & ", " & _
				"LarcenyFromMV1=" & prepIntegerSQL(LarcenyFromMV1) & ", " & _
				"LarcenyFromMV2=" & prepIntegerSQL(LarcenyFromMV2) & ", " & _
				"LarcenyFromMV3=" & prepIntegerSQL(LarcenyFromMV3) & ", " & _
				"LarcenyFromMVParts1=" & prepIntegerSQL(LarcenyFromMVParts1) & ", " & _
				"LarcenyFRomMVParts2=" & prepIntegerSQL(LarcenyFRomMVParts2) & ", " & _
				"LarcenyFromMVParts3=" & prepIntegerSQL(LarcenyFromMVParts3) & ", " & _
				"LarcenyJurisdiction=" & prepIntegerSQL(LarcenyJurisdiction) & ", " & _
				"MVT1=" & prepIntegerSQL(MVT1) & ", " & _
				"MVT2=" & prepIntegerSQL(MVT2) & ", " & _
				"MVT3=" & prepIntegerSQL(MVT3) & ", " & _
				"RecoveryMVT1=" & prepIntegerSQL(RecoveryMVT1) & ", " & _
				"RecoveryMVT2=" & prepIntegerSQL(RecoveryMVT2) & ", " & _
				"RecoveryMVT3=" & prepIntegerSQL(RecoveryMVT3) & ", " & _
				"MVTJurisdiction=" & prepIntegerSQL(MVTJurisdiction) & ", " & _
				"DataProblems=" & prepStringSQL(DataProblems) & ", " & _
				"BudgetCashMatch=" & prepNumberSQL(BudgetCashMatch) & ", "
		If Button = "submit" Then
			sql = sql &	"SubmitID=" & prepIntegerSQL(UserSystemID) & ", " & _
				"SubmitTimestamp=" & prepStringSQL(Timestamp) & ", " & _
				"ConfirmationNumber=" & prepStringSQL(Year(Timestamp) & Right("00" & Month(Timestamp),2) & Right("00" & Day(Timestamp),2) & Right("00000" & AppID,5)) & ", "
		End If
		sql = sql &	"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
				"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
			"WHERE AppID=" & prepIntegerSQL(AppID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
End If

If Debug = True Then
	Response.Write("<pre>AppID='" & AppID & "'</pre>" & vbCrLf)
	Response.Write("<pre>IsNumeric(AppID)=" & IsNumeric(AppID) & "</pre>" & vbCrLf)
	Response.Flush
End If

' Update Participating Agencies. Delete Existing and then add.
If ParticipatingAgenciesChanged = True Then
	ParticipatingAgencies = Request.Form("ParticipatingAgencies")
	If Len(ParticipatingAgencies)>0 Then
		List = ""
		For i = 1 to Request.Form("ParticipatingAgencies").Count
			List = List & "'" & Request.Form("ParticipatingAgencies")(i) & "', "
		Next
		List = Mid(List, 1, Len(List)-2)

		sql = "DELETE FROM Application.ParticipatingAgencies WHERE AppID=" & prepIntegerSQL(AppID) 
		If Len(ParticipatingAgencies)>0 Then
			sql = sql & " AND ORI NOT IN (" & List & ")"
		End If
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)

		For i = 1 To Request.Form("ParticipatingAgencies").Count
			sql = "SELECT AppID, ORI FROM Application.ParticipatingAgencies WHERE AppID=" & _
				prepIntegerSQL(AppID) & " AND ORI=" & prepStringSQL(Request.Form("ParticipatingAgencies")(i))
			If Debug = True Then
				Response.Write("<pre>Item" & i & ": " & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = con.execute(sql)
			IF rs.EOF = True Then
				sql = "INSERT INTO Application.ParticipatingAgencies(AppID, ORI, UpdateID, UpdateTimeStamp) VALUES (" & _
					AppID & ", " & prepStringSQL(Request.Form("ParticipatingAgencies")(i)) & ", " & prepIntegerSQL(UserSystemID) & _
					", " & prepStringSQL(TimeStamp) & ")"
				If Debug = True Then
					Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				con.execute(sql)
			End If
		Next
	Else
		sql = "DELETE FROM Application.ParticipatingAgencies WHERE AppID=" & prepIntegerSQL(AppID) 
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)
	End If
End If

' Update Coverage Agencies. Delete Existing and then add.
If CoverageAgenciesChanged = True Then
	CoverageAgencies = Request.Form("CoverageAgencies")
	If Len(CoverageAgencies)>0 Then
		List = ""
		For i = 1 to Request.Form("CoverageAgencies").Count
			List = List & "'" & Request.Form("CoverageAgencies")(i) & "', "
		Next
		List = Mid(List, 1, Len(List)-2)

		sql = "DELETE FROM Application.CoverageAgencies WHERE AppID=" & prepIntegerSQL(AppID) 
		If Len(CoverageAgencies)>0 Then
			sql = sql & " AND ORI NOT IN (" & List & ")"
		End If
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)

		For i = 1 To Request.Form("CoverageAgencies").Count
			sql = "SELECT AppID, ORI FROM Application.CoverageAgencies WHERE AppID=" & _
				prepIntegerSQL(AppID) & " AND ORI=" & prepStringSQL(Request.Form("CoverageAgencies")(i))
			If Debug = True Then
				Response.Write("<pre>Item" & i & ": " & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = con.execute(sql)
			IF rs.EOF = True Then
				sql = "INSERT INTO Application.CoverageAgencies(AppID, ORI, UpdateID, UpdateTimeStamp) VALUES (" & _
					AppID & ", " & prepStringSQL(Request.Form("CoverageAgencies")(i)) & ", " & prepIntegerSQL(UserSystemID) & _
					", " & prepStringSQL(TimeStamp) & ")"
				If Debug = True Then
					Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				con.execute(sql)
			End If
		Next
	Else
		sql = "DELETE FROM Application.CoverageAgencies WHERE AppID=" & prepIntegerSQL(AppID) 
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)
	End If
End If

' Update the Text Sections
' The Changes variable will hold a list of fields that have changed. See if "Question_" is in the list.
If InStr(1,Changes, "Question_")>0 Then
	If Debug = True Then
		Response.Write("<pre>Changes Found in Section Text</pre>" & vbCrLf)
	End If

	sql = "SELECT A.TextSectionID, B.SectionText, CAST(CASE WHEN B.TextSectionID IS NULL THEN 0 ELSE 1 END AS BIT) AS RecordPresent " & vbCrLf & _
		"FROM Lookup.TextSections AS A WITH (NOLOCK) " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".SectionText AS B WITH (NOLOCK)ON A.TextSectionID=B.TextSectionID AND B.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
		"WHERE A.Version=2 " & vbCrLf & _
		"ORDER BY Section, SubSection "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		SectionTextID = rs.Fields("TextSectionID")
		If InStr(1,Changes, "Question_" & SectionTextID)>0 Then
			SectionText = Request.Form("Question_" & SectionTextID)
			If Debug = True Then
				Response.Write("<pre>SectionText_" & SectionTextID & "=""" & SectionText & """" & "</pre>")
			End If

			If SectionText = "" And rs.Fields("SectionText") = False Then
				' Do nothing
			ElseIf SectionText = rs.Fields("SectionText") Then
				' Do nothing
			ElseIf Len(SectionText)>0 And rs.Fields("RecordPresent") = False Then
				sql = "INSERT INTO Application.SectionText (TextSectionID, AppID, SectionText, UpdateID, UpdateTimestamp)" & vbCrLf & _
					"	VALUES (" & SectionTextID & ", " & prepIntegerSQL(AppID) & ", " & _
					prepStringSQL(SectionText) & ", " & prepIntegerSQL(UserSystemID) & ", " & _
					prepStringSQL(Timestamp) & ")"
				If Debug = True Then
					Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				Con.Execute(sql)
			ElseIf Len(SectionText) > 0 and rs.Fields("RecordPresent") = True Then
				sql = "UPDATE Application.SectionText SET " & vbCrLf & _
					"SectionText=" & prepStringSQL(SectionText) & ", " & _
					"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
					"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
					"WHERE TextSectionID=" & SectionTextID &" AND AppID=" & prepIntegerSQL(AppID)
				If Debug = True Then
					Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				Con.Execute(sql)
			ElseIf Len(SectionText)=0 And IsNull(rs.Fields("SectionText")) = False Then 
				sql = "DELETE FROM Application.SectionText " & vbCrLf & _
					"WHERE TextSectionID=" & SectionTextID &" AND AppID=" & prepIntegerSQL(AppID)
				If Debug = True Then
					Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				Con.Execute(sql)
			End If
		End If
		rs.MoveNext()
	Wend
Else
	If Debug = True Then
		Response.Write("<pre>No Changes Found in Section Text</pre>" & vbCrLf)
	End If
End If

' If using percentage method for budget items, be sure to apply percentages to existing budget items.
If FiscalYEar>2019 And BudgetEntryOption="2" Then
	sql = "DECLARE @BudgetCashMatch FLOAT = " & prepNumberSQL(BudgetCashMatch) & "/100.0;" & vbCrLf & _
		"UPDATE Application.BudgetDetails " & vbCrLf & _
		"SET MVCPAFunds = ROUND(ROUND(LineTotal,0) / (1 + @BudgetCashMatch),0), " & vbCrLf & _
		"	CashMatch = ROUND(ROUND(LineTotal,0) * @BudgetCashMatch / (1 + @BudgetCashMatch),0) " & vbCrLf & _
		"WHERE AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
		"	AND (MVCPAFunds <> ROUND(ROUND(LineTotal,0) / (1 + @BudgetCashMatch),0) " & vbCrLf & _
		"	OR (CashMatch <> ROUND(ROUND(LineTotal,0) * @BudgetCashMatch / (1 + @BudgetCashMatch),0)) " & vbCrLf & _
		"	OR LineTotal <> ROUND(LineTotal,0))"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

If AppID=0 Then
	Response.Write("Error: There was an error in retrieving the new AppID. You may wish to go to the menu first to get back to your application. <a href=""../Home/Default.asp?GranteeID=" & GranteeID & """>homepage</a>")
	sendWarning("Error: There was an error in retrieving the new AppID. You may wish to go to the menu first to get back to your application. <a href=""../Home/Default.asp?GranteeID=" & GranteeID & """>homepage</a>")
	Response.End
End If
If Button = "save" Then
	If Debug = True Then
		Response.Write("<a href=""Application.asp?GranteeID=" & GranteeID & "&AppID=" & AppID & _
			"&FiscalYear=" & FiscalYear & """>return</a>")
	Else
		Response.Redirect("Application.asp?GranteeID=" & GranteeID & "&AppID=" & AppID & _
			"&FiscalYear=" & FiscalYear)
	End If
ElseIf Button = "submit" or Button="home" Then
	If Debug = True Then
		Response.Write("<a href=""../Home/Default.asp?GranteeID=" & GranteeID & """>home</a>")
	Else
		Response.Redirect("../Home/Default.asp?GranteeID=" & GranteeID)
	End If
ElseIf Button = "GSA" Then
	If Debug = True Then
		Response.Write("<a href=""gsa.asp?AppID=" & AppID & "&GrantClassID=1"">Goals Strategies and Activities</a>")
	Else
		Response.Redirect("gsa.asp?AppID=" & AppID & "&GrantClassID=1")
	End If
ElseIf Button="CashMatch" Then
	If Debug = True Then
		Response.Write("<a href=""Matches.asp?AppID=" & AppID & "&MatchTypeID=1"">Cash Match</a>")
	Else
		Response.Redirect("Matches.asp?AppID=" & AppID & "&MatchTypeID=1")
	End If
ElseIf Button="InKindMatch" Then
	If Debug = True Then
		Response.Write("<a href=""Matches.asp?AppID=" & AppID & "&MatchTypeID=2"">Cash Match</a>")
	Else
		Response.Redirect("Matches.asp?AppID=" & AppID & "&MatchTypeID=2")
	End If
ElseIf IsNumeric(Button) Then
	If BudgetEntryOption = "1" Then
		If Debug = True Then
			Response.Write("<a href=""BudgetDetail.asp?AppID=" & AppID & _
				"&BudgetCategoryID=" & CInt(Button) & """>budget detail</a>")
		Else
			Response.Redirect("BudgetDetail.asp?AppID=" & AppID & _
				"&BudgetCategoryID=" & CInt(Button))
		End If
	Else
		If Debug = True Then
			Response.Write("<a href=""BudgetDetail2.asp?AppID=" & AppID & _
				"&BudgetCategoryID=" & CInt(Button) & """>budget detail</a>")
		Else
			Response.Redirect("BudgetDetail2.asp?AppID=" & AppID & _
				"&BudgetCategoryID=" & CInt(Button))
		End If
	End If
End If
%><!--#include file="../includes/prepDB.asp"-->