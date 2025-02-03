<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, TimeStamp, Button, SectionTextID, SectionText, _
	ApplicationSchema, AppID, FiscalYear, GranteeID, ProgramName, CoverageAreaDescription, _
	GrantTypeID, StatewideCoverage, OtherCoverage, OtherCoverageText, LawEnforcementGrant, _
	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, _
	ProgramCategory1, ProgramCategory2, ProgramCategory3, ProgramCategory5, _
	ParticipatingAgencies, CoverageAgencies, List, _
	HistoricalDataYear, LarcenyFromMV1, LarcenyFromMV2, LarcenyFromMV3, _
	LarcenyFromMVParts1, LarcenyFRomMVParts2, LarcenyFromMVParts3, LarcenyJurisdiction, _
	MVT1, MVT2, MVT3, RecoveryMVT1, RecoveryMVT2, RecoveryMVT3, MVTJurisdiction, DataProblems, _
	BudgetEntryOption, BudgetCashMatch, Certification,  _
	ParticipatingAgenciesChanged, CoverageAgenciesChanged, Changes, ChangesArray, foundchange
ReDim ProgramCategory(5)
TimeStamp = Now()

debug = False
'PermitEdit = False
ApplicationSchema = "Negotiation"

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
AppID = CInt(Request.Form("AppID"))
GranteeID = Request.Form("GranteeID")
FiscalYear = Request.Form("FiscalYear")
ProgramName = Request.Form("ProgramName")
CoverageAreaDescription = Request.Form("CoverageAreaDescription")
GrantTypeID = Request.Form("GrantTypeID")
StatewideCoverage = Request.Form("StatewideCoverage")
ParticipatingAgencies = Request.Form("ParticipatingAgencies")
CoverageAgencies = Request.Form("CoverageAgencies")
OtherCoverage = Request.Form("OtherCoverage")
OtherCoverageText = Request.Form("OtherCoverageText")
ProgramCategory1 = Request.Form("ProgramCategory1")
ProgramCategory2 = Request.Form("ProgramCategory2")
ProgramCategory3 = Request.Form("ProgramCategory3")
ProgramCategory5 = Request.Form("ProgramCategory5")
HistoricalDataYear = Request.Form("HistoricalDataYear")
LawEnforcementGrant = Request.Form("LawEnforcementGrant")
NationalInsuranceCrimeBureau = Request.Form("NationalInsuranceCrimeBureau")
TexasDepartmentOfPublicSafety = Request.Form("TexasDepartmentOfPublicSafety")
OtherAgency = Request.Form("OtherAgency")
OtherAgencySpecify = Request.Form("OtherAgencySpecify")
BudgetEntryOption = Request.Form("BudgetEntryOption")
Certification = Request.Form("Certification")
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
		sql = "INSERT INTO Application.IDs (GrantClassID, FiscalYear, GranteeID) VALUES (4, " & _
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
		sql = "INSERT INTO CC." & ApplicationSchema & " (AppID, ProgramName, CoverageAreaDescription, " & vbCrLf & _
			"	GrantTypeID, StatewideCoverage, OtherCoverage, OtherCoverageText, ProgramCategory1, " & vbCrLf & _
			"	ProgramCategory2, ProgramCategory3, ProgramCategory5, HistoricalDataYear, LawEnforcementGrant, " & vbCrLf & _
			"	NationalInsuranceCrimeBureau, TexasDepartmentOfPublicSafety, OtherAgency, OtherAgencySpecify, " & vbCrLf & _
			"	BudgetCashMatch, Certification, " & vbCrLf
		If Button="submit" Then
			sql = sql & "SubmitID, SubmitTimestamp, ConfirmationNumber, "
		End If
		sql = sql & "	UpdateID, UpdateTimestamp) " & vbCrLf & _
			"VALUES (" & _
			prepIntegerSQL(AppID) & ", " & _
			prepStringSQL(ProgramName) & ", " & _
			prepStringSQL(CoverageAreaDescription) & ", " & _
			prepIntegerSQL(GrantTypeID) & ", " & _
			prepBitRequiredSQL(StatewideCoverage) & ", " & _
			prepBitRequiredSQL(OtherCoverage) & ", " & _
			prepStringSQL(OtherCoverageText) & ", " & _
			prepBitRequiredSQL(ProgramCategory1) & ", " & _
			prepBitRequiredSQL(ProgramCategory2) & ", " & _
			prepBitRequiredSQL(ProgramCategory3) & ", " & _
			prepBitRequiredSQL(ProgramCategory5) & ", " & _
			prepIntegerSQL(HistoricalDataYear) & ", " & _
			prepBitRequiredSQL(LawEnforcementGrant) & ", " & _
			prepBitRequiredSQL(NationalInsuranceCrimeBureau) & ", " & _
			prepBitRequiredSQL(TexasDepartmentOfPublicSafety) & ", " & _
			prepBitRequiredSQL(OtherAgency) & ", " & _
			prepStringSQL(OtherAgencySpecify) & ", " & _
			prepNumberSQL(BudgetCashMatch) & ", " & _
			prepBitSQL(Certification) & ", "
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
		sql = "UPDATE CC." & ApplicationSchema & " " & vbCrLf & "SET " & _
				"ProgramName=" & prepStringSQL(ProgramName) & ", " & _
				"CoverageAreaDescription=" & prepStringSQL(CoverageAreaDescription) & ", " & _
				"StatewideCoverage=" & prepBitRequiredSQL(StatewideCoverage) & ", " & _
				"OtherCoverage=" & prepBitRequiredSQL(OtherCoverage) & ", " & _
				"OtherCoverageText=" & prepStringSQL(OtherCoverageText) & ", " & _
				"ProgramCategory1=" & prepBitRequiredSQL(ProgramCategory1) & ", " & _
				"ProgramCategory2=" & prepBitRequiredSQL(ProgramCategory2) & ", " & _
				"ProgramCategory3=" & prepBitRequiredSQL(ProgramCategory3) & ", " & _
				"ProgramCategory5=" & prepBitRequiredSQL(ProgramCategory5) & ", " & _
				"HistoricalDataYear=" & prepIntegerSQL(HistoricalDataYear) & ", " & _
				"LawEnforcementGrant=" & prepBitRequiredSQL(LawEnforcementGrant) & ", " & _
				"NationalInsuranceCrimeBureau=" & prepBitRequiredSQL(NationalInsuranceCrimeBureau) & ", " & _
				"TexasDepartmentOfPublicSafety=" & prepBitRequiredSQL(TexasDepartmentOfPublicSafety) & ", " & _
				"OtherAgency=" & prepBitRequiredSQL(OtherAgency) & ", " & _
				"OtherAgencySpecify=" & prepStringSQL(OtherAgencySpecify) & ", " & _
				"BudgetCashMatch=" & prepNumberSQL(BudgetCashMatch) & ", " & _
				"Certification=" & prepBitSQL(Certification) & ", "
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

		sql = "DELETE FROM " & ApplicationSchema & ".ParticipatingAgencies WHERE AppID=" & prepIntegerSQL(AppID) 
		If Len(ParticipatingAgencies)>0 Then
			sql = sql & " AND ORI NOT IN (" & List & ")"
		End If
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)

		For i = 1 To Request.Form("ParticipatingAgencies").Count
			sql = "SELECT AppID, ORI FROM " & ApplicationSchema & ".ParticipatingAgencies WHERE AppID=" & _
				prepIntegerSQL(AppID) & " AND ORI=" & prepStringSQL(Request.Form("ParticipatingAgencies")(i))
			If Debug = True Then
				Response.Write("<pre>Item" & i & ": " & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = con.execute(sql)
			IF rs.EOF = True Then
				sql = "INSERT INTO " & ApplicationSchema & ".ParticipatingAgencies(AppID, ORI, UpdateID, UpdateTimeStamp) VALUES (" & _
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
		sql = "DELETE FROM " & ApplicationSchema & ".ParticipatingAgencies WHERE AppID=" & prepIntegerSQL(AppID) 
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

		sql = "DELETE FROM " & ApplicationSchema & ".CoverageAgencies WHERE AppID=" & prepIntegerSQL(AppID) 
		If Len(CoverageAgencies)>0 Then
			sql = sql & " AND ORI NOT IN (" & List & ")"
		End If
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)

		For i = 1 To Request.Form("CoverageAgencies").Count
			sql = "SELECT AppID, ORI FROM " & ApplicationSchema & ".CoverageAgencies WHERE AppID=" & _
				prepIntegerSQL(AppID) & " AND ORI=" & prepStringSQL(Request.Form("CoverageAgencies")(i))
			If Debug = True Then
				Response.Write("<pre>Item" & i & ": " & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = con.execute(sql)
			IF rs.EOF = True Then
				sql = "INSERT INTO " & ApplicationSchema & ".CoverageAgencies(AppID, ORI, UpdateID, UpdateTimeStamp) VALUES (" & _
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
		sql = "DELETE FROM " & ApplicationSchema & ".CoverageAgencies WHERE AppID=" & prepIntegerSQL(AppID) 
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
		"LEFT JOIN " & ApplicationSchema & ".SectionText AS B WITH (NOLOCK) ON A.TextSectionID=B.TextSectionID AND B.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
		"WHERE A.Version=3 " & vbCrLf & _
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
				sql = "INSERT INTO " & ApplicationSchema & ".SectionText (TextSectionID, AppID, SectionText, UpdateID, UpdateTimestamp)" & vbCrLf & _
					"	VALUES (" & SectionTextID & ", " & prepIntegerSQL(AppID) & ", " & _
					prepStringSQL(SectionText) & ", " & prepIntegerSQL(UserSystemID) & ", " & _
					prepStringSQL(Timestamp) & ")"
				If Debug = True Then
					Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				Con.Execute(sql)
			ElseIf Len(SectionText) > 0 and rs.Fields("RecordPresent") = True Then
				sql = "UPDATE " & ApplicationSchema & ".SectionText SET " & vbCrLf & _
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
				sql = "DELETE FROM " & ApplicationSchema & ".SectionText " & vbCrLf & _
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
		"UPDATE " & ApplicationSchema & ".BudgetDetails " & vbCrLf & _
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
			Response.Write("<a href=""" & ApplicationSchema & "asp?GranteeID=" & GranteeID & "&AppID=" & AppID & _
				"&FiscalYear=" & FiscalYear & """>return</a>")
		Else
			Response.Redirect(ApplicationSchema & ".asp?GranteeID=" & GranteeID & "&AppID=" & AppID & _
				"&FiscalYear=" & FiscalYear)
		End If
ElseIf Button = "submit" or Button="home" Then
	If Debug = True Then
		Response.Write("<a href=""/Home/Default.asp?GranteeID=" & GranteeID & """>home</a>")
	Else
		Response.Redirect("../Home/Default.asp?GranteeID=" & GranteeID)
	End If
ElseIf Button = "GSA" Then
	If Debug = True Then
		Response.Write("<a href=""/" & ApplicationSchema & "/gsa.asp?AppID=" & AppID & "&GrantClassID=4"">Goals Strategies and Activities</a>")
	Else
		Response.Redirect("/" & ApplicationSchema & "/gsa.asp?AppID=" & AppID & "&GrantClassID=4")
	End If
ElseIf Button="CashMatch" Then
	If Debug = True Then
		Response.Write("<a href=""/" & ApplicationSchema & "/Matches.asp?AppID=" & AppID & "&MatchTypeID=1"">Cash Match</a>")
	Else
		Response.Redirect("/" & ApplicationSchema & "/Matches.asp?AppID=" & AppID & "&MatchTypeID=1")
	End If
ElseIf Button="InKindMatch" Then
	If Debug = True Then
		Response.Write("<a href=""/" & ApplicationSchema & "/Matches.asp?AppID=" & AppID & "&MatchTypeID=2"">Cash Match</a>")
	Else
		Response.Redirect("/" & ApplicationSchema & "/Matches.asp?AppID=" & AppID & "&MatchTypeID=2")
	End If
ElseIf Button="Statistics" Then
	If Debug = True Then
		Response.Write("<a href=""Statistics.asp?AppID=" & AppID & "&ApplicationSchema=Negotiation"">Cash Match</a>")
	Else
		Response.Redirect("Statistics.asp?AppID=" & AppID & "&ApplicationSchema=Negotiation")
	End If
ElseIf IsNumeric(Button) Then
	If BudgetEntryOption = "1" Then
		If Debug = True Then
			Response.Write("<a href=""/" & ApplicationSchema & "/BudgetDetail.asp?AppID=" & AppID & _
				"&BudgetCategoryID=" & CInt(Button) & """>budget detail</a>")
		Else
			Response.Redirect("/" & ApplicationSchema & "/BudgetDetail.asp?AppID=" & AppID & _
				"&BudgetCategoryID=" & CInt(Button))
		End If
	Else
		If Debug = True Then
			Response.Write("<a href=""/" & ApplicationSchema & "/BudgetDetail2.asp?AppID=" & AppID & _
				"&BudgetCategoryID=" & CInt(Button) & """>budget detail</a>")
		Else
			Response.Redirect("/" & ApplicationSchema & "/BudgetDetail2.asp?AppID=" & AppID & _
				"&BudgetCategoryID=" & CInt(Button))
		End If
	End If
End If
%><!--#include file="../includes/prepDB.asp"-->