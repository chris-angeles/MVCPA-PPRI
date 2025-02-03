<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, Timestamp, Button, ISAID, FiscalYear, GranteeID, GranteeName, CoverageArea, _
	HistoricalDataYear, MVT1, MVT2, MVT3, MVTLoss1, MVTLoss2, MVTLoss3, _
	BMV1, BMV2, BMV3, BMVLoss1, BMVLoss2, BMVLoss3, _
	ReceiveAuthorization, RegionalTaskForce, CurrentProgram, CurrentProgramDescription, _
	PreviouslyApplied, PreviouslyAwarded, TerminationExplanation, MeetCashMatchRequirement, _
	DedicateResources, NoSupplantation, ProgramCategoryID, BriefNarrative, GrantRequestRangeID, _
	SubmitID, SubmitTimestamp, DateReviewed, DateResponded, MethodRespondedID, Notes, _
	UpdateID, UpdateTimestamp, Unsubmit
debug = False
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
	Response.Write("</pre>" & vbCrLF)
End If

Timestamp = Now()

Button = Request.Form("Button")
ISAID = Request.Form("ISAID")
FiscalYear = Request.Form("FiscalYear")
GranteeID = Request.Form("GranteeID")
GranteeName = Request.Form("GranteeName")
CoverageArea = Request.Form("CoverageArea")
HistoricalDataYear = Request.Form("HistoricalDataYear")
MVT1 = Request.Form("MVT1")
MVT2 = Request.Form("MVT2")
MVT3 = Request.Form("MVT3")
MVTLoss1 = Request.Form("MVTLoss1")
MVTLoss2 = Request.Form("MVTLoss2")
MVTLoss3 = Request.Form("MVTLoss3")
BMV1 = Request.Form("BMV1")
BMV2 = Request.Form("BMV2")
BMV3 = Request.Form("BMV3")
BMVLoss1 = Request.Form("BMVLoss1")
BMVLoss2 = Request.Form("BMVLoss2")
BMVLoss3 = Request.Form("BMVLoss3")
ReceiveAuthorization = Request.Form("ReceiveAuthorization")
RegionalTaskForce = Request.Form("RegionalTaskForce")
CurrentProgram = Request.Form("CurrentProgram")
CurrentProgramDescription = Request.Form("CurrentProgramDescription")
PreviouslyApplied = Request.Form("PreviouslyApplied")
PreviouslyAwarded = Request.Form("PreviouslyAwarded")
TerminationExplanation = Request.Form("TerminationExplanation")
MeetCashMatchRequirement = Request.Form("MeetCashMatchRequirement")
DedicateResources = Request.Form("DedicateResources")
NoSupplantation = Request.Form("NoSupplantation")
ProgramCategoryID = Request.Form("ProgramCategoryID")
BriefNarrative = Request.Form("BriefNarrative")
GrantRequestRangeID = Request.Form("GrantRequestRangeID")
If Button ="submit" Then
	SubmitID = UserSystemID
	SubmitTimestamp = Timestamp
End If
DateReviewed = Request.Form("DateReviewed")
DateResponded = Request.Form("DateResponded")
MethodRespondedID = Request.Form("MethodRespondedID")
Notes = Request.Form("Notes")
If MVCPARights = True And Request.Form("UnSubmit")="1" Then
	UnSubmit = True
	SubmitID = null
	SubmitTimestamp = null
Else
	UnSubmit = False
End If
UpdateID = UserSystemID
UpdateTimestamp = Timestamp

If IsNull(FiscalYear) = True Or Len(FiscalYear)=0 Then
	FiscalYear = 2018
End If

If ISAID = 0 Then
	sql = "INSERT INTO ISA (FiscalYear, GranteeID, CoverageArea, HistoricalDataYear, " & vbCrLF & _
		"	MVT1, MVT2, MVT3, MVTLoss1, MVTLoss2, MVTLoss3, BMV1, BMV2, BMV3, BMVLoss1, BMVLoss2, BMVLoss3, " & vbCrLF & _
		"	ReceiveAuthorization, RegionalTaskForce, CurrentProgram, CurrentProgramDescription, " & vbCrLf & _
		"	PreviouslyApplied, PreviouslyAwarded, TerminationExplanation, MeetCashMatchRequirement, " & vbCrLF & _
		"	DedicateResources, NoSupplantation, ProgramCategoryID, BriefNarrative, GrantRequestRangeID, " & vbCrLF
	If Button="submit" Or UnSubmit = True Then
		sql = sql & "	SubmitID, SubmitTimestamp, " & vbCrLf
	End If
	If MVCPARights = True Then
		sql = sql +	"	DateReviewed, DateResponded, MethodRespondedID, Notes, " & vbCrLf
	End If
		sql = sql & "	UpdateID, UpdateTimestamp)" & vbCrLF & _
		"OUTPUT Inserted.ISAID " & vbCrLf & _
		"VALUES (" & _
		prepIntegerSQL(FiscalYear) & ", " & _
		prepIntegerSQL(GranteeID) & ", " & _
		prepStringSQL(CoverageArea) & ", " & _
		prepIntegerSQL(HistoricalDataYear) & ", " & _
		prepIntegerSQL(MVT1) & ", " & _
		prepIntegerSQL(MVT2) & ", " & _
		prepIntegerSQL(MVT3) & ", " & _
		prepNumberSQL(MVTLoss1) & ", " & _
		prepNumberSQL(MVTLoss2) & ", " & _
		prepNumberSQL(MVTLoss3) & ", " & _
		prepIntegerSQL(BMV1) & ", " & _
		prepIntegerSQL(BMV2) & ", " & _
		prepIntegerSQL(BMV3) & ", " & _
		prepNumberSQL(BMVLoss1) & ", " & _
		prepNumberSQL(BMVLoss2) & ", " & _
		prepNumberSQL(BMVLoss3) & ", " & _
		prepBitSQL(ReceiveAuthorization) & ", " & _
		prepBitSQL(RegionalTaskForce) & ", " & _
		prepBitSQL(CurrentProgram) & ", " & _
		prepStringSQL(CurrentProgramDescription) & ", " & _
		prepBitSQL(PreviouslyApplied) & ", " & _
		prepBitSQL(PreviouslyAwarded) & ", " & _
		PrepStringSQL(TerminationExplanation) & ", " & _
		prepBitSQL(MeetCashMatchRequirement) & ", " & _
		prepBitSQL(DedicateResources) & ", " & _
		prepBitSQL(NoSupplantation) & ", " & _
		prepIntegerSQL(ProgramCategoryID) & ", " & _
		prepStringSQL(BriefNarrative) & ", " & _
		prepIntegerSQL(GrantRequestRangeID) & " , "
		If Button = "submit" or UnSubmit = True Then
			sql = sql + prepIntegerSQL(SubmitID) & ", " & _
			prepStringSQL(SubmitTimestamp) & ", "
		End If
		If MVCPARights = True Then
			sql = sql & prepDateSQL(DateReviewed) & ", " & _
			prepDateSQL(DateResponded) & ", " & _
			prepDateSQL(MethodRespondedID) & ", " & _
			prepStringSQL(Notes) & ", "
		End If
		sql = sql & UserSystemID & ", " & prepStringSQL(UpdateTimestamp) & ") "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		ISAID = rs.Fields("ISAID")
	Else
		Response.Write("Error: No INSERTED.ISAID.")
		Response.End
	End If
Else
	sql = "UPDATE ISA SET " & _
		"FiscalYear=" & prepIntegerSQL(FiscalYear) & ", " & _
		"GranteeID=" & prepIntegerSQL(GranteeID) & ", " & _
		"CoverageArea=" & prepStringSQL(CoverageArea) & ", " & _
		"HistoricalDataYear=" & prepIntegerSQL(HistoricalDataYear) & ", " & _
		"MVT1=" & prepIntegerSQL(MVT1) & ", " & _
		"MVT2=" & prepIntegerSQL(MVT2) & ", " & _
		"MVT3=" & prepIntegerSQL(MVT3) & ", " & _
		"MVTLoss1=" & prepNumberSQL(MVTLoss1) & ", " & _
		"MVTLoss2=" & prepNumberSQL(MVTLoss2) & ", " & _
		"MVTLoss3=" & prepNumberSQL(MVTLoss3) & ", " & _ 
		"BMV1=" & prepIntegerSQL(BMV1) & ", " & _
		"BMV2=" & prepIntegerSQL(BMV2) & ", " & _
		"BMV3=" & prepIntegerSQL(BMV3) & ", " & _
		"BMVLoss1=" & prepNumberSQL(BMVLoss1) & ", " & _
		"BMVLoss2=" & prepNumberSQL(BMVLoss2) & ", " & _
		"BMVLoss3=" & prepNumberSQL(BMVLoss3) & ", " & _
		"ReceiveAuthorization=" & prepBitSQL(ReceiveAuthorization) & ", " & _
		"RegionalTaskForce=" & prepBitSQL(RegionalTaskForce) & ", " & _
		"CurrentProgram=" & prepBitSQL(CurrentProgram) & ", " & _
		"CurrentProgramDescription=" & prepStringSQL(CurrentProgramDescription) & ", " & _
		"PreviouslyApplied=" & prepIntegerSQL(PreviouslyApplied) & ", " & _
		"PreviouslyAwarded=" & prepIntegerSQL(PreviouslyAwarded) & ", " & _
		"TerminationExplanation=" & prepIntegerSQL(TerminationExplanation) & ", " & _
		"MeetCashMatchRequirement=" & prepIntegerSQL(MeetCashMatchRequirement) & ", " & _
		"DedicateResources=" & prepIntegerSQL(DedicateResources) & ", " & _
		"NoSupplantation=" & prepIntegerSQL(NoSupplantation) & ", " & _
		"ProgramCategoryID=" & prepIntegerSQL(ProgramCategoryID) & ", " & _
		"BriefNarrative=" & prepStringSQL(BriefNarrative) & ", " & _
		"GrantRequestRangeID=" & prepIntegerSQL(GrantRequestRangeID) & ", "
	IF Button = "submit" Or Unsubmit = True Then
		sql = sql & "SubmitID=" & prepIntegerSQL(SubmitID) & ", " & _
		"SubmitTimestamp=" & prepStringSQL(SubmitTimestamp) & ", "
	End If 
	If MVCPARights = True Then
		sql = sql &	"DateReviewed=" & prepDateSQL(DateReviewed) & ", " & _
			"DateResponded=" & prepDateSQL(DateResponded) & ", " & _
			"MethodRespondedID=" & prepIntegerSQL(MethodRespondedID) & ", " & _
			"Notes=" & prepStringSQL(Notes) & ", " 
	End If
		sql = sql & "UpdateID=" & prepIntegerSQL(UpdateID) & ", " & _
		"UpdateTimestamp=" & prepStringSQL(UpdateTimestamp) & " " & vbCrLf & _
		"WHERE ISAID=" & prepIntegerSQL(ISAID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	rs = Con.Execute(sql)
End If

If Debug = True Then
	If Button="save" Then
		Response.Write("<a href=""ISA.asp?ISAID=" & ISAID & """>Return to ISA</a>" & vbCrLf)
	ElseIf Button = "submit" Then
		Response.Write("<a href=""../Home/Default.asp"">Return to Home</a>" & vbCrLf)
	End If
Else
	If Button="save" Then
		Response.Redirect("ISA.asp?ISAID=" & ISAID)
	ElseIf Button = "submit" Then
		Response.Redirect("../Home/Default.asp?GranteeID=" & GranteeID)
	End If
End If

%><!--#include file="../includes/prepDB.asp"-->