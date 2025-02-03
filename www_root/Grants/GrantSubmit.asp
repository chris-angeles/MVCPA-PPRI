<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, TimeStamp, records, change, UpdateSystemID, GrantID, GranteeID, _
	FiscalYear, GrantNumber, DisplayQuarterOffset, CurrentYearAllocation, PriorYearAllocation, _
	ProgramName, MatchAmount, AwardAmount, InLieuOfDPSBudget, InLieuOfNICBBudget, _
	ReportsCompleteDate, ProgramGoalsDate, DeficienciesResolvedDate, FundsReturnedDate, _
	CloseoutID, CloseoutDate, AdministrativeComments
debug = False
Timestamp = Now()
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

If Request.Form.Count>0 Then
	GrantID = CInt(Request.Form("GrantID"))
	GranteeID = CInt(Request.Form("GranteeID"))
	FiscalYear = CInt(Request.Form("FiscalYear"))
	ProgramName = Request.Form("ProgramName")
	GrantNumber = Request.Form("GrantNumber")
	DisplayQuarterOffset = Request.Form("DisplayQuarterOffset")
	CurrentYearAllocation = Request.Form("CurrentYearAllocation")
	PriorYearAllocation = Request.Form("PriorYearAllocation")
	MatchAmount = Request.Form("MatchAmount")
	AwardAmount = Request.Form("AwardAmount")
	InLieuOfDPSBudget = Request.Form("InLieuOfDPSBudget")
	InLieuOfNICBBudget = Request.Form("InLieuOfNICBBudget")
	ReportsCompleteDate = Request.Form("ReportsCompleteDate")
	ProgramGoalsDate = Request.Form("ProgramGoalsDate")
	DeficienciesResolvedDate = Request.Form("DeficienciesResolvedDate")
	FundsReturnedDate = Request.Form("FundsReturnedDate")
	CloseoutID = Request.Form("CloseoutID")
	CloseoutDate = Request.Form("CloseoutDate")
	AdministrativeComments = Request.Form("AdministrativeComments")
End If

If Len(GrantID)=0 Then
	Response.Write("Error: No GrantID was provided.")
	Response.End
End If
If IsNumeric(GrantID) Then
	GrantID = CLng(GrantID)
Else
	Response.Write("Error: GrantID was not numeric.")
	Response.End
End If
If Len(MatchAmount) = 0 Then
	MatchAmount = 0.0
End If
If Len(AwardAmount) = 0 Then
	AwardAmount = 0.0
End If

sql = "SELECT  GrantID, FiscalYear, GranteeID, " & vbCrLF & _
	"	ISNULL(GrantNumber,'') AS GrantNumber, " & vbCrLF & _
	"	ISNULL(DisplayQuarterOffset, 0) AS DisplayQuarterOffset, " & vbCrLf & _
	"	ISNULL(CurrentYearAllocation, 0.0) AS CurrentYearAllocation, " & vbCrLf & _
	"	ISNULL(PriorYearAllocation, 0.0) AS PriorYearAllocation, " & vbCrLf & _
	"	ISNULL(ProgramName,'') As ProgramName,  " & VBcRlf & _
	"	ISNULL(MatchAmount,0.0) AS MatchAmount, " & vbCrLF & _
	"	ISNULL(AwardAmount,0.0) AS AwardAmount, " & vbCrLF & _
	"	ISNULL(InLieuOfDPSBudget,0.0) AS InLieuOfDPSBudget, " & vbCrLf & _
	"	ISNULL(InLieuOfNICBBudget,0.0) AS InLieuOfNICBBudget, " & vbCrLf & _
	"	CONVERT(VARCHAR,ReportsCompleteDate,101) AS ReportsCompleteDate, " & vbCrLf & _
	"	CONVERT(VARCHAR,ProgramGoalsDate,101) AS ProgramGoalsDate, " & vbCrLf & _
	"	CONVERT(VARCHAR,DeficienciesResolvedDate,101) AS DeficienciesResolvedDate, " & vbCrLf & _
	"	CONVERT(VARCHAR,FundsReturnedDate,101) AS FundsReturnedDate, " & vbCrLf & _
	"	ISNULL(CloseoutID,0) AS CloseoutID, " & vbCrLf & _
	"	CloseoutDate, " & vbCrLf & _ 
	"	ISNULL(AdministrativeComments, '') AS AdministrativeComments, " & vbCrLf & _
	"	UpdateID, UpdateTimestamp " & vbCrLf & _
	"FROM [Grants].Main " & vbCrLf & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID) 
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>")
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error retireving record for GrantID=" & GrantID)
	Response.End
End If

change = false
if GrantNumber=rs.Fields("GrantNumber") Then
	If Debug = True Then
		Response.Write("<pre>No change in Grant Number</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("Grant number updated from " & rs.Fields("GrantNumber") & " to " & GrantNumber & "</pre>")
	End If
End If

if DisplayQuarterOffset=rs.Fields("DisplayQuarterOffset") Then
	If Debug = True Then
		Response.Write("<pre>No change in Display Quarter Offset</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("Display Quarter Offset updated from " & rs.Fields("DisplayQuarterOffset") & " to " & DisplayQuarterOffset & "</pre>")
	End If
End If

If ProgramName =rs.Fields("ProgramName") Then
	If Debug = True Then
		Response.Write("<pre>No change in Program Name</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("Grant number updated from " & rs.Fields("ProgramName") & " to " & ProgramName & "</pre>")
	End If
End If

If CurrentYearAllocation = rs.Fields("CurrentYearAllocation") Then
	If Debug = True Then
		Response.Write("<pre>No change in Current Year Amount</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Current Year Amount updated from " & rs.Fields("CurrentYearAllocation") & " to " & CurrentYearAllocation & "</pre>")
	End If
End If

If PriorYearAllocation = rs.Fields("PriorYearAllocation") Then
	If Debug = True Then
		Response.Write("<pre>No change in PriorYearAllocation</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Carried Froward Amount updated from " & rs.Fields("PriorYearAllocation") & " to " & PriorYearAllocation & "</pre>")
	End If
End If

If AwardAmount = rs.Fields("AwardAmount") Then
	If Debug = True Then
		Response.Write("<pre>No change in Award Amount</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Award Amount updated from " & rs.Fields("AwardAmount") & " to " & AwardAmount & "</pre>")
	End If
End If

If MatchAmount = rs.Fields("MatchAmount") Then
	If Debug = True Then
		Response.Write("<pre>No change in Match Amount</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Match Amount updated from " & rs.Fields("MatchAmount") & " to " & MatchAmount & "</pre>")
	End If
End If

If InLieuOfDPSBudget = rs.Fields("InLieuOfDPSBudget") Then
	If Debug = True Then
		Response.Write("<pre>No change in In Lieu Of DPS Budget Amount</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>In Lieu Of DPS Budget Amount updated from " & rs.Fields("InLieuOfDPSBudget") & " to " & InLieuOfDPSBudget & "</pre>")
	End If
End If

If InLieuOfNICBBudget = rs.Fields("InLieuOfNICBBudget") Then
	If Debug = True Then
		Response.Write("<pre>No change in In Lieu Of NICB Budget Amount</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>In Lieu Of NICB Budget Amount updated from " & rs.Fields("InLieuOfNICBBudget") & " to " & InLieuOfNICBBudget & "</pre>")
	End If
End If

If ReportsCompleteDate = rs.Fields("ReportsCompleteDate") Then
	If Debug = True Then
		Response.Write("<pre>No change in Reports Complete Date</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Reports Complete Date updated from " & rs.Fields("ReportsCompleteDate") & " to " & ReportsCompleteDate & "</pre>")
	End If
End If

If ProgramGoalsDate = rs.Fields("ProgramGoalsDate") Then
	If Debug = True Then
		Response.Write("<pre>No change in Program Goals Date</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Program Goals Date updated from " & rs.Fields("ProgramGoalsDate") & " to " & ProgramGoalsDate & "</pre>")
	End If
End If

If DeficienciesResolvedDate = rs.Fields("DeficienciesResolvedDate") Then
	If Debug = True Then
		Response.Write("<pre>No change in Deficiencies Resolved Date</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Deficiencies Resolved Date updated from " & rs.Fields("DeficienciesResolvedDate") & " to " & DeficienciesResolvedDate & "</pre>")
	End If
End If

If FundsReturnedDate = rs.Fields("FundsReturnedDate") Then
	If Debug = True Then
		Response.Write("<pre>No change in Funds Returned Date</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Funds Returned Date updated from " & rs.Fields("FundsReturnedDate") & " to " & FundsReturnedDate & "</pre>")
	End If
End If

If CloseoutDate = "" and IsNull(rs.Fields("CloseoutDate")) = True Then
	' Nothing to change
	CloseoutDate = null
	CloseoutID = null
ElseIf CloseoutDate = "" and IsNull(rs.Fields("CloseoutDate")) = False Then
	change = True
	CloseoutID = null
	CloseoutDate = null
ElseIf Len(CloseoutDate) > 0 and IsNull(rs.Fields("CloseoutDate")) = True Then
	change = True
	CloseoutID = UserSystemID
ElseIf CDate(CloseoutDate) = rs.Fields("CloseoutDate") Then
	' Nothing to update
	CloseoutID = rs.Fields("CloseoutID")
	CloseoutDate = rs.Fields("CloseoutDate")
ElseIf CloseoutDate <> rs.Fields("CloseoutDate") Then
	change = True
	CloseoutID = UserSystemID
	CloseoutDate = rs.Fields("CloseoutDate")
End If

If AdministrativeComments = rs.Fields("AdministrativeComments") Then
	If Debug = True Then
		Response.Write("<pre>No change in Administrative Comments</pre>" & vbCrLf)
	End If
Else
	change = True
	If Debug = True Then
		Response.Write("<pre>Administrative Comments updated from " & rs.Fields("AdministrativeComments") & " to " & AdministrativeComments & "</pre>")
	End If
End If

If Change = False Then
	' nothing to update.
	If Debug = True Then
		Response.Write("There were no updates.")
		Response.Flush
	End If
Else
	rs.Close()
	sql = "UPDATE [Grants].Main SET GrantNumber=" & prepStringSQL(GrantNumber) & ", " & vbCrLf & _
		"DisplayQuarterOffset=" & prepIntegerSQL(DisplayQuarterOffset) & ", " & vbCrLf & _
		"ProgramName=" & prepStringSQL(ProgramName) & ", " & vbCrLf & _
		"CurrentYearAllocation=" & prepNumberSQL(CurrentYearAllocation) & ", " & vbCrLf & _
		"PriorYearAllocation=" & prepNumberSQL(PriorYearAllocation) & ", " & vbCrLf & _
		"MatchAmount=" & prepNumberSQL(MatchAmount) & ", " & vbCrLf & _
		"AwardAmount=" & prepNumberSQL(AwardAmount) & ", " & vbCrLf & _
		"InLieuOfDPSBudget=" & prepNumberSQL(InLieuOfDPSBudget) & ", " & vbCrLf & _
		"InLieuOfNICBBudget=" & prepNumberSQL(InLieuOfNICBBudget) & ", " & vbCrLf & _
		"ReportsCompleteDate=" & prepDateSQL(ReportsCompleteDate) & ", " & vbCrLf & _
		"ProgramGoalsDate=" & prepDateSQL(ProgramGoalsDate) & ", " & vbCrLf & _
		"DeficienciesResolvedDate=" & prepDateSQL(DeficienciesResolvedDate) & ", " & vbCrLf & _
		"FundsReturnedDate=" & prepDateSQL(FundsReturnedDate) & ", " & vbCrLf & _
		"CloseoutID=" & prepIntegerSQL(CloseoutID) & ", " & vbCrLf & _
		"CloseoutDate=" & prepDateSQL(CloseoutDate) & ", " & vbCrLf & _
		"AdministrativeComments=" & prepStringSQL(AdministrativeComments) & ", " & vbCrLf & _
		"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
		"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql, records)
	If records<>1 Then
		Response.Write("Error updating record: " & sql)
		Response.End
	End If

	sql = "UPDATE [Grants].Main " & vbCrLf & _
		"SET ReimbursementRate = 100.0*AwardAmount/(AwardAmount+MatchAmount-ISNULL(InLieuOfDPSBudget,0)-ISNULL(InLieuOfNICBBudget,0)) " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & "AND ReimbursementRate <> 100.0*AwardAmount/(AwardAmount+MatchAmount-ISNULL(InLieuOfDPSBudget,0)-ISNULL(InLieuOfNICBBudget,0))"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Con.Execute(sql)
End If
If Debug = True Then
	Response.Write("<a href=""../Grants/Grant.asp?GrantID=" & GrantID & """>Return</a>" & vbCrLf)
Else
	Response.Redirect("../Grants/Grant.asp?GrantID=" & GrantID)
End If

%><!--#include file="../includes/prepDB.asp"-->