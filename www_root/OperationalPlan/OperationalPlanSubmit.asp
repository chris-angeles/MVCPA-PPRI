<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, Version, Button, UpdateApproval, AppID, _
	Colocation, MeetingsGranteeMethod, MeetingsGranteeFrequency, _
	MeetingsSubGranteeMethod, MeetingsSubGranteeFrequency, _
	MeetingsAllTFMethod, MeetingsAllTFFrequency, _
	MeetingsDescription, CommunicationGranteeMethod, _
	CommunicationGranteeFrequency, _
	CommunicationSubGranteeMethod, CommunicationSubGranteeFrequency, _
	CommunicationAllTFMethod, CommunicationAllTFFrequency, _
	CommunicationDescription, CoverageAgencyMeetings, _
	CoverageAgencyContacts, IntelligenceSharing, _
	OperationalCoordination, DirectOperatations, _
	SubmitID, SubmitTimestamp, OperationalPlanApprovalID, OperationalPlanApprovalDate, UpdateID, UpdateTimestamp

TimeStamp = Now()
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

AppID = Request.Form("AppID")
Button = Request.Form("Button")
Colocation = Request.Form("Colocation")
MeetingsGranteeMethod = Request.Form("MeetingsGranteeMethod")
MeetingsGranteeFrequency = Request.Form("MeetingsGranteeFrequency")
MeetingsSubGranteeMethod = Request.Form("MeetingsSubGranteeMethod")
MeetingsSubGranteeFrequency = Request.Form("MeetingsSubGranteeFrequency")
MeetingsAllTFMethod = Request.Form("MeetingsAllTFMethod")
MeetingsAllTFFrequency = Request.Form("MeetingsAllTFFrequency")
MeetingsDescription = Request.Form("MeetingsDescription")
CommunicationGranteeMethod = Request.Form("CommunicationGranteeMethod")
CommunicationGranteeFrequency = Request.Form("CommunicationGranteeFrequency")
CommunicationSubGranteeMethod = Request.Form("CommunicationSubGranteeMethod")
CommunicationSubGranteeFrequency = Request.Form("CommunicationSubGranteeFrequency")
CommunicationAllTFMethod = Request.Form("CommunicationSubGranteeFrequency")
CommunicationAllTFFrequency = Request.Form("CommunicationAllTFFrequency")
CommunicationDescription = Request.Form("CommunicationDescription")
CoverageAgencyMeetings = Request.Form("CoverageAgencyMeetings")
CoverageAgencyContacts = Request.Form("CoverageAgencyContacts")
IntelligenceSharing = Request.Form("IntelligenceSharing")
OperationalCoordination = Request.Form("OperationalCoordination")
DirectOperatations = Request.Form("DirectOperatations")
If Button = "submit" Then
	SubmitID = UserSystemID
	SubmitTimestamp = Timestamp
Else
	SubmitID = null
	SubmitTimestamp = null
End If
UpdateID = UserSystemID
UpdateTimestamp = Timestamp

sql = "SELECT A.*, B.OperationalPlanApprovalDate, B.OperationalPlanApprovalID  " & vbCrLF & _
	"FROM [Grants].OperationalPlan AS A " & vbCrLf & _
	"LEFT JOIN Application.Admin AS B ON B.AppID=A.AppID " & vbCrLf & _
	"WHERE A.AppID=" & prepIntegerSQL(AppID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

If rs.EOF = False Then
	UpdateApproval = False
	OperationalPlanApprovalID = rs.Fields("OperationalPlanApprovalID")
	OperationalPlanApprovalDate = rs.Fields("OperationalPlanApprovalDate")
	If MVCPARights = True Then
		If IsNull(OperationalPlanApprovalDate) = True And Len(Request.Form("OperationalPlanApprovalDate"))>0 Then
			OperationalPlanApprovalID = UserSystemID
			OperationalPlanApprovalDate = Request.Form("OperationalPlanApprovalDate")
			If Debug = True Then
				Response.Write("<pre>Update Approval</pre>")
			End If
			UpdateApproval = True
		ElseIf IsNull(OperationalPlanApprovalDate) = False And Request.Form("OperationalPlanApprovalDate") = "" Then
			OperationalPlanApprovalID = UserSystemID
			OperationalPlanApprovalDate = null
			If Debug = True Then
				Response.Write("<pre>Update Approval</pre>")
			End If
			UpdateApproval = True
		ElseIf IsNull(OperationalPlanApprovalDate) = False And Len(Request.Form("OperationalPlanApprovalDate"))>0 Then
			If CDate(Request.Form("OperationalPlanApprovalDate")) <> OperationalPlanApprovalDate Then
				OperationalPlanApprovalID = UserSystemID
				OperationalPlanApprovalDate = Request.Form("OperationalPlanApprovalDate")
				If Debug = True Then
					Response.Write("<pre>Update Approval</pre>")
				End If
				UpdateApproval = True
			Else
			OperationalPlanApprovalId = Null
			OperationalPlanApprovalDate = Null
			End If
		End If 
	End If
	If  UpdateApproval = True Then
		sql = "UPDATE Application.Admin SET OperationalPlanApprovalID=" & prepIntegerSQL(OperationalPlanApprovalID) & ", OperationalPlanApprovalDate=" & prepDateSQL(OperationalPlanApprovalDate) & ", UpdateID=" & prepIntegerSQL(UpdateID) & ", " & _
		"UpdateTimestamp=" & prepDateSQL(UpdateTimestamp) & " " & vbCrLf & _
		"WHERE AppID=" & prepIntegerSQL(AppID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
Else
	OperationalPlanApprovalId = Null
	OperationalPlanApprovalDate = Null
End If

If rs.EOF = True Then
	' Insert
	sql = "INSERT INTO [Grants].OperationalPlan (AppID, Colocation, MeetingsGranteeMethod, MeetingsGranteeFrequency, MeetingsSubGranteeMethod, MeetingsSubGranteeFrequency, MeetingsAllTFMethod, MeetingsAllTFFrequency, MeetingsDescription, CommunicationGranteeMethod, CommunicationGranteeFrequency, CommunicationSubGranteeMethod, CommunicationSubGranteeFrequency, CommunicationAllTFMethod, CommunicationAllTFFrequency, CommunicationDescription, CoverageAgencyMeetings, CoverageAgencyContacts, IntelligenceSharing, OperationalCoordination, DirectOperatations, SubmitID, SubmitTimestamp, UpdateID, UpdateTimestamp) " & vbCrLf & _
		"VALUES (" & prepIntegerSQL(AppID) & ", " & _
		prepIntegerSQL(Colocation) & ", " & _
		prepIntegerSQL(MeetingsGranteeMethod) & ", " & _
		prepIntegerSQL(MeetingsGranteeFrequency) & ", " & _
		prepIntegerSQL(MeetingsSubGranteeMethod) & ", " & _
		prepIntegerSQL(MeetingsSubGranteeFrequency) & ", " & _
		prepIntegerSQL(MeetingsAllTFMethod) & ", " & _
		prepIntegerSQL(MeetingsAllTFFrequency) & ", " & _
		prepStringSQL(MeetingsDescription) & ", " & _
		prepIntegerSQL(CommunicationGranteeMethod) & ", " & _
		prepIntegerSQL(CommunicationGranteeFrequency) & ", " & _
		prepIntegerSQL(CommunicationSubGranteeMethod) & ", " & _
		prepIntegerSQL(CommunicationSubGranteeFrequency) & ", " & _
		prepIntegerSQL(CommunicationAllTFMethod) & ", " & _
		prepIntegerSQL(CommunicationAllTFFrequency) & ", " & _
		prepStringSQL(CommunicationDescription) & ", " & _
		prepStringSQL(CoverageAgencyMeetings) & ", " & _
		prepStringSQL(CoverageAgencyContacts) & ", " & _
		prepStringSQL(IntelligenceSharing) & ", " & _
		prepStringSQL(OperationalCoordination) & ", " & _
		prepStringSQL(DirectOperatations) & ", " & _
		prepIntegerSQL(SubmitID) & ", " & _
		prepDateSQL(SubmitTimestamp) & ", " & _
		prepIntegerSQL(UpdateID) & ", " & _
		prepDateSQL(UpdateTimestamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
Else
	' Update
	sql = "UPDATE [Grants].OperationalPlan " & vbCrLF & _
		"SET Colocation=" & prepIntegerSQL(Colocation) & ", " & _
		"MeetingsGranteeMethod=" & prepIntegerSQL(MeetingsGranteeMethod) & ", " & _
		"MeetingsGranteeFrequency=" & prepIntegerSQL(MeetingsGranteeFrequency) & ", " & _
		"MeetingsSubGranteeMethod=" & prepIntegerSQL(MeetingsSubGranteeMethod) & ", " & _
		"MeetingsSubGranteeFrequency=" & prepIntegerSQL(MeetingsSubGranteeFrequency) & ", " & _
		"MeetingsAllTFMethod=" & prepIntegerSQL(MeetingsAllTFMethod) & ", " & _
		"MeetingsAllTFFrequency=" & prepIntegerSQL(MeetingsAllTFFrequency) & ", " & _
		"MeetingsDescription=" & prepStringSQL(MeetingsDescription) & ", " & _
		"CommunicationGranteeMethod=" & prepIntegerSQL(CommunicationGranteeMethod) & ", " & _
		"CommunicationGranteeFrequency=" & prepIntegerSQL(CommunicationGranteeFrequency) & ", " & _
		"CommunicationSubGranteeMethod=" & prepIntegerSQL(CommunicationSubGranteeMethod) & ", " & _
		"CommunicationSubGranteeFrequency=" & prepIntegerSQL(CommunicationSubGranteeFrequency) & ", " & _
		"CommunicationAllTFMethod=" & prepIntegerSQL(CommunicationAllTFMethod) & ", " & _
		"CommunicationAllTFFrequency=" & prepIntegerSQL(CommunicationAllTFFrequency) & ", " & _
		"CommunicationDescription=" & prepStringSQL(CommunicationDescription) & ", " & _
		"CoverageAgencyMeetings=" & prepStringSQL(CoverageAgencyMeetings) & ", " & _
		"CoverageAgencyContacts=" & prepStringSQL(CoverageAgencyContacts) & ", " & _
		"IntelligenceSharing=" & prepStringSQL(IntelligenceSharing) & ", " & _
		"OperationalCoordination=" & prepStringSQL(OperationalCoordination) & ", " & _
		"DirectOperatations=" & prepStringSQL(DirectOperatations) & ", "
	If Button = "submit" Then
		sql = sql & "SubmitID=" & prepIntegerSQL(SubmitID) & ", SubmitTimestamp=" & prepDateSQL(SubmitTimestamp) & ", "
	End If
	sql = sql &	"UpdateID=" & prepIntegerSQL(UpdateID) & ", " & _
		"UpdateTimestamp=" & prepDateSQL(UpdateTimestamp) & " " & _
		"WHERE AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)

End If

If Debug = True Then
	Response.Write("<A href=""OperationalPlan.asp?AppID=" & AppID & """>Operational Plan</a>")
Else
	Response.Redirect("OperationalPlan.asp?AppID=" & AppID)
End If 
%><!--#include file="../includes/prepDB.asp"-->