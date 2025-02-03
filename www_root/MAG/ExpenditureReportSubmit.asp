<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, MAGID, Action, Submit, Unsubmit, NewDirectorApproval, _
	PurchaseLease, CashExpenditure, ExcludedAmount, ReimbursableExpenditure, _
	ReimbursableTotal, ReimbursementRate, Reimbursement, _
	SupplementaryComments, Confirmed, _
	SubmitID, SubmitTimestamp, _
	ReviewID, ReviewDate, _
	AuditApprovalID, AuditApprovalDate, _
	DirectorApprovalID, DirectorApprovalDate, _
	AdministrativeComments, AmountPaid, SerialNumbers, DatePaid, UpdateID, UpdateTimestamp

debug = False
TimeStamp = Now()

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

MAGID = Request.Form("MAGID")
If Len(MAGID) = 0 Then
	Response.Write("Error: No MAGID provided.")
	sendWarning("Error: No MAGID provided.")
	Response.End
Else
	MAGID = CInt(MAGID)
End If

Action = Request.Form("Action")
If Action = "submit" Then
	Submit = True
	SubmitID = UserSystemID
	SubmitTimestamp = Timestamp
Else
	Submit = False
End If

If Request.Form("Confirmed") = "1" Then
	Confirmed = True
Else
	Confirmed = False
End If

PurchaseLease = Request.Form("PurchaseLease")
CashExpenditure = Request.Form("CashExpenditure")
ExcludedAmount = Request.Form("ExcludedAmount")
ReimbursableExpenditure = Request.Form("ReimbursableExpenditure")
ReimbursableTotal = Request.Form("ReimbursableTotal")
ReimbursementRate = Request.Form("ReimbursementRate")
Reimbursement = Request.Form("Reimbursement")
SupplementaryComments = Request.Form("SupplementaryComments")
AdministrativeComments = Request.Form("AdministrativeComments")

ReviewID = Request.Form("ReviewID")
ReviewDate = Request.Form("ReviewDate")
If ReviewDate = "" Then
	ReviewID = null
ElseIf (ReviewID = "" OR ReviewID="0") And Len(ReviewDate)>0 Then
	ReviewID = UserSystemID
End If

AmountPaid = Request.Form("AmountPaid")
SerialNumbers = Request.Form("SerialNumbers")
DatePaid = Request.Form("DatePaid")
AuditApprovalID = Request.Form("AuditApprovalID")
AuditApprovalDate = Request.Form("AuditApprovalDate")
If AuditApprovalDate = "" Then
	AuditApprovalID = null
ElseIf (AuditApprovalID="" Or AuditApprovalID="0") And Len(AuditApprovalDate)>0 Then
	AuditApprovalID = UserSystemID
End If

DirectorApprovalID = Request.Form("DirectorApprovalID")
DirectorApprovalDate = Request.Form("DirectorApprovalDate")
NewDirectorApproval = False
If DirectorApprovalDate = "" Then
	DirectorApprovalID = null
ElseIf (DirectorApprovalID="" Or DirectorApprovalID="0") And Len(DirectorApprovalDate)>0 Then
	DirectorApprovalID = UserSystemID
	NewDirectorApproval = True
End If

If Request.Form("Unsubmit") = "1" Then
	Unsubmit = True
	SubmitID = null
	SubmitTimestamp = null
	Confirmed = null
	ReviewID = null
	ReviewDate = null
	AuditApprovalID = null
	AuditApprovalDate = null
	DirectorApprovalID = null
	DirectorApprovalDate = null
	AmountPaid = null
	DatePaid = null
Else
	Unsubmit = False
End If

If Debug = True Then
	Response.Write("<pre>SubmitID=" & SubmitID & "</pre>")
	Response.Write("<pre>SubmitTimestamp=" & SubmitTimestamp & "</pre>")
End If
UpdateID = UserSystemID
UpdateTimestamp = Timestamp

sql = "SELECT * " & vbCrLf & _
	"FROM MAG.ExpenditureReport " & vbCrLf & _
	"WHERE MAGID=" & prepIntegerSQL(MAGID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	' Initial submit. Insert record
	sql = "INSERT INTO MAG.ExpenditureReport (MAGID, PurchaseLease, SerialNumbers, CashExpenditure, ExcludedAmount, ReimbursableExpenditure, ReimbursableTotal, ReimbursementRate, Reimbursement, SupplementaryComments, Confirmed, SubmitID, SubmitTimestamp, UpdateID, UpdateTimestamp) VALUES (" & vbCrLf & _
		prepIntegerSQL(MAGID) & ", " & _
		prepIntegerSQL(PurchaseLease) & ", " & _
		prepStringSQL(SerialNumbers) & ", " & _
		prepNumberSQL(CashExpenditure) & ", " & _
		prepNumberSQL(ExcludedAmount) & ", " & _
		prepNumberSQL(ReimbursableExpenditure) & ", " & _
		prepNumberSQL(ReimbursableTotal) & ", " & _
		prepNumberSQL(ReimbursementRate) & ", " & _
		prepNumberSQL(Reimbursement) & ", " & _
		prepStringSQL(SupplementaryComments) & ", " & _
		prepBitSQL(Confirmed) & ", " & _
		prepIntegerSQL(SubmitID) & ", " & _
		prepStringSQL(SubmitTimestamp) & ", " & _
		prepIntegerSQL(UpdateID) & ", " & _
		prepStringSQL(Timestamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
Else
	sql = "UPDATE MAG.ExpenditureReport SET " & vbCrLf & _
		"PurchaseLease=" & prepIntegerSQL(PurchaseLease) & ", " & vbCrLf & _
		"SerialNumbers=" & prepStringSQL(SerialNumbers) & ", " & vbCrLf & _
		"CashExpenditure=" & prepNumberSQL(CashExpenditure) & ", " & vbCrLf & _
		"ExcludedAmount=" & prepNumberSQL(ExcludedAmount) & ", " & vbCrLf & _
		"ReimbursableExpenditure=" & prepNumberSQL(ReimbursableExpenditure) & ", " & vbCrLf & _
		"ReimbursableTotal=" & prepNumberSQL(ReimbursableTotal) & ", " & vbCrLf & _
		"ReimbursementRate=" & prepNumberSQL(ReimbursementRate) & ", " & vbCrLf & _
		"Reimbursement=" & prepNumberSQL(Reimbursement) & ", " & vbCrLf & _
		"SupplementaryComments=" & prepStringSQL(SupplementaryComments) & ", " & vbCrLf & _
		"Confirmed=" & prepBitSQL(Confirmed) & ", " & vbCrLf
	If Submit = True Or Unsubmit = True Then
		sql = sql & "SubmitID=" & prepIntegerSQL(SubmitID) & ", " & vbCrLf & _
		"SubmitTimestamp=" & prepStringSQL(SubmitTimestamp) & ", " & vbCrLf
	End If
		sql = sql & "UpdateID=" & prepIntegerSQL(UpdateID) & ", " & vbCrLf & _
		"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE MAGID=" & prepIntegerSQL(MAGID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

IF MVCPARights = True Then
	sql = "UPDATE MAG.ExpenditureReport SET " & vbCrLf & _
		"	AdministrativeComments=" & prepStringSQL(AdministrativeComments) & ", " & vbCrLf & _
		"	ReviewID=" & prepIntegerSQL(ReviewID) & ", " & vbCrLf & _
		"	ReviewDate=" & prepDateSQL(ReviewDate) & ", " & vbCrLf & _
		"	AuditApprovalID=" & prepIntegerSQL(AuditApprovalID) & ", " & vbCrLf & _
		"	AuditApprovalDate=" & prepStringSQL(AuditApprovalDate) & ", " & vbCrLf & _
		"	DirectorApprovalID=" & prepIntegerSQL(DirectorApprovalID) & ", " & vbCrLf & _
		"	DirectorApprovalDate=" & prepDateSQL(DirectorApprovalDate) & ", " & vbCrLf & _
		"	AmountPaid=" & prepNumberSQL(AmountPaid) & ", " & vbCrLf & _
		"	DatePaid=" & prepDateSQL(DatePaid) & ", " & vbCrLf & _
		"	UpdateID=" & prepIntegerSQL(UpdateID) & ", " & vbCrLf & _
		"	UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE MAGID=" & prepIntegerSQL(MAGID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

If Debug = True Then
	Response.Write("<a href=""ExpenditureReport.asp?MAGID=" & MAGID & """>return to expenditure report</a><br />" & vbCrLf)
	Response.Write("<a href=""../Home/Default.asp"">return to Home</a><br />" & vbCrLf)
Else
	Response.Redirect("ExpenditureReport.asp?MAGID=" & MAGID)
End If

%><!--#include file="../includes/prepDB.asp"-->