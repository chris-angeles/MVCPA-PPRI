<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim Debug, i, Timestamp, InventoryID, GranteeID, Submit, PriorFirstApproval, PriorSecondApproval, _
	UseID, SerialNo, OwnerID, Location, ConditionID, DateOfDisposal, _
	DisposalID, SalePrice, OtherInformation, AdditionalInformation, _
	FirstApprovalID, FirstApprovalDate, SecondApprovalID, SecondApprovalDate, _
	AdministrativeNotes, ApplyChanges, Unsubmit, Reject, PhaseII
Debug = False
Timestamp = Now()

If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("<pre>Session(""" & i & """)='" & Session(i) & "'</pre>" & vbCrLf)
	Next
End If

InventoryID = Request.Form("InventoryID")
GranteeID = Request.Form("GranteeID")
UseID = Request.Form("UseID")
OwnerID = Request.Form("OwnerID")
Location = Request.Form("Location")
ConditionID = Request.Form("ConditionID")
DateOfDisposal = Request.Form("DateOfDisposal")
DisposalID = Request.Form("DisposalID")
SalePrice = Request.Form("SalePrice")
AdditionalInformation = Request.Form("AdditionalInformation")
FirstApprovalDate = Request.Form("FirstApprovalDate")
SecondApprovalDate = Request.Form("SecondApprovalDate")
AdministrativeNotes = Request.Form("AdministrativeNotes")
If Request.Form("Action") = "submit" Then
	Submit = True
Else
	Submit = False
End If
If Request.Form("Action") = "PhaseII" Then
	PhaseII = True
Else
	PhaseII = False
End If
If Request.Form("Unsubmit") = "1" Then
	Unsubmit = True
Else
	Unsubmit = False
End If
If Request.Form("Reject") = "1" Then
	Reject = True
Else
	Reject = False
End If
If Request.Form("ApplyChanges") = "1" Then
	ApplyChanges = True
Else
	ApplyChanges = False
End If
If Debug = True Then
	Response.Write("<pre>Sumbit=" & Submit & "</pre>" & vbCrLf)
End If
sql = "SELECT * FROM InventoryUpdate WHERE InventoryID=" & prepIntegerSQL(InventoryID) 
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

If rs.EOF = True Then
	PriorFirstApproval = False
	PriorSecondApproval = False
	sql = "INSERT INTO InventoryUpdate (InventoryID, OwnerID, UseID, Location, " & vbCrLf & _
		"	ConditionID, DateOfDisposal, DisposalID, SalePrice, AdditionalInformation, "
	If Submit = True Then
		sql = sql & "SubmitID, SubmitTimestamp, "
	End If
	sql = sql & "UpdateID, UpdateTimestamp)  VALUES (" & vbCrLf & _
		prepIntegerSQL(InventoryID) & ", " & _
		prepIntegerSQL(OwnerID) & ", " & vbCrLf & _
		prepIntegerSQL(UseID) & ", " & vbCrLf & _
		prepStringSQL(Location) & ", " & vbCrLf & _
		prepIntegerSQL(ConditionID) & ", " & vbCrLf & _
		prepDateSQL(DateOfDisposal) & ", " & vbCrLf & _
		prepIntegerSQL(DisposalID) & ", " & vbCrLf & _
		prepNumberSQL(SalePrice) & ", " & vbCrLf & _
		prepStringSQL(AdditionalInformation) & ", " & vbCrLf
	If Submit = True Then
		sql = sql & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
			prepStringSQL(Timestamp) & ", " & vbCrLf
	End If
		sql = sql & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
		prepStringSQL(Timestamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)
Else
	If IsNull(rs.Fields("FirstApprovalDate")) = False Then
		PriorFirstApproval = True
	Else
		PriorFirstApproval = False
	End If
	If IsNull(rs.Fields("SecondApprovalDate")) = False Then
		PriorSecondApproval = True
	Else
		PriorSecondApproval = False
	End If
	
	sql = "UPDATE InventoryUpdate SET " & vbCrLf & _
		"	UseID=" & prepIntegerSQL(UseID) & ", " & vbCrLf & _
		"	OwnerID=" & prepIntegerSQL(OwnerID) & ", " & vbCrLf & _
		"	Location=" & prepStringSQL(Location) & ", " & vbCrLf & _
		"	ConditionID=" & prepIntegerSQL(ConditionID) & ", " & vbCrLf & _
		"	DateOfDisposal=" & prepStringSQL(DateOfDisposal) & ", " & vbCrLf & _
		"	DisposalID=" & prepStringSQL(DisposalID) & ", " & vbCrLf & _
		"	SalePrice=" & prepNumberSQL(SalePrice) & ", " & vbCrLf & _
		"	AdditionalInformation=" & prepStringSQL(AdditionalInformation) & ", " & vbCrLf
	If Unsubmit = True Then
		sql = sql & "	SubmitID=null, " & vbCrLf & _
			"	SubmitTimestamp=null, " & vbCrLf & _
			"	FirstApprovalID=null, " & vbCrLf & _
			"	FirstApprovalDate=null, " & vbCrLf & _
			"	SecondApprovalID=null, " & vbCrLf & _
			"	SecondApprovalDate=null, " & vbCrLf & _
			"	DisposalUpdateID=null, " & vbCrLf & _
			"	DisposalUpdateTimestamp=null, " & vbCrLf & _
			"	AdministrativeNotes = CASE WHEN AdministrativeNotes IS NOT NULL THEN AdministrativeNotes + ' ' ELSE '' END + 'Update Unsubmitted by ' + ISNULL(" & prepStringSQL(UserName) & ",'') + '" & Timestamp & ".', " & vbCrLf
	ElseIf Submit = True Then
		sql = sql & "	SubmitID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
			"	SubmitTimestamp=" & prepStringSQL(Timestamp) & ", " & vbCrLf
	End If
	If Reject = True Then
		sql = sql & "	RejectedDate=" & prepDateSQL(Date()) & ", " & vbCrLf
	End If
	If PhaseII = True Then
		sql = sql & "	DisposalUpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
			"	DisposalUpdateTimestamp=" & prepStringSQL(Timestamp) & ", " & vbCrLf
	End If
	If MVCPARights = True And Unsubmit = False Then
		If PriorFirstApproval = False And Len(FirstApprovalDate)>0 Then
			sql = sql & "	FirstApprovalID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
				"	FirstApprovalDate=" & prepDateSQL(FirstApprovalDate) & ", " & vbCrLf
		ElseIf PriorFirstApproval = True And Len(FirstApprovalDate)=0 Then
			sql = sql & "	FirstApprovalID=null, " & vbCrLf & _
				"	FirstApprovalDate=null, " & vbCrLf
		End If
		If PriorSecondApproval = False And Len(SecondApprovalDate)>0 Then
			sql = sql & "	SecondApprovalID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
				"	SecondApprovalDate=" & prepDateSQL(SecondApprovalDate) & ", " & vbCrLf
		ElseIf PriorSecondApproval = True And Len(SecondApprovalDate)=0 Then
			sql = sql & "	SecondApprovalID=null, " & vbCrLf & _
				"	SecondApprovalDate=null, " & vbCrLf
		End If
		sql = sql & "	AdministrativeNotes=" & prepStringSQL(AdministrativeNotes) & ", " & vbCrLf
	End If
	sql = sql & "	UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
		"	UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE InventoryID=" & prepIntegerSQL(InventoryID)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)
End If

If Len(FirstApprovalDate)>0 And Len(SecondApprovalDate)>0 And ApplyChanges = True Then
	sql = "EXEC spUpdateInventoryFromInventoryUpdate @InventoryID=" & prepIntegerSQL(InventoryID) & " " 
	If Debug = True Then
		Response.Write("<pre>Update Inventory with """ & sql & """</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)
%><html>
<title>Inventory Update Completed</title>
<body>
<p>The Inventory Record has been updated with the changes.</p>

<p>The Update request has been archived.</p>

<p>You may close this tab.</p>

<form><input type="button" name="Close" value="Close" onclick="window.close();"</form></body>
</html>
<%
	Response.End
End If
If Reject = True Then
	sql = "DELETE FROM InventoryUpdate WHERE InventoryID=" & prepIntegerSQL(InventoryID) & "; " 
	If Debug = True Then
		Response.Write("<pre>Delete InventoryUpdate with """ & sql & """</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)
%><html>
<title>Inventory Update Rejected</title>
<body>
<p>The Inventory Update Record has been deleted and has been archived.</p>

<p>You may close this tab.</p>

<form><input type="button" name="Close" value="Close" onclick="window.close();"</form></body>
</html>
<%
	Response.End
ElseIf PriorSecondApproval = False And Len(SecondApprovalDate)>0 Then
%><html>
<title>Inventory Update Approval</title>
<body>
<p>The Inventory Update Record has been updated with the Second Approval.</p>

<p>You may close this tab.</p>

<form><input type="button" name="Close" value="Close" onclick="window.close();"</form></body>
</html>
<%
	Response.End
End If

If Debug = True Then
	Response.Write("<a href=""GranteeEdit.asp?InventoryID=" & InventoryID & """>Return to Edit</a>" & vbCrLf) 
Else
	Response.Redirect("GranteeEdit.asp?InventoryID=" & InventoryID)
End If
%><!--#include file="../includes/prepDB.asp"-->