<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim Debug, i, Timestamp, GranteeID, GrantID, FiscalYear, Action, GranteeComments, _
	CompleteBy, CompleteByDate, AdministrativeNotes, AcceptanceDate, Unsubmit, RecordExists

Debug = False
Timestamp = Now

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

GranteeID = Request.Form("GranteeID")
GrantID = Request.Form("GrantID")
FiscalYear = Request.Form("FiscalYear")
Action = Request.Form("Action")
GranteeComments = Request.Form("GranteeComments")
AdministrativeNotes = Request.Form("AdministrativeNotes")
AcceptanceDate = Request.Form("AcceptanceDate")
CompleteBy = Request.Form("CompleteBy")
CompleteByDate = Request.Form("CompleteByDate")
If Request.Form("Unsubmit") = "1" Then
	Unsubmit = True
Else
	Unsubmit = False
End If

sql = "SELECT GrantID FROM [Grants].InventoryCertification WHERE GrantID=" & prepIntegerSQL(GrantID)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs=con.execute(sql)
If rs.EOF = True Then
	RecordExists = False
Else
	RecordExists = True
End If

If Action = "submit" Then
	If RecordExists=False Then
		sql = "INSERT INTO [Grants].InventoryCertification(GrantID, GranteeComments, CompleteByDate, CompleteBy, SubmitID, SubmitTimestamp, UpdateID, UpdateTimestamp) " & vbCrLf & _
			"VALUES (" & prepIntegerSQL(GrantID) & ", " & _
			prepStringSQL(GranteeComments) & ", " & _
			prepDateSQL(CompleteByDate) & ", " & _
			prepStringSQL(CompleteBy) & ", " & _
			prepIntegerSQL(UserSystemID) & ", " & _
			prepStringSQL(Timestamp) & ", " & _
			prepIntegerSQL(UserSystemID) & ", " & _
			prepStringSQL(Timestamp) & ") "
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)
	Else
		sql = "UPDATE [Grants].InventoryCertification SET " & vbCrLf & _
			"GranteeComments=" & prepStringSQL(GranteeComments) & ", " & _
			"CompleteByDate=" & prepDateSQL(CompleteByDate) & ", " & _
			"CompleteBy=" & prepStringSQL(CompleteBy) & ", " & _
			"SubmitID=" & prepIntegerSQL(UserSystemID) & ", " & _
			"SubmitTimestamp=" & prepStringSQL(Timestamp) & ", " & _
			"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
			"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
			"WHERE GrantID=" & prepIntegerSQL(GrantID)
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)
	End If

	sql = "INSERT INTO [Grants].InventoryCertificationDetail (InventoryID, GrantID, AssetClassID, ItemDescription, ModelYear, MakeManufacturer, Model, SerialNo, Location) " & vbCrLf & _
		"SELECT InventoryID, " & GrantID & ", AssetClassID, ItemDescription, ModelYear, MakeManufacturer, Model, SerialNo, Location " & vbCrLf & _
		"FROM Inventory " & vbCrLF & _
		"WHERE GranteeID=" & prepIntegerSQL(GranteeID) & " AND ISNULL(AcquisitionDate,'9/1/" & (FiscalYear-1) & "')<'9/1/" & (FiscalYear) & "' AND " & vbCrLF & _
		"	ISNULL(DateOfDisposal,'12/31/2099')>'8/31/" & FiscalYear & "' " & vbCrLf & _
		"	AND NotInventoryItem IS NULL " & vbCrLf & _
		"ORDER BY InventoryID "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)

ElseIf Action = "save" Then
	If Unsubmit = True Then
		sql = "UPDATE [Grants].InventoryCertification SET SubmitID=null, SubmitTimestamp=null, " & vbCrLF & _
			"	AcceptanceID=null, AcceptanceDate=null, " & vbCrLF & _
			"	CompleteByDate=null, CompleteBy=null, " & vbCrLF & _
			"	UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLF & _
				"	UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
			"WHERE GrantID=" & prepIntegerSQL(GrantID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		sql = "DELETE FROM [Grants].InventoryCertificationDetail WHERE GrantID=" & prepIntegerSQL(GrantID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
	Else
		sql = "SELECT AdministrativeNotes, AcceptanceID, AcceptanceDate " & vbCrLF & _
			"FROM [Grants].InventoryCertification " & vbCrLF & _
			"WHERE GrantID=" & prepIntegerSQL(GrantID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		If rs.Eof = False Then
			sql = "UPDATE [Grants].InventoryCertification SET " & vbCrLf & _
				"	AdministrativeNotes=" & prepStringSQL(AdministrativeNotes) & ", " & vbCrLf 
			If Len(AcceptanceDate)>0 and IsNull(rs.Fields("AcceptanceDate")) = True Then
				sql = sql & "	AcceptanceID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
					"	AcceptanceDate=" & prepStringSQL(CDate(Timestamp)) & ", " & vbCrLf
			ElseIf Len(AcceptanceDate)=0 And IsNull(rs.Fields("AcceptanceDate")) = False Then
				sql = sql & "	AcceptanceID=null, " & vbCrLf & _
					"	AcceptanceDate=null, " & vbCrLf
			End If
			sql = sql & "	UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLF & _
				"	UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
			"WHERE GrantID=" & prepIntegerSQL(GrantID)
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = Con.Execute(sql)
		End If
	End If
End If

If Debug = False Then
	Response.Redirect("Certify.asp?GranteeID=" & GranteeID & "&FiscalYear=" & FiscalYear)
Else
	Response.Write("<a href=""Certify.asp?GranteeID=" & GranteeID & "&FiscalYear=" & FiscalYear & """>return</a>")
End If
%><!--#include file="../includes/prepDB.asp"-->