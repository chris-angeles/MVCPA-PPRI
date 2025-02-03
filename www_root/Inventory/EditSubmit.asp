<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim Debug, i, Timestamp, InventoryID, GranteeID, AssetClassID, ItemDescription, ModelYear, _
	MakeManufacturer, Model, UseID, SerialNo, SourceID, OwnerID, AcquisitionDate, Cost, _
	MVCPAPercentOfCost, Location, ConditionID, NotInventoryItem, DateOfDisposal, _
	DisposalID, SalePrice, OtherInformation, AdministrativeNotes
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

InventoryID = Request.Form("InventoryID")
GranteeID = Request.Form("GranteeID")
AssetClassID = Request.Form("AssetClassID")
ItemDescription = Request.Form("ItemDescription")
ModelYear = Request.Form("ModelYear")
MakeManufacturer = Request.Form("MakeManufacturer")
Model = Request.Form("Model")
UseID = Request.Form("UseID")
SerialNo = Request.Form("SerialNo")
SourceID = Request.Form("SourceID")
OwnerID = Request.Form("OwnerID")
AcquisitionDate = Request.Form("AcquisitionDate")
Cost = Request.Form("Cost")
MVCPAPercentOfCost = Request.Form("MVCPAPercentOfCost")
Location = Request.Form("Location")
ConditionID = Request.Form("ConditionID")
If Request.Form("NotInventoryItem")="1" Then
	NotInventoryItem = Date()
ElseIf Len(Request.Form("NotInventoryItem"))>1 Then
	NotInventoryItem = Request.Form("NotInventoryItem")
Else
	NotInventoryItem = null
End If
DateOfDisposal = Request.Form("DateOfDisposal")
DisposalID = Request.Form("DisposalID")
SalePrice = Request.Form("SalePrice")
OtherInformation = Request.Form("OtherInformation")
AdministrativeNotes = Request.Form("AdministrativeNotes")

If InventoryID=0 Then
	sql = "INSERT INTO Inventory ( GranteeID, AssetClassID, ItemDescription, " & vbCrLf & _
		"ModelYear, MakeManufacturer, Model, UseID, SerialNo, SourceID, OwnerID, " & vbCrLf & _
		"AcquisitionDate, Cost, MVCPAPercentOfCost, Location, ConditionID, " & vbCrLf & _
		"NotInventoryItem, DateOfDisposal, DisposalID, SalePrice, " & vbCrLf & _
		"OtherInformation, AdministrativeNotes, UpdateID, UpdateTimestamp) VALUES " & vbCrLf & "(" & _
		prepIntegerSQL(GranteeID) & ", " & _
		prepStringSQL(AssetClassID) & ", " & vbCrLf & _
		prepStringSQL(ItemDescription) & ", " & vbCrLf & _
		prepIntegerSQL(ModelYear) & ", " & vbCrLf & _
		prepStringSQL(MakeManufacturer) & ", " & vbCrLf & _
		prepStringSQL(Model) & ", " & vbCrLf & _
		prepIntegerSQL(UseID) & ", " & vbCrLf & _
		prepStringSQL(SerialNo) & ", " & vbCrLf & _
		prepIntegerSQL(SourceID) & ", " & vbCrLf & _
		prepIntegerSQL(OwnerID) & ", " & vbCrLf & _
		prepStringSQL(AcquisitionDate) & ", " & vbCrLf & _
		prepNumberSQL(Cost) & ", " & vbCrLf & _
		prepIntegerSQL(MVCPAPercentOfCost) & ", " & vbCrLf & _
		prepStringSQL(Location) & ", " & vbCrLf & _
		prepIntegerSQL(ConditionID) & ", " & vbCrLf & _
		prepDateSQL(NotInventoryItem) & ", " & vbCrLf & _
		prepDateSQL(DateOfDisposal) & ", " & vbCrLf & _
		prepIntegerSQL(DisposalID) & ", " & vbCrLf & _
		prepNumberSQL(SalePrice) & ", " & vbCrLf & _
		prepStringSQL(OtherInformation) & ", " & vbCrLf & _
		prepStringSQL(AdministrativeNotes) & ", " & vbCrLf & _
		prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
		prepStringSQL(Timestamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)
	sql = "SELECT IDENT_CURRENT('Inventory') AS AppID"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	InventoryID = rs.Fields(0)
Else
	sql = "UPDATE Inventory SET " & vbCrLf & _
		"AssetClassID=" & prepStringSQL(AssetClassID) & ", " & vbCrLf & _
		"ItemDescription=" & prepStringSQL(ItemDescription) & ", " & vbCrLf & _
		"ModelYear=" & prepIntegerSQL(ModelYear) & ", " & vbCrLf & _
		"MakeManufacturer=" & prepStringSQL(MakeManufacturer) & ", " & vbCrLf & _
		"Model=" & prepStringSQL(Model) & ", " & vbCrLf & _
		"UseID=" & prepIntegerSQL(UseID) & ", " & vbCrLf & _
		"SerialNo=" & prepStringSQL(SerialNo) & ", " & vbCrLf & _
		"SourceID=" & prepIntegerSQL(SourceID) & ", " & vbCrLf & _
		"OwnerID=" & prepIntegerSQL(OwnerID) & ", " & vbCrLf & _
		"AcquisitionDate=" & prepStringSQL(AcquisitionDate) & ", " & vbCrLf & _
		"Cost=" & prepNumberSQL(Cost) & ", " & vbCrLf & _
		"MVCPAPercentOfCost=" & prepIntegerSQL(MVCPAPercentOfCost) & ", " & vbCrLf & _
		"Location=" & prepStringSQL(Location) & ", " & vbCrLf & _
		"ConditionID=" & prepIntegerSQL(ConditionID) & ", " & vbCrLf & _
		"NotInventoryItem=" & prepDateSQL(NotInventoryItem) & ", " & vbCrLf & _
		"DateOfDisposal=" & prepStringSQL(DateOfDisposal) & ", " & vbCrLf & _
		"DisposalID=" & prepStringSQL(DisposalID) & ", " & vbCrLf & _
		"SalePrice=" & prepNumberSQL(SalePrice) & ", " & vbCrLf & _
		"OtherInformation=" & prepStringSQL(OtherInformation) & ", " & vbCrLf & _
		"AdministrativeNotes=" & prepStringSQL(AdministrativeNotes) & ", " & vbCrLf & _
		"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
		"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE InventoryID=" & prepIntegerSQL(InventoryID)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)
End If

If Debug = True Then
	Response.Write("<a href=""Edit.asp?InventoryID=" & InventoryID & """>Return to Edit</a>" & vbCrLf) 
Else
	Response.Redirect("Edit.asp?InventoryID=" & InventoryID)
End If
%><!--#include file="../includes/prepDB.asp"-->