<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim Debug, PermitEdit, InventoryID, GranteeID, AssetClassID, AssetClass, ItemDescription, ModelYear, MakeManufacturer, _
	Model, UseID, SerialNo, SourceID, Source, OwnerID, AcquisitionDate, Cost, MVCPAPercentOfCost, _
	Location, ConditionID, DateOfDisposal, DisposalID, SalePrice, OtherInformation, AdditionalInformation,  _
	SubmitID, SubmitTimestamp, SubmitName, FirstApprovalID, FirstApprovalName, FirstApprovalDate, _
	SecondApprovalID, SecondApprovalName, SecondApprovalDate, AdministrativeNotes, OldAdministrativeNotes, _
	UpdateID, UpdateTimestamp, GranteeNAme, UpdateName, CanSubmit, Submitted, _
	FOEmail,FAEmail, PDEmail, PMEmail, PAEmail, UpdateRecord, _
	DisposalUpdateID, DisposalUpdateName, DisposalUpdateTimestamp, PhaseII
Debug = False

If Len(Request.Form("InventoryID"))>0 Then
	InventoryID = CInt(Request.Form("InventoryID"))
ElseIf Len(Request.QueryString("InventoryID"))>0 Then
	InventoryID = CInt(Request.QueryString("InventoryID"))
Else
	InventoryID = 0
End If

If InventoryID>0 Then
	sql = "SELECT * FROM vwInventoryUpdate WHERE InventoryID=" & prepIntegerSQL(InventoryID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error: No Inventory record retrieved")
		SendMessage "Error: No Inventory record retrieved"
		Response.End
	Else
		InventoryID = rs.Fields("InventoryID")
		GranteeID = rs.Fields("GranteeID")
		GranteeName = rs.Fields("GranteeName")
		AssetClassID = rs.Fields("AssetClassID")
		AssetClass = rs.Fields("AssetClass")
		ItemDescription = rs.Fields("ItemDescription")
		ModelYear = rs.Fields("ModelYear")
		MakeManufacturer = rs.Fields("MakeManufacturer")
		Model = rs.Fields("Model")
		UseID = rs.Fields("UseID")
		SerialNo = rs.Fields("SerialNo")
		SourceID = rs.Fields("SourceID")
		Source = rs.Fields("Source")
		OwnerID = rs.Fields("OwnerID")
		AcquisitionDate = rs.Fields("AcquisitionDate")
		Cost = rs.Fields("Cost")
		MVCPAPercentOfCost = rs.Fields("MVCPAPercentOfCost")
		Location = rs.Fields("Location")
		ConditionID = rs.Fields("ConditionID")
		DateOfDisposal = rs.Fields("DateOfDisposal")
		DisposalID = rs.Fields("DisposalID")
		SalePrice = rs.Fields("SalePrice")
		OtherInformation = rs.Fields("OtherInformation")
		AdditionalInformation = rs.Fields("AdditionalInformation")
		SubmitID = rs.Fields("SubmitID")
		SubmitTimestamp = rs.Fields("SubmitTimestamp")
		SubmitName = rs.Fields("SubmitName")
		DisposalUpdateID = rs.Fields("DisposalUpdateID")
		DisposalUpdateName = rs.Fields("DisposalUpdateName")
		DisposalUpdateTimestamp = rs.Fields("DisposalUpdateTimestamp")
		PhaseII = rs.Fields("PhaseII")
		FirstApprovalID = rs.Fields("FirstApprovalID")
		FirstApprovalName = rs.Fields("FirstApprovalName")
		FirstApprovalDate = rs.Fields("FirstApprovalDate")
		SecondApprovalID = rs.Fields("SecondApprovalID")
		SecondApprovalName = rs.Fields("SecondApprovalName")
		SecondApprovalDate = rs.Fields("SecondApprovalDate")
		AdministrativeNotes = rs.Fields("AdministrativeNotes")
		OldAdministrativeNotes = rs.Fields("OldAdministrativeNotes")
		UpdateID = rs.Fields("UpdateID")
		UpdateName = rs.Fields("UpdateName")
		UpdateTimestamp = rs.Fields("UpdateTimestamp")
		FOEmail = rs.Fields("FOEmail")
		FAEmail = rs.Fields("FAEmail")
		PDEmail = rs.Fields("PDEmail")
		PMEmail = rs.Fields("PMEmail")
		PAEmail = rs.Fields("PAEmail")
		UpdateRecord = rs.Fields("UpdateRecord")
		If UserSystemID = rs.Fields("ProgramDirectorID") Then
			CanSubmit = True
		ElseIf UserSystemID = rs.Fields("ProgramManagerID") Then
			CanSubmit = True
		ElseIf UserSystemID = rs.Fields("FinancialOfficerID") Then
			CanSubmit = True
		ElseIf UserSystemID = rs.Fields("FinancialAdministrativeContactID") Then
			CanSubmit = True
		Else
			CanSubmit = False
		End If
	End If
End If

If IsNull(SubmitID) = True Then
	Submitted = False
ElseIf SubmitID=0 Then
	Submitted = False
Else
	Submitted = True
End If
If MVCPARights = True Then
	PermitEdit = False
ElseIf Submitted = True Then
	PermitEdit = False
	CanSubmit = False
Else
	PermitEdit = CheckPermissions(UserSystemID, GranteeID, True)
End If
'PermitEdit=True

If Debug = True Then
	Response.Write("<pre>PermitEdit='" & PermitEdit & "'; CanSubmit=" & CanSubmit & "; UserSystemID=" & UserSystemID & "</pre>")
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Inventory Item Edit Page</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm(action)
	{
		if (validateForm() == true) {
			document.Inventory.Action.value = action;
			document.Inventory.submit();
		}
	}

	function validateForm()
	{
<% If MVCPARights = False Then %>
		if (document.Inventory.AdditionalInformation.value.length==0)
		{
			alert("You must enter an explanation for the change to save this record.");
			return false;
		}
<%	End If %>
		return true;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body>
<h1>Inventory Item Edit Page</h1>
<%
If Submitted = True Then
	Response.Write("<div style=""text-align: center;"">Submitted by " & SubmitName & ", " & SubmitTimeStamp & ".</div>" & vbCrLf)
	If IsNull(FirstApprovalID) = False Then
		Response.Write("<div style=""text-align: center;"">First MVCPA Approver is " & FirstApprovalName & ", " & FirstApprovalDate & ".</div>" & vbCrLf)
	End If
	If IsNull(SecondApprovalID) = False Then
		Response.Write("<div style=""text-align: center;"">Second MVCPA Approver is " & SecondApprovalName & ", " & SecondApprovalDate & ".</div>" & vbCrLf)
	End If
	If IsNull(DisposalUpdateID) = False Then
		Response.Write("<div style=""text-align: center;"">Stage II Information Submitted by " & DisposalUpdateName & ", " & DisposalUpdateTimestamp & ".</div>" & vbCrLf)
	End If
	Response.Write("<br />" & vbCrLf)
End If

If UpdateRecord = False Then
	Response.Write("<div style=""text-align: center; color: red; font-weight: bold; "">The update record does not exist yet. Saving the record will create a new update record.</div>" & vbCrLf)
End If
%>
<table>
<form name="Inventory" method="post" action="GranteeEditSubmit.asp">
<%=Hiddenfield("InventoryID", InventoryID) %><%=Hiddenfield("Action","save") %>
<tr>
	<td style="text-align: right; font-weight: bold;">InventoryID</td>
	<td><%=InventoryID %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Grantee</td>
	<td><%=GranteeID %>: <%=GranteeName %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Asset Class</td>
	<td><%=AssetClassID %>: <%=AssetClass %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Item Description</td>
	<td><%=ItemDescription %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">ModelYear</td>
	<td><%=ModelYear %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Make or Manufacturer</td>
	<td><%=MakeManufacturer %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Model</td>
	<td><%=Model %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">SerialNo or VIN</td>
	<td><%=SerialNo %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">How Acquired</td>
	<td><%=SourceID %>: <%=Source %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Acquisition Date</td>
	<td><%=AcquisitionDate %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Cost</td>
	<td><%=prepCurrencyWeb(Cost) %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">MVCPA Percent Of Cost</td>
	<td><%=MVCPAPercentOfCost %>%</td>
</tr>
<tr><td colspan="2"><hr /></td></tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Use</td>
	<td><select name="UseID" id="UseID">
		<option value="0">Select Use</option>
<%
	If PermitEdit = True Then
		sql = "SELECT UseID, [Use] FROM Lookup.InventoryUse ORDER BY UseSort"
	Else
		sql = "SELECT UseID, [Use] FROM Lookup.InventoryUse WHERE UseID=" & prepIntegerSQL(UseID) & " ORDER BY UseSort"
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("UseID") & """ " & Selected(rs.Fields("UseID"), UseID) & ">" & rs.Fields("Use") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Who holds title?</td>
	<td><select name="OwnerID" id="OwnerID">
		<option value="0">Select Who holds title?</option>
<%
	If PermitEdit = True Then
		sql = "SELECT OwnerID, Owner FROM Lookup.InventoryOwner ORDER BY OwnerSort"
	Else
		sql = "SELECT OwnerID, Owner FROM Lookup.InventoryOwner WHERE OwnerID=" & prepIntegerSQL(OwnerID) & " ORDER BY OwnerSort"
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("OwnerID") & """ " & Selected(rs.Fields("OwnerID"), OwnerID) & ">" & rs.Fields("Owner") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Location</td>
	<td><%=TextField("Location", Location, 80, 100, PermitEdit, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Condition</td>
	<td><select name="ConditionID" id="ConditionID">
		<option value="0">Select condition</option>
<%
	If PermitEdit = True Then
		sql = "SELECT ConditionID, Condition FROM Lookup.InventoryCondition ORDER BY ConditionSort"
	Else
		sql = "SELECT ConditionID, Condition FROM Lookup.InventoryCondition WHERE ConditionID=" & prepIntegerSQL(ConditionID) & " ORDER BY ConditionSort"
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("ConditionID") & """ " & Selected(rs.Fields("ConditionID"), ConditionID) & ">" & rs.Fields("Condition") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Method of Disposal</td>
	<td><select name="DisposalID" id="DisposalID">
		<option value="0">Select method of disposal</option>
<%
	If PermitEdit = True Then
		sql = "SELECT DisposalID, Disposal FROM Lookup.InventoryDisposal ORDER BY DisposalSort"
	Else
		sql = "SELECT DisposalID, Disposal FROM Lookup.InventoryDisposal WHERE DisposalID=" & prepIntegerSQL(DisposalID) & " ORDER BY DisposalSort"
	End IF
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("DisposalID") & """ " & Selected(rs.Fields("DisposalID"), DisposalID) & ">" & rs.Fields("Disposal") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Date Of Disposal</td>
	<td><%=DateField("DateOfDisposal", DateOfDisposal, (PermitEdit Or PhaseII)) %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Net Sale Price</td>
	<td><%=CurrencyField("SalePrice", SalePrice, 12, 12, (PermitEdit Or PhaseII), "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold; vertical-align: top; ">Explanation of Change</td>
	<td><%	
If IsNull(OtherInformation) = False Then 
	Response.Write("Previous Comments: " & OtherInformation & "<br />" & "<b>Additional Information:</b><br>" & vbCrLf)
End If %>
	<%=TextArea("AdditionalInformation", AdditionalInformation, 3, 80, 1000, PermitEdit, "") %></td>
</tr>
<tr><td colspan="2"><hr /></td></tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Last Update ID</td>
	<td><%=UpdateID %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Last Update Name</td>
	<td><%=UpdateName %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold; white-space: nowrap; ">Last Update Timestamp</td>
	<td><%=UpdateTimestamp %></td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<%
'PermitEdit=True
'CanSubmit=True
'MVCPARights = False
If MVCPARights = True Then
%>
<tr>
	<th colspan="2">Administrative Section</th>
</tr>
<tr>
	<td colspan="2" style="text-align: right"><a href="mailto:<%
	if IsNull(FOEmail) = False Then
		Response.Write(FOEmail & ";")
	End If
	if IsNull(FAEmail) = False Then
		Response.Write(FAEmail & ";")
	End If
	Response.Write("?CC=grantsMVCPA@txdmv.gov;")
	if IsNull(PDEmail) = False Then
		Response.Write(PDEmail & ";")
	End If
	if IsNull(PMEmail) = False Then
		Response.Write(PMEmail & ";")
	End If
	if IsNull(PMEmail) = False Then
		Response.Write(PMEmail & ";")
	End If
	if IsNull(PAEmail) = False Then
		Response.Write(PAEmail & ";")
	End If
	Response.Write("&subject=Inventory Update Request for " & ItemDescription)
	%>">E-Mail</a></td>
</tr>

<%
	If PhaseII And IsNull(DisposalID) = False And IsNull(DisposalUpdateID)=False Then
		Response.Write("<tr><td colspan=""2"" style=""text-align: center; color: red; "">Item in second phase of approval process.</td</tr>" & vbCrLf)
	ElseIf PhaseII And IsNull(DisposalID) = False Then
		Response.Write("<tr><td colspan=""2"" style=""text-align: center; color: red; "">Item in second phase of approval process. Waiting on grantee.</td</tr>" & vbCrLf)
	End If

	If IsNull(SubmitID) = False Then 
%>
<tr>
	<td style="text-align: right; font-weight: bold; white-space: nowrap; ">First Approver</td>
	<td><%
		If IsNull(FirstApprovalID) = False Then
			Response.Write(DateField("FirstApprovalDate", FirstApprovalDate, MVCPARights) & ", " & FirstApprovalName)
		ElseIf MVCPARights = True Then
			Response.Write(DateField("FirstApprovalDate", FirstApprovalDate, MVCPARights))
		End If
%></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold; white-space: nowrap; ">Second Approver</td>
	<td><%
		If IsNull(SecondApprovalID) = False Then
			Response.Write(DateField("SecondApprovalDate", SecondApprovalDate, MVCPARights) & ", " & SecondApprovalName)
		ElseIf MVCPARights = True And (FirstApprovalID<>UserSystemID Or UserSystemID=2 Or UserSystemID=402) Then
			Response.Write(DateField("SecondApprovalDate", SecondApprovalDate, (MVCPARights And ISNULL(FirstApprovalDate)=False)))
		ElseIf MVCPARights = True And FirstApprovalID=UserSystemID Then
			Response.Write(DateField("SecondApprovalDate", SecondApprovalDate, False) & "Second approver must be different from first approver.")
		ElseIf MVCPARights = True Then
			Response.Write(DateField("SecondApprovalDate", SecondApprovalDate, False) & "First Approval must be completed prior to second approval.")
		End If
%></td>
<%	End If %>
</tr>
<%	If IsNull(OldAdministrativeNotes) = False Then %>
<tr style="vertical-align: top; ">
	<td style="text-align: right; font-weight: bold; ">Administrative Notes<br />attached to Inventory Record</td>
	<td><%=OldAdministrativeNotes%></td>
</tr>
<%	End If %>
<tr style="vertical-align: top; ">
	<td style="text-align: right; font-weight: bold; white-space: nowrap; ">Administrative Notes</td>
	<td><%=TextArea("AdministrativeNotes", AdministrativeNotes, 4, 80, 2000, MVCPARights, "")%></td>
</tr>
<%	If IsNull(SubmitID) = False Then %>
<tr style="vertical-align: top; ">
	<td style="text-align: right; font-weight: bold; white-space: nowrap; ">Unsubmit</td>
	<td><%=CheckboxField("Unsubmit", False)%></td>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>
<%
	End If
	If MVCPAAdministrator = True Or MVCPAGrantCoordinator = True Or Developer = True Then
%>
<tr style="vertical-align: top; ">
	<td style="text-align: right; font-weight: bold; white-space: nowrap; ">Reject and Delete</td>
	<td><%=CheckboxField("Reject", False)%></td>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>
<%	
	End If 
	If MVCPARights = True And IsNull(FirstApprovalDate) = False And IsNull(SecondApprovalDate) = False Then
		' Can apply changes if approvals and not phase II or
		' if approvals and phase II and not being disposed of or
		' if approvals and phase II and date of disposal present
		If (PhaseII = False) Or _
			(PhaseII = True And IsNull(DisposalID) = True) Or _
			(PhaseII = True And IsNull(DisposalID) = False And IsNull(DateOfDisposal) = False) Then 
%>
<tr><td colspan="2">&nbsp;</td></tr>
<tr style="vertical-align: top; ">
	<td colspan=""2"" style="text-align: right; font-weight: bold; white-space: nowrap; ">Apply Changes</td>
	<td><%=CheckBoxField("ApplyChanges",False) %></td></tr>
</tr>
<%	
		ElseIf IsNull(DisposalID) = False Then
			If IsNull(DateOfDisposal) = True Then
				Response.Write("<tr><td colspan=""2"" style=""text-align: center; color: red; "">There should be a date of disposal if a method of disposal is chosen.</td></tr>" & vbCrLf)
			End If
			If (DisposalID=1 Or DisposalID=2) And IsNull(SalePrice)=True Then
				Response.Write("<tr><td colspan=""2"" style=""text-align: center; color: red; "">There should be a Net Sales Price if an item is marked as Auction or Sold.</td></tr>" & vbCrLf)
			End If
		End If
	End If
	If MVCPARights = True Then
		Response.Write("<tr><td colspan=""2"" style=""text-align: center;""><a href=""ItemUpdateHistory.asp?InventoryID=" & InventoryID & """ target=""_blank"">Show Update History</a></td>")
	End If
%>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<%
End If
%>
<tr>
	<td colspan="2" style="text-align: center">
<%	If PermitEdit = True Or PhaseII = True Or MVCPARights = True Then %>
		<input type="button" name="SaveBtn" value="Save" onclick="return submitForm('save');"/>&nbsp;&nbsp;
<%		If CanSubmit = True Then %>
		<input type="button" name="SubmitBtn" value="Submit" onclick="return submitForm('submit');" />&nbsp;&nbsp;
<%		ElseIf PhaseII = True And MVCPARights = False Then %>
		<input type="button" name="PhaseIIBtn" value="Submit" title="Update Disposal Date and Net Sales Price" onclick="return submitForm('PhaseII');" />&nbsp;&nbsp;
<%		End If %>
		<input type="reset" name="reset" value="Reset" />&nbsp;&nbsp;
<%	End If %>
		<input type="button" name="close" value="Close" onclick="window.close();" />
	</td>
</tr>
<%	If UpdateRecord = 0 Then %>
<tr>
	<td colspan="2" style="text-align: center">Saving this record will start the Inventory Update Process for this item.</td>
</tr>
<%	End If %>
</form>
</table>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->