<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim Debug, PermitEdit, InventoryID, GranteeID, AssetClassID, ItemDescription, ModelYear, _
	MakeManufacturer, Model, UseID, SerialNo, SourceID, OwnerID, AcquisitionDate, _
	Cost, MVCPAPercentOfCost, Location, ConditionID, NotInventoryItem, _
	DateOfDisposal, DisposalID, SalePrice, OtherInformation, AdministrativeNotes, _
	UpdateID, UpdateTimestamp, GranteeNAme, UpdateName
Debug = False

If Len(Request.Form("InventoryID"))>0 Then
	InventoryID = CInt(Request.Form("InventoryID"))
ElseIf Len(Request.QueryString("InventoryID"))>0 Then
	InventoryID = CInt(Request.QueryString("InventoryID"))
Else
	InventoryID = 0
End If

IF MVCPARights = True Then
	PermitEdit = True
Else
	PermitEdit = False
End If
If InventoryID>0 Then
	sql = "SELECT I.*, G.GranteeName, " & vbCrLf & _
		"	CASE WHEN I.UpdateID=2147483647 THEN 'Imported' ELSE U.Name END AS UpdateName " & vbCrLf & _
		"FROM Inventory AS I" & vbCrLF & _
		"LEFT JOIN System.[Users] AS U ON U.SystemID=I.UpdateID " & vbCrLf & _
		"LEFT JOIN Grantees AS G ON G.GranteeID=I.GranteeID " & vbCrLf & _
		"WHERE InventoryID=" & prepIntegerSQL(InventoryID)
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
		ItemDescription = rs.Fields("ItemDescription")
		ModelYear = rs.Fields("ModelYear")
		MakeManufacturer = rs.Fields("MakeManufacturer")
		Model = rs.Fields("Model")
		UseID = rs.Fields("UseID")
		SerialNo = rs.Fields("SerialNo")
		SourceID = rs.Fields("SourceID")
		OwnerID = rs.Fields("OwnerID")
		AcquisitionDate = rs.Fields("AcquisitionDate")
		Cost = rs.Fields("Cost")
		MVCPAPercentOfCost = rs.Fields("MVCPAPercentOfCost")
		Location = rs.Fields("Location")
		ConditionID = rs.Fields("ConditionID")
		NotInventoryItem = rs.Fields("NotInventoryItem")
		DateOfDisposal = rs.Fields("DateOfDisposal")
		DisposalID = rs.Fields("DisposalID")
		SalePrice = rs.Fields("SalePrice")
		OtherInformation = rs.Fields("OtherInformation")
		AdministrativeNotes = rs.Fields("AdministrativeNotes")
		UpdateID = rs.Fields("UpdateID")
		UpdateName = rs.Fields("UpdateName")
		UpdateTimestamp = rs.Fields("UpdateTimestamp")
	End If
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Inventory Item Edit Page</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body>
<h1>Administrative Inventory Item Edit Page</h1>
<table>
<form name="Inventory" method="post" action="EditSubmit.asp">
<%=Hiddenfield("InventoryID", InventoryID) %>
<tr>
	<td style="text-align: right; font-weight: bold;">InventoryID</td>
	<td><%=InventoryID %></td>
</tr>
<%	If InventoryID=0 Then %>
<tr>
	<td style="text-align: right; font-weight: bold;">Grantee</td>
	<td><select name="GranteeID" id="GranteeID">
		<option value="0">Select Grantee</option>
<%
	sql = "SELECT GranteeID, GranteeName FROM Grantees ORDER BY Replace(GranteeName, 'City Of ', '')"
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("GranteeID") & """ " & Selected(rs.Fields("GranteeID"), GranteeID) & ">" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr><%	Else %>
<tr>
	<td style="text-align: right; font-weight: bold;">Grantee</td>
	<td><%=GranteeID %>: <%=GranteeName %></td>
</tr>
<%	End If %>
<tr>
	<td style="text-align: right; font-weight: bold;">AssetClass</td>
	<td><select name="AssetClassID" id="AssetClassID">
		<option value="0">Select Asset Class</option>
<%
	sql = "SELECT AssetClassID, AssetClassShort FROM Lookup.InventoryAssetClass ORDER BY AssetClassSort"
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("AssetClassID") & """ " & Selected(rs.Fields("AssetClassID"), AssetClassID) & ">" & AssetClassID & ": " & rs.Fields("AssetClassShort") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">ItemDescription</td>
	<td><%=TextField("ItemDescription", ItemDescription, 80, 255, PermitEdit, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">ModelYear</td>
	<td><%=IntegerField("ModelYear", ModelYear, 4, 4, MVCPARights, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Make or Manufacturer</td>
	<td><%=TextField("MakeManufacturer", MakeManufacturer, 50, 50, PermitEdit, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Model</td>
	<td><%=TextField("Model", Model, 80, 255, PermitEdit, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Use</td>
	<td><select name="UseID" id="UseID">
		<option value="0">Select Use</option>
<%
	sql = "SELECT UseID, [Use] FROM Lookup.InventoryUse ORDER BY UseSort"
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("UseID") & """ " & Selected(rs.Fields("UseID"), UseID) & ">" & rs.Fields("Use") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">SerialNo or VIN</td>
	<td><%=TextField("SerialNo", SerialNo, 50, 100, PermitEdit, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">How Acquired</td>
	<td><select name="SourceID" id="SourceID">
		<option value="0">Select How Acquired</option>
<%
	sql = "SELECT SourceID, Source FROM Lookup.InventorySource ORDER BY SourceSort"
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("SourceID") & """ " & Selected(rs.Fields("SourceID"), SourceID) & ">" & rs.Fields("Source") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Who holds title?</td>
	<td><select name="OwnerID" id="OwnerID">
		<option value="0">Select Who holds title?</option>
<%
	sql = "SELECT OwnerID, Owner FROM Lookup.InventoryOwner ORDER BY OwnerSort"
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("OwnerID") & """ " & Selected(rs.Fields("OwnerID"), OwnerID) & ">" & rs.Fields("Owner") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Acquisition Date</td>
	<td><%=DateField("AcquisitionDate", AcquisitionDate, MVCPARights) %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Cost</td>
	<td><%=CurrencyField("Cost", Cost, 12, 12, MVCPARights, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">MVCPA Percent Of Cost</td>
	<td><%=IntegerField("MVCPAPercentOfCost", MVCPAPercentOfCost, 3, 3, MVCPARights, "") %></td>
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
	sql = "SELECT ConditionID, Condition FROM Lookup.InventoryCondition ORDER BY ConditionSort"
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("ConditionID") & """ " & Selected(rs.Fields("ConditionID"), ConditionID) & ">" & rs.Fields("Condition") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Not An InventoryItem Date</td>
	<td><%
If IsNull(NotInventoryItem) Then
	Response.Write(CheckBoxField("NotInventoryItem", False))
Else
	Response.Write(DateField("NotInventoryItem", NotInventoryItem, MVCPARights)) 
End If
%> This item should not have been included as an inventory item.</td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Date Of Disposal</td>
	<td><%=DateField("DateOfDisposal", DateOfDisposal, MVCPARights) %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Method of Disposal</td>
	<td><select name="DisposalID" id="DisposalID">
		<option value="0">Select method of disposal</option>
<%
	sql = "SELECT DisposalID, Disposal FROM Lookup.InventoryDisposal ORDER BY DisposalSort"
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("DisposalID") & """ " & Selected(rs.Fields("DisposalID"), DisposalID) & ">" & rs.Fields("Disposal") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
%></select></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Net Sale Price</td>
	<td><%=CurrencyField("SalePrice", SalePrice, 12, 12, MVCPARights, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold; vertical-align: top; ">Other Information</td>
	<td><%=TextArea("OtherInformation", OtherInformation, 4, 80, 1000, MVCPARights, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold; vertical-align: top; ">Administrative Notes</td>
	<td><%=TextArea("AdministrativeNotes", AdministrativeNotes, 4, 80, 2000, MVCPARights, "") %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">UpdateID</td>
	<td><%=UpdateID %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Update Name</td>
	<td><%=UpdateName %></td>
</tr>
<tr>
	<td style="text-align: right; font-weight: bold;">Update Timestamp</td>
	<td><%=UpdateTimestamp %></td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<tr>
	<td colspan="2" style="text-align: center"><input type="submit" name="Save" value="Save" />&nbsp;&nbsp;
		<input type="reset" name="reset" value="Reset" />&nbsp;&nbsp;
		<input type="button" name="close" value="Close" onclick="window.close();" />
	</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->