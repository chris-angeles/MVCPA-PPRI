<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim Debug, i, PermitEdit, columns, GranteeID, GranteeName, GrantID, FiscalYear, ProgramName, _
	GrantNumber, Submitted, AssetClassID, GranteeComments, CompleteBy, CompleteByDate, _
	SubmitID, SubmitName, SubmitTimestamp, AdministrativeNotes, AcceptanceID, AcceptanceDate, _
	ProgramDirectorID, FinancialOfficerID, FinancialAdministrativeContactID, LostStolenDestroyed, PhysicalInventoryYear
Debug = False
columns = 7

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

If Len(Request.Form("GranteeID")) > 0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID")) > 0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	GranteeID = Session("GranteeID")
End If
If Len(Request.Form("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	FiscalYear = Session("FiscalYear")
End If

sql = "SELECT A.GranteeID, A.GranteeName, B.GrantID, B.FiscalYear, B.ProgramName, B.GrantNumber, " & vbCrLf & _
	"	GranteeComments, CompleteByDate, CompleteBy, SubmitID, D.Name AS SubmitName, SubmitTimestamp, " & vbCrLf & _
	"	AdministrativeNotes, AcceptanceID, AcceptanceDate, " & vbCrLf & _
	"	ISNULL(ProgramDirectorID, 0) AS ProgramDirectorID, ISNULL(A.FinancialOfficerID,0) AS FinancialOfficerID, " & vbCrLF & _
	"	ISNULL(A.FinancialAdministrativeContactID,0) AS FinancialAdministrativeContactID, " & vbCrLf & _
	"	LostStolenDestroyed = ISNULL((SELECT SUM(CASE WHEN DisposalID=7 THEN 1 ELSE 0 END) AS LostStolenDestroyed " & vbCrLf & _
	"		FROM Inventory " & vbCrLf & _
	"		WHERE GranteeID=A.GranteeID AND ISNULL(AcquisitionDate,'9/1/" & (FiscalYear-1) & "')<'9/1/" & (FiscalYear) & "' AND " & vbCrLf & _
	"			ISNULL(DateOfDisposal,'12/31/2099')>'8/31/" & FiscalYear & "' " & vbCrLf & _
	"			AND NotInventoryItem IS NULL),0) " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"JOIN [Grants].Main AS B ON A.GranteeID=B.GranteeID And B.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
	"LEFT JOIN [Grants].InventoryCertification AS C ON C.GrantID=B.GrantID " & vbCrLf & _
	"LEFT JOIN [System].Users AS D ON D.SystemID=C.SubmitID " & vbCrLF & _
	"WHERE A.GranteeID=" & prepIntegerSQL(GranteeID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No record retrieved for Grant/Grantee")
	SendMessage "Error: No record retrieved for Grant/Grantee"
	Response.End
Else
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	GrantID = rs.Fields("GrantID")
	FiscalYear = rs.Fields("FiscalYear")
	ProgramName = rs.Fields("ProgramName")
	GrantNumber = rs.Fields("GrantNumber")
	GranteeComments = rs.Fields("GranteeComments")
	CompleteBy = rs.Fields("CompleteBy")
	CompleteByDate = rs.Fields("CompleteByDate")
	SubmitID = rs.Fields("SubmitID")
	SubmitName = rs.Fields("SubmitName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	AdministrativeNotes = rs.Fields("AdministrativeNotes")
	AcceptanceID = rs.Fields("AcceptanceID")
	AcceptanceDate = rs.Fields("AcceptanceDate")
	ProgramDirectorID = rs.Fields("ProgramDirectorID")
	FinancialOfficerID = rs.Fields("FinancialOfficerID")
	FinancialAdministrativeContactID = rs.Fields("FinancialAdministrativeContactID")
	LostStolenDestroyed = rs.Fields("LostStolenDestroyed")
	If IsNull(SubmitID) = False Then
		Submitted = True
	Else
		Submitted = False
	End If
End If

If FiscalYear MOD 2 = 1 Then
	PhysicalInventoryYear = True
Else
	PhysicalInventoryYear = False
End If

PermitEdit = CheckPermissions(UserSystemID, GranteeID, Submitted)
If UserSystemID<>FinancialOfficerID And UserSystemID<>FinancialAdministrativeContactID And UserSystemID<>ProgramDirectorID Then
	PermitEdit = False
End If

If Debug = True Then
	Response.Write("<pre>FinancialOfficerID=" & FinancialOfficerID & "; FinancialAdministrativeContactID=" & FinancialAdministrativeContactID & "; SystemID=" & UserSystemID & "; GranteeID=" & GranteeID & "; PermitEdit=" & PermitEdit & "; Submitted=" & Submitted & "; PhysicalInventoryYear=" & PhysicalInventoryYear & ".</pre>")
End If

If Submitted = True Then ' Pull from Inventory Certification Detail
	sql = "SELECT B.AssetClassID, B.AssetClassShort, InventoryID, " & prepIntegerSQL(GranteeID) & " AS GranteeID, ItemDescription, ModelYear, MakeManufacturer, Model, SerialNo, Location, 0 AS DisposalID " & vbCrLf & _
		"FROM [Grants].InventoryCertificationDetail AS A " & vbCrLF & _
		"LEFT JOIN Lookup.InventoryAssetClass AS B ON B.AssetClassID=A.AssetClassID " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " " & vbCrLF & _
		"ORDER BY B.AssetClassID, InventoryID "
Else ' Load from Inventory.
	sql = "SELECT B.AssetClassID, B.AssetClassShort, InventoryID, GranteeID, ItemDescription, ModelYear, MakeManufacturer, Model, SerialNo, Location, DisposalID " & vbCrLf & _
		"FROM Inventory AS A " & vbCrLF & _
		"LEFT JOIN Lookup.InventoryAssetClass AS B ON B.AssetClassID=A.AssetClassID " & vbCrLf & _
		"WHERE GranteeID=" & prepIntegerSQL(GranteeID) & " AND ISNULL(AcquisitionDate,'9/1/" & (FiscalYear-1) & "')<'9/1/" & (FiscalYear) & "' AND " & vbCrLF & _
		"	ISNULL(DateOfDisposal,'12/31/2099')>'8/31/" & FiscalYear & "' " & vbCrLf & _
		"	AND NotInventoryItem IS NULL " & vbCrLf & _
		"ORDER BY B.AssetClassID, InventoryID "
End If
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No record retrieved for Grant/Grantee")
	SendMessage "Error: No record retrieved for Grant/Grantee"
	Response.End
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Inventory Certification</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm(action)
	{
<%
If PhysicalInventoryYear = True Then
%>
		if (document.Certification.CompleteByDate.value.length==0) {
			alert("You must indicate date that physical inventory was completed to submit this form.");
			document.Certification.CompleteByDate.focus();
			return false;
		}
		if (document.Certification.CompleteBy.value.length==0) {
			alert("You must indicate the individuals who completed the physical inventory to submit this form.");
			document.Certification.CompleteBy.focus();
			return false;
		}
		if (action == "submit" && document.Certification.Inventory.checked == false) {
			alert("You must check the physical inventory completed box to submit this form.");
			document.Certification.Inventory.focus();
			return false;
		}
		<%
End If
%>
		if (action == "submit" && document.Certification.Certify.checked == false) {
			alert("You must check the Certification box to submit this form.");
			document.Certification.Certify.focus();
			return false;
		}
		if (validateForm() == true) {
			document.Certification.Action.value = action;
			document.Certification.submit();
		}
	}

	function validateForm()
	{
		return true;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body>
<h1>Inventory Certification</h1>
<h2><%=GranteeName %><br />
<%=ProgramName %><br />
<%=GrantNumber %><br />
For the Fiscal Year Ending August 31, <%=FiscalYear %></h2>

<%	If IsNull(SubmitTimestamp) = False Then  %>
<p style="margin: auto; text-align: center;">Submitted by <%=SubmitName %> at <%=SubmitTimestamp %></p>
<%	End If %>
<p>If any corrections need to be made to this inventory certification, use the inventory update process external to this certification to submit the changes for approval.</p>

<%
If Debug = True Then
	Response.Write("<pre>LostStolenDestroyed=" & LostStolenDestroyed & "</pre>" & vbCrLf)
End If
If LostStolenDestroyed=1 Then
	Response.Write("<p style=""margin: auto; text-align: center; color: red; font-weight: bold; "">There is one item that has been marked as lost, stolen or destroyed.</p><br />" & vbCrLf)
ElseIf LostStolenDestroyed>1 Then
	Response.Write("<p style=""margin: auto; text-align: center; color: red; font-weight: bold; "">There are " & LostStolenDestroyed & " items that have been marked as lost, stolen or destroyed.</p><br />" & vbCrLf)
End If
%>
<table>
	<tr style="vertical-align: bottom">
		<th>ID</th>
		<th>Description</th>
		<th>Model Year</th>
		<th>Make/Manufacturer</th>
		<th>Model</th>
		<th>Serial Number / VIN</th>
		<th>Location</th>
	</tr>
<%
If rs.EOF = False Then
	AssetClassID = 0
	While rs.EOF = False
		If AssetClassID <> rs.Fields("AssetClassID") Then
			Response.Write("<tr><th colspan=""" & columns & """ style=""color: blue; "">" & rs.Fields("AssetClassShort") & "</th></tr>" & vbCrLf)
			AssetClassID = rs.Fields("AssetClassID")
		End If
		If rs.Fields("DisposalID") = 7 And FiscalYear>2020 Then
			Response.Write("<tr style=""vertical-align: top; background-color: tomato; "">" & vbCrLf)
		Else
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		End If
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("InventoryID") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("ItemDescription") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; "">" & rs.Fields("ModelYear") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("MakeManufacturer") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("Model") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("SerialNo") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("Location") & "</td>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
End If
%>
</table>
<br />
<form name="Certification" method="post" action="CertifySubmit.asp">
<%=Hiddenfield("GranteeID", GranteeID) %><%=Hiddenfield("FiscalYear",FiscalYear) %><%=Hiddenfield("GrantID",GrantID) %><%=HiddenField("Action","submit") %>

<%
If PhysicalInventoryYear = True Then
	Response.Write(CheckBoxField("Inventory", Submitted))
	Response.Write("I certify that a physical inventory has been conducted and all items shown are within the agency's possession and accurately reflected in this inventory system. ")
	Response.Write("The physical inventory of the above items was completed on (date <i>mm/dd/yyyy</i>): ")
	Response.Write(DateField("CompleteByDate", CompleteByDate, PermitEdit) & "<br />" & vbCrLf)
	Response.Write("The physical inventory was completed by (Please include name(s), title, email address, and phone for each person completing the inventory at each location): <br />")
	Response.Write(TextArea("CompleteBy", CompleteBy, 4, 100, 1024, PermitEdit, "") & "<br />")
	Response.Write("<br />")
End If
%>
<%=CheckBoxField("Certify", Submitted) %> I have reviewed and confirmed the information in this 
inventory report and I attest that this report is correct and complete for purposes set forth in the 
MVCPA Grant Administrative Manual.  I am aware that any false, fictitious, or fraudulent information 
may subject me to criminal, civil, or administrative penalties.<br /><br />
<%
If PermitEdit = True and Submitted = False Then
	Response.Write("Comments:" & TextArea("GranteeComments", GranteeComments, 4, 100, 2000, PermitEdit, "") & "<br />")
ElseIf IsNull(GranteeComments) = False Then
	Response.Write("Comments:" & GranteeComments)
End If
If MVCPARights = True And Submitted = True Then
	Response.Write("<br /><div style=""margin: auto; text-align: center;"">Administrative Section</div>" & vbCrLf)
	Response.Write("Administrative Notes:" & TextArea("AdministrativeNotes", AdministrativeNotes, 4, 100, 2000, MVCPARights, "") & "<br />")
%>MVCPA Acceptance Date: <%=DateField("AcceptanceDate", AcceptanceDate, MVCPARights) %><br />
<%=CheckBoxField("Unsubmit", False) %> Unsubmit
<%
End If

If Submitted = True Then
sql = "SELECT A.InventoryID, B.GrantID, A.AssetClassID, A.ItemDescription, A.ModelYear, " & vbCrLf & _
	"	A.MakeManufacturer, A.Model, A.SerialNo, A.Location, C.AssetClassShort " & vbCrLf & _
	"FROM Inventory AS A " & vbCrLf & _
	"LEFT JOIN [Grants].InventoryCertificationDetail AS B ON A.InventoryID=B.InventoryID AND B.GrantID=" & GrantID & " " & vbCrLf & _
	"LEFT JOIN [Lookup].[InventoryAssetClass] AS C ON C.AssetClassID=A.AssetClassID " & vbCrLf & _
	"WHERE A.GranteeID=" & GranteeID & " AND B.InventoryID IS NULL AND AcquisitionDate<'10/1/" & FiscalYear & "' AND ISNULL(DateOfDisposal,'10/1/" & FiscalYear & "')>='10/1/" & FiscalYear & "' AND NotInventoryItem IS NULL " & vbCrLf & _
	"UNION " & vbCrLf
Else
	sql = ""
End If
sql = sql & "SELECT -A.EquipmentID AS InventoryID, E.GrantID, A.AssetClassID, A.ItemDescription, A.ModelYear, " & vbCrLf & _
	"	A.MakeManufacturer, A.Model, A.SerialNo, A.Location, C.AssetClassShort " & vbCrLf & _
	"FROM ER.EquipmentDetail AS A " & vbCrLf & _
	"LEFT JOIN [Grants].InventoryCertificationDetail AS B ON A.InventoryID=B.InventoryID AND B.GrantID=88 " & vbCrLf & _
	"LEFT JOIN [Lookup].[InventoryAssetClass] AS C ON C.AssetClassID=A.AssetClassID " & vbCrLf & _
	"LEFT JOIN Inventory AS D ON D.InventoryID=A.InventoryID " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS E ON E.GrantID=A.GrantID " & vbCrLf & _
	"WHERE E.GranteeID=" & GranteeID & " AND B.InventoryID IS NULL AND A.AcquisitionDate<'10/1/" & FiscalYear & "' AND ISNULL(DateOfDisposal,'10/1/" & FiscalYear & "')>='10/1/2020' " & vbCrLf & _
	"	AND D.InventoryID IS NULL " & vbCrLf & _
	"ORDER BY InventoryID "

	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
%>
<hr style="margin: 10px;" />
<div style="margin: auto; text-align: center; font-style: italic; ">The following items were added after certification because reimbursement was made on the 4th Quarter Expenditure Report after certification</div>
<table>
	<tr style="vertical-align: bottom">
		<th>ID</th>
		<th>Description</th>
		<th>Model Year</th>
		<th>Make/Manufacturer</th>
		<th>Model</th>
		<th>Serial Number / VIN</th>
		<th>Location</th>
	</tr>
<%

	AssetClassID = 0
	While rs.EOF = False
		If AssetClassID <> rs.Fields("AssetClassID") Then
			Response.Write("<tr><th colspan=""" & columns & """ style=""color: blue; "">" & rs.Fields("AssetClassShort") & "</th></tr>" & vbCrLf)
			AssetClassID = rs.Fields("AssetClassID")
		End If
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("InventoryID") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("ItemDescription") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; "">" & rs.Fields("ModelYear") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("MakeManufacturer") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("Model") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("SerialNo") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("Location") & "</td>" & vbCrLf)
		Response.Write("<tr>" & vbCrLf)
		rs.MoveNext
		Wend
	End If
%>
</table>
<br />
<div style="margin: auto; text-align: center"><%
If PermitEdit = True And Submitted = false And MVCPARights = False And LostStolenDestroyed=0 Then
%><input type="button" name="Submit" value="Submit" onclick="submitForm('submit')" />
<%
ElseIf MVCPARights = True And Submitted = True Then 
%><input type="button" name="Save" value="Save" onclick="submitForm('save')" />
<%
End If
%>


<input type="button" name="close" value="Close" title="Close window without changes" onclick="window.close();" /></div>
</form>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->