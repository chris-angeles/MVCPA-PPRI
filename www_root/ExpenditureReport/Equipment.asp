<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, GrantID, Quarter, EquipmentID, GranteeID, GranteeName, Submitted, _
	AssetClassID, ItemDescription, ModelYear, MakeManufacturer, Model, UseID, SerialNo, _
	SourceID, OwnerID, AcquisitionDate, Cost, MVCPAPercentOfCost, Location, ConditionID, _
	OtherInformation, UpdateID, UpdateTimestamp, AddEdit

Debug = False
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
	Response.Flush
End If

If Len(Request.Form("GrantID"))>0 Then
	GrantID = Request.Form("GrantID")
ElseIf Len(Request.QueryString("GrantID"))>0 Then
	GrantID = Request.QueryString("GrantID")
Else
	Response.Write("Error: No GrantID provided for Equipment page.")
	SendMessage "Error: No GrantID provided for Equipment page."
	Response.End
End If
If Len(Request.Form("Quarter"))>0 Then
	Quarter = Request.Form("Quarter")
ElseIf Len(Request.Querystring("Quarter"))>0 Then
	Quarter = Request.Querystring("Quarter")
Else
	Response.Write("Error: No Quarter provided for Equipment page.")
	SendMessage "Error: No Quarter provided for Equipment page."
	Response.End
End If
If Len(Request.Form("EquipmentID")) > 0 Then
	EquipmentID = CInt(Request.Form("EquipmentID"))
	AddEdit = "Edit"
Else
	EquipmentID = 0
	AddEdit = "Add"
End If

sql = "SELECT B.GrantID, B.GranteeID, A.GranteeName, CAST(CASE WHEN C.SubmitID>0 THEN 1 ELSE 0 END AS BIT) AS Submitted " & vbCrLF & _
	"FROM Grantees AS A " & vbCrLF & _
	"JOIN [Grants].Main AS B ON A.GranteeID=B.GranteeID " & vbCrLf & _
	"JOIN [ER].Main AS C ON C.GrantID=B.GrantID AND C.Quarter=" & prepIntegerSQL(Quarter) & vbCrLf & _
	"WHERE B.GrantID=" & prepIntegerSQL(GrantID)
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	GranteeName = rs.Fields("GranteeName")
	GranteeID = rs.Fields("GranteeID")
	Submitted = rs.Fields("Submitted")
End If
' Add code to check for submission.
If Submitted = True Then 
	PermitEdit = False
Else
	PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
End If

If EquipmentID>0 Then
	sql = "SELECT * FROM ER.EquipmentDetail WHERE EquipmentID=" & prepIntegerSQL(EquipmentID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error: No Equipment Detail record retrieved, " & EquipmentID)
		SendMessage "Error: No Equipment Detail record retrieved, " & EquipmentID
		Response.End
	Else
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
		OtherInformation = rs.Fields("OtherInformation")
		UpdateID = rs.Fields("UpdateID")
		UpdateTimestamp = rs.Fields("UpdateTimestamp")
	End If
Else
	AssetClassID = 0
	ItemDescription = ""
	ModelYear = ""
	MakeManufacturer = ""
	Model = ""
	UseID = 0
	SerialNo = ""
	SourceID = 0
	OwnerID = 0
	AcquisitionDate = ""
	Cost = 0
	MVCPAPercentOfCost = 100
	Location = ""
	ConditionID = 0
	OtherInformation = ""
	UpdateID = UserSystemID
	UpdateTimestamp = ""

End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Expenditure Report Equipment Detail</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function selectEquipment(eid) {
		document.Selection.EquipmentID.value = eid;
		document.Selection.submit();
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body>
<form name="Selection" id="Selection" method="post" action="Equipment.asp">
<%=HiddenField("GrantID", GrantID) %><%=HiddenField("Quarter", Quarter) %><%=HiddenField("EquipmentID", EquipmentID) %>
</form>
<table style="margin: auto; ">
	<thead>
		<tr><th colspan="5">Equipment Detail</th></tr>
		<tr>
			<th colspan="2">Asset Class</th>
			<th>Description</th>
			<th>Serial No. / VIN</th>
			<th>Cost</th>
		</tr>
	</thead>
	<tbody>
<%
sql = "SELECT A.EquipmentID, A.AssetClassID, B.AssetClassShort, A.ItemDescription, A.SerialNo, A.Cost  " & vbCrLF & _
	"FROM ER.EquipmentDetail AS A " & vbCrLf & _
	"LEFT JOIN Lookup.InventoryAssetClass AS B ON B.AssetClassID=A.AssetClassID " & vbCrLf & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"ORDER BY AssetClassID, EquipmentID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("<tr><td colspan=""5"" style=""text-align: center"">No Equipment Detail Records Found</td></tr>" & vbCrLF)
Else
	While rs.EOF = False
		Response.Write("<tr>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("AssetClassID") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("AssetClassShort") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("ItemDescription") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("SerialNo") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("Cost")) & "</td>" & vbCrLf)
		If Submitted = False Then
			Response.Write(vbTab & "<td><a href=""javascript:selectEquipment(" & rs.Fields("EquipmentID") & ");"" class=""plainlink"">Edit</a></td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		End If
		Response.Write("<tr>" & vbCrLf)
		rs.MoveNext
	Wend
End If

%>
	</tbody>
	<tfoot>
		<tr><td colspan="5" style="text-align: center;"><a href="javascript:selectEquipment(0);">Add a new equipment item</a></td></tr>
	</tfoot>
</table>
<br>
<form name="Equipment" method="post" action="EquipmentSubmit.asp">
<%=HiddenField("GrantID", GrantID) %><%=HiddenField("Quarter", Quarter) %><%=HiddenField("EquipmentID",EquipmentID) %>

<table style="margin: auto;">
<thead>
<tr><th colspan="2">Add / Edit Equipment</th></tr>
<tbody>
	<tr>
		<td colspan="2" style="text-align: center; "><%
			If EquipmentID=0 Then
				Response.Write("Add Equipment Item")
			Else
				Response.Write("Edit EquipmentID=" & EquipmentID & " Item.")
			End If
		%></td>
	</tr>
	<tr>
		<td>Asset Class</td>
		<td><select name="AssetClassID" id="AssetClassID" style="width: 750px; ">
			<option value="0">Select Asset Class</option>
<%
sql = "SELECT AssetClassID, CAST(AssetClassID AS VARCHAR) + ' ' + AssetClass AS AssetClass FROM Lookup.InventoryAssetClass ORDER BY AssetClassSort"
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("AssetClassID") & """" & Selected(rs.Fields("AssetClassID"), AssetClassID) & ">" & rs.Fields("AssetClass") & "</option>" & vbCrLf)
	rs.MoveNext
Wend
%>
		    </select></td>
	</tr>
	<tr>
		<td>Item Description</td>
		<td><%=TextField("ItemDescription", ItemDescription, 100, 254, PermitEdit, "")%></td>
	</tr>
	<tr>
		<td>ModelYear</td>
		<td><%=IntegerField("ModelYear", ModelYear, 4, 4, PermitEdit, "")%></td>
	</tr>
	<tr>
		<td>Make or Manufacturer</td>
		<td><%=TextField("MakeManufacturer", MakeManufacturer, 40, 50, PermitEdit, "")%></td>
	</tr>
	<tr>
		<td>Model</td>
		<td><%=TextField("Model", Model, 40, 250, PermitEdit, "")%></td>
	</tr>
	<tr>
		<td>Use</td>
		<td><select name="UseID" id="UseID">
			<option value="0">Select Use</option>
<%
sql = "SELECT UseID, [Use] FROM Lookup.InventoryUse ORDER BY UseSort"
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("UseID") & """" & Selected(rs.Fields("UseID"), UseID) & ">" & rs.Fields("Use") & "</option>" & vbCrLf)
	rs.MoveNext
Wend
%>
		    </select></td>
	</tr>
	<tr>
		<td>SerialNo / VIN</td>
		<td><%=TextField("SerialNo", SerialNo, 17, 50, PermitEdit, "")%></td>
	</tr>
	<tr>
		<td>Source</td>
		<td><select name="SourceID" id="SourceID">
			<option value="0">Select Source</option>
<%
sql = "SELECT SourceID, [Source] FROM Lookup.InventorySource ORDER BY SourceSort"
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("SourceID") & """" & Selected(rs.Fields("SourceID"), SourceID) & ">" & rs.Fields("Source") & "</option>" & vbCrLf)
	rs.MoveNext
Wend
%>
		    </select></td>
	</tr>
	<tr>
		<td>Owner (Title held by)</td>
		<td><select name="OwnerID" id="OwnerID">
			<option value="0">Select Owner</option>
<%
sql = "SELECT OwnerID, [Owner] FROM Lookup.InventoryOwner ORDER BY OwnerSort"
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("OwnerID") & """" & Selected(rs.Fields("OwnerID"), OwnerID) & ">" & rs.Fields("Owner") & "</option>" & vbCrLf)
	rs.MoveNext
Wend
%>
		    </select></td>
	</tr>
	<tr>
		<td>Acquisition Date</td>
		<td><%=DateField("AcquisitionDate", AcquisitionDate, PermitEdit)%></td>
	</tr>
	<tr>
		<td>Cost</td>
		<td><%=CurrencyField("Cost", Cost, 11, 15, PermitEdit, "") %></td>
	</tr>
	<tr>
		<td>MVCPA Percent Of Cost</td>
		<td><%=IntegerField("MVCPAPercentOfCost", MVCPAPercentOfCost, 3, 3, PermitEdit, "")%>%</td>
	</tr>
	<tr>
		<td>Location</td>
		<td><%=TextField("Location", Location, 40, 98, PermitEdit, "")%></td>
	</tr>
	<tr>
		<td>Condition</td>
		<td><select name="ConditionID" id="ConditionID">
			<option value="0">Select Condition</option>
<%
sql = "SELECT ConditionID, [Condition] FROM Lookup.InventoryCondition ORDER BY ConditionSort"
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("ConditionID") & """" & Selected(rs.Fields("ConditionID"), ConditionID) & ">" & rs.Fields("Condition") & "</option>" & vbCrLf)
	rs.MoveNext
Wend
%>
		    </select></td>
	</tr>
	<tr>
		<td>Other Information</td>
		<td><%=TextField("OtherInformation", OtherInformation, 100, 254, PermitEdit, "")%></td>
	</tr>
	<tr>
		<td colspan="2"></td>
	</tr>
</tbody>
	<tfoot>
		<tr><td colspan="5" style="text-align: center;">
			<input type="submit" name="Submit" value="Save" />
			<input type="button" name="Done" value="Done" 
			onclick="location.href='Report.asp?GrantID=<%=GrantID%>&Quarter=<%=Quarter%>';" 
			title="Return to expenditure report losing unsaved changes if any."/>
		</td></tr>
	</tfoot>
</table>
</form>

</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<%
Function ReportingPeriodDates(vFiscalYear, vReportingPeriod)
	If vReportingPeriod = 1 Then
		ReportingPeriodDates = "September 1, " & (vFiscalYear-1) & " - November 30, " & (vFiscalYear-1)
	ElseIf vReportingPeriod = 2 and (vFiscalYear Mod 4 = 0) Then
		ReportingPeriodDates = "December 1, " & (vFiscalYear-1) & " - February 29, " & vFiscalYear
	ElseIf vReportingPeriod = 2 Then
		ReportingPeriodDates = "December 1, " & (vFiscalYear-1) & " - February 28, " & vFiscalYear
	ElseIf vReportingPeriod = 3 Then
		ReportingPeriodDates = "March 1, " & vFiscalYear & " - May 31, " & vFiscalYear
	ElseIf vReportingPeriod = 4 Then
		ReportingPeriodDates = "June 1, " & vFiscalYear & " - August 31, " & vFiscalYear
	Else
		ReportingPeriodDates = "Error in Reporting Period"
	End If
End Function
%>