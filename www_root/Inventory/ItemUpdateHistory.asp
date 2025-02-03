<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, ShowExcel, InventoryID, counter
Debug = False

If Len(Request.Form("InventoryID"))>0 Then
	InventoryID = CInt(Request.Form("InventoryID"))
ElseIf Len(Request.QueryString("InventoryID"))>0 Then
	InventoryID = CInt(Request.QueryString("InventoryID"))
Else
	InventoryID = 0
	Response.Write("No Inventory ID Provided.")
	Response.End
End If

If Len(Request.Form("ShowExcel"))>0 Then
	If Request.Form("ShowExcel")="1" Then 
		ShowExcel = True
	Else
		ShowExcel = False
	End If
ElseIf Len(Request.QueryString("ShowExcel"))>0 Then
	If Request.QueryString("ShowExcel")="1" Then 
		ShowExcel = True
	Else
		ShowExcel = False
	End If
Else
	ShowExcel = False
End If

sql = "SELECT 999999 AS Log_ID, InventoryID AS Inventory_ID, OwnerID AS Owner_ID, " & vbCrLf & _
	"UseID, Location, ConditionID AS Condition_ID, NotInventoryItem AS Not_Inventory_Item, " & vbCrLf & _
	"	DateOfDisposal AS Date_Of_Disposal, DisposalID AS Disposal_ID, " & vbCrLf & _
	"	SalePrice AS Sale_Price, AdditionalInformation AS Additional_Information, " & vbCrLf & _
	"	SubmitID AS Submit_ID, SubmitTimestamp AS Submit_Timestamp, " & vbCrLf & _
	"	DisposalUpdateID AS Disposal_Update_ID, DisposalUpdateTimestamp AS Disposal_Update_Timestamp, " & vbCrLf & _
	"	FirstApprovalID AS First_Approval_ID, FirstApprovalDate AS First_Approval_Date, " & vbCrLf & _
	"	SecondApprovalID AS Second_Approval_ID, SecondApprovalDate AS Second_Approval_Date, " & vbCrLf & _
	"	RejectedDate AS Rejected_Date, " & vbCrLf & _
	"	REPLACE(REPLACE(AdministrativeNotes, CHAR(10),' '),CHAR(13),' ') AS Administrative_Notes, " & vbCrLf & _
	"	ChangesApplied AS Changes_Applied_Date, IU.UpdateID AS Update_ID, " & vbCrLf & _
	"	IU.UpdateTimestamp AS Update_Timestamp, U.Name AS Update_Name " & vbCrLf & _
	"FROM InventoryUpdate AS IU " & vbCrLf & _
	"LEFT JOIN [System].Users AS U ON U.SystemID=IU.UpdateID " & vbCrLf & _
	"WHERE InventoryID=" & prepIntegerSQL(InventoryID) & " " & vbCrLf & _
	"UNION " & vbCrLf & _
	"SELECT LogID AS Log_ID, InventoryID AS Inventory_ID, OwnerID AS Owner_ID, " & vbCrLf & _
	"	UseID, Location, ConditionID AS Condition_ID, NotInventoryItem AS Not_Inventory_Item, " & vbCrLf & _
	"	DateOfDisposal AS Date_Of_Disposal, DisposalID AS Disposal_ID, " & vbCrLf & _
	"	SalePrice AS Sale_Price, AdditionalInformation AS Additional_Information, " & vbCrLf & _
	"	SubmitID AS Submit_ID, SubmitTimestamp AS Submit_Timestamp, " & vbCrLf & _
	"	DisposalUpdateID AS Disposal_Update_ID, DisposalUpdateTimestamp AS Disposal_Update_Timestamp, " & vbCrLf & _
	"	FirstApprovalID AS First_Approval_ID, FirstApprovalDate AS First_Approval_Date, " & vbCrLf & _
	"	SecondApprovalID AS Second_Approval_ID, SecondApprovalDate AS Second_Approval_Date, " & vbCrLf & _
	"	RejectedDate AS Rejected_Date, " & vbCrLf & _
	"	REPLACE(REPLACE(AdministrativeNotes, CHAR(10),' '),CHAR(13),' ') AS Administrative_Notes, " & vbCrLf & _
	"	ChangesApplied AS Changes_Applied_Date, IU.UpdateID AS Update_ID, " & vbCrLf & _
	"	IU.UpdateTimestamp AS Update_Timestamp, U.Name AS Update_Name " & vbCrLf & _
	"FROM InventoryUpdate_Log AS IU " & vbCrLf & _
	"LEFT JOIN [System].Users AS U ON U.SystemID=IU.UpdateID " & vbCrLf & _
	"WHERE InventoryID=" & prepIntegerSQL(InventoryID) & " " & vbCrLf & _
	"ORDER BY InventoryID, Log_ID"

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=InventoryReport" & FiscalYear & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Inventory Update HIstory Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style type="text/css">
	table, tr, th, td {
		border-width: thin;
		border-style: solid;
		border-color: gray;
		border-collapse: collapse;
		padding: inherit;
	}
</style>
</head>
<body style="width: 100%">
<%
End If
%>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr><th colspan=""" & rs.Fields.Count & """>Archived Inventory Update Items for InventoryID=" & InventoryID & "</th></tr>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		If rs.Fields(i).Name <> "GranteeID" Then
			Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
		End If
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLf)

	If Debug = True Then
		Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
		For i = 0 To (rs.Fields.Count-1)
			If rs.Fields(i).Name <> "GranteeID" Then
				Response.Write("<th>" & rs.Fields(i).Type & "</th>")
			End If
		Next
		Response.Write(vbCrLf & "</tr>" & vbCrLf)
	End If

	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If rs.Fields(i).Name = "GranteeID" Then
				' Skip
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf Right(rs.Fields(i).Name,2) = "ID" Then
					Response.Write("<td style=""text-align: right;"">" & rs.Fields(i).value & "</td>")
			ElseIf rs.Fields(i).Type = 202 and IsNull(rs.Fields(i).value)=False Then
				Response.Write("<td style=""text-align: center"">" & formatdatetime(rs.Fields(i), vbShortDate) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name = "Grantee_Notes" Then
				Response.Write("<td style=""text-align: center"" title=""" & rs.Fields(i) & """>Note</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name = "Admin_Notes" Then
				Response.Write("<td style=""text-align: center"" title=""" & rs.Fields(i) & """>Note</td>" & vbCrLf)
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		Response.Write("</tr>" & vbCrLf)
		Counter = Counter + 1
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("<tfoot><tr><td colspan=""" & rs.Fields.count & """ style=""text-align: center;"">" & counter & " records.</td></tr></tfoot>" & vbCrLF)
Else
	Response.Write("<tr><td>No Inventory Update Log Items to show for InventeroyID=" & InventoryID & "</td></tr>" & vbCrLf)
End If
%>
</table>
<%
If ShowExcel = False Then
%>
<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

<br />

</body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->