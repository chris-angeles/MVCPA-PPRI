<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, ShowExcel, GranteeID, _
	OrderBy, OrderByDescription, OrderByField, _
	Counter, ClientIP, PPRIComputer, BOY
OrderByDescription = Array("GranteeID", "Grantee Name", "InventoryID", "Asset Class")
OrderByField = Array("GranteeID, AssetClass, InventoryID", "REPLACE(GranteeName,'City of ',''), Asset_Class, InvID", "InvID", "Asset_Class, InvID")
debug = False
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

If Len(Request.Form("GranteeID"))>0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID"))>0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	GranteeID = 0 
End If
If Len(GranteeID)=0 Then
	GranteeID=0
End If

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
End If
If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If
ClientIP = Request.ServerVariables("REMOTE_ADDR")
If ClientIP = "127.0.0.1" Then
	PPRIComputer = True
ElseIf Left(ClientIP, 11) = "128.194.68." Then
	PPRIComputer = True
Else
	PPRIComputer = False
End If

If Month(Date())<9 Then
	BOY = CDate("9/1/" & (Year(Date())-1))
Else
	BOY = CDate("9/1/" & (Year(Date())))
End If

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=InventoryUpdateReport.xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Inventory Updates Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" action="UpdateReport.asp">
<label for="GranteeID">Grantee:</label> <select name="GranteeID" id="GranteeID" onchange="Selection.submit();">
	
<%
If MVCPARights = True Then
	Response.Write(vbTab & "<option value=""0"">All grantees</option>" & vbCrLf)
	sql = "SELECT GranteeID, GranteeName " & vbCrLf & _
		"FROM Grantees " & vbCrLf & _
		"WHERE GranteeID IN (SELECT DISTINCT GranteeID FROM [Grants].Main WHERE AwardAmount>0) " & vbCrLf & _
		"ORDER BY REPLACE(GranteeName,'City of ','') "
Else
	sql = "SELECT GranteeID, GranteeName " & vbCrLf & _
		"FROM Grantees " & vbCrLf & _
		"WHERE GranteeID IN (SELECT DISTINCT GranteeID FROM [System].[GranteePermissions] WHERE SystemID=" & prepIntegerSQL(UserSystemID) & ") " & vbCrLf & _
		"ORDER BY REPLACE(GranteeName,'City of ','') "
End If
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("GranteeID") & """" & selected(rs.Fields("GranteeID"), GranteeID) & ">" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
	rs.MoveNext()
Wend
%>
</select>&nbsp;&nbsp;
<label for="OrderBy">Order By:</label> <select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%></select>&nbsp;&nbsp;
<a href="UpdateReport.asp?ShowExcel=1&GranteeID=<%=GranteeID %>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a></form>
<br />
<%
End If
Counter = 0
sql = "SELECT InventoryID AS InvID, GranteeName AS Grantee, AssetClassID + ' ' + AssetClass AS Asset_Class,  " & vbCrLf & _
	"	ItemDescription AS Description, SubmitName AS Submit_By, CONVERT(VARCHAR,SubmitTimestamp,1) AS Submit_Date, " & vbCrLF & _
	"	Disposal, NotInventoryItem AS Not_An_Inventory_Item, " & vbCrLf & _
	"	CONVERT(VARCHAR,DateOfDisposal,1) AS Date_of_Disposal, SalePrice AS Sale_Price, " & vbCrLf & _
	"	FirstApprovalName AS First_Approver, " & vbCrLf & _
	"	CONVERT(VARCHAR,FirstApprovalDate,1) AS First_Approval_Date, " & vbCrLf & _
	"	SecondApprovalName as Second_Approver, " & vbCrLf & _
	"	CONVERT(VARCHAR,SecondApprovalDate,1) AS Second_Approval_Date, PhaseII, " & vbCrLf & _
	"	AdministrativeNotes AS Administrative_Notes" & vbCrLf & _
	"FROM dbo.vwInventoryUpdate " & vbCrLf & _
	"WHERE UpdateRecord=1 " & vbCrLf
If GranteeID>0 Then
	sql = sql & vbTab & "AND GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf
End If
sql = sql & "ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

%>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		If rs.Fields(i).Name <> "GranteeID" Then
			Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
		End If
	Next
	'If ShowExcel = False Then
	'	Response.Write("<th>Edit Link</th>")
	'End If
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If rs.Fields(i).Name = "GranteeID" Then
				' Skip
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "InvID" Then
				If ShowExcel = False Then
					Response.Write("<td style=""text-align: right;""><a href=""GranteeEdit.asp?InventoryID=" & rs.Fields(i) & """ target=""_blank"">" & rs.Fields(i).value & "</a></td>")
				Else
					Response.Write("<td style=""text-align: right;"">" & rs.Fields(i).value & "</td>")
				End If
			ElseIf InStr(rs.Fields(i).Name,"Date")>0 Then
				Response.Write("<td style=""text-align: center"">" & rs.Fields(i) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name = "Administrative_Notes" Then
				Response.Write("<td style=""text-align: center"" title=""" & rs.Fields(i) & """>Note</td>" & vbCrLf)
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		'If ShowExcel = False Then
		'	If rs.Fields("Update_Pending") = True Then
		'		Response.Write("<td style=""text-align: right;""><a href=""GranteeEdit.asp?InventoryID=" & rs.Fields("InvID") & """ target=""_blank"">Edit</a> pending</td>")
		'	Else
		'		Response.Write("<td style=""text-align: right;""><a href=""GranteeEdit.asp?InventoryID=" & rs.Fields("InvID") & """ target=""_blank"">Edit</a></td>")
		'	End If
		'End If
		Response.Write("</tr>" & vbCrLf)
		Counter = Counter + 1
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("<tfoot><tr><td colspan=""" & rs.Fields.count & """ style=""text-align: center;"">" & counter & " records.</td></tr></tfoot>" & vbCrLF)
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%
If ShowExcel = False Then
%>
<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->