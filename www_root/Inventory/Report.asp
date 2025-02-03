<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, ShowExcel, FiscalYear, GranteeID, _
	OrderBy, OrderByDescription, OrderByField, _
	ShowWhich, ShowWhichDescription, ShowWhichCondition, ShowAbbreviated, _
	Counter, ClientIP, PPRIComputer, BOY
OrderByDescription = Array("GranteeID", "Grantee Name", "InventoryID", "Asset Class", "MVCPA Cost Percent")
OrderByField = Array("GranteeID, Asset_Class, InvID", "REPLACE(Grantee_Name,'City of ',''), Asset_Class, InvID", "InvID", "Asset_Class, InvID", "MVCPA_Pct_Of_Cost")
ShowWhichDescription = Array("Show Current Inventory", "Show Current and Disposed this year", "Show Only Items with Pending Changes", "Show Only Items with Changes Submitted", "Show All Inventory", "Items Classified as Not Inventory Items")
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

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	If Month(Date()) > 9 Then
		FiscalYear = Year(Date)+1
	Else
		FiscalYear = Year(Date)
	End If
End If
If Len(Request.Form("GranteeID"))>0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID"))>0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	GranteeID = Session("GranteeID")
End If
If Len(GranteeID)=0 Then
	GranteeID=0
End If

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 0
End If

If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If
If Len(Request.Form("ShowWhich"))>0 Then
	ShowWhich = CInt(Request.Form("ShowWhich"))
ElseIf Len(Request.QueryString("ShowWhich"))>0 Then
	ShowWhich = CInt(Request.QueryString("ShowWhich"))
Else
	ShowWhich = 0
End If
If Request.Form("ShowAbbreviated") = "1" Then
	ShowAbbreviated = True
ElseIf Request.QueryString("ShowAbbreviated") = "1" Then
	ShowAbbreviated = True
ElseIf Request.Form.Count=0 And Request.Querystring.Count=0 Then
	ShowAbbreviated = True
Else
	ShowAbbreviated = False
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
	Response.AddHeader "content-disposition", "filename=InventoryReport" & FiscalYear & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Inventory Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" action="Report.asp">
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
<%
If PPRIComputer = True Then
	Response.Write("(GranteeID=" & GranteeID & ") ")
End If
%>
<label for="OrderBy">Order By:</label> <select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%></select>&nbsp;&nbsp;
<label for="ShowWhich">Show:</label> <select name="ShowWhich" id="ShowWhich" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowWhichDescription)
	Response.Write("<option value=""" & i & """" & Selected(ShowWhich, i) & ">" & ShowWhichDescription(i) & "</option>" & vbCrLf)
Next
%></select>&nbsp;&nbsp;
<input type="Checkbox" name="ShowAbbreviated" value="1" <%=Checked(ShowAbbreviated,True) %> onchange="Selection.submit();"/> Abbreviated Display&nbsp;&nbsp;
<a href="Report.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&GranteeID=<%=GranteeID %>&OrderBy=<%=OrderBy %>&ShowWhich=<%=ShowWhich%>&ShowAbbreviated=<%=prepBitRequiredSQL(ShowAbbreviated) %>" target="_blank">Excel</a></form>
<br />
<%
End If
Counter = 0
sql = "SELECT * " & vbCrLf & _
	"FROM dbo.vwInventory " & vbCrLf
If GranteeID>0 Then
	sql = sql & "WHERE GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf
Else
	sql = sql & "WHERE GranteeID IN (SELECT DISTINCT GranteeID FROM [Grants].Main WHERE AwardAmount>0) " & vbCrLf
End If
If ShowWhich = 0 Then
	sql = sql & vbTab & "AND Date_Of_Disposal IS NULL AND Not_Inventory_Item IS NULL " & vbCrLf
ElseIf ShowWhich = 1 Then
	sql = sql & vbTab & "AND (CAST(Date_Of_Disposal AS DATE) > '" & BOY & "') AND Not_Inventory_Item IS NULL " & vbCrLf
ElseIf ShowWhich = 2 Then
	sql = sql & " AND Update_Pending IS NOT NULL AND Not_Inventory_Item IS NULL "
ElseIf ShowWhich = 3 Then
	sql = sql & " AND Update_Pending='Submitted' AND Not_Inventory_Item IS NULL  "
ElseIF ShowWhich = 5 Then
	sql = sql & " AND Not_Inventory_Item IS NOT NULL "
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
		If ShowAbbreviated = True And i>7 And i<21 Then
			' Skip
		ElseIf ShowWhich = 0 And rs.Fields(i).Name = "Not_Inventory_Item" Then
			' skip, don't show this column for current inventory
		ElseIf rs.Fields(i).Name <> "GranteeID" Then
			Response.Write("<th>" & Replace(Replace(rs.Fields(i).Name,"_"," "),"MVCPA","MVCPA") & "</th>")
		End If
	Next
	If ShowExcel = False Then
		Response.Write("<th>Edit Link</th>")
	End If
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If ShowAbbreviated And i>7 And i<21 Then
				' Skip
			ElseIf rs.Fields(i).Name = "GranteeID" Then
				' Skip
			ElseIf ShowWhich = 0 And rs.Fields(i).Name = "Not_Inventory_Item" Then
				' skip, don't show this column for current inventory
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "InvID" Then
				If MVCPAAdministrator = True and ShowExcel = False Then
					Response.Write("<td style=""text-align: right;""><a href=""Edit.asp?InventoryID=" & rs.Fields(i) & """ target=""_blank"">" & rs.Fields(i).value & "</a></td>")
				Else
					Response.Write("<td style=""text-align: right;"">" & rs.Fields(i).value & "</td>")
				End If
			ElseIf rs.Fields(i).Name = "Grantee_ID" Then
				If MVCPARights = True and ShowExcel = False Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grantees\Grantee.asp?GranteeID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "Grant_ID" Then
				If MVCPARights = True And ShowExcel = False Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grants\Grant.asp?GrantID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "Update_Pending" Then
				If IsNull(rs.Fields(i).Value) = False Then
					Response.Write("<td style=""text-align: center; "">" & rs.Fields(i).Value & "</td>")
				Else
					Response.Write("<td></td>")
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Name="Reimbursement_Rate" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,4, true, false, false) & "%</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		If ShowExcel = False Then
			If IsNull(rs.Fields("Date_Of_Disposal")) = False Then
				Response.Write("<td></td>")
			ElseIf IsNull(rs.Fields("Update_Pending")) = False Then
				Response.Write("<td style=""text-align: right;""><a href=""GranteeEdit.asp?InventoryID=" & rs.Fields("InvID") & """ target=""_blank"" tilte=""Edit/View a current inventory update"">Edit</a></td>")
			Else
				Response.Write("<td style=""text-align: right;""><a href=""GranteeEdit.asp?InventoryID=" & rs.Fields("InvID") & """ target=""_blank"" title=""Begin a new inventory update"">New</a></td>")
			End If
		End If
		'Response.Write("<td>" & rs.Fields(18).Type & "</td>")
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
<%
	If MVCPARights = True and ShowExcel = False Then
		Response.Write("<div style=""text-align: right;""><a href=""Edit.asp?InventoryID=0"" target=""_blank"">Create New Inventory Item</a></div>")
	End If
%></body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->