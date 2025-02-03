<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, Quarter, OrderByDescription, QuarterDescription, OrderByField, _
	ShowExcel, ShowExcluded, ShowYTD, ShowInKind, InKindTotal, GranteeID, rs2
OrderByDescription = Array("GrantID", "Grantee", "Grant Number")
QuarterDescription = Array("All", "September 1 - November 30","December 1 - February 28", "March 1 - May 31", "June 1 - August 31")
OrderByField = Array("GrantID", "REPLACE(Grantee,'City of ','')", "Grant_Number")
debug = False
InKindTotal = 0

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
If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
End If
If Len(Request.Form("Quarter"))>0 Then
	Quarter = CInt(Request.Form("Quarter"))
ElseIf Len(Request.QueryString("Quarter"))>0 Then
	Quarter = CInt(Request.QueryString("Quarter"))
Else
	Quarter = 1
End If

If Len(Request.Form("GranteeID"))>0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID"))>0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	GranteeID = 0
End If

If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "SELECT * FROM vwFundsFromDisposal " & vbCrLf
If FiscalYear>0 AND Quarter>0 And GranteeID>0 Then
	sql = sql & "WHERE Fiscal_Year=" & prepIntegerSQL(FiscalYear) & " AND Quarter=" & prepIntegerSQL(Quarter) & " " & " AND GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf
ElseIf FiscalYear>0 AND Quarter>0 Then
	sql = sql & "WHERE Fiscal_Year=" & prepIntegerSQL(FiscalYear) & " AND Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf
ElseIf FiscalYear>0 And GranteeID>0 Then
	sql = sql & "WHERE Fiscal_Year=" & prepIntegerSQL(FiscalYear) & " AND GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf
ElseIf Quarter>0 And GranteeID>0 Then
	sql = sql & "WHERE Quarter=" & prepIntegerSQL(Quarter) & " " & " AND GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf
ElseIf FiscalYear>0 Then
	sql = sql & "WHERE Fiscal_Year=" & FiscalYear & " " & vbCrLf
ElseIf Quarter>0 Then
	sql = sql & "WHERE Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf
ElseIf GranteeID>0 Then
	sql = sql & "WHERE GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf
End If
sql = sql & "ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

If ShowExcel = True and Debug = False Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=FundsFromDisposal" & FiscalYear & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Funds from Disposal of Inventory</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	Response.Write("<option value=""0""" & selected(FiscalYear, 0) & ">All</option>" & vbCrLf)
	For i = 2017 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="Quarter">Quarter:</label> <select name="Quarter" id="Quarter" onchange="Selection.submit();">
<%
	Response.Write("<option value=""0""" & selected(Quarter, 0) & ">All</option>" & vbCrLf)
	For i = 1 to 4
		Response.Write("<option value=""" & i & """" & selected(Quarter, i) & ">" & QuarterDescription(i) & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;&nbsp;
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
Set rs2 = Con.Execute(sql)
While rs2.EOF = False
	Response.Write("<option value=""" & rs2.Fields("GranteeID") & """" & selected(rs2.Fields("GranteeID"), GranteeID) & ">" & rs2.Fields("GranteeName") & "</option>" & vbCrLf)
	rs2.MoveNext()
Wend
%>
</select>&nbsp;&nbsp;
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;&nbsp;<a href="FundsFromDisposal.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&Quarter=<%=Quarter %>&GranteeID=<%=GranteeID %>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a>
</form>

<br />
<%
End If
%>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	Response.Write("<th>Quarter</th>" & vbCrLf)
	Response.Write("<th>Grantee</th>" & vbCrLf)
	For i = 0 To (rs.Fields.Count-3)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		If ShowExcel = False Then
			Response.Write("<td style=""text-align: center; ""><a href=""/ExpenditureReport/Report.asp?GrantID=" & rs.Fields("GrantID") & "&Quarter=" & rs.Fields("Quarter") & """ target=""_blank"">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</a></td>" & vbCrLf)
		Else
			Response.Write("<td style=""text-align: center; "">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</td>" & vbCrLf)
		End If
		If ShowInKind = True Then
			If IsNull(rs.Fields("In_Kind_Expenditure")) = False Then
				InKIndTotal = InKindTotal + rs.Fields("In_Kind_Expenditure")
			End If
		End If
		Response.Write("<td style=""text-align: left; "" title=""" & rs.Fields("Program_Name") & ", " & rs.Fields("Grant_Number") & """>" & rs.Fields("Grantee") & "</td>" & vbCrLf)
		For i = 0 To (rs.Fields.Count-3)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIF rs.Fields(i).Name = "InvID" Then
				Response.Write("<td style=""text-align: right"">" & rs.Fields(i).value & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf InStr(1, rs.Fields(i).Name, "date", vbTextCompare) > 0 Then
				Response.Write("<td style=""text-align: right"">" & formatDate(rs.Fields(i).value) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		'Response.Write("<td>" & rs.Fields("First Approval Date").Type & "</td>")
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
	If ShowInKind = True Then
		Response.Write("<tfoot><tr><td colspan=""" & rs.Fields.Count & """ style=""text-align: center"">Total of In-Kind Expenditures: " & formatnumber(InKindTotal) & "</td></tr></tfoot>")
	End If
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