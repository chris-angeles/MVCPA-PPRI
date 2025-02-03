<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, Quarter, OrderByDescription, QuarterDescription, OrderByField, _
	ShowExcel, ShowAmounts, ShowApprovals, ShowAY, CurrentDate
OrderByDescription = Array("GrantID", "Grantee Name", "Grant Number", "ORI", "Award Amount", "Submit Date", "Approval Date")
QuarterDescription = Array("", "September 1 - November 30","December 1 - February 28", "March 1 - May 31", "June 1 - August 31")
OrderByField = Array("H.GrantID", "REPLACE(G.GranteeName,'City of ','')", "H.GrantNumber", "G.ORI", "H.AwardAmount", "I.SubmitTimestamp DESC", "I.DirectorApprovalDate DESC, I.AuditApprovalDate DESC, I.ReviewDate DESC, I.SubmitTimestamp DESC, H.GrantID")
debug = False
CurrentDate = Date()
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'</pre>" & vbCrLf)
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
ElseIf Len(Request.Querystring("OrderBy"))>0 Then
	OrderBy = CInt(Request.Querystring("OrderBy"))
End If
If Len(Request.Form("Quarter"))>0 Then
	Quarter = CInt(Request.Form("Quarter"))
ElseIf Len(Request.QueryString("Quarter"))>0 Then
	Quarter = CInt(Request.QueryString("Quarter"))
ElseIf CurrentDate < CDate("3/1/" & FiscalYear) Then
	Quarter = 1
ElseIf CurrentDate < CDate("5/1/" & FiscalYear) Then
	Quarter = 2
ElseIf CurrentDate < CDate("8/1/" & FiscalYear) Then
	Quarter = 3
Else
	Quarter = 4
End If
If Request.Form("ShowAmounts") = "1" Then
	ShowAmounts = True
ElseIf Request.QueryString("ShowAmounts") = "1" Then
	ShowAmounts = True
Else
	ShowAmounts = False
End If
If Request.Form("ShowApprovals") = "1" Then
	ShowApprovals = True
ElseIf Request.QueryString("ShowApprovals") = "1" Then
	ShowApprovals = True
Else
	ShowApprovals = False
End If
If Request.Form("ShowAY") = "1" Then
	ShowAY = True
ElseIf Request.QueryString("ShowAY") = "1" Then
	ShowAY = True
Else
	ShowAY = False
End If
If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "SELECT H.FiscalYear AS Fiscal_Year, ISNULL(I.Quarter," & prepIntegerSQL(Quarter) & ") AS Quarter, " & vbCrLF & _
	"	CAST(CASE WHEN I.Quarter IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS ERPresent, H.GrantID AS Grant_ID, G.GranteeName AS Grantee_Name, " & vbCrLf & _
	"	H.ProgramName AS Program_Name, H.GrantNumber AS Grant_Number"
If ShowAmounts = True Then
	sql = sql & "," & vbCrLf & "	I.CashExpenditureTotal AS Cash_Expenditure_Total, I.ReimbursableExpenditures AS Reimbursable_Expenditures, " & vbCrLF & _
		"	I.Reimbursement AS Reimbursement"
End If
If ShowAY = True Then
	sql = sql & ", " & vbCrLf & "	CurrentYearFunds AS AY_" & (FiscalYear MOD 100) & ", PriorYearFunds AS AY_" & ((FiscalYear - 1) MOD 100)
End If
sql = sql & ", " & vbCrLf & "	CASE WHEN N.FirstSubmit < I.SubmitTimestamp THEN CAST(N.FirstSubmit AS DATE) ELSE NULL END AS First_Submit, " & vbCrLf & _
	"	CAST(I.SubmitTimestamp AS DATE) AS Submit_Date, J.Name AS Submit_By"
If ShowApprovals = True Then
	sql = sql & ",	" & vbCrLf & "	I.ReviewDate AS Review_Date, K.Name AS Reviewed_By, " & vbCrLf & _
		"	I.AuditApprovalDate AS Audit_Date, L.Name AS Audit_By, " & vbCrLf & _
		"	I.DirectorApprovalDate AS Director_Approval_Date, M.Name AS Director_Approval_By, " & vbCrLf & _
		"	I.DatePaid AS Date_Paid "
		'"	CASE WHEN I.SubmitTimestamp IS NOT NULL AND I.DatePaid IS NOT NULL THEN DATEDIFF(d,CAST(I.SubmitTimestamp AS DATE),I.DatePaid) ELSE NULL END AS Days_To_Pay"
End If
sql = sql &	" " & vbCrLf & "FROM Grantees G " & vbCrLf & _
	"JOIN [Grants].Main AS H ON H.GranteeID=G.GranteeID " & vbCrLF & _
	"LEFT JOIN [ER].Main AS I ON I.GrantID=H.GrantID AND I.Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"LEFT JOIN [System].[Users] AS J ON J.SystemID=I.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].[Users] AS K ON K.SystemID=I.ReviewID " & vbCrLf & _
	"LEFT JOIN [System].[Users] AS L ON L.SystemID=I.AuditApprovalID " & vbCrLf & _
	"LEFT JOIN [System].[Users] AS M ON M.SystemID=I.DirectorApprovalID " & vbCrLf & _
	"LEFT JOIN (SELECT GrantID, Quarter, MIN(SubmitTimestamp) AS FirstSubmit FROM [ER].Main_Log GROUP BY GrantID, Quarter) AS N ON N.GrantID=I.GrantID and N.Quarter=I.Quarter " & vbCrLf & _
	"WHERE H.FiscalYear=" & FiscalYear & " AND H.GrantClassID=4 " & vbCrLf & _
	"ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)


If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ERStatus" & FiscalYear & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Grant Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2017 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="Quarter">Quarter:</label> <select name="Quarter" id="Quarter" onchange="Selection.submit();">
<%
	For i = 1 to 4
		Response.Write("<option value=""" & i & """" & selected(Quarter, i) & ">" & QuarterDescription(i) & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<input type="checkbox" name="ShowAmounts" id="ShowAmounts" value="1" <%
If ShowAmounts=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowAmounts">Show Amounts</label>&nbsp;&nbsp;
<input type="checkbox" name="ShowAY" id="ShowAY" value="1" <%
If ShowAY=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowAY">Show AY</label>
<input type="checkbox" name="ShowApprovals" id="ShowApprovals" value="1" <%
If ShowApprovals=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowApprovals">Show Approvals</label>&nbsp;&nbsp;
<a href="Status.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&Quarter=<%=Quarter %>&ShowAmounts=<%
If ShowAmounts=True Then 
	Response.Write("1") 
Else 
	Response.Write("0") 
End If 
%>&ShowAY=<%
If ShowAY=True Then 
	Response.Write("1") 
Else 
	Response.Write("0") 
End If 
%>&ShowApprovals=<%
If ShowApprovals=True Then 
	Response.Write("1") 
Else 
	Response.Write("0") 
End If 
%>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a>
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
	For i = 3 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("<thead>" & vbCrLf)

	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLf)
		If rs.Fields("ERPresent") Then
			Response.Write("<td style=""text-align: center; ""><a href=""..\ExpenditureReport\Report.asp?GrantID=" & rs.Fields("Grant_ID") & "&Quarter=" & rs.Fields("Quarter") & " target=""_blank"">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</a></td>" & vbCrLf)
		Else
			Response.Write("<td style=""text-align: center; "">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</td>" & vbCrLf)
		End If
		For i = 3 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "Grant_ID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grants\Grant.asp?GrantID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Name="Grant_Number" Then
				Response.Write("<td style=""text-align: left; white-space: nowrap;"">" & rs.Fields(i).value & "</td>")
			ElseIf rs.Fields(i).Name="Reimbursement_Rate" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,4, true, false, false) & "%</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf InStr(1, rs.Fields(i).Name, "date", vbTextCompare) > 0 Then
				Response.Write("<td style=""text-align: right"">" & formatDate(rs.Fields(i).value) & "</td>")
			ElseIf rs.Fields(i).Name = "First_Submit" Then
				Response.Write("<td style=""text-align: right"">" & formatDate(rs.Fields(i).value) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		'Response.Write("<td>" & rs.Fields(7).Type & "</td>")
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
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