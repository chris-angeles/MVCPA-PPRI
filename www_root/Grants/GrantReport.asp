<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, ShowExcel, FiscalYear, OrderBy, OrderByDescription, OrderByField, GrantClassID, GrantClassField, GrantClassDescription, _
	ShowMatch, ShowAllocations, ShowClosing, LastDistrict, DistrictGrouping, ColumnCount, _
	AwardTotal, CurrentYearTotal, PriorYearTotal, BSEarmarkTotal
OrderByDescription = Array("GrantID", "Grantee Name", "Grant Number", "ORI", "Award Amount", "Closeout Date", "State House District", "State Senate District")
OrderByField = Array("H.GrantID", "REPLACE(G.GranteeName,'City of ','')", "H.GrantNumber", "G.ORI", "H.AwardAmount", "CloseoutDate", "District", "District")
GrantClassField = Array(0, 1, 4)
GrantClassDescription = Array("Both", "TaskForce Grant", "Catalytic Converter Grant")

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

If Len(Request.Form("GrantClassID"))>0 Then
	GrantClassID = CInt(Request.Form("GrantClassID"))
ElseIf Len(Request.QueryString("GrantClassID"))>0 Then
	GrantClassID = CInt(Request.QueryString("GrantClassID"))
Else
	GrantClassID = 1
End If

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
End If

If OrderBy = 6 or OrderBy = 7 Then
	DistrictGrouping = True
Else
	DistrictGrouping = False
End If

If Request.Form("ShowMatch") = "1" Then
	ShowMatch = True
ElseIf Request.QueryString("ShowMatch") = "1" Then
	ShowMatch = True
Else
	ShowMatch = False
End If
If Request.Form("ShowAllocations") = "1" Then
	ShowAllocations = True
ElseIf Request.QueryString("ShowAllocations") = "1" Then
	ShowAllocations = True
Else
	ShowAllocations = False
End If
If Request.Form("ShowClosing") = "1" Then
	ShowClosing = True
ElseIf Request.QueryString("ShowClosing") = "1" Then
	ShowClosing = True
Else
	ShowClosing = False
End If
If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If
LastDistrict = 0

sql = "SELECT H.FiscalYear AS Fiscal_Year, H.GrantID AS Grant_ID, G.GranteeID AS Grantee_ID, G.ORI, " & vbCrLf & _
	"	G.GranteeName AS Grantee_Name, H.ProgramName AS Program_Name, H.GrantNumber AS Grant_Number, " & vbCrLF & _
	"	H.AwardAmount AS Award_Amount"
If DistrictGrouping = True Then
	sql = sql & ", D.District" & vbCrLf
Else
	sql = sql & ", 0 As District" & vbCrLf
End If
If ShowMatch = True Then
	sql = sql & ", H.MatchAmount AS Match_Amount, " & vbCrLf & _
		"	100.0 * H.MatchAmount / H.AwardAmount AS Cash_Match_Percent, " & vbCrLf & _
		"	H.ReimbursementRate AS Reimbursement_Rate "
End If
If ShowAllocations = True Then
	sql = sql & ", CurrentYearAllocation AS Current_Year, PriorYearAllocation AS Prior_Year, " & vbCrLf & _
		"	CASE WHEN BorderCounty=1 OR PortCounty=1 OR Port2County=1 THEN H.CurrentYearAllocation ELSE NULL END AS Border_Security_Earmark"
End If
If ShowClosing = True Then
	sql = sql & ", CONVERT(VARCHAR,I.SubmitTimestamp, 101) AS Inventory_Certification_Date, " & vbCrLf & _
		"CONVERT(VARCHAR,ReportsCompleteDate, 101) AS Reports_Complete_Date, " & vbCrLf & _
		"CONVERT(VARCHAR,ProgramGoalsDate, 101) AS Program_Goals_Date, " & vbCrLf & _
		"CONVERT(VARCHAR,DeficienciesResolvedDate, 101) AS Deficiencies_Resolved_Date, " & vbCrLf & _
		"CONVERT(VARCHAR,FundsReturnedDate, 101) AS Funds_Returned_Date, " & vbCrLF & _
		"CONVERT(VARCHAR,CloseoutDate, 101) AS Closeout_Date "
Else
	sql = sql & ", CONVERT(VARCHAR,CloseoutDate, 101) AS Closeout_Date " & vbCrLf
End If
sql = sql & vbCrLf & "FROM Grantees G " & vbCrLf & _
	"JOIN [Grants].Main AS H ON H.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN [Grants].InventoryCertification AS I ON I.GrantID=H.GrantID " & vbCrLf
If OrderBy = 6 Then
	sql = sql & "LEFT JOIN Lookup.CountyHouseDistrict AS D ON D.CountyID=G.CountyID " & vbCrLf
ElseIf OrderBy = 7 Then
	sql = sql & "LEFT JOIN Lookup.CountySenateDistrict AS D ON D.CountyID=G.CountyID " & vbCrLf
End If

sql = sql & "WHERE FiscalYear=" & FiscalYear & " " & vbCrLf
If GrantClassID>0 Then
	sql = sql & vbTab &  "AND GrantClassID=" & prepIntegerSQL(GrantClassID) & " " & vbCrLf
End If

sql = sql & "ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=GrantReport" & FiscalYear & ".xls"
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


<form name="Selection" id="Selection" method="post" action="GrantReport.asp">
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2017 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="GrantClassID">Grant Class:</label> <select name="GrantClassID" id="GrantClassID" onchange="Selection.submit();">
<%
For i = 0 to UBound(GrantClassDescription)
	Response.Write("<option value=""" & GrantClassField(i) & """" & Selected(GrantClassID, GrantClassField(i)) & ">" & GrantClassDescription(i) & "</option>" & vbCrLf)
Next
%></select>&nbsp;&nbsp;
<label for="OrderBy">Order By:</label> <select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" &i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%></select>&nbsp;&nbsp;
<input type="checkbox" name="ShowMatch" id="ShowMatch" value="1" <%
If ShowMatch=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowMatch">Show Match and Reimbursement Rate</label>&nbsp;&nbsp;
<input type="checkbox" name="ShowAllocations" id="ShowAllocations" value="1" <%
If ShowAllocations=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowAllocations">Show Allocations</label>&nbsp;&nbsp;
<input type="checkbox" name="ShowClosing" id="ShowClosing" value="1" <%
If ShowClosing=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowClosing">Show Closeout Info</label>
<a href="GrantReport.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&ShowMatch=<%
If ShowMatch=True Then 
	Response.Write("1") 
Else 
	Response.Write("0") 
End If 
%>&ShowAllocations=<%
If ShowAllocations=True Then 
	Response.Write("1") 
Else 
	Response.Write("0") 
End If 
%>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a></form>

<br />
<%
End If
%>
<table class="reporttable">
<%
AwardTotal = 0.0
CurrentYearTotal = 0.0
PriorYearTotal = 0.0
BSEarmarkTotal = 0.0

If rs.EOF = False Then
	ColumnCount = rs.Fields.Count
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	For i = 0 To (ColumnCount-1)
		If DistrictGrouping = False And rs.Fields(i).Name = "District" Then
			' Skip
		ElseIf ShowClosing = False and rs.Fields(i).Name = "Closeout_Date" Then
				' Skip the field
		Else
			Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
		End If
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		If DistrictGrouping = True And rs.Fields("District") <> LastDistrict Then
			LastDistrict = rs.Fields("District")
			If OrderBy=6 Then
				Response.Write("<tr><th colspan=""" & ColumnCount & """>State House District: " & LastDistrict & "</th></tr>" & vbCrLf)
			ElseIf OrderBy=7 Then
				Response.Write("<tr><th colspan=""" & ColumnCount & """>State Senate District: " & LastDistrict & "</th></tr>" & vbCrLf)
			End If
		End If
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLf)
		For i = 0 To (ColumnCount-1)
			If ShowClosing = False and rs.Fields(i).Name = "Closeout_Date" Then
				' Skip the field
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
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
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Name="Grant_Number"  Then
				Response.Write("<td style=""text-align: center; white-space: nowrap; "">" &rs.Fields(i).value & "</td>")
			ElseIf DistrictGrouping = False And rs.Fields(i).Name="District" Then
				' Skip District
			ElseIf rs.Fields(i).Name="Reimbursement_Rate" Or rs.Fields(i).Name="Cash_Match_Percent" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,4, true, false, false) & "%</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		If ShowAllocations = True Then
			If IsNull(rs.Fields("Award_Amount")) = False Then
				AwardTotal = AwardTotal + rs.Fields("Award_Amount")
			End If
			If IsNull(rs.Fields("Current_Year")) = False Then
				CurrentYearTotal = CurrentYearTotal + rs.Fields("Current_Year")
			End If
			If IsNull(rs.Fields("Prior_Year")) = False Then
				PriorYearTotal = PriorYearTotal + rs.Fields("Prior_Year")
			End If
			If IsNull(rs.Fields("Border_Security_Earmark")) = False Then
				BSEarmarkTotal = BSEarmarkTotal + rs.Fields("Border_Security_Earmark")
			End If
		End If
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	If ShowAllocations = True Then
		Response.Write("<tfoot>" & vbCrLf)
		Response.Write("<tr>" & vbCrLF)
		For i = 0 To (ColumnCount-1)
			If i = 0 Then
				Response.Write("<td style=""text-align: right"">Total</td>")
			ElseIf DistrictGrouping = False And rs.Fields(i).Name = "District" Then
				' Skip
			ElseIf rs.Fields(i).Name="Award_Amount" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(AwardTotal,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Name="Current_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(CurrentYearTotal,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Name="Prior_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(PriorYearTotal,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Name="Border_Security_Earmark" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(BSEarmarkTotal,2, true, true, true) & "</td>")
			Else
				Response.Write("<td></td>")
			End If
		Next
		Response.Write("</tr>" & vbCrLF)
		Response.Write("</tfoot>" & vbCrLf)
	End If

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