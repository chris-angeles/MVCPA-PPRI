<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, Quarter, OrderByDescription, QuarterDescription, OrderByField, _
	ShowCategoryDetail, ShowCategoryDetailDescription, ShowYTD, Show, ShowDescription, _
	ShowClause, StartQuarter, ShowExcel, CurrentDate
OrderByDescription = Array("GrantID", "Grantee Name", "Grant Number", "ORI", "Grant Award Amount")
QuarterDescription = Array("", "September 1 - November 30","December 1 - February 28", "March 1 - May 31", "June 1 - August 31")
OrderByField = Array("Grant_ID", "REPLACE(Grantee_Name,'City of ','')", "Grant_Number", "A.ORI", "Grant_Award_Amount")
ShowDescription = Array ("All", "Border", "Port", "Port 2", "Border and Port", "Border, Port, and Port 2")
ShowClause = Array ("1=1", "A.BorderCounty=1", "A.PortCounty=1", "A.Port2County=1", "(A.BorderCounty=1 OR A.PortCounty=1)", "(A.BorderCounty=1 OR A.PortCounty=1 OR A.Port2County=1)")
ShowCategoryDetailDescription = Array("Do Not Show Category Details","Show Total and Excluded", "Show Net")
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
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
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
If Len(Request.Form("ShowCategoryDetail")) > 0 Then
	ShowCategoryDetail = CInt(Request.Form("ShowCategoryDetail"))
ElseIf Len(Request.QueryString("ShowCategoryDetail")) > 0 Then
	ShowCategoryDetail = CInt(Request.QueryString("ShowCategoryDetail"))
Else
	ShowCategoryDetail = 0
End If
'If Request.Form("ShowYTD") = "1" Then
	ShowYTD = True
	StartQuarter=1
'ElseIf Request.QueryString("ShowYTD") = "1" Then
'	ShowYTD = True
'	StartQuarter=1
'Else
'	StartQuarter = Quarter
'	ShowYTD = False
'End If
If Len(Request.Form("Show")) > 0 Then
	Show = CInt(Request.Form("Show"))
ElseIf Len(Request.QueryString("Show")) > 0 Then
	Show = CInt(Request.QueryString("Show"))
Else
	Show = 0
End If
If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "SELECT Grantee_ID,  Grant_ID,  Fiscal_Year,  Quarter, Grantee_Name, Program_Name, Grant_Award_Amount,  Grant_Number, Reimbursement_Rate, " & vbCrLf
If ShowCategoryDetail = 1 Then
	sql = sql &	"	Personnel,  Fringe,  Overtime,  Professional_And_Contract_Services,  Travel,  Equipment,  Supplies_And_DOE, " & vbCrLf & _
		"	Personnel_Excluded,  Fringe_Excluded,  Overtime_Excluded,  Professional_And_Contract_Services_Excluded,  Travel_Excluded,  Equipment_Excluded,  Supplies_And_DOE_Excluded, " & vbCrLf
ElseIf ShowCategoryDetail = 2 Then
	sql = sql &	"	Net_Personnel,  Net_Fringe,  Net_Overtime,  Net_Professional_And_Contract_Services,  Net_Travel,  Net_Equipment,  Net_Supplies_And_DOE, " & vbCrLf
End If
sql = sql & "	Total_Expenditures,  Total_Excluded,  [Total_Expenditures_(Less_Excluded)], " & vbCrLf & _
	"	In_Lieu_Of_DPS,  In_Lieu_Of_NICB,  Program_Income_Used,  Unbudgeted_Program_Income, " & vbCrLf & _
	"	Reimbursable_Expenditures, /*Excluded_Over_Budget,*/ Calculate_Reimbursable, Over_Budget_Exclusion, " & vbCrLf & _
	"	Reimbursements, Prior_Year_Funds, Current_Year_Funds " & vbCrLf & _
	"	FROM ER.fnSummary(" & prepIntegerSQL(FiscalYear) & ", " & prepIntegerSQL(Quarter) & ") AS A " & vbCrLf & _
	"	WHERE " & ShowClause(Show) & " " & vbCrLf & _
	"	ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)


If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=MatchReport" & FiscalYear & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Match Report</title>
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
<label for="Show">Detail:</label> <select name="ShowCategoryDetail" id="ShowCategoryDetail" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowCategoryDetailDescription)
	Response.Write("<option value=""" & i & """" & Selected(ShowCategoryDetail, i) & ">" & ShowCategoryDetailDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;<!--<input type="checkbox" name="ShowYTD" id="ShowYTD" value="1" <%
If ShowYTD=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowYTD">YTD</label>&nbsp;&nbsp;-->
<label for="Show">Show:</label> <select name="Show" id="Show" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowDescription)
	Response.Write("<option value=""" & i & """" & Selected(Show, i) & ">" & ShowDescription(i) & "</option>" & vbCrLf)
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
<a href="MatchReport2.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&Quarter=<%=Quarter %>&ShowCategoryDetail=<%=prepBitSQL(ShowCategoryDetail) %>&OrderBy=<%=OrderBy %>&ShowYTD=<%=prepBitSQL(ShowYTD) %>" target="_blank">Excel</a>
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
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("<thead>" & vbCrLf)

	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLf)
		For i = 0 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "Grant_ID" Then
				Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf (InStr(rs.Fields(i).Name,"Rate")>0 OR InStr(rs.Fields(i).Name,"Percent")>0) and rs.Fields(i).Name<>"Reimbursement_Rate_Cash_Match" Then
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