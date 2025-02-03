<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, Quarter, OrderByDescription, QuarterDescription, OrderByField, _
	CurrentDate, YEVersion
OrderByDescription = Array("GrantID", "Grantee Name", "Grant Number", "ORI", "Award Amount", "Submit Date", "Approval Date")
QuarterDescription = Array("", "September 1 - November 30","December 1 - February 28", "March 1 - May 31", "June 1 - August 31", "Year-End")
OrderByField = Array("H.GrantID", "REPLACE(G.GranteeName,'City of ','')", "H.GrantNumber", "G.ORI", "H.AwardAmount", "I.SubmitTimestamp DESC", "I.ApprovalDate DESC, I.SubmitTimestamp DESC")
debug = False
CurrentDate=Date()
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
ElseIf CurrentDate < CDate("3/1/" & FiscalYear) Then
	Quarter = 1
ElseIf CurrentDate < CDate("5/1/" & FiscalYear) Then
	Quarter = 2
ElseIf CurrentDate < CDate("8/1/" & FiscalYear) Then
	Quarter = 3
Else
	Quarter = 4
End If

If FiscalYear > 2023 Then 
	YEVersion = 2
Else
	YEVersion = 1
End If

If Quarter < 5 Then
	sql = "SELECT H.FiscalYear AS Fiscal_Year, ISNULL(I.Quarter," & prepIntegerSQL(Quarter) & ") AS Quarter, H.GrantID, " & vbCrLf & _
		"	G.GranteeName AS Grantee_Name, H.ProgramName, H.GrantNumber AS Grant_Number, " & vbCrLf & _
		"	ISNULL(J.Questions,0) AS Questions, J.Answered AS Responses, " & vbCrLf & _
		"	CONVERT(VARCHAR, I.SubmitTimestamp, 101) AS Submit_Date, I.ApprovalDate AS Approval_Date " & vbCrLf & _
		"FROM Grantees G " & vbCrLf & _
		"JOIN [Grants].Main AS H ON H.GranteeID=G.GranteeID " & vbCrLF & _
		"LEFT JOIN [PR].Main AS I ON I.GrantID=H.GrantID AND I.Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
		"LEFT JOIN ( " & vbCrLf
	If Quarter = 1 Then
		sql = sql & "	SELECT A.GrantID, COUNT(*) AS Questions, " & vbCrLf & _
			"		SUM(CASE WHEN IntegerResponse_Sep IS NOT NULL OR IntegerResponse_Oct IS NOT NULL OR IntegerResponse_Nov IS NOT NULL OR " & vbCrLf & _
			"		DecimalResponse_Sep IS NOT NULL OR DecimalResponse_Oct IS NOT NULL OR DecimalResponse_Nov IS NOT NULL OR " & vbCrLf & _
			"		TextResponse_Q1 IS NOT NULL THEN 1 ELSE 0 END) AS Answered " & vbCrLf & _
			"	FROM PR.GrantQuestions AS A " & vbCrLf & _
			"	LEFT JOIN PR.Responses AS B ON A.GrantID=B.GrantID AND A.QuestionID=B.QuestionID " & vbCrLf & _
			"	GROUP BY A.GrantID "
	ElseIf Quarter = 2 Then
		sql = sql & "	SELECT A.GrantID, COUNT(*) AS Questions, " & vbCrLf & _
			"		SUM(CASE WHEN IntegerResponse_Dec IS NOT NULL OR IntegerResponse_Jan IS NOT NULL OR IntegerResponse_Feb IS NOT NULL OR " & vbCrLf & _
			"		DecimalResponse_Dec IS NOT NULL OR DecimalResponse_Jan IS NOT NULL OR DecimalResponse_Feb IS NOT NULL OR " & vbCrLf & _
			"		TextResponse_Q2 IS NOT NULL THEN 1 ELSE 0 END) AS Answered " & vbCrLf & _
			"	FROM PR.GrantQuestions AS A " & vbCrLf & _
			"	LEFT JOIN PR.Responses AS B ON A.GrantID=B.GrantID AND A.QuestionID=B.QuestionID " & vbCrLf & _
			"	GROUP BY A.GrantID "
	ElseIf Quarter = 3 Then
		sql = sql & "	SELECT A.GrantID, COUNT(*) AS Questions, " & vbCrLf & _
			"		SUM(CASE WHEN IntegerResponse_Mar IS NOT NULL OR IntegerResponse_Apr IS NOT NULL OR IntegerResponse_May IS NOT NULL OR " & vbCrLf & _
			"		DecimalResponse_Mar IS NOT NULL OR DecimalResponse_Apr IS NOT NULL OR DecimalResponse_May IS NOT NULL OR " & vbCrLf & _
			"		TextResponse_Q3 IS NOT NULL THEN 1 ELSE 0 END) AS Answered " & vbCrLf & _
			"	FROM PR.GrantQuestions AS A " & vbCrLf & _
			"	LEFT JOIN PR.Responses AS B ON A.GrantID=B.GrantID AND A.QuestionID=B.QuestionID " & vbCrLf & _
			"	GROUP BY A.GrantID "
	ElseIf Quarter = 4 Then
		sql = sql & "	SELECT A.GrantID, COUNT(*) AS Questions, " & vbCrLf & _
			"		SUM(CASE WHEN IntegerResponse_Jun IS NOT NULL OR IntegerResponse_Jul IS NOT NULL OR IntegerResponse_Aug IS NOT NULL OR " & vbCrLf & _
			"		DecimalResponse_Jun IS NOT NULL OR DecimalResponse_Jul IS NOT NULL OR DecimalResponse_Aug IS NOT NULL OR " & vbCrLf & _
			"		TextResponse_Q4 IS NOT NULL THEN 1 ELSE 0 END) AS Answered " & vbCrLf & _
			"	FROM PR.GrantQuestions AS A " & vbCrLf & _
			"	LEFT JOIN PR.Responses AS B ON A.GrantID=B.GrantID AND A.QuestionID=B.QuestionID " & vbCrLf & _
			"	GROUP BY A.GrantID "
	End If
	sql = sql &	") AS J ON J.GrantID=H.GrantID " & vbCrLf & _
		"WHERE FiscalYear=" & FiscalYear & " AND H.GrantClassID=4 " & vbCrLf
ElseIf Quarter=5 Then
	OrderByField(6)="ApprovalTimestamp"
	sql = "SELECT H.FiscalYear AS Fiscal_Year, 'YE' AS Quarter, H.GrantID, " & vbCrLf & _
	"	G.GranteeName AS Grantee_Name, H.ProgramName, H.GrantNumber AS Grant_Number, " & vbCrLf & _
	"	Questions = (SELECT COUNT(*) FROM YE.Questions WHERE Version=" & YEVersion & "), J.Responses AS Responses, " & vbCrLf & _
	"	CAST(I.SubmitTimestamp AS DATE) AS Submit_Date, CAST(I.ApprovalTimestamp AS DATE) AS Approval_Date, " & vbCrLf & _
	"	CAST(K.SubmitTimestamp AS DATE) AS Inventory_Certification_Date, " & vbCrLf & _
	"	CAST(K.AcceptanceDate AS DATE) AS MVCPA_Acceptance_Date " & vbCrLf & _
	"FROM Grantees G " & vbCrLf & _
	"JOIN [Grants].Main AS H ON H.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN [YE].Main AS I ON I.GrantID=H.GrantID " & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT GrantID, SUM(CASE WHEN LEN(Response)>0 THEN 1 ELSE 0 END) AS Responses " & vbCrLf & _
	"	FROM YE.Responses " & vbCrLf & _
	"	GROUP BY GrantID " & vbCrLf & _
	") AS J ON J.GrantID=H.GrantID " & vbCrLf & _
	"LEFT JOIN [Grants].InventoryCertification AS K ON K.GrantID=H.GrantID " & vbCrLf & _
	"WHERE FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf
End If
sql = sql & "ORDER BY " & OrderByField(OrderBy)

If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If


Set rs=Con.Execute(sql)

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title><%=FiscalYear%> Catalytic Converter Progress Report Status Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">
<div class="sectiontitle" style="white-space: nowrap;"><%=FiscalYear%> Catalytic Converter Progress Report Status Report</div>
<div>
<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2017 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="Quarter">Quarter:</label> <select name="Quarter" id="Quarter" onchange="Selection.submit();">
<%
	For i = 1 to 4
		Response.Write("<option value=""" & i & """" & selected(Quarter, i) & ">" & QuarterDescription(i) & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>
</form>
</div>

<br />
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<head>" & vbCrLf)
	Response.Write("<tr>" & vbCrLF)
	Response.Write("<th>Quarter</th>" & vbCrLf)
	For i = 2 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	If Debug = True Then
	For i = 1 To (rs.Fields.Count-1)
		Response.Write("<th>" & rs.Fields(i).Type & "</th>")
	Next
	End If
	Response.Write("<head>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr>" & vbCrLf)
		If Quarter<5 Then
			Response.Write("<td style=""text-align: center; ""><a href=""/ProgressReport/Report.asp?GrantID=" & rs.Fields("GrantID") & "&Quarter=" & rs.Fields("Quarter") & """ target=""_blank"">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</a></td>" & vbCrLf)
		ElseIf Quarter=5 Then
			Response.Write("<td style=""text-align: center; ""><a href=""/ProgressReport/YearEnd.asp?GrantID=" & rs.Fields("GrantID") & """ target=""_blank"">" & rs.Fields("Fiscal_Year") & "/YE</a></td>" & vbCrLf)
		End If
		For i = 2 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "GrantID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grants\Grant.asp?GrantID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Name="Reimbursement_Rate" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,4, true, false, false) & "%</td>")
			ElseIf rs.Fields(i).Type = 202 Then ' Date
				Response.Write("<td style=""text-align: right"">" & formatdatetime(rs.Fields(i).value,vbShortDate) & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Name="Grant_Number" Then
				Response.Write("<td style=""white-space: nowrap; "">" & rs.Fields(i).value & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
Else
	Response.WRite("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>

<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->