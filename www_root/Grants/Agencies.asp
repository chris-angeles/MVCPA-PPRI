<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, Source, SourceDescription, OrderBy, OrderByDescription, OrderByField, _
	Table1, Table2, JoinTable, JoinField
OrderByDescription = Array("ID", "Grantee Name", "ORI")
OrderByField = Array("1, O.ORI", "REPLACE(G.GranteeName,'City of ',''), 1, O.ORI", "O.ORI, 1")
SourceDescription = Array("Grant Participating Agencies", "Grant Coverage Agencies", _
	"Negotiation Participating Agencies", "Negotiation Coverage Agencies", _
	"Application Participating Agencies", "Application Coverage Agencies")
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
If FiscalYear < 2018 Then
	FiscalYear = 2018
End If
If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
End If
If Len(Request.Form("Source"))>0 Then
	Source = CInt(Request.Form("Source"))
Else 
	Source = 0
End If

If Source = 0 Then
	Table1 = "Grant"
	Table2 = "ParticipatingAgencies"
	JoinTable = "[Grants]"
	JoinField = "GrantID"
ElseIf Source = 1 Then
	Table1 = "Grant"
	Table2 = "CoverageAgencies"
	JoinTable = "[Grants]"
	JoinField = "GrantID"
ElseIf Source = 2 Then
	Table1 = "Negotiation"
	Table2 = "ParticipatingAgencies"
	JoinTable = "[Negotiation]"
	JoinField = "AppID"
ElseIf Source = 3 Then
	Table1 = "Negotiation"
	Table2 = "CoverageAgencies"
	JoinTable = "[Negotiation]"
	JoinField = "AppID"
ElseIf Source = 4 Then
	Table1 = "Application"
	Table2 = "ParticipatingAgencies"
	JoinTable = "[Application]"
	JoinField = "AppID"
ElseIf Source = 5 Then
	Table1 = "Application"
	Table2 = "CoverageAgencies"
	JoinTable = "[Application]"
	JoinField = "AppID"
End If
If Source = 0 Or Source = 1 Or Source = 2 Then
	sql = "SELECT H.GrantID, I.FiscalYear AS Fiscal_Year, G.GranteeName AS Grantee_Name, " & vbCrLf & _
		"	H.ProgramName AS Program_Name, H.GrantNumber AS Grant_Number, A.ORI, O.Agency AS [" & SourceDescription(Source) & "] " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLF & _
		"LEFT JOIN [Grants].Main AS H ON H.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN " & JoinTable & ".Main AS N ON N." & JoinField & "=H." & JoinField & " " & vbCrLf & _
		"LEFT JOIN " & JoinTable & "." & Table2 & " AS A ON A." & JoinField & "=N." & JoinField & " " & vbCrLf & _
		"LEFT JOIN Lookup.ORI AS O ON O.ORI=A.ORI " & vbCrLf & _
		"WHERE H.GrantID IS NOT NULL AND I.FiscalYear=" & FiscalYear & " " & vbCrLF & _
		"ORDER BY " & OrderByField(OrderBy)
Else
	sql = "SELECT I.AppID AS App_ID, I.FiscalYear AS Fiscal_Year, G.GranteeName AS Grantee_Name, " & vbCrLf & _
		"	N.ProgramName AS Program_Name, A.ORI, O.Agency AS [" & SourceDescription(Source) & "] " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLF & _
		"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN " & JoinTable & ".Main AS N ON N." & JoinField & "=I." & JoinField & " " & vbCrLf & _
		"LEFT JOIN " & JoinTable & "." & Table2 & " AS A ON A." & JoinField & "=N." & JoinField & " " & vbCrLf & _
		"LEFT JOIN Lookup.ORI AS O ON O.ORI=A.ORI " & vbCrLf & _
		"WHERE I.FiscalYear=" & FiscalYear & " AND O.ORI IS NOT NULL " & vbCrLF & _
		"ORDER BY " & OrderByField(OrderBy)
End If
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Grant Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" >
<label for="Source">Source for Report:</label><select name="Source" id="Source" onchange="Selection.submit();">
<%
For i = 0 to UBound(SourceDescription)
	Response.Write("<option value=""" & i & """" & Selected(Source, i) & ">" & SourceDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2018 to (Year(Date())+1)
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
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
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("<head>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "GranteeID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grantees\Grantee.asp?GranteeID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "GrantID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grants\Grant.asp?GrantID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Name="Grant_Number" Then
				Response.Write("<td style=""text-align: center; white-space: nowrap; "">" & rs.Fields(i).value & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
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