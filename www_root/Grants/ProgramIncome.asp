<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><%
Dim Debug, LastGranteeID
Debug = False
LastGranteeID = 0

sql = "SELECT A.GranteeID, B.GrantID, B.FiscalYear, A.GranteeName, B.ProgramName, B.GrantNumber, " & vbCrLf & _
	"	C.Quarter, C.BeginningBalance, C.EarnedThisQuarter, C.ExpendedThisQuarter, C.EndingBalance " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS B ON A.GranteeID=B.GranteeID " & vbCrLf & _
	"LEFT JOIN [Grants].ProgramIncome AS C ON C.GrantID=B.GrantID " & vbCrLf & _
	"WHERE B.GrantID>0 " & vbCrLf & _
	"ORDER BY GranteeID, B.FiscalYear, C.[Quarter] "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Program Income Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">
<table>
<caption>Program Income Report</caption>
<thead>
	<tr style="background-color: PowderBlue"><th>Grantee ID</th><th colspan="8">Grantee</th></tr>
	<tr>
		<th>Grant ID</th>
		<th>Program Name</th>
		<th>Grant Number</th>
		<th>Fiscal Year</th>
		<th>Quarter</th>
		<th>Beginning Balance</th>
		<th>Earned In Quarter</th>
		<th>Expended In Quarter</th>
		<th>Ending Balance</th>
	</tr>
</thead>
<tbody>
<%
While rs.EOF = False
	If LastGranteeID <> rs.Fields("GranteeID") Then
		LastGranteeID = rs.Fields("GranteeID")
		Response.Write("<tr style=""background-color: PowderBlue;"">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & LastGranteeID & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td colspan=""8"">" & rs.Fields("GranteeName") & "</td>" & vbCrLF)
		Response.Write("</tr>" & vbCrLf)
	End If
	Response.Write("<tr>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("GrantID") & "</td>" & vbCrLF)
	Response.Write(vbTab & "<td>" & rs.Fields("ProgramName") & "</td>" & vbCrLF)
	Response.Write(vbTab & "<td style=""white-space: nowrap"">" & rs.Fields("GrantNumber") & "</td>" & vbCrLF)
	Response.Write(vbTab & "<td>" & rs.Fields("FiscalYear") & "</td>" & vbCrLF)
	If IsNull(rs.Fields("Quarter")) Then
		Response.Write(vbTab & "<td colspan=""5"" style=""text-align: center; ""> - no values recorded - </td>" & vbCrLF)
	Else
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Quarter") & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("BeginningBalance")) & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("EarnedThisQuarter")) & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("ExpendedThisQuarter")) & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("EndingBalance")) & "</td>" & vbCrLF)
	End If
	Response.Write("</tr>" & vbCrLf)

	rs.MoveNext
Wend
%>
</tbody>
</table>
<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->