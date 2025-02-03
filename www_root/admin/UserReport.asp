<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, OrderBy, OrderByDescription, OrderByField, IncludeDisabled, _
	Filter, FilterDescription, FilterCondition, Counter, ShowExcel 
OrderByDescription = Array("SystemID", "UserID", "LastName")
OrderByField = Array("U.SystemID", "UserID", "LastName, FirstName")
FilterDescription = Array("All", _
	"MVCPA", _
	"Non-MVCPA", _
	"Licensed Peace Officers", _
	"Developer Role", _
	"MVCPA Viewer Role", _
	"MVCPA Administrator Role", _
	"MVCPA Auditor Role", _
	"MVCPA Grant Coordinator Role", _
	"MVCPA Administrative Assistant Role", _
	"Disabled Accounts", _
	"No role in Grantees or MVCPA", _
	"No Grantee Permissions", _
	"Disabled with Grantee Permissions", _
	"Permissions in Grantees without current grant", _
	"emails with .com", _
	"emails with txdmv.gov", _
	"emails with texas.gov or tx.gov", _
	"emails with .gov", _
	"Never Logged In", _
	"No login in last two years")
FilterCondition = Array("1=1", _
	"(MVCPAAdministrator=1 OR MVCPAGrantCoordinator=1 OR MVCPAAdministrativeAssistant=1)", _
	"NOT(MVCPAAdministrator=1 OR MVCPAGrantCoordinator=1 OR MVCPAAdministrativeAssistant=1)", _
	"LicensedPeaceOfficer=1", _
	"Developer=1", _
	"MVCPAViewer=1", _
	"MVCPAAdministrator=1", _
	"MVCPAAuditor=1", _
	"MVCPAGrantCoordinator=1", _
	"MVCPAAdministrativeAssistant=1", _
	"AccountDisabled=1", _
	"U.SystemID NOT IN (SELECT SystemID FROM [System].vwUserInRole)", _
	"U.SystemID NOT IN (SELECT SystemID FROM [System].GranteePermissions)", _
	"ISNULL(AccountDisabled,0)=1 AND U.SystemID IN (SELECT SystemID FROM [System].GranteePermissions)", _
	"U.SystemID IN (SELECT SystemID FROM [System].GranteePermissions WHERE GranteeID NOT IN (SELECT GranteeID FROM [Grants].Main WHERE FiscalYear=" & Application("CurrentFiscalYear") & "))", _
	"[System].[fnGovernmentEMailDomain](UserID)=0 AND ISNULL(AccountDisabled,0)=0 AND UserID IS NOT NULL ", _
	"U.email like '%@txdmv.gov' AND ISNULL(AccountDisabled,0)=0 AND UserID IS NOT NULL ", _
	"(U.email like '%tx.gov' OR U.email like '%texas.gov') AND ISNULL(AccountDisabled,0)=0 AND UserID IS NOT NULL ", _
	"U.email like '%.gov' AND ISNULL(AccountDisabled,0)=0 AND UserID IS NOT NULL ", _
	"[Last Login] IS NULL", _
	"ISNULL([Last Login], '1/1/2000') < '" & DateAdd("yyyy",-2,date()) & "'")
debug = False
Counter = 0
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

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.Querystring("OrderBy"))>0 Then
	OrderBy = CInt(Request.Querystring("OrderBy"))
Else
	OrderBy = 0
End If

If Len(Request.Form("Filter"))>0 Then
	Filter = CInt(Request.Form("Filter"))
ElseIf Len(Request.Querystring("Filter"))>0 Then
	Filter = CInt(Request.Querystring("Filter"))
Else
	Filter=0
End If

If Request.Form("IncludeDisabled")="1" Then
	IncludeDisabled = True
ElseIf Request.Querystring("IncludeDisabled")="1" Then
	IncludeDisabled = True
Else
	IncludeDisabled = False
End If

If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

If InStr(FilterDescription(Filter),"Disabled Accounts")>0 Then
	IncludeDisabled = True
ElseIf FilterDescription(Filter) = "Disabled with Grantee Permissions" Then
	IncludeDisabled = True
End If


sql = "SELECT U.SystemID AS SID, UserID, " & vbCrLf & _
	"	AccountDisabled AS [Account Disabled], " & vbCrLf & _
	"	email, FirstName AS [First Name], MiddleName AS [Middle Name], " & vbCrLf & _
	"	LastName AS [Last Name], Title, " & vbCrLf & _
	"	LicensedPeaceOfficer AS [Licensed Peace Officer], " & vbCrLf & _
	"	Phone, Fax, Mobile AS [Mobile], [Last Login]" & vbCrLf
	If IncludeDisabled = True Then
		sql = sql & ", CASE WHEN AccountDisabled=1 THEN C.[Disabled] ELSE NULL END AS [Disabled] " & vbCrLf
	Else
		sql = sql & " " & vbCrLf
	End If
sql = sql & "FROM System.Users AS U " & vbCrLf & _
	"LEFT JOIN (SELECT SystemID, MAX(LoginTime) AS [Last Login] FROM [System].LoginLog GROUP BY SystemID) AS L ON L.SystemID=U.SystemID " & vbCrLf & _
	"LEFT JOIN (SELECT SystemID, MIN(CAST(UpdateTimestamp AS Date)) AS [Disabled] FROM System.Users WHERE [AccountDisabled]=1 GROUP BY SystemID) AS C ON C.SystemID=U.SystemID " & vbCrLf & _
	"WHERE " & FilterCondition(Filter) & " " & vbCrLf
If IncludeDisabled = False Then
	sql = sql & " AND ISNULL(AccountDisabled,0)=0 " & vbCrLf
End If
sql = sql &	"ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=UserReport" & Date() & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Grantee Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" >
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>
<label for="Filter">Show:</label><select name="Filter" id="Filter" onchange="Selection.submit();">
<%
For i = 0 to UBound(FilterDescription)
	Response.Write("<option value=""" & i & """" & Selected(Filter, i) & ">" & FilterDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;&nbsp;<%=CheckBoxFieldClick("IncludeDisabled", IncludeDisabled, "Selection.submit();") %>
<label for="IncludeDisabled">Include Disabled Accounts</label>
<a href="UserReport.asp?ShowExcel=1&Filter=<%=Filter%>&OrderBy=<%=OrderBy %>&IncludeDisabled=<%=prepBitSQL(IncludeDisabled) %>" target="_blank">Excel</a>
</form>

<br />
<%
End If
%><table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	For i = 0 To (rs.Fields.Count-1)
		If rs.Fields(i).Name = "email" Then
			' skip field
		Else
			Response.Write(vbTab & "<th style=""vertical-align: bottom"">" & Replace(rs.Fields(i).Name,"_"," ") & "</th>" & vbCrLf)
		End If
	Next
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Counter = Counter + 1
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLf)
		For i = 0 To (rs.Fields.Count-1)
			If rs.Fields(i).Name = "email" Then
				' skip field
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf rs.Fields(i).Name = "SID" Then
				If MVCPARights = True Then
					Response.Write(vbTab & "<td style=""text-align: right""><a href=""..\User\UpdateUser3.asp?SystemID=" & rs.Fields(i) & """ target=""Main"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right"">>" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write(vbTab & "<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name="Last Login" Then
				Response.Write(vbTab & "<td style=""text-align: right; white-space: nowrap;"">" & rs.Fields(i).value & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name="Disabled" Then
				Response.Write(vbTab & "<td style=""text-align: right; white-space: nowrap;"">" & rs.Fields(i).value & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write(vbTab & "<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write(vbTab & "<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name="UserID" Then
				Response.Write(vbTab & "<td style=""white-space: nowrap""><a href=""mailto:" & rs.Fields("email") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>" & vbCrLf)
			ElseIf rs.Fields(i).Type = adBoolean Then
				If rs.Fields(i).Value = True Then
					Response.Write(vbTab & "<td style=""text-align: center"">&check;</td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="Phone" Or rs.Fields(i).Name="Fax" Or rs.Fields(i).Name="Mobile" Then
				Response.Write(vbTab & "<td style=""white-space: nowrap; "">" & rs.Fields(i).value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td>" & rs.Fields(i).value & "</td>" & vbCrLf)
			End If
		Next
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("<tfoot><tr><td colspan=""" & rs.Fields.Count & """ style=""text-align: center; "">" & Counter & " users.</td></tr></tfoot>")
	Response.Write("<tbody>" & vbCrLf)
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%
If ShowExcel = False Then
%>
<br />
<div style="text-align: center; "><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->