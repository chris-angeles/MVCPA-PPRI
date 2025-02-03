<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, vrs, vsql, ShowExcel, GrantClassID, GrantClassDescription, GrantClassWhere
GrantClassDescription = Array("All", "Taskforce", "Auxiliary", "Rapid Response Strikeforce", "Catalytic Converter")
GrantClassWhere = Array("", "WHERE TaskforceGrant=1","WHERE AuxiliaryGrant=1","WHERE RapidResponseStrikeforceGrant=1", "WHERE CatalyticConverterGrant=1")

Debug = False

If Debug = True Then
	Response.Write("<pre>Request.Form.Count='" & Request.Form.Count & "'</pre>" & vbCrLf)
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

If Len(Request.Form("GrantClassID"))>0 Then
	GrantClassID=CInt(Request.Form("GrantClassID"))
ElseIf Len(Request.QueryString("GrantClassID"))>0 Then
	GrantClassID=CInt(Request.QueryString("GrantClassID"))
Else
	GrantClassID=0
End If

If Request.QueryString("ShowExcel")="1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = sql & "SELECT G.GranteeID AS [Grantee ID], GranteeName AS [Grantee Name], C.County, ORI, " & vbCrLf & _
	"	StatePayeeIDNo AS [State Payee ID No], VendorOrganizationalUnit AS [Vendor Organizational Unit], " & vbCrLf & _
	"	VendorAddress1 AS [Vendor Address 1], VendorAddress2 AS [Vendor Address 2], " & vbCrLF & _
	"	VendorCity AS [Vendor City], VendorState AS [Vendor State], VendorZIP AS [Vendor ZIP], " & vbCrLf & _
	"	TaskforceGrant AS [TF], AuxiliaryGrant AS [Aux], RapidResponseStrikeforceGrant AS [RRS], CatalyticConverterGrant AS [CC], " & vbCrLf & _
	"	Y.[Earliest Year], Y.[Latest Year] " & vbCrLf & _
	"FROM [Grantees] AS G " & vbCrLf & _
	"LEFT JOIN Lookup.Counties AS C ON C.CountyID=G.CountyID " & vbCrLf & _
	"LEFT JOIN (SELECT GranteeID, MIN(FiscalYear) AS [Earliest Year], MAX(FiscalYear) AS [Latest Year] FROM Grants.vwCombinedGrants GROUP BY GranteeID) AS Y ON Y.GranteeID=G.GranteeID " & vbCrLf
If GrantClassID>0 Then
	sql = sql & GrantClassWhere(GrantClassID)
End If
sql = sql & "ORDER BY GranteeNameSort "

If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

If ShowExcel = True and Debug = False Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=PayeeReport.xls"
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Payee Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<div class="sectiontitle" style="white-space: nowrap;">Payee Report</div>
<form name="Selection" id="Selection" method="post" ><input type="hidden" name="ShowExcel" value="0" />
Grant Class: <select name="GrantClassID" onchange="Selection.submit();">
<%
	For i = 0 To UBound(GrantClassDescription)
		Response.Write(SelectOption(i, GrantClassDescription(i), GrantClassID))
	Next
%>
</select>&nbsp;&nbsp;<a href="PayeeReport.asp?ShowExcel=1&GrantCLassID=<%=GrantClassID %>" target="_blank">Excel</a>
</form>
<br />
<table class="reporttable">
<%
End If ' End of html only code.

If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "Grantee ID" Then
				If MVCPARights = True And ShowExcel = False Then
					Response.Write("<td style=""text-align: right""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\Grantees\Grantee.asp?GranteeID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf INStr(rs.Fields(i).Name, "Year") >0 Then
				If IsNull(rs.Fields(i)) = True Then
					Response.Write("<tdX</td>")
				Else
					Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value, 0, true, false, false) & "</td>")
				End If
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type = adBoolean Then
				If rs.Fields(i).Value = True Then
					Response.Write("<td style=""text-align: center; "">X</td>")
				Else
					Response.Write("<td style=""text-align: center; ""></td>")
				End If
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%	If ShowExcel = False Then %>
<div style="width: 100%; text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%	End If %>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->