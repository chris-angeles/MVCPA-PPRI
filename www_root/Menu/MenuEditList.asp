<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j
debug = false
If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLf)
	For each i in Request.Form
		Response.Write("Request.Form(""" & i & """)='" & Request.Form(i) & "'" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("Session(""" & i & """)='" & Session(i) & "'" & vbCrLf)
	Next
	for each i in Request.Cookies
		if Request.Cookies(i).HasKeys then
			for each j in Request.Cookies(x)
				response.write("Cookies(" & i & ":" & j & ")=" & Request.Cookies(i)(j))
			next
		else
			Response.Write("Cookies(""" & i & """)=" & Request.Cookies(i) & "<br>")
		end if
	next
	Response.Write("</pre>" & vbCrLf)
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Select Menu Item To Edit</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="margin: 0px; width: 100%">

<div class="sectiontitle">Select Menu Item To Edit</div>

<%
Dim Category, ColumnCount
Category = ""

sql = "SELECT C.Category, I.ItemID, I.MenuText AS [Menu Text], I.Page, I.Directory, " & vbCrLf & _
	"	I.MenuDescription AS [Menu Description], " & vbCrLf & _
	"	I.StartFiscalYear AS [Start FY], I.EndFiscalYear AS [End FY], " & vbCrLf & _
	"	CASE WHEN TaskforceGrantee=1 THEN 'X' ELSE '' END AS [TF], " & vbCrLf & _
	"	CASE WHEN I.AuxiliaryGrantee=1 THEN 'X' ELSE '' END AS [MAG], " & vbCrLf & _
	"	CASE WHEN I.CCGrantee=1 THEN 'X' ELSE '' END AS [CC], " & vbCrLf & _
	"	P.PermissionLevelDescription AS [Permission Level], " & vbCrLf & _
	"	I.ItemSort AS Sort, " & vbCrLf & _
	"	L.LinkID, L.LinkDescription AS [Link Description], " & vbCrLf & _
	"	CASE WHEN Inactive=1 THEN 'X' ELSE '' END AS [Inactive], " & vbCrLf & _
	"	CASE WHEN I.CategoryAndLink=1 THEN 'X' ELSE '' END AS [Cat. And Link], " & vbCrLf & _
	"	CASE WHEN I.NewWindow=1 THEN 'X' ELSE '' END AS [New Window], " & vbCrLf & _
	"	CASE WHEN I.GranteeRequired=1 THEN 'X' ELSE '' END AS [Grantee Req], " & vbCrLf & _
	"	CASE WHEN I.GrantRequired=1 THEN 'X' ELSE '' END AS [Grant Req], " & vbCrLf & _
	"	CASE WHEN I.ISARequired=1 THEN 'X' ELSE '' END AS [ISA Req], " & vbCrLf & _
	"	CASE WHEN I.RRSRequired=1 THEN 'X' ELSE '' END AS [RRS Req], " & vbCrLf & _
	"	CASE WHEN I.CCRequired=1 THEN 'X' ELSE '' END AS [CC Req], " & vbCrLf & _
	"	CASE WHEN I.GranteeLink=1 THEN 'X' ELSE '' END AS [Grantee Link], " & vbCrLf & _
	"	CASE WHEN I.GrantLink=1 THEN 'X' ELSE '' END AS [Grant Link], " & vbCrLf & _
	"	CASE WHEN I.ISALink=1 THEN 'X' ELSE '' END AS [ISA Link], " & vbCrLf & _
	"	CASE WHEN I.AppLink=1 THEN 'X' ELSE '' END AS [App Link], " & vbCrLf & _
	"	CASE WHEN I.RRSLink=1 THEN 'X' ELSE '' END AS [RRS Link], " & vbCrLf & _
	"	CASE WHEN I.CCLink=1 THEN 'X' ELSE '' END AS [CC Link], " & vbCrLf & _
	"	CASE WHEN I.AppRequired=1 THEN 'X' ELSE '' END AS [App Req], " & vbCrLf & _
	"	CASE WHEN I.NegotiationRequired=1 THEN 'X' ELSE '' END AS [Neg Req], " & vbCrLf & _
	"	CASE WHEN I.MAGRequired=1 THEN 'X' ELSE '' END AS [MAG Req], " & vbCrLf & _
	"	CASE WHEN I.SendGranteeID=1 THEN 'X' ELSE '' END AS [Send Grantee ID], " & vbCrLf & _
	"	CASE WHEN I.SendGrantID=1 THEN 'X' ELSE '' END AS [Send Grant ID], " & vbCrLf & _
	"	CASE WHEN I.SendISAID=1 THEN 'X' ELSE '' END AS [Send ISA ID], " & vbCrLf & _
	"	CASE WHEN I.SendAppID=1 THEN 'X' ELSE '' END AS [Send Appl ID], " & vbCrLf & _
	"	CASE WHEN I.SendMAGID=1 THEN 'X' ELSE '' END AS [Send MAG ID], " & vbCrLf & _
	"	CASE WHEN I.SendFiscalYear=1 THEN 'X' ELSE '' END AS [Send Fiscal Year] " & vbCrLf & _
	"FROM Menu.Categories AS C " & vbCrLf & _
	"LEFT JOIN Menu.Items AS I ON I.CategoryID=C.CategoryID " & _
	"LEFT JOIN Menu.PermissionLevels AS P ON P.PermissionLevelID=I.PermissionLevelID " & vbCrLf & _
	"LEFT JOIN Menu.Links AS L ON L.LinkID=I.LinkID " & vbCrLf & _
	"ORDER BY C.CategorySort ASC, I.ItemSort ASC "
Set rs=Con.Execute(sql)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Response.Write("<table class=""reporttable"" style=""width: 100%"">" & vbCrLf)

Category = ""
If rs.EOF = False Then
	ColumnCount = rs.Fields.Count-1
	Response.Write("<head>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom"">" & vbCrLf)
	For i = 2 To ColumnCount
		If rs.Fields(i).Name = "Directory" Then
			' skip
		ElseIf rs.Fields(i).Name = "Menu Description" Then
			' skip
		Else
			Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>" & vbCrLf)
		End If
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLf)
	Response.Write("<head>" & vbCrLf)

	While rs.EOF = False
		If Category<>rs.Fields("Category") Then
			Category = rs.Fields("Category")
			Response.Write("<tr><td colspan=""" & ColumnCount & """ style=""text-aligh: left; font-weight: bold"">" & Category & "</td></tr>" & vbCrLf)
		End If
		Response.Write("<tr>" & vbCrLf)
		For i = 2 To ColumnCount
			If InStr(rs.Fields(i).Name, "Menu Description")>0 Then
				' skip field
			ElseIf InStr(rs.Fields(i).Name, "Directory")>0 Then
				' skip field
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>" & vbCrLf)
			ElseIf InStr(rs.Fields(i).Name, "FY")>0 Then
				Response.Write("<td>" & rs.Fields(i).value & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name = "Menu Text" Then
				If rs.Fields("Inactive") = True Then
					Response.Write("<td style=""text-align: left; white-space: nowrap; text-decoration: line-through; ""><a href=""MenuEdit.asp?ItemID=" & rs.Fields("ItemID") & """ class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: left; white-space: nowrap; ""><a href=""MenuEdit.asp?ItemID=" & rs.Fields("ItemID") & """ class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "Page" Then
				If rs.Fields("Inactive") = True Then
					Response.Write("<td style=""text-align: left; white-space: nowrap; text-decoration: line-through;"" title=""" & rs.Fields("Menu Description") & """>" & rs.Fields("Directory") & rs.Fields(i) & "</td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: left; white-space: nowrap; "" title=""" & rs.Fields("Menu Description") & """>" & rs.Fields("Directory") & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "Permission Level" Then
				If rs.Fields("Inactive") = True Then
					Response.Write("<td style=""text-align: left; white-space: nowrap; text-decoration: line-through; "">" & rs.Fields(i) & "</td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: left; white-space: nowrap; "">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name = "Page" Then
				Response.Write("<td title=""" & rs.Fields(i).Name & """>" & rs.Fields(i).value & "</td>" & vbCrLf)
			Else
				Response.Write("<td title=""" & rs.Fields(i).Name & """ style=""text-align: center"">" & rs.Fields(i).value & "</td>" & vbCrLf)
			End If
		Next
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("<tr><td colspan=""" & ColumnCount & """><a href=""MenuEdit.asp?ItemID=0"">Create New Item</td></tr>" & vbCrLf)
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
Response.WRite("</table>" & vbCrLf)
%>

<div style="text-align: center; width: 100%; "><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->