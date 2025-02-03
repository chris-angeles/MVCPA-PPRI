<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, ShowExcel
debug = False
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

If Len(Request.Form("ShowExcel"))>0 Then
	If Request.Form("ShowExcel")="1" Then 
		ShowExcel = True
	Else
		ShowExcel = False
	End If
ElseIf Len(Request.QueryString("ShowExcel"))>0 Then
	If Request.QueryString("ShowExcel")="1" Then 
		ShowExcel = True
	Else
		ShowExcel = False
	End If
Else
	ShowExcel = False
End If

If ShowExcel = True Then
	If Debug = False Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "content-disposition", "filename=ChangedOfficials.xls"
	End If
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Changed Officials</title>
<style>
	td {
		padding-left: 6px;
		padding-right: 6px;
	}
</style>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">
<h2>Recent Grantee Official Changes</h2>
<%
End If
sql = "SELECT * " & vbCrLf & _
	"FROM vwChangedOfficials " & vbCrLf & _
	"WHERE UpdateTimestamp > '" & Month(Date()) & "/" & Day(Date()) & "/" & (Year(Date())-1) & "' " & vbCrLf & _
	"ORDER BY UpdateTimestamp DESC"
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
If ShowExcel = False Then
%>
<div style=" text-align: right; margin-right: 150px;"><a href="ChangedOfficials.asp?ShowExcel=1" target="_blank">Excel</a></div><br />
<div style="width: 100%; text-align: center; "><input type="button" value="Close" onclick="window.close();" /></div>
</body>
</html>
<%
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Nothing to report")
Else
	Response.Write("<table style=""margin: auto; padding: 8px; "">" & vbCrLf)
	Response.Write("<thead><tr>" & vbCrLf)
	Response.Write("<th>Date and Time</th>" & vbCrLf)
	Response.Write("<th>Grantee</th>" & vbCrLf)
	Response.Write("<th>Position</th>" & vbCrLf)
	Response.Write("<th>Old Official</th>" & vbCrLf)
	Response.Write("<th>New Official</th>" & vbCrLf)
	Response.Write("<th>Changed By</th>" & vbCrLf)
	Response.Write("</tr></thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr>" & vbCrLf)
		Response.Write("<td style=""text-align: right; white-space: nowrap; "">" & rs.Fields("UpdateTimestamp") & "</td>" & vbCrLf)
		Response.Write("<td style=""text-align: left; "">" & rs.Fields("Grantee_Name") & "</td>" & vbCrLf)
		Response.Write("<td style=""text-align: left; "">" & rs.Fields("Position") & "</td>" & vbCrLf)
		Response.Write("<td style=""text-align: left; "">" & rs.Fields("Original_Official") & "</td>" & vbCrLf)
		Response.Write("<td style=""text-align: left; "">" & rs.Fields("New_Official") & "</td>" & vbCrLf)
		Response.Write("<td style=""text-align: left; "">" & rs.Fields("Update_By") & "</td>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext()
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>")
End If


%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->