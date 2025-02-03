<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j

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

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Activate/Deactivate Grantees</title>
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
<h2>List of Grantees and Active / Inactive Status</h2>
<!--Pageload time=<%=Now() %>-->
<br />
<table style="margin: auto; ">
<thead>
	<tr style="vertical-align: bottom; ">
		<th>Grantee<br />ID</th>
		<th>Grantee<br />Name</th>
		<th>ORI</th>
		<th>TFG</th>
		<th>MAG</th>
		<th>CC</th>
		<th>RRS</th>
		<th>Active /<br />Inactive</th>
	</tr>
</thead>
<%

sql = "SELECT GranteeID, GranteeName, ORI, ISNULL(Inactive, 0) AS Inactive, " & vbCrLf & _
	"	CASE WHEN [TaskforceGrant]=1 THEN 'X' ELSE '' END AS TFG, " & vbCrLf & _ 
	"	CASE WHEN [AuxiliaryGrant]=1 THEN 'X' ELSE '' END AS MAG,  " & vbCrLf & _
	"	CASE WHEN [CatalyticConverterGrant]=1 THEN 'X' ELSE '' END AS CC, " & vbCrLf & _ 
	"	CASE WHEN [RapidResponseStrikeforceGrant]=1 THEN 'X' ELSE '' END AS RRS " & vbCrLf & _
	"FROM Grantees " & vbCrLf & _
	"ORDER BY GranteeNameSort"
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

While rs.EOF = False
	Response.Write("<tr>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("GranteeID") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("GranteeName") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("ORI") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("TFG") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("MAG") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("CC") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("RRS") & "</td>" & vbCrLf)
	If rs.Fields("Inactive") = True Then
		Response.Write(vbTab & "<td style=""text-align: center; ""><a href=""#"" onclick=""window.open('ChangeActivation.asp?GranteeID=" & rs.Fields("GranteeID") & "&Inactive=0', '_blank','width=400,height=200')"">Inactive</a></td>" & vbCrLf)
	Else
		Response.Write(vbTab & "<td style=""text-align: center; ""><a href=""#"" onclick=""window.open('ChangeActivation.asp?GranteeID=" & rs.Fields("GranteeID") & "&Inactive=1', '_blank','width=400,height=200')"">Active</a></td>" & vbCrLf)
	End If
	Response.Write("</tr>" & vbCrLf)
	rs.MoveNext
Wend
%>
</table>
</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->