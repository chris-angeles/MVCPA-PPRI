<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i
debug = false
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Login Log</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<table class="reporttable">
<%
sql = "SELECT top 100 L.LoginTime AS [Login Time], L.LogoutTime AS Logout_Time, L.SystemID AS System_ID, U.UserID, " & vbCrLF & _
	"	U.Name, L.ipaddress AS IP_Address, L.ScreenSize AS Screen_Size, L.AvailableSize AS Available_Size, L.appName AS [Browser], SessionRecovery AS Session_Recovery " & vbCrLf & _
	"FROM System.LoginLog AS L " & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=L.SystemID " & vbCrLf & _
	"ORDER BY id desc"
Set rs=Con.Execute(sql)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If
If rs.EOF = False Then
	Response.Write("<head>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write("</tr>" & vbCrLF)
	Response.Write("<head>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr>" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If rs.Fields(i).Name = "Browser" Then
				Response.Write("<td title=""" & rs.Fields(i).value & """  style=""text-align: center;"">app</td>")
			ElseIf rs.Fields(i).Name="System_ID" Then 
				Response.Write("<td style=""text-align: right; white-space: nowrap; "">" & rs.Fields(i).value & "</td>")
			ElseIf rs.Fields(i).Name="IP_Address" Or rs.Fields(i).Name="Screen_Size" Or rs.Fields(i).Name="Available_Size" Or rs.Fields(i).Name="Session_Recovery" Then 
				Response.Write("<td style=""text-align: center; white-space: nowrap; "">" & rs.Fields(i).value & "</td>")
			Else
				Response.Write("<td style=""white-space: nowrap; "">" & rs.Fields(i).value & "</td>")
			End If
		Next
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
End If
%>
</table>

<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html><!--#include file="../includes/PrepDB.asp"-->