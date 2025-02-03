<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, GranteeID, Inactive, Timestamp, rowsAffected

debug = False
Timestamp = Now()

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

GranteeID = Request.QueryString("GranteeID")
Inactive = Request.QueryString("Inactive")

If GranteeID="" Then
	Response.Write("No Grantee ID Provided")
	Response.End
ElseIf Inactive="" Then
	Response.Write("No Active Status Provided")
	Response.End
Else
	If Debug = True Then
		Response.Write(GranteeID & "<br />")
		Response.Write(Inactive & "<br />")
		Response.Flush
	End If
	GranteeID = CInt(GranteeID)
	Inactive = CInt(Inactive)
End If

sql = "UPDATE Grantees SET Inactive=" & prepBitSQL(Inactive) & ", UpdateID=" & prepIntegerSQL(UserSystemID) & ", UpdateTimestamp=" & prepStringSQL(Timestamp) & " WHERE GranteeID=" & prepIntegerSQL(GranteeID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Activate/Deactivate Grantees</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
    function refreshAndClose() {
        window.opener.location.reload(true);
        window.close();
    }
</script>
</head>
<body style="width: 100%; text-align: center; ">
<br />
<br />
<br />
<h3>Active Status has been changed for <%=GranteeID %></h3>
<br />
<form><input type="button" name="Close" value="Close" onclick="refreshAndClose();" /></form>
<script></script>
</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->