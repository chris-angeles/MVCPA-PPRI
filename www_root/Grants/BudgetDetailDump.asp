<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, FiscalYear, ShowExcel
debug = False
If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLF)
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
	Response.Write("</pre>" & vbCrLF)
End If

IF Request.QueryString("ShowExcel")="1" Then
	ShowExcel = True
Else
	ShowExcel = false
End If

FiscalYear = Request.QueryString("FiscalYear")
IF Debug =True Then
	Response.Write("FiscalYear=" & FiscalYear)
End If

If Len(FiscalYear)=0 Then
	Response.Write("No FiscalYear specified.")
	Response.End
End If
ShowExcel = True
'FiscalYear=2018

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=BudgetDetailDump" & FiscalYear & ".xls"
Else ' Start of Web only code
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%;">
<%
End If

sql = "SELECT * FROM [Grants].vwBudgetDetail WHERE [Fiscal Year]=" & prepIntegerSQL(FiscalYear) & " ORDER BY 3,2 "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write("<table>" & vbCrLf)
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom;"">" & vbCrLf)
	For i = 0 to rs.Fields.Count-2
		Response.Write("<th>" & rs.Fields(i).Name & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLf)
		For i = 0 to rs.Fields.Count-2
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf (InStr(rs.Fields(i).Name,"ID")>0 or rs.Fields(i).Name="Seq") And (rs.Fields(i).Type = adInteger Or rs.Fields(i).Type = adSmallInt) Then
				Response.Write("<td style=""text-align: right;"">" & formatnumber(rs.Fields(i).Value,0,true,false,false) & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right;"">" & formatnumber(rs.Fields(i).Value,2,true,false,true) & "</td>")
			ElseIf rs.Fields(i).Type = adInteger Or rs.Fields(i).Type = adSmallInt Then
				Response.Write("<td style=""text-align: right;"">" & formatnumber(rs.Fields(i).Value,0,true,false,true) & "</td>")
			Else
				Response.Write("<td style=""text-align: left;"">" & rs.Fields(i).Value & "</td>")
			End if
		Next
		Response.Write(vbCrLF & "</tr>" & vbCrLf)
		rs.MoveNext()
	Wend
	Response.Write("</tbody>" & vbCrLf)
	Response.Write("</table>" & vbCrLf)
Else
	Response.WRite("No Data To Show")
End If
IF ShowExcel = False Then
%>
</body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->