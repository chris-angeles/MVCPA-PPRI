<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j
Debug = True

sql = "SELECT * FROM Negotiation.vwReimbursementRates WHERE Fiscal_Year=2022"
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write("<tr>")
	Response.Write(vbTab & "<td>" & rs.Fields("AppID") & "/td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("Grantee_Name") & "/td>" & vbCrLf)
	Response.Write(vbTab & "</td>")
	Response.Write("</tr>")
	rs.MoveNext
Wend
%>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->