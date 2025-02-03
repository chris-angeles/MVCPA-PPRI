<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 

Dim debug, i, j, LastCategory, Timestamp, PermitEdit, Button, GranteeID, FiscalYear, OptionID, _
	StolenVehicles, StolenVehicleValue, Certification, SubmitID, SubmitTimestamp, MAGID
Debug = False

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
	Response.Write("Now=" & Now() & vbCrLf)
	Response.Write("</pre>" & vbCrLf)
End If

Timestamp = Now()
GranteeID = Request.Form("GranteeID")
FiscalYear = Request.Form("FiscalYear")
OptionID = Request.Form("OptionID")
StolenVehicles = Request.Form("StolenVehicles")
StolenVehicleValue = Request.Form("StolenVehicleValue")
Certification = Request.Form("Certification")
Button = Request.Form("Button")

sql = "SELECT MagID, GranteeID, FiscalYear, OptionID, StolenVehicles, StolenVehicleValue, " & vbCrLf & _
	"	Certification, SubmitID, SubmitTimestamp " & vbCrLf & _
	"FROM MAG.Main " & vbCrLF & _
	"WHERE GranteeID=" & prepIntegerSQL(GranteeID) & " AND FiscalYear=" & prepIntegerSQL(FiscalYear)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(SQL)
If rs.EOF = True Then
	' Insert
	If Button = "submit" Then
		SubmitID = UserSystemID
		SubmitTimeStamp = Timestamp
	Else
		SubmitID = null
		SubmitTimestamp = Null
	End If
	sql = "INSERT INTO MAG.Main(GranteeID, FiscalYear, OptionID, StolenVehicles, StolenVehicleValue, Certification, SubmitID, SubmitTimestamp, UpdateID, UpdateTimestamp) VALUES " & vbCrLf & _
		"(" & prepIntegerSQL(GranteeID) & ", " & _
		prepIntegerSQL(FiscalYear) & ", " & _
		prepIntegerSQL(OptionID) & ", " & _
		prepIntegerSQL(StolenVehicles) & ", " & _
		prepNumberSQL(StolenVehicleValue) & ", " & _
		prepBitRequiredSQL(Certification) & ", " & _
		prepIntegerSQL(SubmitID) & ", " & _
		prepDateSQL(SubmitTimestamp) & ", " & _
		prepIntegerSQL(UserSystemID) & ", " & _
		prepDateSQL(Timestamp) & ")"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
Else
	' Update
	If Button = "submit" Then
		SubmitID = UserSystemID
		SubmitTimeStamp = Timestamp
	Else
		SubmitID = rs.Fields("SubmitID")
		SubmitTimestamp = rs.Fields("SubmitTimestamp")
	End If
	sql = "UPDATE MAG.Main SET " & _
		"GranteeID=" & prepIntegerSQL(GranteeID) & ", " & _
		"FiscalYear=" & prepIntegerSQL(FiscalYear) & ", " & _
		"OptionID=" & prepIntegerSQL(OptionID) & ", " & _
		"StolenVehicles=" & prepIntegerSQL(StolenVehicles) & ", " & _
		"StolenVehicleValue=" & prepNumberSQL(StolenVehicleValue) & ", " & _
		"Certification=" & prepBitRequiredSQL(Certification) & ", " & _
		"SubmitID=" & prepIntegerSQL(SubmitID) & ", " & _
		"SubmitTimestamp=" & prepDateSQL(SubmitTimestamp) & ", " & _
		"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
		"UpdateTimestamp=" & prepDateSQL(Timestamp) & " " & vbCrLf & _
		"WHERE GranteeID=" & prepIntegerSQL(GranteeID) & " AND FiscalYear=" & prepIntegerSQL(FiscalYear)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
End If

If Debug = True Then
	Response.Write("<a href=MAGApplication.asp?GranteeID=" & GranteeID & "&FiscalYear=" & FiscalYear & ">Go To Page</a>")
Else
	Response.Redirect("MAGApplication.asp?GranteeID=" & GranteeID & "&FiscalYear=" & FiscalYear)
End if
%><!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/prepDB.asp"-->