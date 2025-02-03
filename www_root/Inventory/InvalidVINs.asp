<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, Counter, OrderBy, OrderByDescription, OrderByField, ShowAll, ShowGood, ShowExcel, Columns, _
	ShowOnlySubmitted, ApplicationSchema
OrderByDescription = Array("Grantee ID", "Grantee Name", "Inventory ID")
OrderByField = Array("[Grantee ID]", "REPLACE([GranteeName],'City of ','')", "A.InventoryID")
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
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
End If

If Request.Form("ShowAll")="1" Then
	ShowAll = True
ElseIf Request.QueryString("ShowAll")="1" Then
	ShowAll = True
Else
	ShowAll = False
End If

If Request.Form("ShowGood")="1" Then
	ShowGood = True
ElseIf Request.QueryString("ShowGood")="1" Then
	ShowGood = True
Else
	ShowGood = False
End If

If Request.QueryString("ShowExcel")="1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "SELECT A.InventoryID AS [Inventory ID], A.GranteeID AS [Grantee ID], C.GranteeName AS [Grantee], ItemDescription, ModelYear AS [Model Year], MakeManufacturer AS [Make], Model, " & vbCrLf & _
	"	serialNo AS VIN, B.Is_Valid, B.Length_Valid, B.Characters_Valid, B.Check_Digit, B.Product, B.Remainder " & vbCrLf
If ShowAll = True Then
	sql = sql & ", NotInventoryItem AS [Not Inventory Item Date], DateOfDisposal AS [Date Of Disposal] " & vbCrLf
End If
sql = sql & "FROM Inventory AS A " & vbCrLf & _
	"CROSS APPLY dbo.fnValidateVIN(SerialNo) AS B " & vbCrLf & _
	"LEFT JOIN Grantees AS C ON C.GranteeID=A.GranteeID " & vbCrLf & _
	"WHERE AssetClassID='01-01' "
If ShowGood = False Then
	sql = sql & "AND B.IS_VALID=0 "
End If
If ShowAll = False Then
	sql = sql & "AND NotInventoryItem IS NULL AND DateOfDisposal IS NULL"
End If
sql = sql & vbCrLf & "ORDER BY " & OrderByField(OrderBy)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

Set rs=Con.Execute(sql)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=InvalidVINs.xls"
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Vehicle Report</title>
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
</select>&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowGood" id="ShowGood" value="1" <%
If ShowGood=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowAll">Show Valid VINs also</label>&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowAll" id="Checkbox1" value="1" <%
If ShowAll=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowAll">Show All Bad VINs</label>&nbsp;&nbsp;&nbsp;
<a href="InvalidVINs.asp?ShowExcel=1&OrderBy=<%=OrderBy %>&ShowGood=<%If ShowGood=True Then Response.Write("1") Else Response.Write("0") End If %>&ShowAll=<%If ShowAll=True Then Response.Write("1") Else Response.Write("0") End If %>" target="_blank">Excel</a>
</form>
<table class="reporttable">
<%
End if

If rs.EOF = False Then
	Columns = rs.Fields.count
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom"">" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top"">" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write(vbTab & "<td></td>")
			ElseIf rs.Fields(i).Name = "Inventory ID" Then
				Response.Write("<td style=""text-align: right; "">" & rs.Fields(i).value & "</td>")
			ElseIf rs.Fields(i).Name = "VIN" Then
				If rs.Fields("Is_Valid") = True Then
					Response.Write("<td style=""text-align: left; "">" & rs.Fields(i).value & "</td>")
				Else
					Response.Write("<td style=""text-align: left; color: red; "">" & rs.Fields(i).value & "</td>")
				End If
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf InStr(1, rs.Fields(i).Name, "Date")>0 Then
				Response.Write("<td style=""text-align: right"">" & formatdatetime(rs.Fields(i).value, vbGeneralDate) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right; "">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adNumeric Then
				Response.Write("<td style=""text-align: right; "">" & formatnumber(rs.Fields(i).value,2, true, false, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBoolean Then
				Response.Write("<td style=""text-align: center; "">" & BitAsX(rs.Fields(i).value) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		'Response.Write("<td>" & rs.Fields("Cash Match Pct Chg").Type & "</td>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
		Counter = Counter + 1
		rs.MoveNext
	Wend
	Response.Write("<tr><td style=""text-align: center"" colspan=""" & Columns & """>" & Counter & " records</td></tr>")
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If

If ShowExcel = False Then %>
<tr><th style="width: 100%; text-align: center" colspan="<%=columns %>"><input type="button" value="Close" onclick="window.close();" /></th></tr>
<%	
End If 
%>
</table>
<% If ShowExcel = False Then %>
</body>
</html>
<%	End If %>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<%
Function BitAsX(vValue)
	If vValue = True Then
		BitAsX = "X"
	Else
		BitAsX = ""
	End If
End Function
%>