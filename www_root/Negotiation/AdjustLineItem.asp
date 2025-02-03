<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, FiscalYear, AppID, NewAppID, ProgramName, GrantType, GranteeName, _
	BudgetItemID, UnallowedItem, AllowedAmount, Issue, GrantAwardAmount, GranteeID
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

If MVCPARights = False Then
	Response.Write("You do not have access rights to this page.")
	Response.End
End If
If Len(Request.Form("FiscalYear"))>0 Then 
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then 
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
ElseIf Len(Session("FiscalYear"))>0 Then 
	FiscalYear = CInt(Session("FiscalYear"))
Else
	FiscalYear=Year(Date())
End If

GranteeID = Session("GranteeID")
If Len(Request.Form("NewAppID"))>0 Then 
	NewAppID = CInt(Request.Form("NewAppID"))
ElseIf Len(Request.QueryString("NewAppID"))>0 Then 
	NewAppID = CInt(Request.QueryString("NewAppID"))
ElseIf GranteeID>0 And FiscalYear>0 Then
	sql = "SELECT AppID FROM Negotiation.Main WHERE GranteeID=" & prepIntegerSQL(GranteeID) & " AND FiscalYear=" & prepIntegerSQL(FiscalYear)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		NewAppID = rs.Fields("AppID")
	Else
		AppID = 0
	End If
Else
	NewAppID = 0
End If

If Len(Request.Form("AppID"))>0 Then 
	AppID = CInt(Request.Form("AppID"))
ElseIf Len(Request.QueryString("AppID"))>0 Then 
	AppID = CInt(Request.QueryString("AppID"))
Else
	AppID = 0
End If

For Each i in Request.Form
	If Left(i,14)="AllowedAmount_" Then
		BudgetItemID = CLng(Mid(i, 15)) 
		If Request.Form("UnallowedItem_" & BudgetItemID) = "1" Then
			UnallowedItem = True
		Else
			UnallowedItem = False
		End If
		AllowedAmount = Request.Form("AllowedAmount_" & BudgetItemID)
		If Len(AllowedAmount)>0 Then
			AllowedAmount = CDbl(AllowedAmount)
		Else
			AllowedAmount = null
		End If
		If Request.Form("Issue_" & BudgetItemID) = "1" Then
			Issue = True
		Else
			Issue = False
		End If
		sql = "SELECT BudgetItemID, ISNULL(UnallowedItem,0) AS UnallowedItem, AllowedAmount, ISNULL(Issue,0) AS Issue FROM Negotiation.BudgetDetails WHERE BudgetItemID=" & BudgetItemID
		If Debug = True Then
			Response.Write("<pre>" & i & ": " & BudgetItemID & ", " & UnallowedItem & ", " & AllowedAmount & ", " & Issue & vbCrLf & sql & "</pre>" & vbCrLf)
		End If
		Set rs = Con.Execute(sql)
		If rs.EOF = True Then
			Response.Write("Error: The Budget line item " & BudgetItemID & " does not exist.")
			Response.End
		ElseIf rs.Fields("UnallowedItem")=UnallowedItem And rs.Fields("Issue")=Issue And (rs.Fields("AllowedAmount") = AllowedAmount Or (IsNull(rs.Fields("AllowedAmount") = True And IsNull(AllowedAmount) = True))) Then
			'Response.Write("Don't Update")
		Else
			'Response.Write("Update")
			sql = "UPDATE Negotiation.BudgetDetails " & vbCrLf & _
				"SET UnallowedItem=" & prepBitSQL(UnallowedItem) & _
				", AllowedAmount=" & prepNumberSQL(AllowedAmount) & _
				", Issue=" & prepBitSQL(Issue) & " " & vbCrLf & _
				"WHERE BudgetItemID=" & BudgetItemID
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		End If
	End If
Next

' Be sure to do any processing of Application Infornation before AppID is updated to NewAppID.

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Adjust Line Items</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm()
	{
		document.AdjustLineItem.submit();
		return false;
	}

	function checkCurrency(field)
	{
		var fieldvalue;

		if (field.value == "") {
			return true;
		}
		fieldvalue = stripFormatting(field.value)
		if (isNaN(fieldvalue)) {
			alert("The value entered is not a valid number");
			field.focus();
			return false;
		}
		field.value = currency(fieldvalue);
		return true;
	}

	function currency(num)
	{
		var prefix = "$";
		var suffix = "";
		var result = "";
		if (num < 0) {
			prefix = "($";
			suffix = ")";
			num = -num;
		}
		var temp = Math.round(num * 100.0); // convert to pennies!
		if (temp < 10) return prefix + "0.0" + temp + suffix;
		if (temp < 100) return prefix + "0." + temp + suffix;

		temp = temp.toString()
		if (temp.length > 11) {
			return prefix + temp.substring(0, temp.length - 11) + "," + temp.substring(temp.length - 11, temp.length - 8) + "," + temp.substring(temp.length - 8, temp.length - 5) + "," + temp.substring(temp.length - 5, temp.length - 2) + "." + temp.substring(temp.length - 2, temp.length) + suffix;

		}
		if (temp.length > 8) {
			return prefix + temp.substring(0, temp.length - 8) + "," + temp.substring(temp.length - 8, temp.length - 5) + "," + temp.substring(temp.length - 5, temp.length - 2) + "." + temp.substring(temp.length - 2, temp.length) + suffix;

		}
		if (temp.length > 5) {
			return prefix + temp.substring(0, temp.length - 5) + "," + temp.substring(temp.length - 5, temp.length - 2) + "." + temp.substring(temp.length - 2, temp.length) + suffix;

		}
		return prefix + temp.substring(0, temp.length - 2) + "." + temp.substring(temp.length - 2) + suffix;
	}

	function stripFormatting(fieldvalue)
	{
		var num
		num = fieldvalue
		if (fieldvalue != "") {
			if (num.charAt(0) == "(" && num.charAt(num.length - 1) == ")") {
				num = "-" + num.substring(1, num.length - 1)
			}
			num = num.replace(/[^1234567890.-]/g, "")
		}
		if (num.charAt(1) == '(') {
			return parseFloat('-' + num)
		}
		else {
			return parseFloat(num);
		}
	}

</script>
</head>
<body style="width: 100%">
<div class="sectiontitle">Adjust Line Items for Negotiation Records</div>
<form name="AdjustLineItem" method="post" action="AdjustLineItem.asp">
<%
AppID = NewAppID
sql = "SELECT A.AppID, A.ProgramName, C.GranteeName, A.GrantTypeID, B.GrantType, ISNULL(D.GrantAwardAmount,0.0) AS GrantAwardAmount " & vbCrLf & _
	"FROM Negotiation.Main AS A " & vbCrLf & _
	"LEFT JOIN Lookup.GrantType AS B ON B.GrantTypeID=A.GrantTypeID AND B.Version=1 " & vbCrLf & _
	"LEFT JOIN Grantees AS C ON C.GranteeID=A.GranteeID " & vbCrLf & _
	"LEFT JOIN Application.Admin AS D ON D.AppID=A.AppID " & vbCrLf & _
	"WHERE A.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
	"ORDER BY REPLACE(C.GranteeName,'City of ',''), A.GrantTypeID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	Response.Write(FiscalYear & "Applications: <select name=""NewAppID"" onchange=""submitForm();"">" & vbCrLf)
	Response.Write("<option value=""0"">Select Application</option>" & vbCrLf)
	While rs.EOF = False
		If rs.Fields("AppID") = AppID Then
			If rs.Fields("GrantAwardAmount")>0 Then
				Response.Write("<option value=""" & rs.Fields("AppID") & """ selected>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & ", " & formatcurrency(rs.Fields("GrantAwardAmount"),2,true,true,true) & "</option>" & vbCrLf)
			Else
				Response.Write("<option value=""" & rs.Fields("AppID") & """ selected>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</option>" & vbCrLf)
			End If
			ProgramName = rs.Fields("ProgramName")
			GranteeName = rs.fields("GranteeName")
			GrantType = rs.Fields("GrantType")
			GrantAwardAmount = rs.Fields("GrantAwardAmount")
		Else
			If rs.Fields("GrantAwardAmount")>0 Then
				Response.Write("<option value=""" & rs.Fields("AppID") & """>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & ", " & formatcurrency(rs.Fields("GrantAwardAmount"),2,true,true,true) & "</option>" & vbCrLf)
			Else
				Response.Write("<option value=""" & rs.Fields("AppID") & """>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</option>" & vbCrLf)
			End If
		End If
		rs.MoveNext()
	Wend
	Response.Write("</select>" & vbCrLf)
End If

Response.Write("<input type=""hidden"" name=""AppID"" value=""" & AppID & """ />" & vbCrLf)

If AppID>0 Then
	Response.Write("<br>Program: " & GranteeName & ", " & ProgramName & ", " & GrantType)
	If GrantAwardAmount > 0 Then
		Response.Write(", Award: " & formatcurrency(GrantAwardAmount,2, true, true, true))
	End If
	Response.Write("<br /><br />" & vbCrLf)

	sql = "SELECT B.BudgetItemID,  " & vbCrLf & _
		"	C.BudgetCategory, B.Description, S.SubCategory, PctTime,  " & vbCrLf & _
		"	MVCPAFunds, CashMatch, LineTotal, InKindMatch , " & vbCrLf & _
		"	Issue, UnallowedItem, AllowedAmount, FiscalYear " & vbCrLf & _
		"FROM Negotiation.BudgetDetails AS B " & vbCrLf & _
		"LEFT JOIN Negotiation.Main AS A ON A.AppID=B.AppID " & vbCrLf & _
		"LEFT JOIN Lookup.BudgetCategories AS C ON C.BudgetCategoryID=B.BudgetCategoryID " & vbCrLf & _
		"LEFT JOIN Lookup.BudgetSubcategories AS S ON S.BudgetCategoryID=B.BudgetCategoryID AND S.SubCategoryID=B.SubCategoryID " & vbCrLf & _
		"LEFT JOIN Grantees AS G ON G.GranteeID=A.GranteeID " & vbCrLf & _
		"LEFT JOIN Lookup.GrantType AS T ON T.GrantTypeID=A.GrantTypeID AND T.Version=1 " & vbCrLf & _
		"WHERE B.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
		"ORDER BY C.BudgetCategoryID, B.SubCategoryID "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		Response.Write("<table>" & vbCrLf)
		Response.Write("<thead>" & vbCrLf)
		Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
		Response.Write("<th>Seq</th>" & vbCrLf)
		Response.Write("<th>Category</th>" & vbCrLf)
		Response.Write("<th>Description</th>" & vbCrLf)
		Response.Write("<th>SubCategory</th>" & vbCrLf)
		Response.Write("<th>MVCPA Funds</th>" & vbCrLf)
		Response.Write("<th>Cash Match</th>" & vbCrLf)
		Response.Write("<th>Line Total</th>" & vbCrLf)
		Response.Write("<th>Unallowed Line Item</th>" & vbCrLf)
		Response.Write("<th>Allowed Amount</th>" & vbCrLf)
		Response.Write("<th>Issue Requiring Adjustment</th>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
		Response.Write("</thead>" & vbCrLf)

		Response.Write("<tbody>" & vbCrLf)
		While rs.EOF = False
			Response.Write("<tr>" & vbCrLf)
			Response.Write("<td>" & rs.Fields("BudgetItemID") & "</td>" & vbCrLf)
			Response.Write("<td>" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
			Response.Write("<td>" & rs.Fields("Description") & "</td>" & vbCrLf)
			Response.Write("<td>" & rs.Fields("SubCategory") & "</td>" & vbCrLf)
			If IsNull(rs.Fields("MVCPAFunds")) = True Then
				Response.Write("<td></td>" & vbCrLf)
			Else
				Response.Write("<td style=""text-align: right; "">" & formatcurrency(rs.Fields("MVCPAFunds"),2,true,true,true) & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("CashMatch")) = True Then
				Response.Write("<td></td>" & vbCrLf)
			Else
				Response.Write("<td style=""text-align: right; "">" & formatcurrency(rs.Fields("CashMatch"),2,true,true,true) & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("LineTotal")) = True Then
				Response.Write("<td></td>" & vbCrLf)
			Else
				Response.Write("<td style=""text-align: right; "">" & formatcurrency(rs.Fields("LineTotal"),2,true,true,true) & "</td>" & vbCrLf)
			End If
			If rs.Fields("UnallowedItem") = True Then
				Response.Write("<td style=""text-align: center; ""><input type=""checkbox"" name=""UnallowedItem_" & rs.Fields("BudgetItemID") & """ value=""1"" checked></td>" & vbCrLf)
			Else
				Response.Write("<td style=""text-align: center; ""><input type=""checkbox"" name=""UnallowedItem_" & rs.Fields("BudgetItemID") & """ value=""1""></td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("AllowedAmount")) = True Then
				Response.Write("<td><input type=""text"" name=""AllowedAmount_" & rs.Fields("BudgetItemID") & """ size=""8"" maxlength=""10"" value="""" onchange=""return checkCurrency(this);"" style=""text-align: right; ""></td>" & vbCrLf)
			Else
				Response.Write("<td><input type=""text"" name=""AllowedAmount_" & rs.Fields("BudgetItemID") & """ size=""8"" maxlength=""10"" value=""" & formatcurrency(rs.Fields("AllowedAmount")) & """ onchange=""return checkCurrency(this);"" style=""text-align: right; ""></td>" & vbCrLf)
			End If
			If rs.Fields("Issue") = True Then
				Response.Write("<td style=""text-align: center; ""><input type=""checkbox"" name=""Issue_" & rs.Fields("BudgetItemID") & """ value=""1"" checked></td>" & vbCrLf)
			Else
				Response.Write("<td style=""text-align: center; ""><input type=""checkbox"" name=""Issue_" & rs.Fields("BudgetItemID") & """ value=""1""></td>" & vbCrLf)
			End If
			Response.Write("</tr>" & vbCrLf)
			rs.MoveNext
		Wend
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
	End If

End If

%>
<div style="width: 100%; text-align: center;">
<input type="button" value="submit" onclick="submitForm();" />
<input type="button" value="close" onclick="window.close();" title="Close window without saving current edits."/>
</div>
</form>

</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->