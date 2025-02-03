<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, k, PermitEdit, AppID, GrantClassID, GrantClass, MatchTypeID, MatchType, GranteeID, GranteeName, FiscalYear, _
	SubmitID, SubmitName, SubmitTimestamp, NoOfOptions, RoundCurrency, AppURL

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

If Request.Form.Count>0 Then
	AppID = Request.Form("AppID")
	MatchTypeID = Request.Form("MatchTypeID")
Else
	AppID = Request.QueryString("AppID")
	MatchTypeID = Request.QueryString("MatchTypeID")
End If
If Len(AppID)>0 Then
	AppID = CInt(AppID)
Else
	AppID=0
End If
If Len(MatchTypeID)>0 Then
	MatchTypeID = CInt(MatchTypeID)
Else
	MatchTypeID=0
End If

If AppID=0 Then	
	Response.Write("An Application ID must be provided to display this page.")
	Response.End
End If
If MatchTypeID=0 Then	
	Response.Write("An Match Type ID must be provided to display this page.")
	Response.End
End If

sql = "SELECT I.AppID, I.GrantClassID, D.GrantClass, I.FiscalYear, G.GranteeID, G.GranteeName, " & vbCrLf & _
	"	A.SubmitID, A.SubmitTimestamp, U.Name AS SubmitName, T.MatchType, " & vbCrLf & _
	"	NoOfOptions = (SELECT ISNULL(MAX(MatchSourceID),0) FROM Lookup.MatchSources) " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN Application.Main AS A ON A.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=A.SubmitID " & vbCrLf & _
	"LEFT JOIN Lookup.MatchTypes AS T ON T.MatchTypeID=" & prepStringSQL(MatchTypeID) & " " & vbCrLf & _
	"LEFT JOIN Lookup.GrantClass AS D ON D.GrantClassID=I.GrantClassID " & vbCrLf & _
	"WHERE I.AppID=" & AppID 
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)
If rs.EOF = False Then
	AppID = rs.Fields("AppID")
	GrantClassID = rs.Fields("GrantClassID")
	GrantClass = rs.Fields("GrantClass")
	MatchType = rs.Fields("MatchType")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	SubmitID = rs.Fields("SubmitID")
	SubmitName = rs.Fields("SubmitName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	NoOfOptions = rs.Fields("NoOfOptions")
Else
	Response.Write("Error retrieving grant application record")
	Response.End
End If

' Start rounding dollar amounts as of 2020.
If FiscalYear>=2020 Then
	RoundCurrency = True
Else
	RoundCurrency = False
End If

If FiscalYear<2022 Then
	AppURL = "Application.asp"
Else
	AppURL = getHomeApplicationReferenceByGrantClass(GrantClassID, AppID)
End If

REDIM Options(NoOfOptions)
If FiscalYear>2021 Then
	sql = "SELECT MatchSourceID, MatchSource " & vbCrLf & _
		"FROM Lookup.MatchSources " & vbCrLf & _
		"WHERE MatchSource <> 'DPS' " & vbCrLf & _
		"ORDER BY MatchSourceID "
Else
	sql = "SELECT MatchSourceID, MatchSource " & vbCrLf & _
		"FROM Lookup.MatchSources " & vbCrLf & _
		"ORDER BY MatchSourceID "
End If
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs=Con.Execute(sql)
while rs.EOF = False
	Options(rs.Fields("MatchSourceID")) = rs.Fields("MatchSource")
	rs.MoveNext()
Wend

If GranteeID>0 Then
	If IsNull(SubmitID) = True Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
	ElseIf ISNull(SubmitID) = False Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, True)
	Else
		PermitEdit = False
	End If
Else
		PermitEdit = False
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Grant Application for <%=GranteeName %>: <%=MatchType %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function chosenButton(buttonchoice)
	{
		document.Matches.ButtonChoice.value=buttonchoice;
		if (validateForm() == true)
			document.Matches.submit();
	}
	function validateForm()
	{
		//alert("validate!");
		for (i = 1; i <= document.Matches.RowCount.value; i++) {
			if (document.Matches["Source_" + i].value.length>0)
			{
				if (document.Matches["Amount_" + i].value.length ==0)
				{
					alert("If you have a source for a row, you should also have an amount.");
					document.Matches["Amount_" + i].focus();
					return false;
				}
				if (document.Matches["MatchSourceID_" + i].selectedIndex == 0) {
					alert("You must select an item from the dropdown menu");
					document.Matches["MatchSourceID_" + i].focus();
					return false;
				}
			}
		}
		return true;
	}

	function clearValues(row){
		document.Matches["Source_" + row].value = "";
		document.Matches["Amount_" + row].value = "";
		updateTotals();
	}

	function changedCurrencyField(field)
	{
		if (checkCurrencyRound(field, <%=LCase(CStr(RoundCurrency))%>) == false)
			return false;
		updateTotals();
		return true;
	}

	function updateTotals()
	{
		//alert("Update Totals");

		var total = 0;
		for (i = 1; i <= document.Matches.RowCount.value; i++)
		{
			total = total + getNumericValue(document.Matches["Amount_" + i].value);
		}
		document.Matches.Total.value = currencyRound(total, <%=LCase(CStr(RoundCurrency))%>);
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body style="width: 100%" onload="updateTotals()">
<h1><%=GranteeName %> MVCPA <%=GrantClass %> Application for Fiscal Year <%=FiscalYear %></h1>
<h2><%=MatchType %> Detail</h2>
<%	If SubmitID>0 Then %>
<p style="text-align: center; font-weight: bold; ">The Application was submitted by <%=SubmitName%> at <%=SubmitTimestamp %> and is now locked.</p>
<%	End If %>
<form name="Matches" method="post" action="MatchesSubmit.asp" onsubmit="return validateForm()">
<%=HiddenField("AppID", AppID) %><%=HiddenField("GrantClassID", GrantClassID)%><%=HiddenField("MatchTypeID",MatchTypeID) %><%=HiddenField("FiscalYear",FiscalYear) %><%=HiddenField("ButtonChoice", "save") %>
<h2>Source of <%=MatchType %></h2>
<table style="margin: auto">
<thead>
	<tr>
		<th>Source</th>
		<th><%=MatchType %> Type</th>
		<th>Amount</th>
	</tr>
</thead>
<tbody>
<%
sql = "SELECT MatchID, AppID, MatchTypeID, Source, MatchSourceID, Amount " & vbCrLf & _
	"FROM Application.Matches " & vbCrLf & _
	"WHERE AppID=" & prepIntegerSQL(AppID) & " AND MatchTypeID=" & prepIntegerSQL(MatchTypeID) & " " & vbCrLf & _
	"ORDER BY MatchTypeID, MatchID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
i=0
Set rs = Con.Execute(sql)
While rs.EOF = False
	i = i + 1
	WriteBudgetRow rs.Fields("MatchID"), i, rs.Fields("Source"), rs.Fields("MatchSourceID"), rs.Fields("Amount"), PermitEdit
	rs.MoveNext()
Wend

For j = 1 to 5
	i = i + 1
	WriteBudgetRow 0, i, "", "", "", PermitEdit
Next
%>
</tbody>
<tfoot>
<tr>
	<td>Total</td>
	<td></td>
	<td><input type=text name="Total" value="" size="12" maxLength="14" style="text-align: right; border-style: none" readonly="readonly" /></td>
</tr>
</tfoot>
</table>
<br />

<div style="text-align: center">
<%	If PermitEdit = True Then %>
	<input type="button" value="Save" onclick="chosenButton('save');" title="Click to save and return to this page." />
	<input type="button" value="Done" onclick="chosenButton('done');" title="Click to save and return to the application." />
	<input type="button" value="Cancel" onclick="location.href = '<%=AppURL %>';" 
		title="Return to the application without saving latest changes." />
<%	Else %>
	<input type="button" value="Return" onclick="location.href = '<%=AppURL%>';" 
		title="Return to the application main page." />
<%	End If %>
</div>

<%=HiddenField("RowCount", i) %>
</form>
</body>
</html>
<%
Sub WriteBudgetRow(vMatchID, i, vSource, vMatchSourceID, vAmount, vPermitEdit)
	Response.Write("<tr>" & vbCrLf)
	Response.Write(WriteInCell(TextField("Source_" & i, vSource, 50, 255, vPermitEdit, "")))
	Response.Write("<td><select name=""MatchSourceID_" & i & """>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<option value=""0""></option>" & vbCrLf)
	For k = 1 to UBound(Options)
		If Len(Options(k))>0 Then
			If k = vMatchSourceID Then
				Response.Write(vbTab & vbTab & "<option value=""" & k & """ selected=""selected"">" & Options(k) & "</option>" & vbCrLf)
			Else
				Response.Write(vbTab & vbTab & "<option value=""" & k & """>" & Options(k) & "</option>" & vbCrLf)
			End If
		End If
	Next
	Response.Write("</select></td>" & vbCrLf)
	Response.Write(WriteInCell(CurrencyFieldRound("Amount_" & i, vAmount, 12, 14, RoundCurrency, vPermitEdit, "changedCurrencyField(this);")))
	If PermitEdit = True Then
		Response.Write(WriteInCell("<img style=""border: none"" src=""../images/delete.gif"" onclick=""clearValues(" & _
			i & ")"" title=""Clear values and delete record on next save"" />" & HiddenField("MatchID_" & i, vMatchID)))
	End If
	Response.Write("</tr>" & vbCrLf)
End Sub
%>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/HomeRef.asp"-->