<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, k, PermitEdit, AppID, GranteeID, GranteeName, FiscalYear, HistoricalDataYear,  _
	SubmitID, SubmitName, SubmitTimestamp, RoundCurrency, AppURL, ApplicationSchema, GrantClassID

debug = False
GrantClassID=4

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

If Request.Form.Count>0 Then
	AppID = Request.Form("AppID")
Else
	AppID = Request.QueryString("AppID")
End If
If Len(AppID)>0 Then
	AppID = CInt(AppID)
Else
	AppID=0
End If

If Len(Request.Form("ApplicationSchema")) > 0 Then
	ApplicationSchema = Request.Form("ApplicationSchema")
ElseIf Len(Request.QueryString("ApplicationSchema")) > 0 Then
	ApplicationSchema = Request.QueryString("ApplicationSchema")
End If

If ApplicationSchema = "Application" Or ApplicationSchema = "Negotiation" Then
	If Debug = True Then
		Response.Write("<pre>ApplicationSchema'" & ApplicationSchema & "</pre>" & vbCrLf)
	End If
Else
	Response.Write("Error in Application Schema: " & ApplicationSchema)
	Response.End
End If

If AppID=0 Then	
	Response.Write("An Application ID must be provided to display this page.")
	Response.End
End If

sql = "SELECT I.AppID, I.FiscalYear, G.GranteeID, G.GranteeName, ISNULL(A.HistoricalDataYear,(I.FiscalYear-1)) AS HistoricalDataYear, " & vbCrLf & _
	"	A.SubmitID, A.SubmitTimestamp, U.Name AS SubmitName " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLF & _
	"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN CC." & ApplicationSchema & " AS A ON A.AppID=I.AppID " & vbCrLF & _
	"LEFT JOIN System.Users AS U ON U.SystemID=A.SubmitID " & vbCrLf & _
	"WHERE I.AppID=" & AppID 
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)
If rs.EOF = False Then
	AppID = rs.Fields("AppID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	HistoricalDataYear = rs.Fields("HistoricalDataYear")
	SubmitID = rs.Fields("SubmitID")
	SubmitName = rs.Fields("SubmitName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
Else
	Response.Write("Error retrieving grant application record")
	Response.End
End If

RoundCurrency = True

If ApplicationSchema = "Application" Then
	If GrantClassID = 1 And FiscalYear<2022 Then
		AppURL = "Application.asp"
	Else
		AppURL = getHomeApplicationReferenceByGrantClass(GrantClassID, AppID)
	End If
ElseIf ApplicationSchema = "Negotiation" Then
	If GrantClassID = 1 And FiscalYear<2022 Then
		AppURL = "Negotiation.asp"
	Else
		AppURL = getHomeNegotiationReferenceByGrantClass(GrantClassID, AppID)
	End If
Else
	Response.Write("Error in Application Schema")
	Response.End
End If

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
<title>MVCPA Grant <%=ApplicationSchema %> for <%=GranteeName %>: Statistics to Support Application</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function chosenButton(buttonchoice)
	{
		document.Statistics.ButtonChoice.value=buttonchoice;
		if (validateForm() == true)
			document.Statistics.submit();
	}
	function validateForm()
	{
		//alert("validate!");
		for (i = 1; i <= document.Statistics.RowCount.value; i++) {
			if (document.Statistics["Jurisdiction_" + i].value.length>0)
			{
				if (document.Statistics["CCTheft1_" + i].value.length==0)
				{
					alert("If you have a jurisdiction for a row, you should also have a value for each column.");
					document.Statistics["CCTheft1_" + i].focus();
					return false;
				}
				if (document.Statistics["CCTheft2_" + i].value.length==0)
				{
					alert("If you have a jurisdiction for a row, you should also have a value for each column.");
					document.Statistics["CCTheft2_" + i].focus();
					return false;
				}
			}
			if (document.Statistics["Jurisdiction_" + i].value.length==0)
			{
				if (document.Statistics["CCTheft1_" + i].value.length>0)
				{
					alert("If you have a value entered in a row, you should also have a Jurisdiction entered for the row.");
					document.Statistics["CCTheft1_" + i].focus();
					return false;
				}
				if (document.Statistics["CCTheft2_" + i].value.length>0)
				{
					alert("If you have a value entered in a row, you should also have a Jurisdiction entered for the row.");
					document.Statistics["CCTheft2_" + i].focus();
					return false;
				}
			}
		}
		return true;
	}

	function clearValues(row){
		document.Statistics["Jurisdiction_" + row].value = "";
		document.Statistics["CCTheft1_" + row].value = "";
		document.Statistics["CCTheft2_" + row].value = "";
		updateTotals();
	}

	function changedCurrencyField(field)
	{
		if (checkCurrencyRound(field, <%=LCase(CStr(RoundCurrency))%>) == false)
			return false;
		return true;
	}

	function changedIntegerField(field)
	{
		if (checkIntegerComma(field) == false)
			return false;
		return true;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body style="width: 100%">
<h1><%=GranteeName %> MVCPA Grant <%=ApplicationSchema %> for Fiscal Year <%=FiscalYear %></h1>
<h2>Statistics to Support Application</h2>
<%	If SubmitID>0 Then %>
<p style="text-align: center; font-weight: bold; ">The Application was submitted by <%=SubmitName%> at <%=SubmitTimestamp %> and is now locked.</p>
<%	End If %>
<form name="Statistics" method="post" action="StatisticsSubmit.asp" onsubmit="return validateForm();">
<%=HiddenField("AppID", AppID) %><%=HiddenField("FiscalYear",FiscalYear) %><%=HiddenField("ButtonChoice", "save") %><%=HiddenField("ApplicationSchema", ApplicationSchema) %>
<table style="margin: auto">
<thead>
	<tr>
		<th>Reported Cases</th>
		<th colspan="1" style="border: solid black thin; "><%=(HistoricalDataYear-1) %></th>
		<th colspan="1" style="border: solid black thin; "><%=(HistoricalDataYear) %></th>
	</tr>
	<tr style="vertical-align: bottom; ">
		<th style="width: 175px; ">Jurisdiction</th>
		<th style="width: 110px; ">Catalytic Converter Theft</th>
		<th style="width: 110px; ">Catalytic Converter Theft</th>
	</tr>
</thead>
<tbody>
<%
sql = "SELECT StatisticsID, AppID, Jurisdiction, CCTheft1, CCTheft2" & vbCrLf & _
	"FROM CC." & ApplicationSchema & "Statistics " & vbCrLF & _
	"WHERE AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"ORDER BY AppID, StatisticsID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
i=0
Set rs = Con.Execute(sql)
While rs.EOF = False
	i = i + 1
	WriteBudgetRow rs.Fields("StatisticsID"), i, rs.Fields("Jurisdiction"), rs.Fields("CCTheft1"), rs.Fields("CCTheft2"), PermitEdit
	rs.MoveNext()
Wend

For j = 1 to 5
	i = i + 1
	WriteBudgetRow 0, i, "", "", "", PermitEdit
Next
%>
</tbody>
</table>
<br />

<div style="text-align: center">
<%	If PermitEdit = True Then %>
	<input type="button" value="Save" onclick="chosenButton('save');" title="Click to save and return to this page." />
	<input type="button" value="Done" onclick="chosenButton('done');" title="Click to save and return to the application." />
	<input type="button" value="Cancel" onclick="location.href='<%=AppURL %>';" 
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
Sub WriteBudgetRow(vStatisticsID, i, vJurisdiction, vCCTheft1, vCCTheft2, vPermitEdit)
	Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
	Response.Write(WriteInCell(TextField("Jurisdiction_" & i, vJurisdiction, 40, 255, vPermitEdit, "")))
	Response.Write(WriteInCell(IntegerField("CCTheft1_" & i, formatInteger(vCCTheft1), 8, 11, vPermitEdit, "changedIntegerField(this);")))
	Response.Write(WriteInCell(IntegerField("CCTheft2_" & i, formatInteger(vCCTheft2), 8, 11, vPermitEdit, "changedIntegerField(this);")))
	If PermitEdit = True Then
		Response.Write(WriteInCell("<img style=""border: none"" src=""../images/delete.gif"" onclick=""clearValues(" & _
			i & ")"" title=""Clear values and delete record on next save"" />" & HiddenField("StatisticsID_" & i, vStatisticsID)))
	End If
	Response.Write("</tr>" & vbCrLf)
End Sub
%>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/HomeRef.asp"-->