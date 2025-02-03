<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, k, PermitEdit, AppID, GrantClassID, GranteeID, GranteeName, FiscalYear, _
	BudgetCategoryID, BudgetCategory, CategoryDescription, Narrative, NoOfOptions, _
	SubmitID, SubmitByName, SubmitTimestamp, RoundCurrency, AppURL, NegotiationLocked

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

If Request.Form.Count>0 Then
	AppID = Request.Form("AppID")
	BudgetCategoryID = Request.Form("BudgetCategoryID")
Else
	AppID = Request.QueryString("AppID")
	BudgetCategoryID = Request.QueryString("BudgetCategoryID")
End If
IF Len(AppID)>0 Then
	AppID = CInt(AppID)
Else
	AppID=0
End If
If Len(BudgetCategoryID)>0 Then 
	BudgetCategoryID = CInt(BudgetCategoryID)
Else
	BudgetCategoryID = 0
End If
If AppID=0 Or BudgetCategoryID=0 Then	
	Response.Write("An Application ID and a Budget Category ID must be provided to display this page.")
	Response.End
End If
sql = "SELECT I.AppID, I.GrantClassID, D.GrantClass, I.FiscalYear, G.GranteeID, G.GranteeName, " & vbCrLf & _
	"	B.BudgetCategory, B.CategoryDescription, C.Narrative, U.Name AS SubmitByName, " & vbCrLf & _
	"	CASE WHEN I.GrantClassID=1 THEN A1.BudgetCashMatch WHEN I.GrantClassID=4 THEN A4.BudgetCashMatch ELSE NULL END AS BudgetCashMatch, " & vbCrLf & _
	"	CASE WHEN I.GrantClassID=1 THEN A1.SubmitID WHEN I.GrantClassID=4 THEN A4.SubmitID ELSE NULL END AS SubmitID, " & vbCrLf & _
	"	CASE WHEN I.GrantClassID=1 THEN A1.SubmitTimestamp WHEN I.GrantClassID=4 THEN A4.SubmitTimestamp ELSE NULL END AS SubmitTimestamp, " & vbCrLf & _
	"	CAST(ISNULL(NegotiationLocked,0) AS Bit) AS NegotiationLocked, " & vbCrLf & _
	"	NoOfOptions = (SELECT ISNULL(MAX(SubCategoryID),0) FROM Lookup.BudgetSubcategories WHERE BudgetCategoryID=B.BudgetCategoryID) " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN Negotiation.Main AS A1 ON A1.AppID=I.AppID " & vbCrLF & _
	"LEFT JOIN CC.Negotiation AS A4 ON A4.AppID=I.AppID " & vbCrLF & _
	"LEFT JOIN CC.Admin AS L ON L.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN Lookup.BudgetCategories AS B ON B.BudgetCategoryID=" & prepIntegerSQL(BudgetCategoryID) & " " & vbCrLf & _
	"LEFT JOIN Negotiation.BudgetDetailNarrative AS C ON C.BudgetCategoryID=" & prepIntegerSQL(BudgetCategoryID) & _
		" AND C.AppID=I.AppID " & vbCrLF & _
	"LEFT JOIN Lookup.GrantClass AS D ON D.GrantClassID=I.GrantClassID " & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=CASE WHEN I.GrantClassID=1 THEN A1.SubmitID WHEN I.GrantClassID=4 THEN A4.SubmitID ELSE NULL END " & vbCrLf & _
	"WHERE I.AppID=" & AppID
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)
If rs.EOF = False Then
	AppID = rs.Fields("AppID")
	GrantClassID = rs.Fields("GrantClassID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	BudgetCategory = rs.Fields("BudgetCategory")
	CategoryDescription = rs.Fields("CategoryDescription")
	Narrative = rs.Fields("Narrative")
	NoOfOptions = rs.Fields("NoOfOptions")
	SubmitID = rs.Fields("SubmitID")
	SubmitByName = rs.Fields("SubmitByName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	NegotiationLocked = rs.Fields("NegotiationLocked")
Else
	Response.Write("Error retrieving grant application record")
	Response.End
End If
IF GrantClassID = 1 And FiscalYear<2022 Then
	AppURL = "Application.asp"
Else
	AppURL = getHomeNegotiationReferenceByGrantClass(GrantClassID, AppID)
End If

' Start rounding dollar amounts as of 2020.
If FiscalYear>=2020 Then
	RoundCurrency = True
Else
	RoundCurrency = False
End If

If GranteeID>0 Then
	If NegotiationLocked = True Then
		PermitEdit = False
	ElseIf IsNull(SubmitID) = True Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
	ElseIf IsNull(SubmitID) = False Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, True)
	Else
		PermitEdit = False
	End If
Else
	PermitEdit = False
End If

If BudgetCategoryID=1 Or BudgetCategoryID=2 Or BudgetCategoryID=3 Or BudgetCategoryID=4 Or BudgetCategoryID=5 Then
	REDIM Options(NoOfOptions)
	sql = "SELECT SubCategoryID, SubCategory " & vbCrLF & _
		"FROM Lookup.BudgetSubcategories " & vbCrLF & _
		"WHERE BudgetCategoryID=" & prepIntegerSQL(BudgetCategoryID) & " " & _
		"ORDER BY SubCategoryID "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs=Con.Execute(sql)
	while rs.EOF = False
		Options(rs.Fields("SubcategoryID")) = rs.Fields("SubCategory")
		rs.MoveNext()
	Wend
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Grant Application Negotiation for <%=GranteeName %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function chosenButton(buttonchoice)
	{
		document.BudgetDetail.ButtonChoice.value=buttonchoice;
		if (validateForm() == true)
			document.BudgetDetail.submit();
	}
	function validateForm()
	{
		//alert("validate!");
		for (i = 1; i <= document.BudgetDetail.RowCount.value; i++) {
			if (document.BudgetDetail["Description_" + i].value.length>0)
			{
				document.BudgetDetail["Description_" + i].value = replaceWordChars(document.BudgetDetail["Description_" + i].value);
				if (document.BudgetDetail["MVCPAFunds_" + i].value.length==0 && document.BudgetDetail["InKindMatch_" + i].value.length==0)
				{
					alert("If you have a description for a row, you should also have a MVCPA Funds value.");
					document.BudgetDetail["MVCPAFunds_" + i].focus();
					return false;
				}
			}
			if (document.BudgetDetail["LineTotal_" + i].value.length>0)
			{
				if (document.BudgetDetail["Description_" + i].value==0)
				{
					alert("If you have a total for a row, you should also have a description.");
					document.BudgetDetail["Description_" + i].focus();
					return false;
				}
			}
			<% If BudgetCategoryID=1 Then %>
				if (getNumericValue(document.BudgetDetail["PctTime_" + i].value) > 100) {
					alert("Percent values must not be greater than 100");
					document.BudgetDetail["PctTime_" + i].focus();
					return false;
				}
			<%	End If
			If BudgetCategoryID=0  and True=False Then %>
					if (getNumericValue(document.BudgetDetail["PctSalary_" + i].value) > 100) {
						alert("Percent values must not be greater than 100");
						document.BudgetDetail["PctSalary_" + i].focus();
						return false;
					}
			<%	End If
			If BudgetCategoryID=1 Or BudgetCategoryID=2 Or BudgetCategoryID=3 Or BudgetCategoryID=4 Or BudgetCategoryID=5 Then %>
					if (document.BudgetDetail["Description_" + i].value.length>0)
			{
				if (document.BudgetDetail["SubCategoryID_" + i].selectedIndex==0)
				{
					alert("You must select a Subcategory");
					document.BudgetDetail["SubCategoryID_" + i].focus();
					return false;
				}
			}
			<%	End If 
			If BudgetCategoryID=4 And FiscalYear > 2021 Then %>
					if (document.BudgetDetail["SubCategoryID_" + i].value==4 || document.BudgetDetail["SubCategoryID_" + i].value==5 || document.BudgetDetail["SubCategoryID_" + i].value==6) 
			{
				if (getNumericValue(document.BudgetDetail["LineTotal_" + i].value) > 0.0) 
				{
					alert("DPS salaries, fringe, and overtime cannot be used as a cash match beginning in FY2022.");
					document.BudgetDetail["LineTotal_" + i].focus();
					return false;
				}
			}
			<%	End If %>
		}
		document.BudgetDetail.Narrative.value = replaceWordChars(document.BudgetDetail.Narrative.value);
		return true;
	}

	function clearValues(row){
		document.BudgetDetail["Description_" + row].value = "";
		document.BudgetDetail["MVCPAFunds_" + row].value = "";
		document.BudgetDetail["CashMatch_" + row].value = "";
		document.BudgetDetail["LineTotal_" + row].value = "";
		document.BudgetDetail["InKindMatch_" + row].value = "";
		<%	If BudgetCategoryID=1 Or BudgetCategoryID=2 Or BudgetCategoryID=3 Or BudgetCategoryID=4 Or BudgetCategoryID=5 Then %>
				document.BudgetDetail["SubCategoryID_" + row].selectedIndex=0;
		<%	End If 
		If BudgetCategoryID=1 Then %>
		document.BudgetDetail["PctTime_" + row].value="";
		<%	End If 
		If (BudgetCategoryID=1 or BudgetCategoryID=2 Or BudgetCategoryID=3 Or BudgetCategoryID=4) And True=False Then %>
	document.BudgetDetail["PctSalary_" + row].value="100";
		<%	End If %>
				updateTotals();
	}

	function changedCurrencyField(field)
	{
		if (checkCurrencyRound(field, <%=Lcase(CStr(RoundCurrency))%>) == false)
			return false;
		updateTotals();
		return true;
	}

	function updateTotals()
	{
		//alert("Update Totals");

		var grandtotal = 0;
		var MVCPAFunds_Total = 0;
		var CashMatch_Total = 0;
		var InKindMatch_Total = 0;
		for (i = 1; i <= document.BudgetDetail.RowCount.value; i++)
		{
			var rowtotal = 0;
			rowtotal = rowtotal + getNumericValue(document.BudgetDetail["MVCPAFunds_"+i].value);
			rowtotal = rowtotal + getNumericValue(document.BudgetDetail["CashMatch_" + i].value);
			MVCPAFunds_Total = MVCPAFunds_Total + getNumericValue(document.BudgetDetail["MVCPAFunds_" + i].value);
			CashMatch_Total = CashMatch_Total + getNumericValue(document.BudgetDetail["CashMatch_" + i].value);
			InKindMatch_Total = InKindMatch_Total + getNumericValue(document.BudgetDetail["InKindMatch_" + i].value);
			if (rowtotal==0)
				document.BudgetDetail["LineTotal_" + i].value = "";
			else
			{
				document.BudgetDetail["LineTotal_" + i].value = currencyRound(rowtotal, <%=LCase(CStr(RoundCurrency))%>);
				if (document.BudgetDetail["MVCPAFunds_"+i].value.length==0)
					document.BudgetDetail["MVCPAFunds_"+i].value = "$0";
				if (document.BudgetDetail["CashMatch_"+i].value.length==0)
					document.BudgetDetail["CashMatch_"+i].value = "$0";
			}
			grandtotal = grandtotal + rowtotal;
		}
		document.BudgetDetail.MVCPAFunds_Total.value = currencyRound(MVCPAFunds_Total, <%=LCase(CStr(RoundCurrency))%>);
		document.BudgetDetail.CashMatch_Total.value = currencyRound(CashMatch_Total, <%=LCase(CStr(RoundCurrency))%>);
		document.BudgetDetail.InKindMatch_Total.value = currencyRound(InKindMatch_Total, <%=LCase(CStr(RoundCurrency))%>);
		document.BudgetDetail.Total.value = currencyRound(grandtotal, <%=LCase(CStr(RoundCurrency))%>);
		//alert(grandtotal);
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body style="width: 100%" onload="updateTotals()">
<h1><%=GranteeName %> MVCPA Grant Application Negotiation for Fiscal Year <%=FiscalYear %></h1>
<h2><%=BudgetCategory %> Budget Detail</h2>
<%	If SubmitID>0 Then %>
<p style="text-align: center; font-weight: bold; ">The Application was submitted by <%=SubmitByName%> at <%=SubmitTimestamp %> and is now locked.</p>
<%	End If %>
<form name="BudgetDetail" method="post" action="BudgetDetailSubmit.asp" onsubmit="return validateForm()">
<%
Response.Write(HiddenField("AppID", AppID))
Response.Write(HiddenField("GrantClassID", GrantClassID))
Response.Write(HiddenField("GranteeID", GranteeID))
Response.Write(HiddenField("FiscalYear", FiscalYear))
Response.Write(HiddenField("BudgetCategoryID", BudgetCategoryID))
Response.Write(HiddenField("ButtonChoice", "save"))
If IsNull(CategoryDescription) = False Then
	Response.Write("<p style=""margin: auto; width: 75%; text-align: left"">" & CategoryDescription & "</p>" & vbCrLf)
	Response.Write("<br />" & vbCrLf)
End If

sql = "SELECT A.BudgetItemID, A.AppID, A.BudgetCategoryID, A.Description, A.NoOfItems, A.SubCategoryID, " & vbCrLF & _
	"	B.SubCategory, A.PctTime, A.PctSalary, A.MVCPAFunds, A.CashMatch, A.InKindMatch, A.LineTotal, " & vbCrLF & _
	"	CAST(CASE WHEN ISNULL(A.UnallowedItem,0)=1 THEN 1 WHEN AllowedAmount IS NOT NULL THEN 1 WHEN ISNULL(issue,0)=1 THEN 1 ELSE 0 END AS BIT) AS UnallowedItem " & vbCrLF & _
	"FROM Negotiation.BudgetDetails AS A" & vbCrLF & _
	"LEFT JOIN Lookup.BudgetSubcategories AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND B.SubCategoryID=A.SubCategoryID " & vbCrLF & _
	"WHERE A.AppID=" & prepIntegerSQL(AppID) & " AND A.BudgetCategoryID=" & prepIntegerSQL(BudgetCategoryID) & " " & vbCrLf & _
	"ORDER BY A.BudgetCategoryID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If


Response.Write("<table style=""margin: auto;"">" & vbCrLf)
Response.Write("<tr style=""vertical-align: bottom"">" & vbCrLf)
If BudgetCategoryID=1 or BudgetCategoryID=2 Or BudgetCategoryID=3 Then
	Response.Write(vbTab & "<th>Title or Position (No names)</th>" & vbCrLf)
ElseIf BudgetCategoryID=4 Then 
	Response.Write(vbTab & "<th>Description of Professional<br />or Contracted Services</th>" & vbCrLf)
ElseIf BudgetCategoryID=5 Then 
	Response.Write(vbTab & "<th>Description of Travel</th>" & vbCrLf)
ElseIf BudgetCategoryID=6 Then 
	Response.Write(vbTab & "<th>Equipment Name or Description and Quantity<br />(Do not list brand names)</th>" & vbCrLf)
ElseIf BudgetCategoryID=7 Then 
	Response.Write(vbTab & "<th>Description of Supplies and Other<br />Operating Expenses</th>" & vbCrLf)
Else
	Response.Write(vbTab & "<th>Description</th>" & vbCrLf)
End If
If BudgetCategoryID = 7 Then
	Response.Write(vbTab & "<th>No. Of Items</th>" & vbCrLf)
End If
If BudgetCategoryID=1 Or BudgetCategoryID=2 Or BudgetCategoryID=3 Or BudgetCategoryID=4 Or BudgetCategoryID=5 Then
	Response.Write(vbTab & "<th>Subcategory</th>" & vbCrLf)
End If
If BudgetCategoryID=1 Or BudgetCategoryID=4 Then
	Response.Write(vbTab & "<th>Percent<br />Time<br />Spent on<br />Grant<br />Activity</th>" & vbCrLf)
End If
If BudgetCategoryID=0 Then
	Response.Write(vbTab & "<th>Percent<br />Total<br />Salary<br />Paid by<br />MVCPA</th>" & vbCrLf)
ElseIf BudgetCategoryID=0 Then
	Response.Write(vbTab & "<th>Percent<br />Total<br />Fringe<br />Paid by<br />MVCPA</th>" & vbCrLf)
End If
Response.Write(vbTab & "<th>MVCPA Funds</th>" & vbCrLf)
Response.Write(vbTab & "<th>Cash Match</th>" & vbCrLf)
Response.Write(vbTab & "<th>Total</th>" & vbCrLf)
Response.Write(vbTab & "<th>In-Kind Match</th>" & vbCrLf)
Response.Write("</tr>" & vbCrLf)

Set rs=Con.Execute(sql)
i=0
while rs.EOF = False
	i = i + 1
	WriteBudgetRow rs.Fields("BudgetItemID"), rs.Fields("BudgetCategoryID"), i, rs.Fields("Description"), rs.Fields("NoOfItems"), rs.Fields("SubCategoryID"), rs.Fields("PctTime"), rs.Fields("PctSalary"), rs.Fields("MVCPAFunds"), rs.Fields("CashMatch"), rs.Fields("InKindMatch"), rs.Fields("LineTotal"), rs.Fields("UnallowedItem"), PermitEdit
	rs.MoveNext
Wend

For j = 1 to 5
	i = i + 1
	WriteBudgetRow 0, BudgetCategoryID, i, "", "", "", "100", "", "", "", "", 0, False, PermitEdit
Next

' Total Row
Response.Write("<tfoot><tr>" & vbCrLf)
Response.Write(WriteInCell("") & vbCrLf) ' Description
If BudgetCategoryID=7 Then
	Response.Write(WriteInCell("") & vbCrLf) ' No. Of Items
End If
If BudgetCategoryID=1 Or BudgetCategoryID=2 Or BudgetCategoryID=3 Or BudgetCategoryID=4 Or BudgetCategoryID=5 Then
	Response.Write(WriteInCell("") & vbCrLf)
End If	
If BudgetCategoryID=1 Or BudgetCategoryID=4 Then
	Response.Write(WriteInCell("") & vbCrLf)
End If	
If BudgetCategoryID=0 Then
	Response.Write(WriteInCell("") & vbCrLf)
End If	
Response.Write(WriteInCell("<input type=text name=""MVCPAFunds_Total"" value="""" size=""12"" maxLength=""14"" style=""text-align:right;border-style:none"" readonly=""readonly"" />") & vbCrLf)	
Response.Write(WriteInCell("<input type=text name=""CashMatch_Total"" value="""" size=""12"" maxLength=""14"" style=""text-align:right;border-style:none"" readonly=""readonly"" />") & vbCrLf)	
Response.Write(WriteInCell("<input type=text name=""Total"" value="""" size=""12"" maxLength=""14"" style=""text-align:right;border-style:none"" readonly=""readonly"" />") & vbCrLf)	
Response.Write(WriteInCell("<input type=text name=""InKindMatch_Total"" value="""" size=""12"" maxLength=""14"" style=""text-align:right;border-style:none"" readonly=""readonly"" />") & vbCrLf)	
Response.Write("</tr></tfoot>" & vbCrLf)

Response.Write("</table>" & vbCrLF)
Response.Write(HiddenField("RowCount",i))
%>
	<br />
	<div style="text-align: center">Narrative: <%= TextArea("Narrative", Narrative, 15, 120, 20000, PermitEdit, "") %></div>
<br />
	<div style="text-align: center">
<%	If PermitEdit = True Then %>
	<input type="button" value="Save" onclick="chosenButton('save');" title="Click to save and return to this page." />
	<input type="button" value="Done" onclick="chosenButton('done');" title="Click to save and return to the application." />
	<input type="button" value="Next Category" onclick="chosenButton('next');" title="Click to save and return to the application." />
	<input type="button" value="Cancel" onclick="location.href='<%=AppURL%>';" 
		title="Return to the application without saving latest changes." />
<%	Else %>
	<input type="button" value="Return" onclick="location.href='<%=AppURL%>';" 
		title="Return to the application main page." />
<%	End If %>
  </div>

</form>
<p style="text-align:center; font-size: small;">Each time you save form, more blank rows will be added if needed. Click on save if you have more items to add. To delete a row, erase all of the data in that row and save it.</p>
</body>
</html><%
Sub WriteBudgetRow(vBudgetItemID, vBudgetCategoryID, i, vTitle, vNoOfItems, vSubCategoryID, vPctTime, vPctSalary, vMVCPAFunds, vCashMatch, vInKindMatch, vLineTotal, vAddColor, vPermitEdit)
	Dim vColor
	If vAddColor = True Then
		Response.Write("<tr style=""background-color: Yellow; "">" & vbCrLf)
		vColor = "Yellow"
	Else
		Response.Write("<tr>" & vbCrLf)
		vColor = ""
	End If
	Response.Write(WriteInCell(TextFieldColor("Description_" & i, vTitle, 38, 255, vPermitEdit, vColor, "")))
	If vBudgetCategoryID=7 Then
		Response.Write(WriteInCell(IntegerFieldColor("NoOfItems_" & i, vNoOfItems, 2, 3, vPermitEdit, vColor, "return checkInteger(this);")))
	End If
	If vBudgetCategoryID=1 Or vBudgetCategoryID=2 Or vBudgetCategoryID=3 Or vBudgetCategoryID=4 Or vBudgetCategoryID=5 Then
		Response.Write(vbTab & "<td style=""text-align: center""><select name=""SubCategoryID_" & i & """>")
		Response.Write(vbTab & vbTab & "<option value=""0""></option>" & vbCrLf)
		For k = 1 to UBound(Options)
			If Len(Options(k))>0 Then
				If k = vSubCategoryID Then
					Response.Write(vbTab & vbTab & "<option value=""" & k & """ selected=""selected"">" & Options(k) & "</option>" & vbCrLf)
				Else
					Response.Write(vbTab & vbTab & "<option value=""" & k & """>" & Options(k) & "</option>" & vbCrLf)
				End If
			End If
		Next
		Response.Write("</select></td>" & vbCrLf)
	End If	
	If BudgetCategoryID=1 Or BudgetCategoryID=4 Then
		Response.Write(WriteInCell(IntegerFieldColor("PctTime_" & i, vPctTIme, 2, 3, vPermitEdit, vColor, "return checkInteger(this);")&"%"))
	End If
	If vBudgetCategoryID=0 Then
		Response.Write(WriteInCell(IntegerFieldColor("PctSalary_" & i, vPctSalary, 2, 3, vPermitEdit, vColor, "return checkInteger(this);")&"%"))
	End If
	Response.Write(WriteInCell(CurrencyFieldRoundColor("MVCPAFunds_" & i, vMVCPAFunds, 10, 11, RoundCurrency, vPermitEdit, vColor, "changedCurrencyField(this);")))
	Response.Write(WriteInCell(CurrencyFieldRoundColor("CashMatch_" & i, vCashMatch, 10, 11, RoundCurrency, vPermitEdit, vColor, "changedCurrencyField(this);")))
	Response.Write(WriteInCell(CurrencyFieldRoundColor("LineTotal_" & i, vLineTotal, 10, 11, RoundCurrency, False, vColor, "") & vbCrLf _
		& HiddenField("BudgetItemID_" & i, vBudgetItemID)))
	Response.Write(WriteInCell(CurrencyFieldRoundColor("InKindMatch_" & i, vInKindMatch, 10, 11, RoundCurrency, vPermitEdit, vColor, "changedCurrencyField(this);")))
	If PermitEdit = True Then
		Response.Write(WriteInCell("<img style=""border: none"" src=""../images/delete.gif"" onclick=""clearValues(" & i & ")"" title=""Clear values and delete record on next save"" />"))
	End If
	Response.Write("</tr>" & vbCrLf)
End Sub
%>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/HomeRef.asp"-->