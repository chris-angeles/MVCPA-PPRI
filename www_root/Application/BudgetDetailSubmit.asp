<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, AppID, GrantClassID, BudgetCategoryID, RowCount, ButtonChoice,  _
	BudgetItemID, Description, NoOfItems, SubCategoryID, PctTime, PctSalary, MVCPAFunds, CashMatch, _
	InKindMatch, LineTotal, UpdateID, UpdateTimestamp, Narrative, NextCategory, _
	BudgetCashMatch, FiscalYear, AppURL
ReDim ProgramCategory(5)
TimeStamp = Now()

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

AppID = Request.Form("AppID")
GrantClassID = CInt(Request.Form("GrantClassID"))
FiscalYear = CInt(Request.Form("FiscalYear"))
BudgetCategoryID = Request.Form("BudgetCategoryID")
BudgetCashMatch = Request.Form("BudgetCashMatch")
RowCount = Request.Form("RowCount")
ButtonChoice = Request.Form("ButtonChoice")
Narrative = Request.Form("Narrative")
UpdateID = UserSystemID
UpdateTimestamp = Timestamp

IF GrantClassID = 1 And FiscalYear<2022 Then
	AppURL = "Application.asp"
Else
	AppURL = getHomeApplicationReferenceByGrantClass(GrantClassID, AppID)
End If
If Debug = True Then
	Response.Write("<pre>AppURL='" & AppURL & "'</pre>" & vbCrLf)
	Response.Flush
End If

For i = 1 to RowCount
	BudgetItemID = Request.Form("BudgetItemID_" & i)
	Description = Request.Form("Description_" & i)
	NoOfItems = Request.Form("NoOfItems_" & i)
	SubCategoryID = Request.Form("SubCategoryID_" & i)
	PctTime = Request.Form("PctTime_" & i)
	PctSalary = Request.Form("PctSalary_" & i)
	MVCPAFunds = Request.Form("MVCPAFunds_" & i)
	CashMatch = Request.Form("CashMatch_" & i)
	InKindMatch = Request.Form("InKindMatch_" & i)
	LineTotal = Request.Form("LineTotal_" & i)
	If Len(MVCPAFunds)>0 Or Len(CashMatch)>0 Or Len(Description)>0 Or Len(InKindMatch)>0 Then
		If CInt(BudgetItemID) = 0 Then ' Insert
			sql  = "INSERT INTO Application.BudgetDetails (AppID, BudgetCategoryID, Description, NoOfItems, " & vbCrLf & _
				"	SubCategoryID, PctTime, PctSalary, MVCPAFunds, CashMatch, InKindMatch, " & vbCrLF & _
				"	UpdateID, UpdateTimestamp)" & vbCrLf & _
				"VALUES (" & _
				prepIntegerSQL(AppID) & ", " & _
				prepIntegerSQL(BudgetCategoryID) & ", " & _
				prepStringSQL(Description) & ", " & _
				prepIntegerSQL(NoOfItems) & ", " & _
				prepIntegerSQL(SubCategoryID) & ", " & _
				prepIntegerSQL(PctTIme) & ", " & _
				prepIntegerSQL(PctSalary) & ", " & _
				prepNumberSQL(MVCPAFunds) & ", " & _
				prepNumberSQL(CashMatch) & ", " & _
				prepNumberSQL(InKindMatch) & ", " & _
				prepIntegerSQL(UpdateID) & ", " & _
				prepStringSQL(TimeStamp) & ")"
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		Else ' Update
			sql = "UPDATE Application.BudgetDetails SET " & vbCrLF & _
				"AppID=" & prepIntegerSQL(AppID) & ", " & _
				"BudgetCategoryID=" & prepIntegerSQL(BudgetCategoryID) & ", " & _
				"Description=" & prepStringSQL(Description) & ", " & _
				"NoOfItems=" & prepIntegerSQL(NoOfItems) & ", " & _
				"SubCategoryID=" & prepIntegerSQL(SubCategoryID) & ", " & _
				"PctTIme=" & prepIntegerSQL(PctTIme) & ", " & _
				"PctSalary=" & prepIntegerSQL(PctSalary) & ", " & _
				"MVCPAFunds=" & prepNumberSQL(MVCPAFunds) & ", " & _
				"CashMatch=" & prepNumberSQL(CashMatch) & ", " & _
				"InKindMatch=" & prepNumberSQL(InKindMatch) & ", " & _
				"UpdateID=" & prepIntegerSQL(UpdateID) & ", " & _
				"UpdateTimeStamp=" & prepStringSQL(TimeStamp) & " " & _
				"WHERE BudgetItemID=" & prepIntegerSQL(BudgetItemID)
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		End If
	ElseIf Cint(BudgetItemID)>0 Then ' Delete
		sql = "DELETE FROM Application.BudgetDetails WHERE BudgetItemID=" & prepIntegerSQL(BudgetItemID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
Next

' Update Narrative
If Len(Narrative)>0 Then
	sql = "UPDATE Application.BudgetDetailNarrative SET " & vbCrLF & _
		"Narrative=" & prepStringSQL(Narrative) & ", UpdateID=" & prepIntegerSQL(UserSystemID) & _
		", UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLF & _
		"WHERE AppID=" & prepIntegerSQL(AppID) & " AND BudgetCategoryID=" & BudgetCategoryID & " " & vbCrLF & _
		"IF @@ROWCOUNT=0 " & vbCrLF & _
		"INSERT INTO Application.BudgetDetailNarrative (BudgetCategoryID, AppID, Narrative, UpdateID, UpdateTimestamp) " & vbCrLF& _
		"VALUES (" & prepIntegerSQL(BudgetCategoryID) & ", " & prepIntegerSQL(AppID) & ", " & _
		prepStringSQL(Narrative) & ", " & _
		prepIntegerSQL(UserSystemID) & ", " & prepStringSQL(Timestamp) & ")"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs=Con.Execute(sql)
Else
	sql = "DELETE FROM Application.BudgetDetailNarrative WHERE AppID=" & prepIntegerSQL(AppID) & " AND BudgetCategoryID=" & BudgetCategoryID
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs=Con.Execute(sql)
End If

If Debug = True Then
Response.Write("<pre>ButtonChoice=" & ButtonChoice & "</pre>" & vbCrLf)
End If

If ButtonChoice = "save" Then
	If Len(BudgetCashMatch)>0 Then
		If Debug = True Then
			Response.Write("<a href=""BudgetDetail2.asp?AppID=" & AppID & "&BudgetCategoryID=" & BudgetCategoryID & _
			""">return</a>")
		Else
			Response.Redirect("BudgetDetail2.asp?AppID=" & AppID & "&BudgetCategoryID=" & BudgetCategoryID)
		End If	
	Else
		If Debug = True Then
			Response.Write("<a href=""BudgetDetail.asp?AppID=" & AppID & "&BudgetCategoryID=" & BudgetCategoryID & _
			""">return</a>")
		Else
			Response.Redirect("BudgetDetail.asp?AppID=" & AppID & "&BudgetCategoryID=" & BudgetCategoryID)
		End If
	End If
ElseIf ButtonChoice = "next" Then
	If Len(BudgetCashMatch)>0 Then
		BudgetCategoryID = BudgetCategoryID + 1
		If BudgetCategoryID=8 Then BudgetCategoryID=1
		If Debug = True Then
			Response.Write("<a href=""BudgetDetail2.asp?AppID=" & AppID & "&BudgetCategoryID=" & BudgetCategoryID & _
			""">return</a>")
		Else
			Response.Redirect("BudgetDetail2.asp?AppID=" & AppID & "&BudgetCategoryID=" & BudgetCategoryID)
		End If
	Else
		BudgetCategoryID = BudgetCategoryID + 1
		If BudgetCategoryID=8 Then BudgetCategoryID=1
		If Debug = True Then
			Response.Write("<a href=""BudgetDetail.asp?AppID=" & AppID & "&BudgetCategoryID=" & BudgetCategoryID & _
			""">return</a>")
		Else
			Response.Redirect("BudgetDetail.asp?AppID=" & AppID & "&BudgetCategoryID=" & BudgetCategoryID)
		End If
	End If
Else
	If Debug = True Then
		Response.Write("<a href=""" & AppURL & """>return</a>")
	Else
		Response.Redirect(AppURL)
	End If
End If
%><!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/HomeRef.asp"-->

