<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i , GrantID, FiscalYear, QuestionID, ResponseTypeID, IntegerTarget, DecimalTarget, DatabaseTarget, Change
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

GrantID = CInt(Request.Form("GrantID"))
FiscalYear = CInt(Request.Form("FiscalYear"))
sql = "SELECT A.QuestionID, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.NoTarget, A.ResponseTypeID, " & vbCrLf & _
	"	IntegerTarget, DecimalTarget " & vbCrLf & _
	"FROM PR.Goals AS G " & vbCrLf & _
	"LEFT JOIN PR.Strategies AS S ON S.GoalID=G.GoalID " & vbCrLf & _
	"LEFT JOIN PR.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID " & vbCrLf & _
	"LEFT JOIN PR.GrantQuestions AS Q ON Q.GrantID=" & prepIntegerSQL(GrantID) & " AND Q.QuestionID=A.QuestionID " & vbCrLf & _
	"LEFT JOIN PR.Responses AS R ON R.GrantID=" & prepIntegerSQL(GrantID) & " AND R.QuestionID=A.QuestionID " & vbCrLf & _
	"WHERE Q.GrantID=" & prepIntegerSQL(GrantID) & " AND Mandatory=0 AND NoTarget=0 AND ResponseTypeID IN (1,2,3) " & vbCrLF & _
	"ORDER BY A.Mandatory DESC, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

While rs.EOF = False
	QuestionID = rs.Fields("QuestionID")
	ResponseTypeID = rs.Fields("ResponseTypeID")
	IntegerTarget = null
	DecimalTarget = null
	DatabaseTarget = null
	If ResponseTypeID = 1 Then
		IntegerTarget = Request.Form("IntegerTarget_" & QuestionID)
		DatabaseTarget = rs.Fields("IntegerTarget")
		If IntegerTarget = "" Then
			IntegerTarget = null
		End If
		If IsNull(IntegerTarget) = True And IsNull(DatabaseTarget) = True Then
			Change = False
		ElseIf IsNull(IntegerTarget) <> IsNull(DatabaseTarget) Then
			Change = True
		ElseIf Clng(IntegerTarget) = CLng(DatabaseTarget) Then
			Change = False
		Else
			Change = True
		End If
		If Debug = True Then 
			Response.Write("<pre>Request.Form(IntegerTarget_" & QuestionID & " = " & IntegerTarget & "; Database value = " & DatabaseTarget & "; Change=" & Change & ".</pre>")
		End If
		If Change = True Then
			sql = "UPDATE PR.GrantQuestions SET IntegerTarget=" & prepIntegerSQL(IntegerTarget) & " WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND QuestionID=" & prepIntegerSQL(QuestionID)
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			cON.Execute(sql)
		End If
	ElseIf ResponseTypeID = 1 Or ResponseTypeID = 3 Then
		DecimalTarget  = Request.Form("IntegerTarget_" & QuestionID)
		If DecimalTarget = "" Then
			DecimalTarget = null
		End If
		If IsNull(DecimalTarget) = True And IsNull(DatabaseTarget) = True Then
			Change = False
		ElseIf IsNull(DecimalTarget) <> IsNull(DatabaseTarget) Then
			Change = True
		ElseIf CDbl(DecimalTarget) = CDbl(DatabaseTarget) Then
			Change = False
		Else
			Change = True
		End If
		If Debug = True Then 
			Response.Write("<pre>Request.Form(DecimalTarget_" & QuestionID & " = " & DecimalTarget & "; Database value = " & DatabaseTarget & "; Change=" & Change & ".</pre>")
		End If
		If Change = True Then
			sql = "UPDATE PR.GrantQuestions SET DecimalTarget=" & prepIntegerSQL(DecimalTarget) & " WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND QuestionID=" & prepIntegerSQL(QuestionID)
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			cON.Execute(sql)
		End If
	End If
	rs.MoveNext
Wend
If Debug = True Then
	Response.Write("<a href=""Targets.asp?FiscalYear=" & FiscalYEar & "&GrantID=" & GrantID & """>Return</a>")
Else
	Response.Redirect("Targets.asp?FiscalYear=" & FiscalYEar & "&GrantID=" & GrantID)
End If

%><!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->