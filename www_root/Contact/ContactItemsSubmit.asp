<%@ Language=VBScript %><% Option Explicit %>
<!--#include file="../includes/adovbs.asp"--> 
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><%

Dim Debug, ContactPhoneCallID, CallDateTime, CallLength, PhoneNumber, ContactID, ContactName, _
	ContactTitleID,	GranteeID, Organization, Questions, MVCPAContactID, Answer, EMail, ContactIssueID, _
	ContactTypeID, Positive, Negative, DateComplete, LongText, i
Debug = False
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

If Request.Form.Count < 10 Then
	Response.Write("Error: No data posted")
	Response.End
End If

ContactPhoneCallID = CLng(Request.Form("ContactPhoneCallID"))
CallDateTime = Request.Form("CallDateTime")
CallLength = Request.Form("CallLength")
PhoneNumber = Left(Request.Form("PhoneNumber"),24)
ContactID = Request.Form("ContactID")
If ContactID = 0 Then ContactID = null
ContactName = Request.Form("ContactName")
GranteeID = Request.Form("GranteeID")
ContactTitleID = Request.Form("ContactTitleID")
If ContactTitleID = 0 Then ContactTitleID = null
Organization = Request.Form("Organization")
Questions = Request.Form("Questions")
MVCPAContactID = Request.Form("MVCPAContactID")
Answer = Request.Form("Answer")
EMail = Request.Form("EMail")
ContactIssueID = Request.Form("ContactIssueID")
If ContactIssueID = 0 Then ContactIssueID = null
ContactTypeID = Request.Form("ContactTypeID")
If ContactTypeID = 0 Then ContactTypeID = null
DateComplete = Request.Form("DateComplete")
If Request.Form("Positive") = 1 Then
	Positive = True
Else
	Positive = False
End If
If Request.Form("Negative") = 1 Then
	Negative = True
Else
	Negative = False
End If
LongText = Request.Form("LongText")

If ContactPhoneCallID=0 Then
	sql = "INSERT INTO Contact.PhoneCalls (CallDateTime, CallLength, PhoneNumber, ContactID, ContactName, ContactTitleID, GranteeID, Organization, Questions, MVCPAContactID, Answer, EMail, ContactIssueID, ContactTypeID, Positive, Negative, LongText, DateComplete, UpdateID, UpdateTimeStamp) VALUES " & vbCrLf & _
	"(" & prepStringSQL(CallDateTime) & ", " & vbCrLF & _
	prepIntegerSQL(CallLength) & ", " & vbCrLF & _ 
	prepStringSQL(PhoneNumber) & ", " & vbCrLF & _
	prepIntegerSQL(ContactID) & ", " & vbCrLF & _
	prepStringSQL(ContactName) & ", " & vbCrLF & _
	prepIntegerSQL(ContactTitleID) & ", " & vbCrLF & _
	prepIntegerSQL(GranteeID) & ", " & vbCrLF & _
	prepStringSQL(Organization) & ", " & vbCrLF & _
	prepStringSQL(Questions) & ", " & vbCrLF & _
	prepIntegerSQL(MVCPAContactID) & ", " & vbCrLF & _
	prepStringSQL(Answer) & ", " & vbCrLF & _
	prepStringSQL(EMail) & ", " & vbCrLF & _
	prepIntegerSQL(ContactIssueID) & ", " & vbCrLF & _
	prepIntegerSQL(ContactTypeID) & ", " & vbCrLF & _
	prepBitSQL(Positive) & ", " & vbCrLF & _
	prepBitSQL(Negative) & ", " & vbCrLF & _
	prepStringSQL(LongText) & ", " & vbCrLf & _
	prepStringSQL(DateComplete) & ", " & vbCrLF & _
	prepIntegerSQL(UserSystemID) & ", " & vbCrLF & _
	prepStringSQL(Now()) & ")"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
Else
	sql = "UPDATE Contact.PhoneCalls " & vbCrLf & _
		"SET " & vbCrLf & _ 
		"	CallDateTime=" & prepStringSQL(CallDateTime) & ", " & vbCrLf & _
		"	CallLength=" & prepIntegerSQL(CallLength) & ", " & vbCrLf & _
		"	PhoneNumber=" & prepStringSQL(PhoneNumber) & ", " & vbCrLf & _
		"	ContactID=" & prepIntegerSQL(ContactID) & ", " & vbCrLf & _
		"	ContactName=" & prepStringSQL(ContactName) & ", " & vbCrLf & _
		"	ContactTitleID=" & prepIntegerSQL(ContactTitleID) & ", " & vbCrLf & _
		"	GranteeID=" & prepIntegerSQL(GranteeID) & ", " & vbCrLf & _
		"	Organization=" & prepStringSQL(Organization) & ", " & vbCrLf & _
		"	Questions=" & prepStringSQL(Questions) & ", " & vbCrLf & _
		"	MVCPAContactID=" & prepIntegerSQL(MVCPAContactID) & ", " & vbCrLf & _
		"	Answer=" & prepStringSQL(Answer) & ", " & vbCrLf & _
		"	EMail=" & prepStringSQL(EMail) & ", " & vbCrLf & _
		"	ContactIssueID=" & prepIntegerSQL(ContactIssueID) & ", " & vbCrLf & _
		"	ContactTypeID=" & prepIntegerSQL(ContactTypeID) & ", " & vbCrLf & _
		"	Positive=" & prepBitSQL(Positive) & ", " & vbCrLf & _
		"	Negative=" & prepBitSQL(Negative) & ", " & vbCrLf & _
		"	LongText=" & prepStringSQL(LongText) & ", " & vbCrLf & _
		"	DateComplete=" & prepStringSQL(DateComplete) & ", " & vbCrLf & _
		"	UpdateID=" & prepStringSQL(UserSystemID) & ", " & vbCrLf & _
		"	UpdateTimeStamp=" & prepStringSQL(now()) & " " & vbCrLF & _
		"WHERE ContactPhoneCallID=" & prepIntegerSQL(ContactPhoneCallID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If


If ContactPhoneCallID = 0 Then
	sql = "select IDENT_CURRENT('Contact.PhoneCalls')"
	Set rs = con.Execute(sql)
	If rs.BOF = True Then
		Response.Write("Error retreiving ID")
		Response.End
	Else
		ContactPhoneCallID = rs.Fields(0)
	End If
End If

If Debug = True Then
	Response.Write("<a href=""ContactItems.asp?ContactPhoneCallID=" & ContactPhoneCallID & """>return</a>")
Else
	'Response.Redirect("ContactItemsSearch.asp")
	Response.Redirect("ContactItems.asp?ContactPhoneCallID=" & ContactPhoneCallID)
End If	

%>
<!--#include file="../includes/prepDB.asp"--> 
