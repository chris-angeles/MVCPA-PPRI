<%
function prepStringSQL(StringValue)
	If IsNull(StringValue) = True Then
		prepStringSQL = "null"
	else
		prepStringSQL = Trim(StringValue)
		If len(prepStringSQL) = 0 Then 
			prepStringSQL = "null"
		else
			prepStringSQL = Replace(prepStringSQL,"'","''")
			prepStringSQL = "'" & prepStringSQL & "'"
		End If
	End If
end function

function prepUnicodeSQL(StringValue)
	If IsNull(StringValue) = True Then
		prepUnicodeSQL = "null"
	else
		prepUnicodeSQL = Trim(StringValue)
		If len(prepUnicodeSQL) = 0 Then 
			prepUnicodeSQL = "null"
		else
			prepUnicodeSQL = Replace(prepUnicodeSQL,"'","''")
			prepUnicodeSQL = Replace(prepUnicodeSQL,"""","""""")
			prepUnicodeSQL = "N'" & prepUnicodeSQL & "'"
		End If
	End If	
end function

function prepDateSQL(StringValue)
	If IsNull(StringValue) = True Then
		prepDateSQL = "null"
	ElseIf StringValue="" Then
		prepDateSQL = "null"
	ElseIf isDate(StringValue) = false Then
		prepDateSQL = "null"
	else
		prepDateSQL = "'" & CDate(StringValue) & "'"
	End If
end function

function cleanMS(StringValue)
	if IsNull(StringValue) = True Then
		cleanMS = null
	elseif len(StringValue)=0 Then
		cleanMS = ""
	else
		cleanMS = Replace(StringValue, Chr(226)&Chr(128)&Chr(147),"-")  ' en-dash
		cleanMS = Replace(cleanMS, Chr(226)&Chr(128)&Chr(148),"-")      ' em-dash
		cleanMS = Replace(cleanMS, Chr(226)&Chr(128)&Chr(152),"'")      ' left curly apostrophe
		cleanMS = Replace(cleanMS, Chr(226)&Chr(128)&Chr(153),"'")      ' right curly apostrophe
		cleanMS = Replace(cleanMS, Chr(226)&Chr(128)&Chr(156),"""")     ' left curly quote
		cleanMS = Replace(cleanMS, Chr(226)&Chr(128)&Chr(157),"""")     ' right curly quote
		'cleanMS = Replace(cleanMS, Chr(226)&Chr(128)&Chr(162),"&bull;") ' bullet
		'cleanMS = Replace(cleanMS, Chr(194)&Chr(167),"&sect;")          ' legal section symbol
	end if
end function

function prepIntegerSQL(NumericValue)
	If IsNull(NumericValue) = True Then
		prepIntegerSQL = "null"
	ElseIf NumericValue = "" Then
		prepIntegerSQL = "null"
	ElseIf IsNumeric(NumericValue) Then
		prepIntegerSQL = CStr(CLng(NumericValue))
	Else
		prepIntegerSQL = "null"
	End If
End Function

function prepNumberSQL(NumericValue)
	If IsNull(NumericValue) = True Then
		prepNumberSQL = "null"
	ElseIf NumericValue = "" Then
		prepNumberSQL = "null"
	ElseIf IsNumeric(Replace(NumericValue,"%","",1,1)) Then
		prepNumberSQL = CDbl(Replace(NumericValue,"%","",1,1))
	Else
		prepNumberSQL = "null"
	End If
End Function

function prepBitSQL(BitValue)
	If IsEmpty(BitValue) = True Then
		prepBitSQL = "null"
	ElseIf BitValue="Y" Then
		prepBitSQL = 1
	ElseIf BitValue="N" Then
		prepBitSQL = 0
	ElseIf BitValue="1" Then
		prepBitSQL = 1
	ElseIf BitValue = "0" Then
		prepBitSQL = 0
	ElseIf BitValue = True Then
		prepBitSQL = 1
	ElseIf isNumeric(BitValue) Then
		If BitValue=1 Then
			prepBitSQL = 1
		ElseIf BitValue = 0 Then
			prepBitSQL = 0
		Else
			prepBitSQL = "null"
		End If
	Else
		prepBitSQL = "null"
	End If
End Function

function prepBitRequiredSQL(BitValue)
	If IsEmpty(BitValue) = True Then
		prepBitRequiredSQL = 0
	ElseIf BitValue="Y" Then
		prepBitRequiredSQL = 1
	ElseIf BitValue="N" Then
		prepBitRequiredSQL = 0
	ElseIf BitValue="1" Then
		prepBitRequiredSQL = 1
	ElseIf BitValue = "0" Then
		prepBitRequiredSQL = 0
	ElseIf BitValue = True Then
		prepBitRequiredSQL = 1
	ElseIf isNumeric(BitValue) Then
		If BitValue=1 Then
			prepBitRequiredSQL = 1
		ElseIf BitValue = 0 Then
			prepBitRequiredSQL = 0
		Else
			prepBitRequiredSQL = 0
		End If
	Else
		prepBitRequiredSQL = 0
	End If
End Function

Function checkEmail(sEmail)
  checkEmail = false
  Dim regEx, retVal
  Set regEx = New RegExp

  ' Create regular expression:
  'regEx.Pattern ="^[\w-\.+#$]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"		
  regEx.Pattern ="^[\w-\.+#$%'_~^|]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"		

  ' Set pattern:
  regEx.IgnoreCase = true

  ' Set case sensitivity.
  retVal = regEx.Test(sEmail)

  ' Execute the search test.
  If not retVal Then
    exit function
  End If

  checkEmail = true
End Function
%>