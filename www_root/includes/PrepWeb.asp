<%
function prepStringWeb(StringValue)
	If isnull(StringValue) = True Then 
		prepStringWeb = ""
	Else
		prepStringWeb = Trim(StringValue)
		If isnull(StringValue) = True Then 
			prepStringWeb = ""
		ElseIf Len(prepStringWeb) = 0 Then
			prepStringWeb = ""
		Else
			prepStringWeb = server.HTMLEncode(prepStringWeb)
		End If
	End If
end function

function prepIntegerWeb(IntegerValue)
	If isnull(IntegerValue) Then 
		prepIntegerWeb = ""
	ElseIf isnumeric(IntegerValue) = False Then
		prepIntegerWeb = ""
	Else
		prepIntegerWeb = CLng(IntegerValue)
	End If
end function

function prepIntegerWebNZ(IntegerValue)
	If isnull(IntegerValue) Then 
		prepIntegerWebNZ = ""
	ElseIf isnumeric(IntegerValue) = False Then
		prepIntegerWebNZ = ""
	ElseIf CLng(IntegerValue) = 0 Then
		prepIntegerWebNZ = ""
	Else
		prepIntegerWebNZ = CLng(IntegerValue)
	End If
end function

function prepCurrencyWeb(DecimalValue)
	If isnull(DecimalValue) Then 
		prepCurrencyWeb = ""
	Else
		prepCurrencyWeb = FormatCurrency(DecimalValue,2,True,True,True)
	End If
end function

function prepCurrencyWebRound(DecimalValue, vRound)
	If isnull(DecimalValue) Then 
		prepCurrencyWebRound = ""
	Else
		If isnumeric(DecimalValue) = False Then
			prepCurrencyWebRound = ""
		ElseIf vRound = True Then
			prepCurrencyWebRound = FormatCurrency(DecimalValue,0,True,True,True)
		Else
			prepCurrencyWebRound = FormatCurrency(DecimalValue,2,True,True,True)
		End If
	End If
end function

function prepCurrencyWebNZ(DecimalValue)
	If isnull(DecimalValue) Then 
		prepCurrencyWebNZ = ""
	Else
		If isnumeric(DecimalValue) = False Then
			prepCurrencyWebNZ = ""
		ElseIf DecimalValue=0 Then
			prepCurrencyWebNZ = ""
		Else
			prepCurrencyWebNZ = FormatCurrency(DecimalValue,2,True,True,True)
		End If
	End If
end function

function prepCurrencyWebNZRound(DecimalValue, vRound)
	If isnull(DecimalValue) Then 
		prepCurrencyWebNZRound = ""
	Else
		If isnumeric(DecimalValue) = False Then
			prepCurrencyWebNZRound = ""
		ElseIf DecimalValue=0 Then
			prepCurrencyWebNZRound = ""
		ElseIf vRound = True Then 
			prepCurrencyWebNZRound = FormatCurrency(DecimalValue,0,True,True,True)
		Else
			prepCurrencyWebNZRound = FormatCurrency(DecimalValue,2,True,True,True)
		End If
	End If
end function

function prepCurrencyWebNoNull(DecimalValue)
	If isnull(DecimalValue) Then 
		prepCurrencyWebNoNull = "$0.00"
	Else
		If isnumeric(DecimalValue) = False Then
			prepCurrencyWebNoNull = "$0.00"
		ElseIf DecimalValue=0 Then
			prepCurrencyWebNoNull = "$0.00"
		Else
			prepCurrencyWebNoNull = FormatCurrency(DecimalValue,2,True,True,True)
		End If
	End If
end function

function prepNumberWeb(NumberValue, DecimalDigits)
	NumberValue = Trim(NumberValue)
	If isnull(NumberValue) Then 
		prepNumberWeb = ""
	Else
		If IsNumeric(NumberValue) = False Then
			prepNumberWeb = ""
		Else
			prepNumberWeb = FormatNumber(NumberValue, DecimalDigits,True,False,True)
		End If
	End If
end function

function prepBitWise(NumberValue)
	If IsNull(NumberValue) = True Then
		prepBitwise = 0
	Else
		prepBitwise = CInt(NumberValue)
	End If
end function

function prepNumberWebNZ(NumberValue, DecimalDigits)
	NumberValue = Trim(NumberValue)
	If isnull(NumberValue) Then 
		prepNumberWebNZ = ""
	ElseIf IsNumeric(NumberValue) = True Then
		If NumberValue = 0 Then
			prepNumberWebNZ = ""
		ElseIf NumberValue = 0.0 Then
			prepNumberWebNZ = ""		
		Else
			prepNumberWebNZ = FormatNumber(NumberValue, DecimalDigits,True,False,True)
		End If
	Else
			prepNumberWebNZ = ""
	End If
end function

Function formatInteger(value)
	If value="" then
		formatInteger = ""
	Elseif IsNull(value) = True then
		formatInteger = ""
	ElseIf IsNumeric(CStr(value)) = false then
		formatInteger = ""
	Else
		formatInteger = formatnumber(value,0,true,true,true)
	End If
End Function

function formatCurrencyRound(value, vRound)
	If vRound = True Then
		formatCurrencyRound = FormatCurrency(value, 0, True, True, True)
	Else
		formatCurrencyRound = FormatCurrency(value, 2, True, True, True)
	End If
End Function

Function formatDecimal(value, digits)
	If value="" then
		formatDecimal = ""
	Elseif value=null then
		formatDecimal = ""
	ElseIf isnumeric(value) = false then
		formatDecimal = ""
	Else
		formatDecimal = formatnumber(value,digits,false,true,true)
	End If
End Function

Function formatDate(value)
	If Isnull(Value) Then
		formatDate = ""
	ElseIf ISDate(value) = false Then
		formatDate = ""
	Else
		formatDate = FormatDateTime(CDate(value), vbShortDate)
	End If
End Function

Function Checked(vVariable, vValue)
	If vVariable = vValue Then
		Checked = "checked=""checked"""
	Else
		Checked = ""
	End If
End Function

Function Selected(vVariable, vValue)
	If vVariable = vValue Then
		Selected = "selected=""selected"""
	Else
		Selected = ""
	End If
End Function
%>