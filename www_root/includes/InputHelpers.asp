<%
function TextFieldColor(Name, Value, Size, MaxLength, Editable, Color, onChange)
	TextFieldColor = "<input type=""text"" name=""" & name & """ id=""" & name & """ value=""" 
	If Len(Value)>0 Then
		TextFieldColor = TextFieldColor & Replace(value, """", "&quot;") 
	End If
	TextFieldColor = TextFieldColor & """ size=" & size & " maxLength=" & MaxLength
	If Len(Color)>0 Then
		If Editable=True Then
			TextFieldColor = TextFieldColor & " style=""text-align:left; background-color:" & Color & ";"" "
		Else
			TextFieldColor = TextFieldColor & " style=""text-align:left;border-style:none; background-color:" & Color & ";"" READONLY"
		End If
	Else
		If Editable=True Then
			TextFieldColor = TextFieldColor & " style=""text-align:left;"" "
		Else
			TextFieldColor = TextFieldColor & " style=""text-align:left;border-style:none"" READONLY tabindex=""-1"" "
		End If
	End If
	If Len(onChange) > 0 Then
		TextFieldColor = TextFieldColor & " onchange=""" & onChange & """"
	End If
	TextFieldColor = TextFieldColor & ">"
End Function

function TextField(Name, Value, Size, MaxLength, Editable, onChange)
	TextField = TextFieldColor(Name, Value, Size, MaxLength, Editable, "", onChange)
End Function

function TextFieldDblClick(Name, Value, Size, MaxLength, Editable, onChange, onDblClick)
	TextFieldDblClick = "<input type=""text"" name=""" & name & """ id=""" & name & """ value=""" 
	If Len(Value)>0 Then
		TextFieldDblClick = TextFieldDblClick & Replace(value, """", "&quot;") 
	End If
	TextFieldDblClick = TextFieldDblClick & """ size=" & size & " maxLength=" & MaxLength
	If Editable=True Then
		TextFieldDblClick = TextFieldDblClick & " style=""text-align:left;"" "
	Else
		TextFieldDblClick = TextFieldDblClick & " style=""text-align:left;border-style:none"" READONLY tabindex=""-1"" "
	End If
	If Len(onChange) > 0 Then
		TextFieldDblClick = TextFieldDblClick & " onchange=""" & onChange & """"
	End If
	If Len(onDblClick) > 0 Then
		TextFieldDblClick = TextFieldDblClick & " ondblclick=""" & onDblClick & """"
	End If
	TextFieldDblClick = TextFieldDblClick & ">"

End Function

function IntegerFieldColor(Name, Value, Size, MaxLength, Editable, Color, onChange)
	IntegerFieldColor = "<input type=""text"" name=""" & name & """ id=""" & name & """ value=""" & value & """ size=" & size & " maxLength=" & MaxLength
	If Len(Color) > 0 Then
		If Editable=True Then
			IntegerFieldColor = IntegerFieldColor & " style=""text-align:right; background-color:" & Color & """ "
		Else
			IntegerFieldColor = IntegerFieldColor & " style='text-align:right;border-style:none background-color:" & Color & "' READONLY tabindex=""-1"" "
		End If
	Else
		If Editable=True Then
			IntegerFieldColor = IntegerFieldColor & " style=""text-align:right;"" "
		Else
			IntegerFieldColor = IntegerFieldColor & " style='text-align:right;border-style:none' READONLY tabindex=""-1"" "
		End If
	End If
	If Len(onChange) > 0 Then
		IntegerFieldColor = IntegerFieldColor & " onchange=""" & onChange & """"
	Else
		IntegerFieldColor = IntegerFieldColor & " onchange=""return checkInteger(this);"""
	End If
	IntegerFieldColor = IntegerFieldColor & ">"
End Function

function IntegerField(Name, Value, Size, MaxLength, Editable, onChange)
	IntegerField = IntegerFieldColor(Name, Value, Size, MaxLength, Editable, "", onChange)
End Function

function NumberFieldColor(Name, Value, Size, MaxLength, Editable, Color, onChange)
	NumberFieldColor = "<input type=""text"" name=""" & name & """ id=""" & name & """ value=""" & value & """ size=" & size & " maxLength=" & MaxLength
	If Len(Color) > 0 Then
		If Editable=True Then
			NumberFieldColor = NumberFieldColor & " style=""text-align:right; background-color:" & Color & """ "
		Else
			NumberFieldColor = NumberFieldColor & " style='text-align:right;border-style:none background-color:" & Color & "' READONLY tabindex=""-1"" "
		End If
	Else
		If Editable=True Then
			NumberFieldColor = NumberFieldColor & " style=""text-align:right;"" "
		Else
			NumberFieldColor = NumberFieldColor & " style='text-align:right;border-style:none' READONLY tabindex=""-1"" "
		End If
	End If
	If Len(onChange) > 0 Then
		NumberFieldColor = NumberFieldColor & " onchange=""" & onChange & """"
	Else
		NumberFieldColor = NumberFieldColor & " onchange=""return checkDecimal(this);"""
	End If
	NumberFieldColor = NumberFieldColor & ">"
End Function

function NumberField(Name, Value, Size, MaxLength, Editable, onChange)
	NumberField = NumberFieldColor(Name, Value, Size, MaxLength, Editable, "", onChange)
End Function

function CurrencyFieldColor(Name, Value, Size, MaxLength, Editable, Color, onChange)
	If IsNull(Value)=True Or Value="" Then
		CurrencyFieldColor = "<input type=""text"" name=""" & name & """ id=""" & name & """ value="""" size=" & size & " maxLength=""" & MaxLength & """"
	Else
		CurrencyFieldColor = "<input type=""text"" name=""" & name & """ id=""" & name & """ value=""" & formatcurrency(value,2,True,False,True) & """ size=""" & size & """ maxLength=""" & MaxLength & """"
	End If
	If Len(Color)>0 Then
		If Editable=True Then
			CurrencyFieldColor = CurrencyFieldColor & " style=""text-align:right; background-color:" & Color & ";"" "
		Else
			CurrencyFieldColor = CurrencyFieldColor & " style=""text-align:right;border-style:none; background-color:" & Color & ";"" readonly=""readonly"" tabindex=""-1"" "
		End If
	Else
		If Editable=True Then
			CurrencyFieldColor = CurrencyFieldColor & " style=""text-align:right;"" "
		Else
			CurrencyFieldColor = CurrencyFieldColor & " style=""text-align:right;border-style:none"" readonly=""readonly"" tabindex=""-1"" "
		End If
	End If
	If Len(onChange) > 0 Then
		CurrencyFieldColor = CurrencyFieldColor & " onchange=""" & onChange & """"
	Else
		CurrencyFieldColor = CurrencyFieldColor & " onchange=""return checkCurrency(this);"""
	End If
	CurrencyFieldColor = CurrencyFieldColor & " />"
End Function

function CurrencyField(Name, Value, Size, MaxLength, Editable, onChange)
	CurrencyField = CurrencyFieldColor(Name, Value, Size, MaxLength, Editable, "", onChange)
End Function

function CurrencyFieldRoundColor(Name, Value, Size, MaxLength, vRound, Editable, Color, onChange)
	If IsNull(Value)=True Or Value="" Then
		CurrencyFieldRoundColor = "<input type=""text"" name=""" & name & """ id=""" & name & """ value="""" size=" & size & " maxLength=""" & MaxLength & """"
	ElseIf vRound = True Then
		CurrencyFieldRoundColor = "<input type=""text"" name=""" & name & """ id=""" & name & """ value=""" & formatCurrency(value,0,True,False,True) & """ size=""" & size & """ maxLength=""" & MaxLength & """"
	Else
		CurrencyFieldRoundColor = "<input type=""text"" name=""" & name & """ id=""" & name & """ value=""" & formatCurrency(value,2,True,False,True) & """ size=""" & size & """ maxLength=""" & MaxLength & """"
	End If
	If Len(Color)>0 Then
		If Editable=True Then
			CurrencyFieldRoundColor = CurrencyFieldRoundColor & " style=""text-align:right; background-color:" & Color & ";"" "
		Else
			CurrencyFieldRoundColor = CurrencyFieldRoundColor & " style=""text-align:right;border-style:none; background-color:" & Color & ";"" readonly=""readonly"" tabindex=""-1"" "
		End If
	Else
		If Editable=True Then
			CurrencyFieldRoundColor = CurrencyFieldRoundColor & " style=""text-align:right;"" "
		Else
			CurrencyFieldRoundColor = CurrencyFieldRoundColor & " style=""text-align:right;border-style:none"" readonly=""readonly"" tabindex=""-1"" "
		End If
	End If
	If Len(onChange) > 0 Then
		CurrencyFieldRoundColor = CurrencyFieldRoundColor & " onchange=""" & onChange & """"
	Else
		CurrencyFieldRoundColor = CurrencyFieldRoundColor & " onchange=""return checkCurrencyRound(this, " & LCase(CStr(vRound)) & ");"""
	End If
	CurrencyFieldRoundColor = CurrencyFieldRoundColor & " />"
End Function

function CurrencyFieldRound(Name, Value, Size, MaxLength, vRound, Editable, onChange)
	CurrencyFieldRound = CurrencyFieldRoundColor(Name, Value, Size, MaxLength, vRound, Editable, "", onChange)
End Function

function SelectOption(Value, Text, SelectedValue)
	If Value = SelectedValue Then
		SelectOption = "<option value=""" & value & """ selected>" & Text & "</option>" & vbCrLF
	Else
		SelectOption = "<option value=""" & value & """>" & Text & "</option>" & vbCrLF
	End If
End Function

function CheckBoxField(Name, Checked)
	If Checked = True Then
		CheckBoxField = "<input type=""checkbox"" name=""" & Name & """ id=""" & Name & """ value=""1"" checked />"
	Else
		CheckBoxField = "<input type=""checkbox"" name=""" & Name & """ id=""" & Name & """ value=""1"" />"
	End If
End Function

function CheckBoxField2(Name, Checked, Editable)
	If Editable = False Then
		If Checked = True Then
			CheckBoxField2 = "<input type=""checkbox"" name=""" & Name & """ id=""" & Name & """ value=""1"" checked onclick=""this.checked=true;"" tabindex=""-1"" />"
		Else
			CheckBoxField2 = "<input type=""checkbox"" name=""" & Name & """ id=""" & Name & """ value=""1"" onclick=""this.checked=false;"" tabindex=""-1""/>"
		End If
	Else
		If Checked = True Then
			CheckBoxField2 = "<input type=""checkbox"" name=""" & Name & """ id=""" & Name & """ value=""1"" checked />"
		Else
			CheckBoxField2 = "<input type=""checkbox"" name=""" & Name & """ id=""" & Name & """ value=""1"" />"
		End If
	End If
End Function

function CheckBoxFieldClick(Name, Checked, onclick)
	If Checked = True Then
		CheckBoxFieldClick = "<input type=""checkbox"" name=""" & Name & """ id=""" & Name & """ value=""1"" checked onclick=""" & onclick & """ />"
	Else
		CheckBoxFieldClick = "<input type=""checkbox"" name=""" & Name & """ id=""" & Name & """ value=""1"" onclick=""" & onclick & """ />"
	End If
End Function

function RadioInputField(Name, variable, value)
	If VarType(variable)=vbBoolean And ((variable=True and value=1) Or (variable=False and value=0)) Then
		RadioInputField = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ checked />"
	ElseIf variable = value Then
		RadioInputField = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ checked />"
	Else
		RadioInputField = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ />"
	End If
End Function

function RadioInputField2(Name, variable, value, Editable)
	If Editable = True Then
		If VarType(variable)=vbBoolean And ((variable=True and value=1) Or (variable=False and value=0)) Then
			RadioInputField2 = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ checked />"
		ElseIf variable = value Then
			RadioInputField2 = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ checked />"
		Else
			RadioInputField2 = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ />"
		End If
	Else
		If VarType(variable)=vbBoolean And ((variable=True and value=1) Or (variable=False and value=0)) Then
			RadioInputField2 = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ checked onclick=""this.checked=true;"" />"
		ElseIf variable = value Then
			RadioInputField2 = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ checked onclick=""this.checked=true;"" />"
		Else
			RadioInputField2 = "<input type=""radio"" name=""" & Name & """ id=""" & Name & value & """ value=""" & value & """ onclick=""this.checked=false;"" />"
		End If
	End If
End Function

function HiddenField(Name, Value)
	HiddenField = "<input type=""hidden"" name=""" & Name & """ id=""" & Name & """ value=""" & Value & """ />"
End Function

function Text2Bool(Value)
	If Value="1" Then
		Text2Bool = True
	Else
		Text2Bool = False
	End If
End Function

function TextArea(Name, Value, rows, cols, MaxLength, Editable, onChange)
	If Editable = True Then
		TextArea = "<textarea name=""" & name & """  id=""" & name & """ class=""screenOnly"" rows=""" & rows & """ cols=""" & cols & """ maxLength=""" & MaxLength & """"
			TextArea = TextArea & " style=""text-align: left; vertical-align: top; "" "
		If Len(onChange) > 0 Then
			TextArea = TextArea & " onchange=""document.getElementById('" & Name & "Div').innerText=this.value; " & onChange & """"
		Else
			TextArea = TextArea & " onchange=""document.getElementById('" & Name & "Div').innerText=this.value; """
		End If
		TextArea = TextArea & ">" & Value & "</textarea>"
		If IsNull(Value) = False Then
			TextArea = TextArea & vbCrLf & "<div id=""" & Name & "Div"" class=""printOnly"">" & Replace(Value, vbCrLf, "<br />") & "</div>" & vbCrLf
		Else
			TextArea = TextArea & vbCrLf & "<div id=""" & Name & "Div"" class=""printOnly""><br /></div>" & vbCrLf
		End If
	ElseIf IsNull(Value) = True Then
		TextArea = "<font style=""font-style: italic; "">No text provided.</font>" & _
		HiddenField(Name, "")
	Else
		TextArea = Replace(Value, vbCrLf, "<br />") & HiddenField(Name, server.htmlencode(Value))
	End If
End Function

function TextArea2(Name, Value, rows, pixels, MaxLength, Editable, onChange)
	If Editable=True Then
		TextArea2 = "<textarea name=""" & name & """ id=""" & name & """ class=""screenOnly"" rows=""" & rows & """ maxLength=""" & MaxLength & """"
			TextArea2 = TextArea2 & " style=""text-align: left; vertical-align: top; width: " & pixels & "px; "" "
		If Len(onChange) > 0 Then
			TextArea2 = TextArea2 & " onchange=""document.getElementById('" & Name & "Div').innerText=this.value; " & onChange & """"
		Else
			TextArea2 = TextArea2 & " onchange=""document.getElementById('" & Name & "Div').innerText=this.value; """
		End If
		TextArea2 = TextArea2 & ">" & Value & "</textarea>"
		If IsNull(Value) = False Then
			TextArea2 = TextArea2 & vbCrLf & vbTab & "<div id=""" & Name & "Div"" class=""printOnly"">" & Replace(Value, vbCrLf, "<br />") & "</div>" & vbCrLf
		Else
			TextArea2 = TextArea2 & vbCrLf & vbTab & "<div id=""" & Name & "Div"" class=""printOnly""><br /></div>" & vbCrLf
		End If
	ElseIf IsNull(Value) Then
		TextArea2 = HiddenField(Name, "") & "<font style=""font-style: italic; "">No text provided.</font>"
	Else
		TextArea2 = HiddenField(Name, Server.htmlencode(Value)) & Value
	End If
End Function

function DateField(Name, Value, Editable)
	If IsDate(Value) Then
		Value = FormatDateTime(Value, vbShortDate)
	End If
	If Editable = True Then
		DateField = "<input type=text name=""" & Name & """ id=""" & Name & """ size=""10"" maxlength=""10"" value=""" & Value & _
		""" style=""text-align: right; "" onchange=""checkDate(this);"" ondblclick=""this.value='" & date & "';"" title=""Double-click for today's date."" />"
	Else
		DateField = Value & HiddenField(Name, Value)
	End If
End Function

function DateField2(Name, Value, startdate, enddate, Editable)
	If IsDate(Value) Then
		Value = FormatDateTime(Value, vbShortDate)
	End If
	If Editable = True Then
		DateField2 = "<input type=text name=""" & Name & """ id=""" & Name & """ size=""10"" maxlength=""10"" value=""" & Value & _
		""" style=""text-align: right; "" onchange=""checkDate2(this,'" & startdate & "','" & enddate & "');"" ondblclick=""this.value='" & date & "';"" title=""Double-click for today's date."" />"
	Else
		DateField2 = Value & HiddenField(Name, Value)
	End If
End Function

function DateFieldChange(Name, Value, onchange, Editable)
	If IsDate(Value) Then
		Value = FormatDateTime(Value, vbShortDate)
	End If
	If Editable = True Then
		DateFieldChange = "<input type=text name=""" & Name & """ id=""" & Name & """ size=""10"" maxlength=""10"" value=""" & Value & _
		""" style=""text-align: right; "" onchange=""checkDate(this); " & onchange & """ ondblclick=""this.value='" & date & "';"" title=""Double-click for today's date."" />"
	Else
		DateFieldChange = Value & HiddenField(Name, Value)
	End If
End Function

Function WriteInCell(vText)
	WriteInCell = vbTab & "<td style=""text-align: center; "">" & vText & "</td>" & vbCrLF
End Function
%>