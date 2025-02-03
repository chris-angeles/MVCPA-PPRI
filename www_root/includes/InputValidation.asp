<script type="text/javascript">

	function checkCurrency(field)
	{
		var fieldvalue;

		if (field.value == "") {
			field.setAttribute("class", "");
			return true;
		}
		fieldvalue = stripFormatting(field.value)
		if (isNaN(fieldvalue)) {
			field.setAttribute("class", "failed");
			alert("The value entered is not a valid number");
			field.focus();
			return false;
		}
		field.value = currency(fieldvalue);
		field.setAttribute("class", "");
		return true;
	}

	function checkCurrencyRound(field, vRound)
	{
		var fieldvalue;

		if (field.value == "") {
			field.setAttribute("class", "");
			return true;
		}
		fieldvalue = stripFormatting(field.value)
		if (isNaN(fieldvalue)) {
			field.setAttribute("class", "failed");
			alert("The value entered is not a valid number");
			field.focus();
			return false;
		}
		field.value = currencyRound(fieldvalue, vRound);
		field.setAttribute("class", "");
		return true;
	}

	function checkDecimal(field)
	{
		var fieldvalue;
		if (field.value == "") {
			field.setAttribute("class", "");
			return true;
		}

		fieldvalue = field.value.replace(/[^1234567890.-]/g, "")
		fieldvalue = parseFloat(fieldvalue)
		if (isNaN(fieldvalue)) {
			field.setAttribute("class", "failed");
			alert("The value must be a numeric value.");
			field.focus();
			return false;
		}
		field.value = fieldvalue;
		field.setAttribute("class", "");
		return true;
	}

	function checkInteger(field)
	{
		var fieldvalue;
		if (field.value == "") {
			field.setAttribute("class", "");
			return true;
		}
		fieldvalue = field.value.replace(/[^1234567890-]/g, "");
		fieldvalue = parseInt(fieldvalue)
		if (isNaN(fieldvalue)) {
			field.setAttribute("class", "failed");
			alert("The value must be an integer value.");
			field.focus();
			return false;
		}
		field.value = fieldvalue;
		field.setAttribute("class", "");
		return true;
	}

	function checkIntegerComma(field)
	{
		var fieldvalue;
		if (field.value == "") {
			field.setAttribute("class", "");
			return true;
		}
		fieldvalue = field.value.replace(/[^1234567890-]/g, "");
		fieldvalue = parseInt(fieldvalue)
		if (isNaN(fieldvalue)) {
			field.setAttribute("class", "failed");
			alert("The value must be an integer value.");
			field.focus();
			return false;
		}
		fieldvalue = fieldvalue.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,')
		field.value = fieldvalue;
		field.setAttribute("class", "");
		return true;
	}

	function getNumericValue(fieldvalue)
	{
		if (fieldvalue == null) {
			return 0;
		}
		var stringvalue;
		stringvalue = stripFormatting(fieldvalue);
		if (stringvalue == "") {
			return 0;
		}
		if (isNaN(stringvalue)) {
			return 0;
		}
		return parseFloat(stringvalue);
	}

	function stripFormatting(fieldvalue)
	{
		var num;
		num = fieldvalue
		if (fieldvalue != "") {
			if (num.toString().charAt(0) == "(" && num.toString().charAt(num.length - 1) == ")") {
				num = "-" + num.substring(1, num.length - 1)
			}
			num = num.toString().replace(/[^1234567890.-]/g, "")
		}
		if (num.toString().charAt(1) == '(') {
			return parseFloat('-' + num)
		}
		else {
			return parseFloat(num);
		}
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

	function currencyRound(num, vRound)
	{
		var prefix = "$";
		var suffix = "";
		var result = "";
		var temp;
		if (num < 0) {
			prefix = "($";
			suffix = ")";
			num = -num;
		}
		if (vRound == true) {
			temp = Math.round(num); // round to nearest dollar!
			temp = temp.toString()
			if (temp.length > 9) {
				return prefix + temp.substring(0, temp.length - 9) + "," + temp.substring(temp.length - 9, temp.length - 6) + "," + temp.substring(temp.length - 6, temp.length - 3) + "," + temp.substring(temp.length - 3, temp.length) + suffix;
			}
			if (temp.length > 6) {
				return prefix + temp.substring(0, temp.length - 6) + "," + temp.substring(temp.length - 6, temp.length - 3) + "," + temp.substring(temp.length - 3, temp.length) + suffix;
			}
			if (temp.length > 3) {
				return prefix + temp.substring(0, temp.length - 3) + "," + temp.substring(temp.length - 3, temp.length) + suffix;
			}
			return prefix + temp.substring(0, temp.length) + suffix;
		}
		else {
			temp = Math.round(num * 100.0); // convert to pennies!
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
	}

	function checkDate(fld)
	{
		var testMo, testDay, testYr, inpMo, inpDay, inpYr, msg;
		var inp = fld.value;
		if (inp == "") return true;
		var testDate = new Date(inp);
		// extract pieces from date object
		testMo = testDate.getMonth() + 1;
		testDay = testDate.getDate();
		testYr = testDate.getFullYear();
		// extract components of input data
		inpMo = parseInt(inp.substring(0, inp.indexOf("/")), 10);
		inpDay = parseInt(inp.substring((inp.indexOf("/") + 1), inp.lastIndexOf("/")), 10);
		inpYr = parseInt(inp.substring((inp.lastIndexOf("/") + 1), inp.length), 10);
		// make sure parseInt() succeeded on input components
		if (isNaN(inpMo) || isNaN(inpDay) || isNaN(inpYr)) {
			msg = "Invalid Date. Enter using form mm/dd/yyyy";
		}
		// make sure that year is resonable, not just valid!
		if (parseInt(inpYr) < 1980 || parseInt(inpYr) > 2050) {
			msg = "Check the year. The value does not seem reasonable.";
		}
		// make sure conversion to date object succeeded
		if (isNaN(testMo) || isNaN(testDay) || isNaN(testYr)) {
			msg = "Couldn't convert your entry to a valid date. Enter using form mm/dd/yyyy. Try again.";
		}
		// make sure values match
		if (testMo != inpMo || testDay != inpDay || testYr != inpYr) {
			msg = "Check the range of your date value. Enter using form mm/dd/yyyy";
		}
		if (msg) {
			// there's a message, so something failed
			alert(msg);
			fld.focus();
			return false;
		}
		else {
			fld.value = "" + inpMo + "/" + inpDay + "/" + inpYr
			return true;
		}
	}

	function checkDate2(fld, d1, d2)
	{
		if (fld.value == "")
			return true;
		if (checkDate(fld) == false)
			return false;
		var checkdate = new Date(fld.value);
		var startdate = new Date(d1);
		var enddate = new Date(d2);
		if (checkdate >= startdate && checkdate <= enddate) {
			return true;
		}
		alert(fld.value + " not in the valid date range between " + d1 + " and " + d2);
		fld.focus();
		return false;
	}

	/// Replaces commonly-used Windows 1252 encoded chars that do not exist in ASCII or ISO-8859-1 with ISO-8859-1 cognates.
	function replaceWordChars(text)
	{
		var s = text;
		// smart single quotes and apostrophe
		s = s.replace(/[\u2018\u2019\u201A]/g, "\'");
		// smart double quotes
		s = s.replace(/[\u201C\u201D\u201E]/g, "\"");
		// ellipsis
		s = s.replace(/\u2026/g, "...");
		// dashes
		s = s.replace(/[\u2013\u2014]/g, "-");
		// circumflex
		s = s.replace(/\u02C6/g, "^");
		// open angle bracket
		s = s.replace(/\u2039/g, "<");
		// close angle bracket
		s = s.replace(/\u203A/g, ">");
		// spaces
		s = s.replace(/[\u02DC\u00A0]/g, " ");

		return s;
	}
</script>